"""
WFM Solver — Hybrid (MILP-first + Heuristic Slices fallback)

Arquitetura:
  hybrid_solve()
    1. MILP (HiGHS) com time_limit + mip_rel_gap → resultado ótimo em <1s
    2. Fallback → heurístico "fatias" (alocação distribuída, sem picos absurdos)
    3. Post-trim usando SLA de 30min (consistente com engine.compute_day_indicators)

Correções vs versão anterior:
  - Post-trim usa cálculo de SLA em intervalos de 30min (não 10min)
    → elimina discrepância engine 97% vs post-trim 90.5%
  - Heurístico "fatias": seleciona slots por score, aloca 1 agente por vez
    → nunca gera alocação de 64 agentes num único horário
  - MILP agora usa mip_rel_gap=0.02 (aceita solução 2% acima do ótimo)
"""
import math
import numpy as np
from typing import List, Tuple, Dict, Optional
from scipy.optimize import milp, LinearConstraint, Bounds
import scipy.sparse as sp
from .models import SHIFTS, INTERVALS_PER_DAY, INTERVALS_30MIN


# ── Utilitários ───────────────────────────────────────────────────────────

def pl_efetivo(shift_type: str, shrinkage: float) -> float:
    return SHIFTS[shift_type]["pl_base"] * (1.0 - shrinkage)


def valid_slots(shift_type: str, first_slot: int = 0, last_slot_close: int = 144,
                wrap: bool = False,
                entry_allowed_mask: Optional[np.ndarray] = None) -> List[int]:
    """
    Retorna slots de entrada válidos para um turno.
    Em modo wrap-around (operação 24h), qualquer slot é válido pois o turno
    pode atravessar meia-noite.

    `entry_allowed_mask` (opcional): array booleano (144,) onde True = slot
    permitido como início de turno. None = sem restrição adicional.
    """
    s = SHIFTS[shift_type]
    if wrap:
        # Operação 24h: turnos podem começar em qualquer slot e atravessar meia-noite.
        slots = list(range(0, INTERVALS_PER_DAY))
    else:
        abs_last   = s["last_entry_slot"]
        hours_last = (last_slot_close - s["n_intervals"]) if last_slot_close < 144 else abs_last
        last  = min(abs_last, hours_last)
        first = max(s["first_entry_slot"], first_slot)
        slots = list(range(first, last + 1)) if first <= last else []

    # Aplica máscara de janelas de entrada (se configurada)
    if entry_allowed_mask is not None:
        slots = [x for x in slots if entry_allowed_mask[x]]
    return slots


def coverage_mask(shift_type: str, start_slot: int, wrap: bool = False) -> np.ndarray:
    """
    Máscara booleana dos slots cobertos por um turno iniciando em start_slot.
    Com wrap=True, o turno atravessa a meia-noite (operação 24h).
    """
    n = SHIFTS[shift_type]["n_intervals"]
    mask = np.zeros(INTERVALS_PER_DAY, dtype=bool)
    end = start_slot + n
    if end <= INTERVALS_PER_DAY or not wrap:
        mask[start_slot:min(end, INTERVALS_PER_DAY)] = True
    else:
        # Wrap-around: parte final do dia + parte inicial do dia seguinte
        mask[start_slot:] = True
        mask[:end - INTERVALS_PER_DAY] = True
    return mask


def coverage_from_schedule(
    schedule: List[Tuple[int, int]], shift_type: str, pl: float, wrap: bool = False
) -> np.ndarray:
    cov = np.zeros(INTERVALS_PER_DAY)
    for start, n in schedule:
        cov[coverage_mask(shift_type, start, wrap)] += n * pl
    return cov


def schedule_to_total_agents(schedule: List[Tuple[int, int]]) -> int:
    return sum(n for _, n in schedule)


# ── SLA de 30min (consistente com engine.compute_day_indicators) ──────────

def weighted_sla_30min(
    schedule: List[Tuple[int, int]],
    shift_type: str,
    pl: float,
    vol_curve: np.ndarray,
    tmo_curve: np.ndarray,
    t_target: float,
    sla_target: float,
    first_slot: int = 0,
    last_slot: int = 144,
    wrap: bool = False,
) -> float:
    """
    SLA ponderado em intervalos de 30min — consistente com engine.compute_day_indicators.
    Só considera slots dentro da janela operacional [first_slot, last_slot).
    Excluir slots fora da janela evita drag de SLA=0% em slots sem cobertura intencional.
    """
    from .erlang import calc_sla, calc_traffic
    cov = coverage_from_schedule(schedule, shift_type, pl, wrap)
    num = den = 0.0
    for slot in range(INTERVALS_30MIN):
        i0       = slot * 3
        slot_end = i0 + 3
        # Pula slots completamente fora da janela operacional
        if slot_end <= first_slot or i0 >= last_slot:
            continue
        vol30  = float(vol_curve[i0:i0+3].sum())
        tmo30  = float((tmo_curve[i0:i0+3] * vol_curve[i0:i0+3]).sum() / (vol30 + 1e-9))
        cov30  = float(cov[i0:i0+3].mean())
        u_avg  = sum(calc_traffic(vol_curve[i0+k], tmo_curve[i0+k]) for k in range(3)) / 3
        m      = max(1, math.ceil(cov30)) if cov30 > 0 else 0
        s30    = calc_sla(m, u_avg, t_target, tmo30) if m > 0 else (1.0 if vol30 == 0 else 0.0)
        num   += vol30 * s30
        den   += vol30
    return num / den if den > 0 else 0.0


def physical_coverage_ok(
    schedule: List[Tuple[int, int]],
    shift_type: str,
    hc_liq: np.ndarray,
    min_physical: int,
    wrap: bool = False,
) -> bool:
    """Verifica cobertura física mínima (agentes brutos, sem fator pl)."""
    if min_physical <= 0:
        return True
    phys = np.zeros(INTERVALS_PER_DAY)
    for s, n in schedule:
        phys[coverage_mask(shift_type, s, wrap)] += n
    return bool(np.all(phys[hc_liq > 0] >= min_physical))


# ── Post-trim (usa SLA de 30min) ──────────────────────────────────────────

def _no_zero_gap(
    trial: List[Tuple[int, int]],
    shift_type: str,
    pl: float,
    hc_liq: np.ndarray,
    wrap: bool = False,
) -> bool:
    """
    Garante que nenhum intervalo com demanda vai para zero cobertura.
    Proteção mínima — o controle fino fica por conta de min_interval_sla_ok().
    """
    cov    = coverage_from_schedule(trial, shift_type, pl, wrap)
    active = hc_liq > 0.01
    return bool(np.all((cov >= 0.01) | ~active))


def _min_interval_sla_ok(
    trial: List[Tuple[int, int]],
    shift_type: str,
    pl: float,
    vol_curve: np.ndarray,
    tmo_curve: np.ndarray,
    t_target: float,
    sla_target: float,
    first_slot: int = 0,
    last_slot: int = 144,
    wrap: bool = False,
) -> bool:
    """
    Verifica se NENHUM intervalo de 30min na janela operacional fica abaixo do SLA target.
    Impede que o post-trim produza buracos em intervalos de alta demanda que são
    mascarados pela SLA ponderada global.
    """
    from .erlang import calc_sla, calc_traffic
    cov = coverage_from_schedule(trial, shift_type, pl, wrap)
    for slot in range(INTERVALS_30MIN):
        i0       = slot * 3
        slot_end = i0 + 3
        if slot_end <= first_slot or i0 >= last_slot:
            continue
        vol30 = float(vol_curve[i0:i0+3].sum())
        if vol30 < 0.01:
            continue
        tmo30  = float((tmo_curve[i0:i0+3] * vol_curve[i0:i0+3]).sum() / (vol30 + 1e-9))
        cov30  = float(cov[i0:i0+3].mean())
        u_avg  = sum(calc_traffic(vol_curve[i0+k], tmo_curve[i0+k]) for k in range(3)) / 3
        m      = max(1, math.ceil(cov30)) if cov30 > 0 else 0
        s30    = calc_sla(m, u_avg, t_target, tmo30) if m > 0 else 0.0
        # Tolerância por intervalo: 5pp abaixo do target.
        # Protege contra buracos graves (ex: 57%) mas permite variações
        # naturais de ±2-3pp causadas por rounding de ceil(cov30).
        if s30 < sla_target - 0.05:
            return False
    return True


def _post_trim(
    schedule: List[Tuple[int, int]],
    shift_type: str,
    pl: float,
    hc_liq: np.ndarray,
    vol_curve: np.ndarray,
    tmo_curve: np.ndarray,
    t_target: float,
    sla_target: float,
    min_physical: int = 0,
    first_slot: int = 0,
    last_slot: int = 144,
    wrap: bool = False,
) -> List[Tuple[int, int]]:
    """
    Remove 1 agente por vez enquanto:
      - SLA_30min (só janela operacional) >= target + trim_margin
      - nenhum intervalo com demanda vai a zero cobertura
      - cobertura física >= min_physical (se configurado)
    """
    # Threshold: com Erlang C correto, SLA já é próxima do real.
    # Margem pequena e fixa evita over-trim.
    trim_margin = 0.005

    for _ in range(500):
        sla_now = weighted_sla_30min(schedule, shift_type, pl,
                                     vol_curve, tmo_curve, t_target, sla_target,
                                     first_slot, last_slot, wrap)
        if sla_now < sla_target + trim_margin:
            break

        best_idx, best_sla = None, -1.0
        for idx, (s, n) in enumerate(schedule):
            if n <= 0:
                continue
            trial = list(schedule)
            trial[idx] = (s, n - 1)
            trial = [(s2, n2) for s2, n2 in trial if n2 > 0]

            # Verificação tripla: SLA, no-zero-gap, cobertura mínima física
            if not _no_zero_gap(trial, shift_type, pl, hc_liq, wrap):
                continue
            if not physical_coverage_ok(trial, shift_type, hc_liq, min_physical, wrap):
                continue
            # Critério triplo: SLA média OK + nenhum intervalo individual abaixo + física OK
            if not _min_interval_sla_ok(trial, shift_type, pl,
                                        vol_curve, tmo_curve, t_target, sla_target,
                                        first_slot, last_slot, wrap):
                continue
            ts = weighted_sla_30min(trial, shift_type, pl,
                                    vol_curve, tmo_curve, t_target, sla_target,
                                    first_slot, last_slot, wrap)
            if ts >= sla_target and ts > best_sla:
                best_sla, best_idx = ts, idx

        if best_idx is None:
            break

        s, n = schedule[best_idx]
        if n == 1:
            schedule.pop(best_idx)
        else:
            schedule[best_idx] = (s, n - 1)

    return [(s, n) for s, n in schedule if n > 0]


# ── Heurístico "fatias" ───────────────────────────────────────────────────

def _heuristic_slices(
    hc_liq: np.ndarray,
    shift_type: str,
    pl: float,
    slots: List[int],
    max_horarios: int,
    vol_curve: np.ndarray,
    tmo_curve: np.ndarray,
    sla_target: float,
    t_target: float,
    min_physical: int,
    wrap: bool = False,
) -> List[Tuple[int, int]]:
    """
    Heurístico "fatias" v5 — SET-COVER + alocação incremental:

    0. SET-COVER: garante que todo slot com demanda seja coberto por
       pelo menos um slot de entrada selecionado. Inclui primeiro slots
       "obrigatórios" (únicos que cobrem algum i) e depois greedy-cover.
       Isso elimina buracos em operações 24h que o heurístico antigo gerava.
    1. Completa com os melhores slots por score (volume × cobertura),
       até atingir `max_horarios` (ou mais, se o cover exigir).
    2. Aloca 1 agente por slot selecionado, incrementa até cobrir o gap.
    3. Post-trim para remover folga mantendo SLA.

    Quando o set-cover força mais slots que max_horarios, prioriza cobertura
    sobre o limite — preferível abrir um horário a mais do que deixar buraco
    com SLA 0%.
    """
    if not slots:
        return []

    # ── 0. SET-COVER: seleciona slots obrigatórios para cobertura completa ────
    required = _greedy_set_cover(hc_liq, shift_type, slots, vol_curve, wrap)
    # Se required sozinho já ultrapassa max_horarios, usa required (mais slots > buraco)
    n_cover = len(required)

    # ── 1. Score dos slots para completar até max_horarios ───────────────────
    slot_scores: Dict[int, float] = {}
    for s in slots:
        if s in required:
            continue
        mask = coverage_mask(shift_type, s, wrap)
        need = hc_liq[mask]
        if need.max() <= 0.01:
            continue
        gain  = np.minimum(need, pl)
        vol_w = vol_curve[mask].sum() + 1e-9
        slot_scores[s] = (gain * vol_curve[mask]).sum() / vol_w * need.sum()

    # Completa até max_horarios (ou mantém só required se cover já extrapolou)
    extra_budget = max(0, max_horarios - n_cover)
    extras       = sorted(slot_scores, key=slot_scores.get, reverse=True)[:extra_budget]
    selected     = list(required) + extras

    if not selected:
        return []

    # ── 2. Aloca 1 agente inicial por slot ───────────────────────────
    schedule: Dict[int, int] = {s: 1 for s in selected}

    # ── 3. Incrementa 1 agente por vez até cobrir o gap ──────────────
    max_iters = int(hc_liq.sum() / max(pl, 1e-9)) + len(selected) * 2 + 20
    for _ in range(max_iters):
        cov_arr = np.zeros(INTERVALS_PER_DAY)
        for s, n in schedule.items():
            cov_arr[coverage_mask(shift_type, s, wrap)] += n * pl
        gap = np.maximum(hc_liq - cov_arr, 0.0)
        if gap.max() <= 0.01:
            break

        best_s, best_score = None, -1.0
        for s in selected:
            mask  = coverage_mask(shift_type, s, wrap)
            rg    = gap[mask]
            if rg.max() <= 0.01:
                continue
            score = (np.minimum(rg, pl) * vol_curve[mask]).sum()
            if score > best_score:
                best_score, best_s = score, s

        if best_s is None:
            break
        schedule[best_s] = schedule.get(best_s, 0) + 1

    sched = [(s, n) for s, n in schedule.items() if n > 0]

    # ── 4. Post-trim ─────────────────────────────────────────────────
    sched = _post_trim(sched, shift_type, pl, hc_liq,
                       vol_curve, tmo_curve, t_target, sla_target,
                       min_physical=min_physical,
                       first_slot=0, last_slot=144, wrap=wrap)
    return sched


def _greedy_set_cover(
    hc_liq: np.ndarray,
    shift_type: str,
    slots: List[int],
    vol_curve: np.ndarray,
    wrap: bool = False,
) -> List[int]:
    """
    Seleciona um conjunto mínimo de slots de entrada que cobre TODOS
    os i com hc_liq[i] > 0. Greedy: em cada passo, escolhe o slot que cobre
    mais demanda ainda descoberta, ponderada por volume.

    Inclui primeiro os slots "forçados" — aqueles que são o único cobridor
    de algum i com demanda.
    """
    demand_mask = hc_liq > 0.01
    if not demand_mask.any():
        return []

    # Pre-computa máscara de cobertura de cada slot
    cov_by_slot: Dict[int, np.ndarray] = {
        s: coverage_mask(shift_type, s, wrap) for s in slots
    }

    # Identifica slots "forçados": únicos que cobrem algum i com demanda
    forced: set = set()
    for i in np.where(demand_mask)[0]:
        cobridores = [s for s in slots if cov_by_slot[s][i]]
        if len(cobridores) == 1:
            forced.add(cobridores[0])
        elif len(cobridores) == 0:
            # Ninguém pode cobrir este i — não há slot de entrada que o inclua.
            # Isso não é problema do set-cover, é uma restrição inviável.
            pass

    selected: List[int] = list(forced)
    covered = np.zeros(INTERVALS_PER_DAY, dtype=bool)
    for s in selected:
        covered |= cov_by_slot[s]

    # Greedy cover enquanto houver demanda descoberta
    uncovered = demand_mask & ~covered
    safety = len(slots) + 5
    while uncovered.any() and safety > 0:
        safety -= 1
        best_s, best_score = None, -1.0
        for s in slots:
            if s in selected:
                continue
            mask = cov_by_slot[s]
            newly = uncovered & mask
            if not newly.any():
                continue
            # Score: demanda descoberta + volume (desempate)
            score = float(hc_liq[newly].sum()) + 0.01 * float(vol_curve[newly].sum())
            if score > best_score:
                best_score, best_s = score, s
        if best_s is None:
            break
        selected.append(best_s)
        covered |= cov_by_slot[best_s]
        uncovered = demand_mask & ~covered

    return selected


# ── MILP (HiGHS via scipy) ────────────────────────────────────────────────

def _milp_solve(
    hc_liq: np.ndarray,
    shift_type: str,
    pl: float,
    slots: List[int],
    max_horarios: int,
    time_limit: float = 5.0,
    mip_rel_gap: float = 0.02,
    wrap: bool = False,
) -> Optional[List[Tuple[int, int]]]:
    """
    MILP turno único (HiGHS).
    Minimiza Σ n_s  s.a.:
      Σ_{s covers i} n_s * pl  >= hc_liq[i]   ∀ i com demanda
      Σ y_s                    <= max_horarios
      n_s - M*y_s              <= 0
      n_s inteiro >= 0,  y_s ∈ {0,1}

    mip_rel_gap=0.02: aceita solução dentro de 2% do ótimo
    → solver para quando prova qualidade, não necessariamente otimalidade.
    """
    if not slots:
        return None

    n_s    = len(slots)
    n_vars = 2 * n_s
    c      = np.zeros(n_vars);  c[:n_s] = 1.0
    integ  = np.ones(n_vars)
    lb     = np.zeros(n_vars)
    ub     = np.full(n_vars, np.inf);  ub[n_s:] = 1.0
    M      = int(hc_liq.max() / max(pl, 1e-9)) + 10

    rows, cols, vals, lbs, ubs = [], [], [], [], []
    row = 0

    for i in np.where(hc_liq > 0.01)[0]:
        for j, s in enumerate(slots):
            if coverage_mask(shift_type, s, wrap)[i]:
                rows.append(row);  cols.append(j);  vals.append(pl)
        lbs.append(float(hc_liq[i]));  ubs.append(np.inf);  row += 1

    for j in range(n_s):
        rows += [row, row];  cols += [j, n_s + j];  vals += [1.0, -float(M)]
        lbs.append(-np.inf);  ubs.append(0.0);  row += 1

    for j in range(n_s):
        rows.append(row);  cols.append(n_s + j);  vals.append(1.0)
    lbs.append(-np.inf);  ubs.append(float(max_horarios));  row += 1

    A = sp.csc_matrix((vals, (rows, cols)), shape=(row, n_vars))
    res = milp(
        c=c,
        constraints=LinearConstraint(A, lbs, ubs),
        integrality=integ,
        bounds=Bounds(lb=lb, ub=ub),
        options={"time_limit": time_limit, "mip_rel_gap": mip_rel_gap, "disp": False},
    )

    if res.x is None:
        return None

    sched = [(s, int(round(res.x[j]))) for j, s in enumerate(slots)
             if int(round(res.x[j])) > 0]
    return sched if sched else None


# ── Ponto de entrada: hybrid_solve ───────────────────────────────────────

def hybrid_solve(
    hc_liq: np.ndarray,
    shift_type: str,
    pl: float,
    max_horarios: int,
    vol_curve: np.ndarray,
    sla_target: float,
    tmo_curve: np.ndarray,
    t_target: float,
    max_overstaffing: float = 0.20,
    min_physical: int = 0,
    first_slot: int = 0,
    last_slot: int = 144,
    milp_time_limit: float = 5.0,
    mip_rel_gap: float = 0.02,
    force_mode: str = "hybrid",
    wrap: bool = False,
) -> List[Tuple[int, int]]:
    """
    Solver híbrido por turno único:

      "hybrid"    → MILP(5s, gap=2%) → fallback heurístico fatias
      "milp"      → só MILP
      "heuristic" → só heurístico fatias

    Passos:
      1. Mascara demanda fora de [first_slot, last_slot)
      2. Tenta MILP com time_limit e mip_rel_gap
         - Se HiGHS retorna solução (status 0 ou 1 com x!=None) → usa
      3. Fallback para heurístico fatias se MILP falhar/timeout
      4. Post-trim com SLA_30min (consistente com engine)
    """
    slots = valid_slots(shift_type, first_slot, last_slot, wrap)
    if not slots:
        return []

    hc_win = hc_liq.copy().astype(float)
    if not wrap:
        hc_win[:first_slot] = 0.0
        hc_win[last_slot:]  = 0.0
    if hc_win.max() <= 0.01:
        return []

    # ── MILP ──────────────────────────────────────────────────────────
    if force_mode in ("hybrid", "milp"):
        result = _milp_solve(
            hc_win, shift_type, pl, slots,
            max_horarios, milp_time_limit, mip_rel_gap, wrap
        )
        if result is not None:
            result = _post_trim(result, shift_type, pl, hc_win,
                                vol_curve, tmo_curve, t_target, sla_target,
                                min_physical=min_physical,
                                first_slot=first_slot, last_slot=last_slot,
                                wrap=wrap)
            return result

    if force_mode == "milp":
        return []

    # ── Fallback: heurístico fatias ───────────────────────────────────
    return _heuristic_slices(
        hc_win, shift_type, pl, slots, max_horarios,
        vol_curve, tmo_curve, sla_target, t_target, min_physical, wrap
    )


# ── Retrocompatibilidade ──────────────────────────────────────────────────

def heuristic_solve(
    hc_liq, shift_type, pl, max_horarios, vol_curve,
    sla_target=0.0, tmo_curve=None, t_target=20.0,
    max_overstaffing=0.20, min_physical=0,
    first_slot=0, last_slot=144,
):
    """Alias retrocompatível → chama hybrid no modo heurístico."""
    return hybrid_solve(
        hc_liq, shift_type, pl, max_horarios, vol_curve,
        sla_target,
        tmo_curve if tmo_curve is not None else np.ones(INTERVALS_PER_DAY),
        t_target, max_overstaffing, min_physical, first_slot, last_slot,
        force_mode="heuristic",
    )


def exact_solve(
    hc_liq, available, pl_map, max_horarios, max_hc_total, time_limit=120.0
):
    """MILP multi-turno — mantido para solver_mode='exact' explícito."""
    x_vars = [(t, s) for t, slots in available.items() for s in slots]
    y_vars = list(x_vars)
    n_x = len(x_vars); n_y = len(y_vars); n_vars = n_x + n_y
    x_idx = {v: i for i, v in enumerate(x_vars)}
    y_idx = {v: n_x + i for i, v in enumerate(y_vars)}
    M = int(hc_liq.max() * 5) + 20

    c = np.zeros(n_vars); c[:n_x] = 1.0
    integrality = np.ones(n_vars)
    lb = np.zeros(n_vars); ub = np.full(n_vars, np.inf); ub[n_x:] = 1.0

    rows, cols, vals, lbs, ubs = [], [], [], [], []
    row = 0
    for i in np.where(hc_liq > 0.01)[0]:
        for t, slots in available.items():
            pl = pl_map[t]
            for s in slots:
                if coverage_mask(t, s)[i]:
                    rows.append(row); cols.append(x_idx[(t, s)]); vals.append(pl)
        lbs.append(float(hc_liq[i])); ubs.append(np.inf); row += 1

    for v in x_vars:
        rows += [row, row]; cols += [x_idx[v], y_idx[v]]; vals += [1.0, -float(M)]
        lbs.append(-np.inf); ubs.append(0.0); row += 1

    for t, slots in available.items():
        for s in slots:
            rows.append(row); cols.append(y_idx[(t, s)]); vals.append(1.0)
        lbs.append(-np.inf); ubs.append(float(max_horarios)); row += 1

    for i in range(n_x):
        rows.append(row); cols.append(i); vals.append(1.0)
    lbs.append(-np.inf); ubs.append(float(max_hc_total)); row += 1

    A = sp.csc_matrix((vals, (rows, cols)), shape=(row, n_vars))
    res = milp(c=c, constraints=LinearConstraint(A, lbs, ubs),
               integrality=integrality, bounds=Bounds(lb=lb, ub=ub),
               options={"time_limit": time_limit, "disp": False})

    result = {t: [] for t in available}
    if res.x is not None:
        for v in x_vars:
            t, s = v
            n = int(round(res.x[x_idx[v]]))
            if n > 0:
                result[t].append((s, n))
    return result


# ── Solver unificado de pools (arquitetura 6×1 correta) ──────────────────────

def unified_pool_solve(
    hc_util: np.ndarray,
    hc_sab:  np.ndarray,
    hc_dom:  np.ndarray,
    pl_620: float,
    pl_812: float,
    max_horarios: int,
    vol_util: np.ndarray,
    tmo_util: np.ndarray,
    vol_sab:  np.ndarray,
    tmo_sab:  np.ndarray,
    vol_dom:  np.ndarray,
    tmo_dom:  np.ndarray,
    sla_target: float,
    t_target: float,
    first_slot: int = 0,
    last_slot:  int = 144,
    milp_time_limit: float = 10.0,
    mip_rel_gap: float = 0.02,
    min_physical: int = 0,
    wrap: bool = False,
    entry_allowed_mask: Optional[np.ndarray] = None,
    allow_shift_620: bool = True,
    allow_shift_812: bool = True,
) -> Dict[str, List[Tuple[int, int]]]:
    """
    Resolve os 3 pools simultaneamente num único MILP respeitando as
    regras do 6×1:

      x_sab[s]: agentes 6:20 que TRABALHAM SÁBADO (folgam Domingo)
                → cobrem dia útil E sábado
      x_dom[s]: agentes 6:20 que TRABALHAM DOMINGO (folgam Sábado)
                → cobrem dia útil E domingo
      x_812[s]: agentes 8:12 → cobrem só dia útil (CLT 5×2)

    Restrições:
      Dia útil: Σ(x_sab + x_dom)*pl_620 + Σ x_812*pl_812 ≥ hc_util[i]
      Sábado:   Σ x_sab * pl_620 ≥ hc_sab[i]
      Domingo:  Σ x_dom * pl_620 ≥ hc_dom[i]

    Objetivo: minimizar Σ x_sab + Σ x_dom + Σ x_812

    Com wrap=True (operação 24h), turnos podem atravessar meia-noite e
    qualquer slot de entrada é válido.
    """
    slots_620 = valid_slots("6:20", first_slot, last_slot, wrap, entry_allowed_mask) if allow_shift_620 else []
    slots_812 = valid_slots("8:12", first_slot, last_slot, wrap, entry_allowed_mask) if allow_shift_812 else []

    empty = {"sab": [], "dom": [], "812": []}
    if not slots_620 and not slots_812:
        return empty

    # Mascara demanda fora da janela (não aplica em wrap-around)
    hu = hc_util.copy().astype(float)
    hs = hc_sab.copy().astype(float)
    hd = hc_dom.copy().astype(float)
    if not wrap:
        hu[:first_slot] = 0; hu[last_slot:] = 0
        hs[:first_slot] = 0; hs[last_slot:] = 0
        hd[:first_slot] = 0; hd[last_slot:] = 0

    has_sab = hs.max() > 0.01
    has_dom = hd.max() > 0.01

    # ── Índices de variáveis ──────────────────────────────────────────────
    n_sab = len(slots_620)
    n_dom = len(slots_620)
    n_812 = len(slots_812)
    n_x   = n_sab + n_dom + n_812          # variáveis contínuas (int)
    n_y   = n_x                             # binárias de abertura de turno
    n_vars = n_x + n_y

    # offsets
    OFF_xsab, OFF_xdom, OFF_x812 = 0, n_sab, n_sab + n_dom
    OFF_ysab  = n_x
    OFF_ydom  = n_x + n_sab
    OFF_y812  = n_x + n_sab + n_dom

    c = np.zeros(n_vars)
    c[OFF_xsab:OFF_xsab+n_sab] = 1.0   # minimiza agentes sab
    c[OFF_xdom:OFF_xdom+n_dom] = 1.0   # minimiza agentes dom
    c[OFF_x812:OFF_x812+n_812] = 1.0   # minimiza agentes 812

    integ = np.ones(n_vars)
    lb    = np.zeros(n_vars)
    ub    = np.full(n_vars, np.inf)
    ub[OFF_ysab:] = 1.0   # binárias

    M = int(max(hu.max(), hs.max(), hd.max()) / min(pl_620, pl_812, 0.5)) + 20

    rows, cols, vals, lbs, ubs = [], [], [], [], []
    row = 0

    # Pré-calcula máscaras (acelera construção e evita recomputar com wrap)
    mask_620 = {s: coverage_mask("6:20", s, wrap) for s in slots_620}
    mask_812 = {s: coverage_mask("8:12", s, wrap) for s in slots_812}

    # ── 1. Dia Útil: (x_sab + x_dom)*pl_620 + x_812*pl_812 ≥ hu[i] ──────
    for i in np.where(hu > 0.01)[0]:
        for j, s in enumerate(slots_620):
            if mask_620[s][i]:
                rows.append(row); cols.append(OFF_xsab + j); vals.append(pl_620)
                rows.append(row); cols.append(OFF_xdom + j); vals.append(pl_620)
        for j, s in enumerate(slots_812):
            if mask_812[s][i]:
                rows.append(row); cols.append(OFF_x812 + j); vals.append(pl_812)
        lbs.append(float(hu[i])); ubs.append(np.inf); row += 1

    # ── 2. Sábado: x_sab*pl_620 ≥ hs[i] ─────────────────────────────────
    if has_sab:
        for i in np.where(hs > 0.01)[0]:
            for j, s in enumerate(slots_620):
                if mask_620[s][i]:
                    rows.append(row); cols.append(OFF_xsab + j); vals.append(pl_620)
            lbs.append(float(hs[i])); ubs.append(np.inf); row += 1
    else:
        # Sábado fechado: x_sab deve ser 0
        for j in range(n_sab):
            rows.append(row); cols.append(OFF_xsab + j); vals.append(1.0)
        lbs.append(-np.inf); ubs.append(0.0); row += 1

    # ── 3. Domingo: x_dom*pl_620 ≥ hd[i] ────────────────────────────────
    if has_dom:
        for i in np.where(hd > 0.01)[0]:
            for j, s in enumerate(slots_620):
                if mask_620[s][i]:
                    rows.append(row); cols.append(OFF_xdom + j); vals.append(pl_620)
            lbs.append(float(hd[i])); ubs.append(np.inf); row += 1
    else:
        # Domingo fechado: x_dom deve ser 0
        for j in range(n_dom):
            rows.append(row); cols.append(OFF_xdom + j); vals.append(1.0)
        lbs.append(-np.inf); ubs.append(0.0); row += 1

    # ── 3b. Piso físico: ≥ min_physical agentes LOGADOS em cada slot ───
    # Garante que todos os slots da janela operacional tenham, no mínimo,
    # `min_physical` pessoas presentes — não importa o pl_efetivo.
    # Aplica por tipo de dia (util / sáb / dom) e respeita slots fechados.
    if min_physical > 0:
        slot_range = range(INTERVALS_PER_DAY) if wrap else range(first_slot, last_slot)
        # Dia útil: x_sab + x_dom + x_812 cobrindo slot i ≥ min_physical
        for i in slot_range:
            added = False
            for j, s in enumerate(slots_620):
                if mask_620[s][i]:
                    rows.append(row); cols.append(OFF_xsab + j); vals.append(1.0)
                    rows.append(row); cols.append(OFF_xdom + j); vals.append(1.0)
                    added = True
            for j, s in enumerate(slots_812):
                if mask_812[s][i]:
                    rows.append(row); cols.append(OFF_x812 + j); vals.append(1.0)
                    added = True
            if added:
                lbs.append(float(min_physical)); ubs.append(np.inf); row += 1
        # Sábado: só x_sab ≥ min_physical
        if has_sab:
            for i in slot_range:
                added = False
                for j, s in enumerate(slots_620):
                    if mask_620[s][i]:
                        rows.append(row); cols.append(OFF_xsab + j); vals.append(1.0)
                        added = True
                if added:
                    lbs.append(float(min_physical)); ubs.append(np.inf); row += 1
        # Domingo: só x_dom ≥ min_physical
        if has_dom:
            for i in slot_range:
                added = False
                for j, s in enumerate(slots_620):
                    if mask_620[s][i]:
                        rows.append(row); cols.append(OFF_xdom + j); vals.append(1.0)
                        added = True
                if added:
                    lbs.append(float(min_physical)); ubs.append(np.inf); row += 1

    # ── 4. Big-M: x_* ≤ M * y_* ─────────────────────────────────────────
    for j in range(n_sab):
        rows += [row, row]; cols += [OFF_xsab+j, OFF_ysab+j]; vals += [1.0, -float(M)]
        lbs.append(-np.inf); ubs.append(0.0); row += 1
    for j in range(n_dom):
        rows += [row, row]; cols += [OFF_xdom+j, OFF_ydom+j]; vals += [1.0, -float(M)]
        lbs.append(-np.inf); ubs.append(0.0); row += 1
    for j in range(n_812):
        rows += [row, row]; cols += [OFF_x812+j, OFF_y812+j]; vals += [1.0, -float(M)]
        lbs.append(-np.inf); ubs.append(0.0); row += 1

    # ── 5. Máximo de horários por pool ───────────────────────────────────
    # Em operação 24h (wrap), max_horarios pode ser insuficiente para cobrir tudo.
    # Aumenta o limite mínimo para garantir cobertura viável.
    effective_max = max_horarios
    if wrap:
        # Turno de 6:20 cobre 38 slots; ~4 turnos sem overlap cobrem 24h.
        # Com curva não-uniforme e PL, geralmente 6-8 é o mínimo prático.
        effective_max = max(max_horarios, 8)

    for off_y, n in [(OFF_ysab, n_sab), (OFF_ydom, n_dom), (OFF_y812, n_812)]:
        for j in range(n):
            rows.append(row); cols.append(off_y + j); vals.append(1.0)
        lbs.append(-np.inf); ubs.append(float(effective_max)); row += 1

    if not rows:
        return empty

    # Em wrap-mode o MILP tem muito mais variáveis (144 slots vs ~100),
    # então precisa de mais tempo para convergir.
    effective_time_limit = milp_time_limit * 3.0 if wrap else milp_time_limit

    A = sp.csc_matrix((vals, (rows, cols)), shape=(row, n_vars))
    res = milp(
        c=c,
        constraints=LinearConstraint(A, lbs, ubs),
        integrality=integ,
        bounds=Bounds(lb=lb, ub=ub),
        options={"time_limit": effective_time_limit, "mip_rel_gap": mip_rel_gap, "disp": False},
    )

    # ── Fallback heurístico se MILP não convergir OU solução inviável ────
    # Status 1 (time limit) pode retornar solução parcial que viola restrições —
    # valida explicitamente a cobertura antes de aceitar.
    def _milp_feasible(x):
        if x is None:
            return False
        def extract(off, slots):
            return [(s, int(round(x[off + j])))
                    for j, s in enumerate(slots) if int(round(x[off + j])) > 0]
        sab_t = extract(OFF_xsab, slots_620)
        dom_t = extract(OFF_xdom, slots_620)
        x812_t = extract(OFF_x812, slots_812)
        cov_u = (coverage_from_schedule(sab_t, "6:20", pl_620, wrap) +
                 coverage_from_schedule(dom_t, "6:20", pl_620, wrap) +
                 coverage_from_schedule(x812_t, "8:12", pl_812, wrap))
        if np.any((hu > 0.01) & (cov_u < hu - 0.01)):
            return False
        if has_sab:
            cov_s = coverage_from_schedule(sab_t, "6:20", pl_620, wrap)
            if np.any((hs > 0.01) & (cov_s < hs - 0.01)):
                return False
        if has_dom:
            cov_d = coverage_from_schedule(dom_t, "6:20", pl_620, wrap)
            if np.any((hd > 0.01) & (cov_d < hd - 0.01)):
                return False
        return True

    if not _milp_feasible(res.x):
        return _unified_heuristic_fallback(
            hu, hs, hd, pl_620, pl_812, slots_620, slots_812,
            effective_max, vol_util, tmo_util, vol_sab, tmo_sab,
            vol_dom, tmo_dom, sla_target, t_target, min_physical, wrap,
        )

    # Extrai schedule
    def extract(off, slots):
        return [(s, int(round(res.x[off + j])))
                for j, s in enumerate(slots) if int(round(res.x[off + j])) > 0]

    sched_sab = extract(OFF_xsab, slots_620)
    sched_dom = extract(OFF_xdom, slots_620)
    sched_812 = extract(OFF_x812, slots_812)

    # Post-trim unificado
    sched_sab, sched_dom, sched_812 = _unified_post_trim(
        sched_sab, sched_dom, sched_812,
        pl_620, pl_812,
        hu, hs, hd,
        vol_util, tmo_util, vol_sab, tmo_sab, vol_dom, tmo_dom,
        sla_target, t_target, first_slot, last_slot, min_physical, wrap,
    )

    return {"sab": sched_sab, "dom": sched_dom, "812": sched_812}


def _unified_heuristic_fallback(
    hu, hs, hd, pl_620, pl_812, slots_620, slots_812,
    max_horarios, vol_util, tmo_util, vol_sab, tmo_sab,
    vol_dom, tmo_dom, sla_target, t_target, min_physical, wrap=False,
) -> Dict[str, List[Tuple[int, int]]]:
    """
    Fallback sequencial correto:
      1. Resolve sábado  → N_sab (cobre sábado E dias úteis)
      2. Resolve domingo → N_dom (cobre domingo E dias úteis)
      3. Verifica se (N_sab + N_dom) já cobre dias úteis
      4. Se não, adiciona 8:12 para fechar o gap
    """
    # Sábado
    sched_sab = _heuristic_slices(
        hs, "6:20", pl_620, slots_620, max_horarios,
        vol_sab, tmo_sab, sla_target, t_target, min_physical, wrap,
    ) if hs.max() > 0.01 else []

    # Domingo
    sched_dom = _heuristic_slices(
        hd, "6:20", pl_620, slots_620, max_horarios,
        vol_dom, tmo_dom, sla_target, t_target, min_physical, wrap,
    ) if hd.max() > 0.01 else []

    # Cobertura útil já disponível dos dois pools 6:20
    cov_util_620 = (
        coverage_from_schedule(sched_sab, "6:20", pl_620, wrap) +
        coverage_from_schedule(sched_dom, "6:20", pl_620, wrap)
    )
    gap_util = np.maximum(hu - cov_util_620, 0.0)

    # Fecha gap com 8:12
    sched_812 = _heuristic_slices(
        gap_util, "8:12", pl_812, slots_812, max_horarios,
        vol_util, tmo_util, sla_target, t_target, min_physical, wrap,
    ) if gap_util.max() > 0.05 else []

    return {"sab": sched_sab, "dom": sched_dom, "812": sched_812}


def _unified_post_trim(
    sched_sab, sched_dom, sched_812,
    pl_620, pl_812,
    hu, hs, hd,
    vol_util, tmo_util, vol_sab, tmo_sab, vol_dom, tmo_dom,
    sla_target, t_target, first_slot, last_slot, min_physical, wrap=False,
):
    """
    Remove 1 agente por vez de qualquer pool enquanto:
      - SLA_util  ≥ target (cobertura dos três pools no dia útil)
      - SLA_sab   ≥ target (cobertura x_sab no sábado)
      - SLA_dom   ≥ target (cobertura x_dom no domingo)
      - Nenhum intervalo individual fica abaixo de target - 5pp
    """
    def wsla(sched, shift, pl, vol, tmo, fs=first_slot, ls=last_slot):
        return weighted_sla_30min(sched, shift, pl, vol, tmo, t_target, sla_target, fs, ls, wrap)

    def cov_util():
        return (coverage_from_schedule(sched_sab, "6:20", pl_620, wrap) +
                coverage_from_schedule(sched_dom, "6:20", pl_620, wrap) +
                coverage_from_schedule(sched_812, "8:12", pl_812, wrap))

    def sla_util():
        from .erlang import calc_sla, calc_traffic
        cov = cov_util()
        num = den = 0.0
        for slot in range(INTERVALS_30MIN):
            i0 = slot * 3
            if not wrap:
                if i0 + 3 <= first_slot or i0 >= last_slot: continue
            v30 = float(vol_util[i0:i0+3].sum())
            if v30 < 0.01: continue
            t30 = float((tmo_util[i0:i0+3]*vol_util[i0:i0+3]).sum()/(v30+1e-9))
            c30 = float(cov[i0:i0+3].mean())
            u   = sum(calc_traffic(vol_util[i0+k],tmo_util[i0+k]) for k in range(3))/3
            m   = max(1, math.ceil(c30)) if c30 > 0 else 0
            s   = calc_sla(m, u, t_target, t30) if m > 0 else (1.0 if v30==0 else 0.0)
            num += v30*s; den += v30
        return num/den if den > 0 else 0.0

    def iv_ok_after(candidate_sab, candidate_dom, candidate_812):
        """Verifica que nenhum intervalo individual fica abaixo de target-5pp
        em NENHUMA das três escalas (util, sab, dom)."""
        from .erlang import calc_sla, calc_traffic

        def _check(cov_arr, vol_arr, tmo_arr):
            for slot in range(INTERVALS_30MIN):
                i0 = slot*3
                if not wrap:
                    if i0+3<=first_slot or i0>=last_slot: continue
                v30=float(vol_arr[i0:i0+3].sum())
                if v30<0.01: continue
                t30=float((tmo_arr[i0:i0+3]*vol_arr[i0:i0+3]).sum()/(v30+1e-9))
                c30=float(cov_arr[i0:i0+3].mean())
                u=sum(calc_traffic(vol_arr[i0+k],tmo_arr[i0+k]) for k in range(3))/3
                m=max(1,math.ceil(c30)) if c30>0 else 0
                s=calc_sla(m,u,t_target,t30) if m>0 else 0.0
                if s < sla_target - 0.05: return False
            return True

        # Util = cobertura dos 3 pools
        cov_u = (coverage_from_schedule(candidate_sab, "6:20", pl_620, wrap) +
                 coverage_from_schedule(candidate_dom, "6:20", pl_620, wrap) +
                 coverage_from_schedule(candidate_812, "8:12", pl_812, wrap))
        if not _check(cov_u, vol_util, tmo_util): return False

        # Sábado = só candidate_sab
        if hs.max() > 0.01:
            cov_s = coverage_from_schedule(candidate_sab, "6:20", pl_620, wrap)
            if not _check(cov_s, vol_sab, tmo_sab): return False

        # Domingo = só candidate_dom
        if hd.max() > 0.01:
            cov_d = coverage_from_schedule(candidate_dom, "6:20", pl_620, wrap)
            if not _check(cov_d, vol_dom, tmo_dom): return False

        return True

    # Limita sched_812 a slots no início e fim do dia (não usa 6:20 do pool 8:12)
    for _ in range(500):
        su = sla_util()
        if su < sla_target + 0.005: break

        best_pool, best_idx, best_sla = None, None, -1.0

        # Tenta remover de cada pool
        for pool_name, sched, shift, pl_v in [
            ("sab", sched_sab, "6:20", pl_620),
            ("dom", sched_dom, "6:20", pl_620),
            ("812", sched_812, "8:12", pl_812),
        ]:
            for idx, (s, n) in enumerate(sched):
                if n <= 0: continue
                trial = list(sched); trial[idx] = (s, n-1)
                trial = [(s2,n2) for s2,n2 in trial if n2>0]

                # Escolhe os schedules candidatos
                t_sab = trial if pool_name=="sab" else sched_sab
                t_dom = trial if pool_name=="dom" else sched_dom
                t_812 = trial if pool_name=="812" else sched_812

                # Verifica util, sab, dom e por intervalo
                if not iv_ok_after(t_sab, t_dom, t_812): continue
                if pool_name=="sab" and hs.max()>0.01:
                    if wsla(t_sab,"6:20",pl_620,vol_sab,tmo_sab) < sla_target: continue
                if pool_name=="dom" and hd.max()>0.01:
                    if wsla(t_dom,"6:20",pl_620,vol_dom,tmo_dom) < sla_target: continue

                # Piso físico mínimo (HC bruto por intervalo) — 3 frentes:
                #   - util: soma dos 3 pools em slots com demanda útil
                #   - sab:  só x_sab em slots com demanda de sábado
                #   - dom:  só x_dom em slots com demanda de domingo
                if min_physical > 0:
                    def _phys(sched, shift):
                        phys = np.zeros(INTERVALS_PER_DAY)
                        for s2, n2 in sched:
                            phys[coverage_mask(shift, s2, wrap)] += n2
                        return phys
                    phys_u = _phys(t_sab,"6:20") + _phys(t_dom,"6:20") + _phys(t_812,"8:12")
                    if np.any((hu>0.01) & (phys_u < min_physical)): continue
                    if hs.max()>0.01:
                        phys_s = _phys(t_sab,"6:20")
                        if np.any((hs>0.01) & (phys_s < min_physical)): continue
                    if hd.max()>0.01:
                        phys_d = _phys(t_dom,"6:20")
                        if np.any((hd>0.01) & (phys_d < min_physical)): continue

                # SLA util com o trial
                cov_t = (coverage_from_schedule(t_sab,"6:20",pl_620, wrap) +
                         coverage_from_schedule(t_dom,"6:20",pl_620, wrap) +
                         coverage_from_schedule(t_812,"8:12",pl_812, wrap))
                from .erlang import calc_sla, calc_traffic
                num=den=0.0
                for slot in range(INTERVALS_30MIN):
                    i0=slot*3
                    if not wrap:
                        if i0+3<=first_slot or i0>=last_slot: continue
                    v30=float(vol_util[i0:i0+3].sum())
                    if v30<0.01: continue
                    t30=float((tmo_util[i0:i0+3]*vol_util[i0:i0+3]).sum()/(v30+1e-9))
                    c30=float(cov_t[i0:i0+3].mean())
                    u=sum(calc_traffic(vol_util[i0+k],tmo_util[i0+k]) for k in range(3))/3
                    m=max(1,math.ceil(c30)) if c30>0 else 0
                    s30=calc_sla(m,u,t_target,t30) if m>0 else (1.0 if v30==0 else 0.0)
                    num+=v30*s30; den+=v30
                ts_util = num/den if den>0 else 0.0

                if ts_util >= sla_target and ts_util > best_sla:
                    best_sla=ts_util; best_pool=pool_name; best_idx=idx

        if best_pool is None: break

        target_sched = {"sab": sched_sab, "dom": sched_dom, "812": sched_812}[best_pool]
        s, n = target_sched[best_idx]
        if n == 1: target_sched.pop(best_idx)
        else:      target_sched[best_idx] = (s, n-1)

    sched_sab = [(s,n) for s,n in sched_sab if n>0]
    sched_dom = [(s,n) for s,n in sched_dom if n>0]
    sched_812 = [(s,n) for s,n in sched_812 if n>0]
    return sched_sab, sched_dom, sched_812
