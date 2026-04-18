"""WFM Engine — orchestrator com suporte a Erlang C e Erlang A."""
import math, time, calendar as cal_mod
import numpy as np
from typing import List, Dict, Optional, Tuple

from .models import (WFMInput, WFMOutput, Turno, Alerta, DiaOut, IntervaloOut,
                     SHIFTS, INTERVALS_PER_DAY, INTERVALS_30MIN)
from .erlang  import (min_hc_for_sla, calc_sla, calc_tme, calc_occupancy,
                      erlang_c, calc_traffic,
                      calc_sla_auto, calc_tme_auto, min_hc_for_sla_auto,
                      calc_p_abandon)
from .demand  import (build_day_curves, default_curva, gen_dias_mes, slot_to_time,
                      to_30min_mean, volume_from_peso, HORARIOS,
                      time_to_slot, detect_operating_hours)
from .solver  import (pl_efetivo, valid_slots, coverage_mask, coverage_from_schedule,
                      schedule_to_total_agents, heuristic_solve, hybrid_solve, exact_solve,
                      unified_pool_solve)

DIAS_PT = ["Segunda","Terça","Quarta","Quinta","Sexta","Sábado","Domingo"]

def _dow(data_str):
    import datetime
    return DIAS_PT[datetime.date.fromisoformat(data_str).weekday()]


def erlang_curve_day(vol_10m, tmo_10m, t_target, sla_target, min_hc=0,
                     erlang_mode="erlang_c", patience=300.0):
    """Curva de HC mínimo por slot de 10min usando Erlang C ou A."""
    hc = np.zeros(INTERVALS_PER_DAY)
    sla = np.zeros(INTERVALS_PER_DAY)
    traf = np.zeros(INTERVALS_PER_DAY)
    for i in range(INTERVALS_PER_DAY):
        u = calc_traffic(vol_10m[i], tmo_10m[i])
        traf[i] = u
        m = max(min_hc, min_hc_for_sla_auto(u, t_target, tmo_10m[i], sla_target,
                                             erlang_mode, patience))
        hc[i] = m
        sla[i] = calc_sla_auto(m, u, t_target, tmo_10m[i], erlang_mode, patience)
    return hc, sla, traf


def compute_day_indicators(vol_10m, tmo_10m, cov_10m, t_target, sla_target,
                            first_slot_op, last_slot_close, pl_eff,
                            erlang_mode="erlang_c", patience=300.0) -> List[IntervaloOut]:
    out = []
    for slot in range(INTERVALS_30MIN):
        i0 = slot * 3
        slot_start = slot * 3
        slot_end   = slot_start + 3
        in_window = not (slot_end <= first_slot_op or slot_start >= last_slot_close)
        if not in_window:
            if last_slot_close < 144 or first_slot_op > 0:
                continue

        vol30  = float(vol_10m[i0:i0+3].sum())
        tmo30  = float((tmo_10m[i0:i0+3] * vol_10m[i0:i0+3]).sum() /
                       (vol_10m[i0:i0+3].sum() + 1e-9))
        cov30  = float(cov_10m[i0:i0+3].mean())
        u_avg  = float(sum(calc_traffic(vol_10m[i0+k], tmo_10m[i0+k]) for k in range(3)) / 3)
        m_bruto = max(1, math.ceil(cov30)) if cov30 > 0 else 0
        hc_liq_slot = max(0, round(float(max(
            min_hc_for_sla_auto(calc_traffic(vol_10m[i0+k], tmo_10m[i0+k]),
                                t_target, tmo_10m[i0+k], sla_target, erlang_mode, patience)
            for k in range(3)))))

        pw   = erlang_c(m_bruto, u_avg) if m_bruto > 0 else 0.0
        sla  = calc_sla_auto(m_bruto, u_avg, t_target, tmo30, erlang_mode, patience) \
               if m_bruto > 0 else (1.0 if vol30 == 0 else 0.0)
        tme  = calc_tme_auto(m_bruto, u_avg, tmo30, erlang_mode, patience) \
               if m_bruto > 0 else 0.0
        occ  = calc_occupancy(m_bruto, u_avg) if m_bruto > 0 else 0.0
        ns   = vol30 * sla
        # Erlang A: taxa de abandono estimada para o intervalo
        p_ab = calc_p_abandon(m_bruto, u_avg, tmo30, patience) \
               if (erlang_mode == "erlang_a" and m_bruto > 0 and patience > 0) else 0.0

        out.append(IntervaloOut(
            horario     = HORARIOS[slot],
            volume      = round(vol30, 1),
            tmo         = round(tmo30, 1),
            hc_liq      = hc_liq_slot,
            hc_bruto    = round(cov30, 2),
            trafico_erl = round(u_avg, 3),
            fila_pw     = round(pw, 4),
            tme_seg     = round(tme, 2),
            ocupacao    = round(occ, 4),
            sla_pct     = round(sla, 4),
            ns          = round(ns, 1),
            p_abandon   = round(p_ab, 4),
        ))
    return out


def run_engine(inp: WFMInput) -> WFMOutput:
    t0 = time.time()

    emode   = inp.erlang_mode
    patience = inp.patience_time

    # ── Shrinkage & PL ────────────────────────────────────────────────
    shrink = inp.pausas.total
    pl = {"6:20": pl_efetivo("6:20", shrink), "8:12": pl_efetivo("8:12", shrink)}
    pl_base = {"6:20": SHIFTS["6:20"]["pl_base"], "8:12": SHIFTS["8:12"]["pl_base"]}

    # ── Curvas ────────────────────────────────────────────────────────
    curva_sem = inp.curva_semana  or default_curva("Util")
    curva_sab = inp.curva_sabado  or default_curva("Sabado")
    curva_dom = inp.curva_domingo or default_curva("Domingo")

    # ── Operating hours ───────────────────────────────────────────────
    if inp.horario_abertura and inp.horario_fechamento:
        first_slot_op   = time_to_slot(inp.horario_abertura)
        last_slot_close = time_to_slot(inp.horario_fechamento)
    else:
        first_slot_op, last_slot_close = detect_operating_hours(curva_sem)

    is_24h = (first_slot_op == 0 and last_slot_close == 144)

    if inp.min_agentes_intervalo == 0 and not inp.horario_abertura and is_24h:
        _min_hc = 1
    else:
        _min_hc = inp.min_agentes_intervalo

    # ── Dias do mês ───────────────────────────────────────────────────
    dias = inp.dias_mes if inp.dias_mes else gen_dias_mes(
        inp.ano, inp.mes, 22,
        inp.vol_sabado_pct, inp.vol_domingo_pct, inp.volume_mes
    )

    # ── V5: Filtrar dias por dias_funcionamento ──────────────────────
    # Mapa de dia_semana PT → chave curta usada no front
    _DOW_KEY = {0:"seg", 1:"ter", 2:"qua", 3:"qui", 4:"sex", 5:"sab", 6:"dom"}
    import datetime as _dt
    active_keys = set(inp.dias_funcionamento or [])
    if active_keys:
        def _is_active(d):
            try:
                dow = _dt.date.fromisoformat(d.data).weekday()
                return _DOW_KEY[dow] in active_keys
            except Exception:
                return True
        # Remove dias inativos e redistribui o peso proporcionalmente
        dias_ativos = [d for d in dias if _is_active(d)]
        dias_removidos = [d for d in dias if not _is_active(d)]
        if dias_removidos and dias_ativos:
            peso_removido = sum(d.peso_pct for d in dias_removidos)
            peso_total_ativo = sum(d.peso_pct for d in dias_ativos)
            if peso_total_ativo > 0:
                fator = (peso_total_ativo + peso_removido) / peso_total_ativo
                for d in dias_ativos:
                    d.peso_pct *= fator
        dias = dias_ativos

    util_dias = [d for d in dias if d.tipo == "Util"]
    sab_dias  = [d for d in dias if d.tipo == "Sabado"]
    dom_dias  = [d for d in dias if d.tipo == "Domingo"]

    # ── Average volumes per type ──────────────────────────────────────
    vol_util_avg = sum(volume_from_peso(d.peso_pct, inp.volume_mes) for d in util_dias) / max(len(util_dias),1)
    vol_sab_avg  = sum(volume_from_peso(d.peso_pct, inp.volume_mes) for d in sab_dias)  / max(len(sab_dias), 1)
    vol_dom_avg  = sum(volume_from_peso(d.peso_pct, inp.volume_mes) for d in dom_dias)  / max(len(dom_dias), 1)

    vol_util_10m, tmo_util_10m = build_day_curves(vol_util_avg, curva_sem, inp.tmo_base)
    vol_sab_10m,  tmo_sab_10m  = build_day_curves(vol_sab_avg,  curva_sab, inp.tmo_base)
    vol_dom_10m,  tmo_dom_10m  = build_day_curves(vol_dom_avg,  curva_dom, inp.tmo_base)

    # ── Erlang curves ─────────────────────────────────────────────────
    hc_util, _, _ = erlang_curve_day(vol_util_10m, tmo_util_10m, inp.tempo_target,
                                      inp.sla_target, _min_hc, emode, patience)
    hc_sab,  _, _ = erlang_curve_day(vol_sab_10m,  tmo_sab_10m,  inp.tempo_target,
                                      inp.sla_target, _min_hc, emode, patience)
    hc_dom,  _, _ = erlang_curve_day(vol_dom_10m,  tmo_dom_10m,  inp.tempo_target,
                                      inp.sla_target, _min_hc, emode, patience)

    # ── Fix 24h: unificar curvas de fim de semana ─────────────────────
    # Garante cobertura completa em operações 24h e elimina gaps entre
    # os pools de sábado e domingo (escala unificada = max de ambos).
    if is_24h and (sab_dias or dom_dias):
        hc_weekend = np.maximum(hc_sab, hc_dom)
        # Forçar pelo menos 1 HC líquido em todos os slots para 24h
        hc_weekend = np.maximum(hc_weekend, _min_hc)
        hc_sab = hc_weekend.copy()
        hc_dom = hc_weekend.copy()

    # ── Solve pools — arquitetura unificada 6×1 ────────────────────────
    milp_limit = 5.0 if inp.solver_mode == "heuristic" else 120.0

    pools = unified_pool_solve(
        hc_util, hc_sab, hc_dom,
        pl["6:20"], pl["8:12"],
        inp.max_horarios_entrada,
        vol_util_10m, tmo_util_10m,
        vol_sab_10m,  tmo_sab_10m,
        vol_dom_10m,  tmo_dom_10m,
        inp.sla_target, inp.tempo_target,
        first_slot_op, last_slot_close,
        milp_time_limit  = milp_limit,
        mip_rel_gap      = 0.02,
        min_physical     = _min_hc,
        wrap             = is_24h,
    )

    sched_b_sab = pools["sab"]
    sched_b_dom = pools["dom"]
    sched_a     = pools["812"]

    cov_b_sab   = coverage_from_schedule(sched_b_sab, "6:20", pl["6:20"], wrap=is_24h)
    cov_b_dom   = coverage_from_schedule(sched_b_dom, "6:20", pl["6:20"], wrap=is_24h)
    cov_a       = coverage_from_schedule(sched_a,     "8:12", pl["8:12"], wrap=is_24h)

    cov_weekday  = cov_b_sab + cov_b_dom + cov_a
    cov_saturday = cov_b_sab
    cov_sunday   = cov_b_dom

    n_sab       = schedule_to_total_agents(sched_b_sab)
    n_dom       = schedule_to_total_agents(sched_b_dom)
    n_812       = schedule_to_total_agents(sched_a)
    total_bruto = n_sab + n_dom + n_812

    n_a_liq     = round(n_812 * pl["8:12"] / pl_base["8:12"], 1)
    n_b_sab_liq = round(n_sab * pl["6:20"] / pl_base["6:20"], 1)
    n_b_dom_liq = round(n_dom * pl["6:20"] / pl_base["6:20"], 1)

    # ── Day-by-day output ─────────────────────────────────────────────
    cov_map   = {"Util": cov_weekday, "Sabado": cov_saturday, "Domingo": cov_sunday}
    curva_map = {"Util": curva_sem,   "Sabado": curva_sab,    "Domingo": curva_dom}
    dias_out: List[DiaOut] = []

    all_sla_num = all_vol = all_ns = 0.0
    all_iv_ok = all_iv_tot = 0

    for dia in dias:
        vol_dia  = volume_from_peso(dia.peso_pct, inp.volume_mes)
        vol_10m, tmo_10m = build_day_curves(vol_dia, curva_map[dia.tipo], inp.tmo_base)
        cov_10m  = cov_map[dia.tipo]

        intervalos = compute_day_indicators(
            vol_10m, tmo_10m, cov_10m,
            inp.tempo_target, inp.sla_target,
            first_slot_op, last_slot_close, pl["8:12"],
            emode, patience,
        )

        vol_total = sum(iv.volume for iv in intervalos)
        ns_total  = sum(iv.ns     for iv in intervalos)
        sla_pond  = sum(iv.volume * iv.sla_pct for iv in intervalos) / (vol_total + 1e-9)
        tme_med   = sum(iv.volume * iv.tme_seg  for iv in intervalos) / (vol_total + 1e-9)
        occ_med   = sum(iv.volume * iv.ocupacao  for iv in intervalos) / (vol_total + 1e-9)
        fila_med  = sum(iv.volume * iv.fila_pw   for iv in intervalos) / (vol_total + 1e-9)
        tmo_med   = sum(iv.volume * iv.tmo       for iv in intervalos) / (vol_total + 1e-9)
        hc_liq_max   = max((iv.hc_liq   for iv in intervalos), default=0)
        hc_bruto_max = max((iv.hc_bruto for iv in intervalos), default=0)
        ok_count  = sum(1 for iv in intervalos if iv.sla_pct >= inp.sla_target)

        all_sla_num += vol_total * sla_pond
        all_vol     += vol_total
        all_ns      += ns_total
        all_iv_ok   += ok_count
        all_iv_tot  += len(intervalos)

        dias_out.append(DiaOut(
            data=dia.data, dia_semana=_dow(dia.data), tipo=dia.tipo,
            peso_pct=dia.peso_pct, volume_total=round(vol_total,0),
            tmo_medio=round(tmo_med,1), hc_liq_max=hc_liq_max,
            hc_bruto_max=round(hc_bruto_max,1),
            sla_ponderado=round(sla_pond,4), ns_total=round(ns_total,0),
            tme_medio=round(tme_med,2), ocupacao_media=round(occ_med,4),
            fila_media=round(fila_med,4),
            intervalos_ok_pct=round(ok_count/(len(intervalos)+1e-9),4),
            status_sla="ok" if sla_pond >= inp.sla_target else "abaixo",
            intervalos=intervalos,
        ))

    # ── Monthly KPIs ──────────────────────────────────────────────────
    sla_mes   = all_sla_num / (all_vol + 1e-9)
    floor_pct = all_iv_ok  / (all_iv_tot + 1e-9) if all_iv_tot else 0

    def _ov(tipo, cov_curve):
        typed = [d for d in dias_out if d.tipo == tipo]
        if not typed: return 0.0
        ratios, vols = [], []
        for _d in typed:
            peak_req = max((iv.hc_liq for iv in _d.intervalos), default=0)
            if peak_req > 0:
                ratios.append(float(cov_curve.max()) / peak_req)
                vols.append(_d.volume_total)
        if not ratios: return 0.0
        return sum(r*v for r,v in zip(ratios,vols)) / (sum(vols)+1e-9)

    ov_u  = _ov("Util",    cov_weekday)
    ov_s  = _ov("Sabado",  cov_saturday)
    ov_d  = _ov("Domingo", cov_sunday)
    vol_u  = sum(d.volume_total for d in dias_out if d.tipo=="Util")
    vol_s  = sum(d.volume_total for d in dias_out if d.tipo=="Sabado")
    vol_d_ = sum(d.volume_total for d in dias_out if d.tipo=="Domingo")
    vol_t  = vol_u + vol_s + vol_d_ + 1e-9
    overstaffing = max(0.0, (ov_u*vol_u + ov_s*vol_s + ov_d*vol_d_)/vol_t - 1.0)

    # ── Status & alerts ───────────────────────────────────────────────
    alertas, status = [], "optimal"
    if sla_mes < inp.sla_target - 0.01:
        status = "infeasible"
        alertas.append(Alerta("SLA_INSUFICIENTE",
            f"SLA mensal {sla_mes*100:.1f}% abaixo do target {inp.sla_target*100:.0f}%. "
            f"Considere reduzir o SLA target ou aumentar o overstaffing máximo.",
            int(total_bruto*0.1)))
    if overstaffing > inp.max_overstaffing + 0.005 and status == "optimal":
        status = "constrained"
        alertas.append(Alerta("OVERSTAFFING_EXCEDIDO",
            f"Overstaffing no pico {overstaffing*100:.1f}% acima do limite de {inp.max_overstaffing*100:.0f}%."))
    if (inp.sla_mode=="interval_floor" and
        floor_pct < inp.interval_floor_pct - 0.005 and status == "optimal"):
        status = "constrained"
        hc_ad = max(1, int((inp.interval_floor_pct - floor_pct) * total_bruto * 1.5))
        alertas.append(Alerta("FLOOR_NAO_ATINGIVEL",
            f"Floor de {inp.interval_floor_pct*100:.0f}% não atingível. "
            f"Intervalos OK: {floor_pct*100:.1f}%.", hc_ad))
    if status == "optimal":
        dias_abaixo = sum(1 for d in dias_out if d.status_sla=="abaixo")
        modelo_tag = f"[{emode.upper()}]" if emode == "erlang_a" else ""
        alertas.append(Alerta("OPTIMAL",
            f"{modelo_tag} SLA mensal {sla_mes*100:.1f}% ✓ | "
            f"{dias_abaixo} dia(s) abaixo do SLA individual | "
            f"Overstaffing no pico {overstaffing*100:.1f}%."))

    # ── Named turnos ──────────────────────────────────────────────────
    def mk_turnos(sched, tipo, pool):
        dur = SHIFTS[tipo]["duration_min"]
        pl_t = pl[tipo]; pl_b = pl_base[tipo]
        result = []
        for s, n in sorted(sched, key=lambda x: x[0]):
            entrada = slot_to_time(s)
            sm = (s*10 + dur) % (24*60)
            saida = f"{sm//60:02d}:{sm%60:02d}"
            result.append(Turno(
                pool=pool, tipo=tipo, entrada=entrada, saida=saida,
                agentes_bruto=n,
                agentes_liq=round(n * pl_t / pl_b, 1),
            ))
        return result

    turnos = (mk_turnos(sched_a,     "8:12", "pool_812") +
              mk_turnos(sched_b_sab, "6:20", "pool_b_sab") +
              mk_turnos(sched_b_dom, "6:20", "pool_b_dom"))

    return WFMOutput(
        status=status, mes=inp.mes, ano=inp.ano,
        sla_ponderado_mes=round(sla_mes,4), sla_target=inp.sla_target,
        overstaffing=round(overstaffing,4), max_overstaffing=inp.max_overstaffing,
        ns_total_mes=round(all_ns,0), volume_total_mes=round(all_vol,0),
        intervalos_ok_pct=round(floor_pct,4) if inp.sla_mode=="interval_floor" else None,
        interval_floor_target=inp.interval_floor_pct if inp.sla_mode=="interval_floor" else None,
        hc_fisico={"pool_812":n_812,"pool_b_sab":n_sab,"pool_b_dom":n_dom,"total":total_bruto},
        hc_liquido_ref={"pool_812":n_a_liq,"pool_b_sab":n_b_sab_liq,"pool_b_dom":n_b_dom_liq,
                        "total":round(n_a_liq+n_b_sab_liq+n_b_dom_liq,1)},
        pl_efetivo={"8:12":round(pl["8:12"],4),"6:20":round(pl["6:20"],4),
                    "8:12_base":pl_base["8:12"],"6:20_base":pl_base["6:20"]},
        turnos=turnos, dias=dias_out, alertas=alertas,
        solver_mode=inp.solver_mode, erlang_mode=emode, patience_time=patience,
        elapsed_sec=time.time()-t0,
        horario_abertura=inp.horario_abertura or slot_to_time(first_slot_op),
        horario_fechamento=inp.horario_fechamento or slot_to_time(last_slot_close),
    )
