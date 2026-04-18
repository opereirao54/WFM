"""Erlang C and Erlang A core mathematics + all derived indicators."""
import math


# ── Erlang C (M/M/c) ─────────────────────────────────────────────────────

def erlang_c(m: int, u: float) -> float:
    """Erlang C — P(wait) em fila M/M/c. Log-space para evitar overflow."""
    if u <= 0: return 0.0
    if m <= u: return 1.0
    log_Am_m = m * math.log(u) - sum(math.log(k) for k in range(1, m + 1))
    log_num = log_Am_m - math.log(1.0 - u / m)
    log_ak = 0.0
    log_terms = [0.0]
    for k in range(1, m):
        log_ak += math.log(u) - math.log(k)
        log_terms.append(log_ak)
    base     = max(log_terms + [log_num])
    partial  = sum(math.exp(lt - base) for lt in log_terms)
    num_norm = math.exp(log_num - base)
    return float(min(num_norm / (partial + num_norm), 1.0))


def calc_sla(m: int, u: float, t_target: float, tmo: float) -> float:
    if u <= 0: return 1.0
    if m <= 0 or m <= u: return 0.0
    ec = erlang_c(m, u)
    return max(0.0, min(1.0, 1.0 - ec * math.exp((u - m) * (t_target / tmo))))


def calc_tme(m: int, u: float, tmo: float) -> float:
    if u <= 0: return 0.0
    if m <= u: return 9999.0
    ec = erlang_c(m, u)
    denom = m - u
    return (ec * tmo) / denom if denom > 0 else 9999.0


def calc_occupancy(m: int, u: float) -> float:
    return min(0.99, u / m) if m > 0 else 0.0


def calc_traffic(volume_10m: float, tmo: float) -> float:
    return (volume_10m * tmo) / 600.0


def min_hc_for_sla(u: float, t_target: float, tmo: float, sla_target: float) -> int:
    if u <= 0: return 0
    m = max(1, math.ceil(u) + 1)
    for _ in range(500):
        if calc_sla(m, u, t_target, tmo) >= sla_target:
            return m
        m += 1
    return m


# ── Erlang A (M/M/c+M) — abandono com tempo de paciência ─────────────────

def _erlang_a_states(m: int, A: float, gamma: float, N_extra: int = 150) -> list:
    """
    Probabilidades de estado normalizadas para fila M/M/c+M.

    Parâmetros:
      A     = tráfego oferecido em Erlangs  (lambda / mu)
      gamma = beta / mu = AHT / patience_time  (adimensional)

    Recursão:
      n < m : p[n+1] = A / (n+1)           * p[n]
      n >= m: p[n+1] = A / (m + (n-m+1)*γ) * p[n]
    """
    if A <= 0 or m <= 0:
        return []
    N = m + N_extra
    p = [0.0] * (N + 1)
    p[0] = 1.0
    for n in range(N):
        if n < m:
            p[n + 1] = (A / (n + 1)) * p[n]
        else:
            denom = m + (n - m + 1) * gamma
            p[n + 1] = (A / denom) * p[n] if denom > 0 else 0.0
    total = sum(p)
    return [x / total for x in p] if total > 0 else p


def calc_p_abandon(m: int, u: float, tmo: float, patience: float) -> float:
    """Probabilidade de abandono na fila M/M/c+M."""
    if u <= 0 or patience <= 0 or m <= 0:
        return 0.0
    gamma = tmo / patience
    p = _erlang_a_states(m, u, gamma)
    if not p:
        return 0.0
    abandon_sum = sum((n - m) * p[n] for n in range(m + 1, len(p)))
    return min(1.0, max(0.0, gamma * abandon_sum / u))


def calc_sla_a(m: int, u: float, t_target: float, tmo: float, patience: float) -> float:
    """
    SLA para Erlang A (M/M/c+M).

    Fórmula: SLA ≈ 1 - P(wait>0) * exp(-θ * T)
    onde θ = c*μ - λ_eff   (capacidade de segurança após abandono)

    Quando patience → ∞ recai em Erlang C.
    """
    if u <= 0:
        return 1.0
    if m <= 0:
        return 0.0
    if patience <= 0 or patience > 9000:
        return calc_sla(m, u, t_target, tmo)

    mu    = 1.0 / tmo
    lam   = u * mu
    gamma = tmo / patience  # = beta/mu

    p = _erlang_a_states(m, u, gamma)
    if not p:
        return 0.0

    p_queue   = sum(p[m:])
    aband_sum = sum((n - m) * p[n] for n in range(m + 1, len(p)))
    p_aband   = min(1.0, max(0.0, gamma * aband_sum / u))

    lam_eff = lam * (1.0 - p_aband)
    theta   = m * mu - lam_eff

    if theta <= 0:
        sla = max(0.0, 1.0 - p_queue)
    else:
        sla = 1.0 - p_queue * math.exp(-theta * t_target)

    return max(0.0, min(1.0, sla))


def calc_tme_a(m: int, u: float, tmo: float, patience: float) -> float:
    """TME (ASA) médio para Erlang A — aproximação de fila estável."""
    if u <= 0:
        return 0.0
    if patience <= 0 or patience > 9000:
        return calc_tme(m, u, tmo)

    mu    = 1.0 / tmo
    lam   = u * mu
    gamma = tmo / patience
    p     = _erlang_a_states(m, u, gamma)
    if not p:
        return 0.0

    p_queue   = sum(p[m:])
    aband_sum = sum((n - m) * p[n] for n in range(m + 1, len(p)))
    p_aband   = min(1.0, max(0.0, gamma * aband_sum / u))

    lam_eff = lam * (1.0 - p_aband)
    theta   = m * mu - lam_eff

    if theta <= 0 or p_queue <= 0:
        return 9999.0
    return p_queue / theta


def min_hc_for_sla_a(u: float, t_target: float, tmo: float, sla_target: float,
                      patience: float) -> int:
    """HC mínimo para atingir SLA usando Erlang A."""
    if u <= 0:
        return 0
    if patience <= 0 or patience > 9000:
        return min_hc_for_sla(u, t_target, tmo, sla_target)
    m = max(1, math.ceil(u) + 1)
    for _ in range(500):
        if calc_sla_a(m, u, t_target, tmo, patience) >= sla_target:
            return m
        m += 1
    return m


# ── Wrappers automáticos (erlang_mode) ────────────────────────────────────

def calc_sla_auto(m: int, u: float, t_target: float, tmo: float,
                  erlang_mode: str = "erlang_c", patience: float = 300.0) -> float:
    if erlang_mode == "erlang_a" and patience > 0:
        return calc_sla_a(m, u, t_target, tmo, patience)
    return calc_sla(m, u, t_target, tmo)


def calc_tme_auto(m: int, u: float, tmo: float,
                  erlang_mode: str = "erlang_c", patience: float = 300.0) -> float:
    if erlang_mode == "erlang_a" and patience > 0:
        return calc_tme_a(m, u, tmo, patience)
    return calc_tme(m, u, tmo)


def min_hc_for_sla_auto(u: float, t_target: float, tmo: float, sla_target: float,
                         erlang_mode: str = "erlang_c", patience: float = 300.0) -> int:
    if erlang_mode == "erlang_a" and patience > 0:
        return min_hc_for_sla_a(u, t_target, tmo, sla_target, patience)
    return min_hc_for_sla(u, t_target, tmo, sla_target)
