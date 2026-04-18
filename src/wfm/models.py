from dataclasses import dataclass, field
from typing import List, Dict, Optional

SHIFTS = {
    "6:20": {"duration_min":380,"n_intervals":38,"breaks_min":40,"pl_base":0.8947,"first_entry_slot":0,"last_entry_slot":106},
    "8:12": {"duration_min":492,"n_intervals":49,"breaks_min":80,"pl_base":0.8374,"first_entry_slot":0,"last_entry_slot":94},
}
INTERVALS_PER_DAY = 144
INTERVALS_30MIN   = 48

@dataclass
class Shrinkage:
    absenteismo: float = 0.05
    ferias:      float = 0.05
    treinamento: float = 0.03
    reunioes:    float = 0.02
    aderencia:   float = 0.03
    outras:      float = 0.00

    @property
    def total(self) -> float:
        return min(0.70, self.absenteismo + self.ferias + self.treinamento
                       + self.reunioes + self.aderencia + self.outras)

class PausasAdicionais(Shrinkage):
    def __init__(self, treinamento=0.0, feedback=0.0, outras=0.0, **kwargs):
        super().__init__(
            treinamento = treinamento,
            aderencia   = kwargs.pop("aderencia", feedback),
            outras      = outras,
            absenteismo = kwargs.pop("absenteismo", 0.0),
            ferias      = kwargs.pop("ferias",      0.0),
            reunioes    = kwargs.pop("reunioes",    0.0),
        )

@dataclass
class CurvaIntraday:
    pesos:       List[float]
    fatores_tmo: List[float]

@dataclass
class DiaMes:
    data:     str
    tipo:     str
    peso_pct: float

@dataclass
class WFMInput:
    volume_mes:           int
    tmo_base:             float
    sla_target:           float
    tempo_target:         float
    mes:                  int   = 5
    ano:                  int   = 2025
    sla_mode:             str   = "weighted"
    interval_floor_pct:   float = 0.85
    max_overstaffing:     float = 0.20
    max_horarios_entrada: int   = 6
    solver_mode:          str   = "heuristic"
    # ── Erlang A ──────────────────────────────────────────────────────
    erlang_mode:          str   = "erlang_c"   # "erlang_c" | "erlang_a"
    patience_time:        float = 300.0        # segundos — relevante só p/ Erlang A
    # ── Horários ──────────────────────────────────────────────────────
    pausas: PausasAdicionais    = field(default_factory=PausasAdicionais)
    horario_abertura:     str   = ""
    horario_fechamento:   str   = ""
    # ── Janela de entrada permitida ──────────────────────────────────
    # Restringe o horário em que novos turnos podem começar (ex: "06:00"
    # a "17:40" numa operação 24h, impedindo entradas de madrugada ou noite).
    # Strings vazias = sem restrição (turnos podem entrar em qualquer horário).
    janela_entrada_inicio: str  = ""
    janela_entrada_fim:    str  = ""
    # Lista de janelas BLOQUEADAS de entrada (ex: ["00:00-05:59","17:41-23:59"]).
    # Quando preenchida, tem precedência sobre janela_entrada_inicio/fim.
    janelas_bloqueadas: List[str] = field(default_factory=list)
    min_agentes_intervalo: int  = 0
    # ── Curvas ────────────────────────────────────────────────────────
    curva_semana:  Optional[CurvaIntraday] = None
    curva_sabado:  Optional[CurvaIntraday] = None
    curva_domingo: Optional[CurvaIntraday] = None
    dias_mes: List[DiaMes] = field(default_factory=list)
    vol_sabado_pct:  float = 0.35
    vol_domingo_pct: float = 0.15
    # ── V5: Dias de funcionamento ─────────────────────────────────────
    # Lista de dias ativos: seg, ter, qua, qui, sex, sab, dom
    # Dias ausentes desta lista recebem volume zero e não entram na escala.
    dias_funcionamento: List[str] = field(default_factory=lambda:
        ["seg","ter","qua","qui","sex","sab","dom"])

@dataclass
class IntervaloOut:
    horario:      str
    volume:       float
    tmo:          float
    hc_liq:       int
    hc_bruto:     float
    trafico_erl:  float
    fila_pw:      float
    tme_seg:      float
    ocupacao:     float
    sla_pct:      float
    ns:           float
    p_abandon:    float = 0.0   # Erlang A — taxa de abandono estimada

@dataclass
class DiaOut:
    data:              str
    dia_semana:        str
    tipo:              str
    peso_pct:          float
    volume_total:      float
    tmo_medio:         float
    hc_liq_max:        int
    hc_bruto_max:      float
    sla_ponderado:     float
    ns_total:          float
    tme_medio:         float
    ocupacao_media:    float
    fila_media:        float
    intervalos_ok_pct: float
    status_sla:        str
    intervalos:        List[IntervaloOut]

@dataclass
class Turno:
    pool: str; tipo: str; entrada: str; saida: str
    agentes_bruto: int
    agentes_liq:   float

@dataclass
class Alerta:
    codigo: str; mensagem: str; hc_adicional_necessario: int = 0

@dataclass
class WFMOutput:
    status:               str
    mes:                  int
    ano:                  int
    sla_ponderado_mes:    float
    sla_target:           float
    overstaffing:         float
    max_overstaffing:     float
    ns_total_mes:         float
    volume_total_mes:     float
    intervalos_ok_pct:    Optional[float]
    interval_floor_target: Optional[float]
    hc_fisico:            Dict[str, int]
    hc_liquido_ref:       Dict[str, float]
    pl_efetivo:           Dict[str, float]
    turnos:               List[Turno]
    dias:                 List[DiaOut]
    alertas:              List[Alerta]
    solver_mode:          str   = "heuristic"
    erlang_mode:          str   = "erlang_c"
    patience_time:        float = 300.0
    elapsed_sec:          float = 0.0
    horario_abertura:     str   = ""
    horario_fechamento:   str   = ""
