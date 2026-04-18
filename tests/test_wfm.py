"""Testes unitários para o WFM Engine."""
import pytest
import numpy as np
from wfm.erlang import erlang_c, calc_sla, calc_tme, calc_traffic, min_hc_for_sla
from wfm.models import WFMInput, PausasAdicionais, CurvaIntraday
from wfm.engine import run_engine


class TestErlangC:
    """Testes para fórmulas Erlang C."""

    def test_erlang_c_zero_traffic(self):
        """Tráfego zero deve resultar em P(wait)=0."""
        assert erlang_c(5, 0.0) == 0.0

    def test_erlang_c_high_traffic(self):
        """Tráfego >= capacidade deve resultar em P(wait)=1."""
        assert erlang_c(5, 5.0) == 1.0
        assert erlang_c(5, 6.0) == 1.0

    def test_erlang_c_typical(self):
        """Caso típico: m=10, u=7 deve ter P(wait) entre 0 e 1."""
        p = erlang_c(10, 7.0)
        assert 0.0 < p < 1.0

    def test_calc_sla_no_traffic(self):
        """Sem tráfego, SLA deve ser 100%."""
        assert calc_sla(5, 0.0, 20, 240) == 1.0

    def test_calc_sla_insufficient_agents(self):
        """Agentes insuficientes deve resultar em SLA baixo."""
        sla = calc_sla(1, 10.0, 20, 240)
        assert sla < 0.5

    def test_calc_traffic(self):
        """Cálculo de tráfego: 100 chamadas de 240s em 10min = 40 Erlangs."""
        u = calc_traffic(100, 240)
        assert abs(u - 40.0) < 0.001


class TestMinHC:
    """Testes para cálculo de HC mínimo."""

    def test_min_hc_zero_volume(self):
        """Volume zero requer HC zero."""
        assert min_hc_for_sla(0.0, 20, 240, 0.8) == 0

    def test_min_hc_increases_with_sla(self):
        """SLA maior requer mais agentes."""
        hc_70 = min_hc_for_sla(10.0, 20, 240, 0.70)
        hc_90 = min_hc_for_sla(10.0, 20, 240, 0.90)
        assert hc_90 > hc_70


class TestWFMInput:
    """Testes para modelo de input."""

    def test_default_values(self):
        """Valores padrão devem ser corretos."""
        inp = WFMInput(volume_mes=100000, tmo_base=240, sla_target=0.8, tempo_target=20)
        assert inp.mes == 5
        assert inp.ano == 2025
        assert inp.solver_mode == "heuristic"
        assert inp.erlang_mode == "erlang_c"

    def test_shrinkage_total(self):
        """Shrinkage total deve ser soma limitada a 70%."""
        pausas = PausasAdicionais(
            absenteismo=0.1,
            ferias=0.1,
            treinamento=0.1,
            reunioes=0.1,
            aderencia=0.1,
            outras=0.1
        )
        assert pausas.total == 0.6

    def test_shrinkage_cap(self):
        """Shrinkage total não pode exceder 70%."""
        pausas = PausasAdicionais(
            absenteismo=0.3,
            ferias=0.3,
            treinamento=0.3,
            reunioes=0.3,
            aderencia=0.3,
            outras=0.3
        )
        assert pausas.total == 0.7


class TestEngine:
    """Testes para engine principal."""

    def test_run_engine_basic(self):
        """Execução básica da engine."""
        inp = WFMInput(
            volume_mes=50000,
            tmo_base=180,
            sla_target=0.8,
            tempo_target=20,
            mes=6,
            ano=2025,
            solver_mode="heuristic",
            erlang_mode="erlang_c"
        )
        out = run_engine(inp)
        
        assert out.status in ["optimal", "infeasible", "constrained"]
        assert out.mes == 6
        assert out.ano == 2025
        assert len(out.dias) > 0
        assert len(out.turnos) > 0

    def test_run_engine_erlang_a(self):
        """Execução com Erlang A."""
        inp = WFMInput(
            volume_mes=50000,
            tmo_base=180,
            sla_target=0.8,
            tempo_target=20,
            mes=6,
            ano=2025,
            solver_mode="heuristic",
            erlang_mode="erlang_a",
            patience_time=300.0
        )
        out = run_engine(inp)
        
        assert out.erlang_mode == "erlang_a"
        assert out.patience_time == 300.0


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
