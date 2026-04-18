# WFM Engine — Sistema de Dimensionamento de Força de Trabalho

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![Flask](https://img.shields.io/badge/Flask-2.0+-green.svg)](https://flask.palletsprojects.com/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Sistema completo de **Workforce Management (WFM)** para dimensionamento ótimo de equipes em call centers e operações de atendimento. Utiliza modelos matemáticos avançados (Erlang C e Erlang A) combinados com algoritmos de otimização híbrida (MILP + heurísticas) para gerar escalas eficientes que atendem metas de SLA minimizando custos.

---

## 📋 Índice

- [Funcionalidades](#-funcionalidades)
- [Arquitetura](#-arquitetura)
- [Requisitos](#-requisitos)
- [Instalação](#-instalação)
- [Uso](#-uso)
- [API REST](#-api-rest)
- [Modelos Matemáticos](#-modelos-matemáticos)
- [Estrutura do Projeto](#-estrutura-do-projeto)
- [Configurações](#-configurações)
- [Exemplos](#-exemplos)
- [Contribuição](#-contribuição)
- [Licença](#-licença)

---

## ✨ Funcionalidades

### 🎯 Dimensionamento Inteligente
- **Cálculo Erlang C e Erlang A**: Suporte a ambos os modelos para cenários com e sem abandono
- **Otimização Híbrida**: Combina MILP (Programação Linear Inteira Mista) com heurísticas "fatias" para soluções ótimas em <1 segundo
- **Multi-turnos**: Gerencia turnos de 6h20 e 8h12 com janelas de entrada flexíveis
- **Shrinkage Ajustável**: Considera absenteísmo, férias, treinamento, reuniões e aderência

### 📊 Análise Intraday
- **Curvas de Demanda**: Perfis personalizáveis por dia da semana (util, sábado, domingo)
- **SLA Ponderado**: Métricas de Service Level em intervalos de 30 minutos
- **KPIs Completos**: TMO, ocupação, fila esperada, tempo médio de espera, taxa de abandono

### 🔧 Recursos Avançados
- **Detecção Automática de Período**: Extrai mês/ano de arquivos Excel
- **Validação de Dados**: Verifica consistência de curvas de demanda antes do processamento
- **Exportação Excel**: Gera relatórios detalhados prontos para uso operacional
- **Interface Web Moderna**: Dashboard responsivo com visualização em tempo real

### 🛡️ Robustez
- **Post-Trim Inteligente**: Remove folga excessiva mantendo SLA dentro de margens seguras
- **Set-Cover Guarantee**: Garante cobertura completa em operações 24h
- **Tratamento de Erros**: Mensagens claras e logs detalhados para debugging

---

## 🏗️ Arquitetura

```
┌─────────────────────────────────────────────────────────────┐
│                     Interface Web (HTML/CSS/JS)             │
│                    Dashboard interativo                     │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                      API Flask (app.py)                     │
│              Rotas REST para cálculo e exportação           │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                  Engine de Dimensionamento                  │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐      │
│  │   Solver     │  │   Erlang     │  │   Demand     │      │
│  │  (MILP/Heur) │  │  (C / A)     │  │  (Curvas)    │      │
│  └──────────────┘  └──────────────┘  └──────────────┘      │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                   Modelos de Dados (Dataclasses)            │
│            Input/Output tipados e validados                 │
└─────────────────────────────────────────────────────────────┘
```

---

## 📦 Requisitos

### Dependências Principais
- **Python** 3.8+
- **Flask** — Servidor web e API REST
- **NumPy** — Computação numérica vetorializada
- **SciPy** — Otimização MILP (HiGHS solver)
- **OpenPyXL** — Leitura/escrita de arquivos Excel

### Dependências Opcionais
- **PuLP** — Solver alternativo para modo exact (instalável via `pip install pulp`)

---

## 🚀 Instalação

### 1. Clone o Repositório
```bash
git clone https://github.com/seu-usuario/wfm-engine.git
cd wfm-engine
```

### 2. Crie um Ambiente Virtual (Recomendado)
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Linux/Mac
source venv/bin/activate
```

### 3. Instale as Dependências
```bash
pip install -r requirements.txt
```

Ou instale manualmente:
```bash
pip install flask numpy scipy openpyxl
```

### 4. Inicie o Servidor
```bash
# Via linha de comando
python app.py

# Ou use o script batch (Windows)
INICIAR.bat
```

O servidor estará disponível em: **http://localhost:5000**

---

## 💻 Uso

### Interface Web

1. Acesse `http://localhost:5000` no seu navegador
2. Preencha os parâmetros de entrada:
   - Volume mensal de chamadas
   - TMO base (Tempo Médio de Ocupação)
   - Meta de SLA (ex: 80% em 20 segundos)
   - Fatores de shrinkage
   - Curvas de demanda (upload de Excel ou padrão)
3. Clique em **"Calcular"**
4. Visualize os resultados:
   - KPIs consolidados do mês
   - Dimensionamento por pool de atendimento
   - Detalhamento diário com intervalos de 30min
5. Exporte o resultado em Excel

### Programa (Python)

```python
from wfm import WFMInput, run_engine, make_bimodal

# Configurar entrada
inp = WFMInput(
    volume_mes=150000,
    tmo_base=240,
    sla_target=0.80,
    tempo_target=20,
    mes=5,
    ano=2025,
    sla_mode="weighted",
    max_overstaffing=0.20,
    solver_mode="heuristic",
    erlang_mode="erlang_c",
    pausas={
        "absenteismo": 0.05,
        "ferias": 0.05,
        "treinamento": 0.03,
        "reunioes": 0.02,
        "aderencia": 0.03,
    },
    curva_semana=make_bimodal(),
    curva_sabado=None,
    curva_domingo=None,
)

# Executar engine
out = run_engine(inp)

# Acessar resultados
print(f"Status: {out.status}")
print(f"SLA Ponderado: {out.sla_ponderado_mes:.2%}")
print(f"HC Físico Total: {sum(out.hc_fisico.values())}")
print(f"Tempo de Processamento: {out.elapsed_sec:.3f}s")

# Iterar sobre dias
for dia in out.dias:
    print(f"\n{dia.data} ({dia.dia_semana})")
    print(f"  Volume: {dia.volume_total:.0f}")
    print(f"  SLA: {dia.sla_ponderado:.2%}")
    print(f"  HC Máx: {dia.hc_bruto_max:.1f}")
```

---

## 🔌 API REST

### Endpoints

#### `GET /`
Retorna a interface web HTML.

#### `GET /templates`
Baixa planilha Excel modelo para preenchimento de curvas de demanda.

#### `POST /validar`
Valida arquivo Excel de curvas de demanda.

**Request:** `multipart/form-data`
- `curvas_excel`: Arquivo Excel

**Response:**
```json
{
  "valido": true,
  "issues": []
}
```

#### `POST /calcular`
Executa o dimensionamento.

**Request:** `multipart/form-data` ou `application/json`

Parâmetros principais:
| Campo | Tipo | Padrão | Descrição |
|-------|------|--------|-----------|
| `volume_mes` | int | 150000 | Volume total de chamadas no mês |
| `tmo_base` | float | 240 | TMO base em segundos |
| `sla_target` | float | 0.80 | Meta de SLA (0-1) |
| `tempo_target` | float | 20 | Tempo alvo em segundos |
| `mes` | int | 5 | Mês de referência (1-12) |
| `ano` | int | 2025 | Ano de referência |
| `sla_mode` | str | "weighted" | "weighted" ou "daily" |
| `solver_mode` | str | "heuristic" | "heuristic" ou "exact" |
| `erlang_mode` | str | "erlang_c" | "erlang_c" ou "erlang_a" |
| `patience_time` | float | 300 | Tempo de paciência (Erlang A) |
| `absenteismo` | float | 0.05 | Taxa de absenteísmo |
| `ferias` | float | 0.05 | Taxa de férias |
| `treinamento` | float | 0.03 | Taxa de treinamento |
| `vol_sabado_pct` | float | 0.35 | Volume de sábado (% da média) |
| `vol_domingo_pct` | float | 0.15 | Volume de domingo (% da média) |

**Response:**
```json
{
  "status": "optimal",
  "mes": 5,
  "ano": 2025,
  "kpis": {
    "sla_ponderado_mes": 0.8123,
    "sla_target": 0.80,
    "overstaffing": 0.0845,
    "ns_total_mes": 142500
  },
  "hc_fisico": {"pool_a": 45, "pool_b_util": 12},
  "turnos": [...],
  "dias": [...],
  "elapsed_sec": 0.847
}
```

#### `POST /exportar`
Exporta resultados em formato Excel.

**Request:** `application/json`
- Corpo JSON com estrutura de output da engine

**Response:** Arquivo Excel para download

#### `POST /detectar_periodo`
Extrai mês e ano de arquivo Excel.

**Request:** `multipart/form-data`
- `curvas_excel`: Arquivo Excel

**Response:**
```json
{
  "mes": 5,
  "ano": 2025
}
```

---

## 📐 Modelos Matemáticos

### Erlang C
Fórmula clássica para sistemas de filas sem abandono:

```
P(wait > t) = 1 - C(m, u) × exp(-(m - u) × t / TMO)
```

Onde:
- `m` = número de agentes
- `u` = tráfego em Erlangs (volume × TMO / 3600)
- `C(m, u)` = probabilidade de espera

### Erlang A
Extensão do Erlang C com taxa de abandono (paciência):

```
P(abandono) = f(m, u, θ)
SLA_auto = (1 - P_abandon) × SLA_ErlangC
```

Onde `θ = 1 / patience_time` é a taxa de abandono.

### Otimização Híbrida

#### Fase 1: MILP (HiGHS)
Minimiza o número total de agentes sujeito a:
```
Min Σ n_s

Sujeito a:
  Σ_{s covers i} n_s × pl ≥ hc_liq[i]   ∀ i
  Σ y_s ≤ max_horarios
  n_s - M×y_s ≤ 0
  n_s ∈ ℤ₊, y_s ∈ {0,1}
```

Configurações:
- `time_limit`: 5 segundos
- `mip_rel_gap`: 2% (aceita solução dentro de 2% do ótimo)

#### Fase 2: Heurístico "Fatias" (Fallback)
Quando MILP não converge:
1. **Set-Cover**: Seleciona slots mínimos para cobertura completa
2. **Score-based**: Completa com slots de maior impacto (volume × cobertura)
3. **Alocação Incremental**: Adiciona 1 agente por vez até cobrir gap
4. **Post-Trim**: Remove folga mantendo SLA ≥ target + margem

#### Post-Trim
Remove agentes excedentes verificando:
- SLA ponderado (30min) ≥ target + 0.5pp
- Nenhum intervalo individual < target - 5pp
- Cobertura física mínima (se configurado)
- Sem buracos de cobertura (zero-gap protection)

---

## 📁 Estrutura do Projeto

```
wfm-engine/
├── __init__.py          # Pacote wfm, exports principais
├── app.py               # API Flask e rotas HTTP
├── models.py            # Dataclasses de input/output
├── engine.py            # Orchestrator principal
├── solver.py            # Algoritmos de otimização (MILP/heur)
├── erlang.py            # Fórmulas Erlang C e Erlang A
├── demand.py            # Geração e parse de curvas de demanda
├── excel_export.py      # Exportação de resultados para Excel
├── interface.html       # Frontend web (HTML/CSS/JS)
├── requirements.txt     # Dependências Python
├── INICIAR.bat          # Script de inicialização (Windows)
└── README.md            # Este arquivo
```

### Módulos Principais

| Módulo | Responsabilidade |
|--------|------------------|
| `models.py` | Definição de tipos (WFMInput, WFMOutput, Turno, Alerta, etc.) |
| `engine.py` | Orquestração do fluxo: demanda → Erlang → solver → KPIs |
| `solver.py` | Algoritmos de dimensionamento (hybrid_solve, heuristic_slices) |
| `erlang.py` | Cálculos de filas (SLA, TME, ocupação, tráfego) |
| `demand.py` | Curvas de demanda, templates Excel, detecção de período |
| `excel_export.py` | Geração de relatórios formatados em Excel |
| `app.py` | Servidor web, endpoints REST, interface HTML |

---

## ⚙️ Configurações

### Parâmetros de Shrinkage

| Parâmetro | Padrão | Descrição |
|-----------|--------|-----------|
| `absenteismo` | 5% | Ausências não planejadas |
| `ferias` | 5% | Férias coletivas/individuais |
| `treinamento` | 3% | Capacitações e onboarding |
| `reunioes` | 2% | Reuniões de equipe/1:1 |
| `aderencia` | 3% | Desvios de schedule (feedback) |
| `outras` | 0% | Outros fatores (customizável) |

**Total máximo:** 70% (limitado no código)

### Tipos de Turno

| Tipo | Duração | Intervalo | PL Base | Janela Entrada |
|------|---------|-----------|---------|----------------|
| `6:20` | 6h20 (380min) | 40min | 0.8947 | 0-106 slots |
| `8:12` | 8h12 (492min) | 80min | 0.8374 | 0-94 slots |

*PL = Productive Load (carga produtiva líquida)*

### Modos de Operação

#### SLA Mode
- **`weighted`**: SLA ponderado pelo volume (padrão, recomendado)
- **`daily`**: SLA mínimo por dia (mais conservador)

#### Solver Mode
- **`heuristic`**: Híbrido MILP + fallback heurístico (rápido, <1s)
- **`exact`**: MILP puro com PuLP (lento, garantidamente ótimo)

#### Erlang Mode
- **`erlang_c`**: Modelo clássico (sem abandono)
- **`erlang_a`**: Modelo com abandono (requer `patience_time`)

---

## 📝 Exemplos

### Exemplo 1: Call Center Comercial (Seg-Sex, 8h-20h)

```python
inp = WFMInput(
    volume_mes=200000,
    tmo_base=180,
    sla_target=0.85,
    tempo_target=15,
    mes=6,
    ano=2025,
    horario_abertura="08:00",
    horario_fechamento="20:00",
    vol_sabado_pct=0.0,    # Não opera sábado
    vol_domingo_pct=0.0,   # Não opera domingo
    dias_funcionamento=["seg","ter","qua","qui","sex"],
    pausas=PausasAdicionais(
        absenteismo=0.04,
        ferias=0.06,
        treinamento=0.02,
    ),
)
```

### Exemplo 2: Suporte 24x7 com Erlang A

```python
inp = WFMInput(
    volume_mes=500000,
    tmo_base=300,
    sla_target=0.75,
    tempo_target=30,
    erlang_mode="erlang_a",
    patience_time=180.0,   # 3 minutos de paciência média
    solver_mode="heuristic",
    max_horarios_entrada=10,
    min_agentes_intervalo=3,  # Mínimo de 3 agentes por intervalo
)
```

### Exemplo 3: Upload de Curva Personalizada

```python
import requests

url = "http://localhost:5000/calcular"
files = {"curvas_excel": open("curvas_junho.xlsx", "rb")}
data = {
    "volume_mes": 180000,
    "tmo_base": 220,
    "sla_target": 0.80,
    "tempo_target": 20,
}

response = requests.post(url, files=files, data=data)
resultado = response.json()
```

---

## 🤝 Contribuição

Contribuições são bem-vindas! Para contribuir:

1. **Fork** o repositório
2. Crie uma branch para sua feature (`git checkout -b feature/nova-feature`)
3. Commit suas mudanças (`git commit -am 'Adiciona nova feature'`)
4. Push para a branch (`git push origin feature/nova-feature`)
5. Abra um **Pull Request**

### Guidelines de Desenvolvimento

- Siga o estilo PEP 8
- Adicione type hints em novas funções
- Escreva docstrings no formato Google
- Inclua testes para novas funcionalidades
- Mantenha a compatibilidade com Python 3.8+

---

## 📄 Licença

Este projeto está licenciado sob a licença **MIT**. Veja o arquivo [LICENSE](LICENSE) para detalhes.

---

## 📞 Suporte

Para dúvidas, bugs ou sugestões:

- **Issues**: Abra uma issue no GitHub
- **Email**: seu-email@exemplo.com
- **Documentação**: Consulte este README e os docstrings dos módulos

---

## 🙏 Agradecimentos

- **SciPy Team**: Pelo solver HiGHS integrado
- **Flask Team**: Pela simplicidade e robustez do framework
- **Comunidade Python**: Pelas bibliotecas NumPy e OpenPyXL

---

<div align="center">

**WFM Engine** — Dimensionamento inteligente para operações de excelência

Made with ❤️ using Python & Flask

</div>
