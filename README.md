# WFM Engine - Workforce Management

Sistema de dimensionamento de equipes (Workforce Management) baseado nas fórmulas de Erlang C e Erlang A.

## Estrutura do Projeto

```
wfm-engine/
├── src/                    # Código fonte principal
│   ├── app.py             # API Flask
│   └── wfm/               # Pacote principal WFM
│       ├── __init__.py
│       ├── models.py      # Modelos de dados (dataclasses)
│       ├── erlang.py      # Fórmulas Erlang C e Erlang A
│       ├── demand.py      # Distribuição de demanda e curvas
│       ├── solver.py      # Otimizador de escalas (MILP/Heurístico)
│       ├── engine.py      # Motor principal de cálculo
│       └── excel_export.py # Exportação e validação Excel
├── frontend/
│   └── index.html         # Interface web
├── tests/                 # Testes unitários
├── docs/                  # Documentação
├── requirements.txt       # Dependências Python
└── README.md             # Este arquivo
```

## Instalação

```bash
pip install -r requirements.txt
```

### Dependências

- **numpy** >= 1.24: Computação numérica
- **scipy** >= 1.10: Otimização MILP (HiGHS solver)
- **flask**: API web
- **openpyxl**: Manipulação de arquivos Excel
- **pulp** (opcional): Solver alternativo para modo exato

## Uso

### Via linha de comando

```bash
python src/app.py --port 5000 --host 0.0.0.0
```

Opções:
- `--port`: Porta do servidor (padrão: 5000)
- `--host`: Host para bind (padrão: 0.0.0.0)
- `--debug`: Ativa modo debug do Flask

### Acessando a aplicação

Após iniciar o servidor, acesse no navegador:
```
http://localhost:5000
```

## Funcionalidades

### Modelos Erlang

- **Erlang C (M/M/c)**: Modelo clássico de filas sem abandono
- **Erlang A (M/M/c+M)**: Modelo com taxa de abandono (patience time)

### Métricas Calculadas

- SLA (Service Level Agreement)
- TME (Tempo Médio de Espera)
- Ocupação dos agentes
- Probabilidade de fila (Pw)
- Taxa de abandono (Erlang A)

### Solver de Escalas

- **Modo Heurístico**: Alocação por "fatias" distribuídas
- **Modo MILP**: Otimização exata usando HiGHS via scipy
- **Modo Híbrido**: Tenta MILP primeiro, fallback para heurístico

### Recursos

- Curvas de demanda intraday personalizáveis (Excel)
- Múltiplos tipos de turno (6x20, 8x12)
- Shrinkage configurável (absenteísmo, férias, treinamento, etc.)
- Operação 24h com wrap-around de turnos
- Horários de abertura/fechamento customizáveis
- Floor mínimo de agentes por intervalo
- Controle de overstaffing máximo

## API Endpoints

| Endpoint | Método | Descrição |
|----------|--------|-----------|
| `/` | GET | Interface web |
| `/calcular` | POST | Executa dimensionamento |
| `/validar` | POST | Valida arquivo Excel |
| `/templates` | GET | Baixa template Excel |
| `/exportar` | POST | Exporta resultado em Excel |
| `/detectar_periodo` | POST | Detecta mês/ano do Excel |

## Exemplo de Uso da API

```python
import requests

payload = {
    "volume_mes": 150000,
    "tmo_base": 240,
    "sla_target": 0.80,
    "tempo_target": 20,
    "mes": 5,
    "ano": 2025,
    "solver_mode": "heuristic",
    "erlang_mode": "erlang_c"
}

response = requests.post("http://localhost:5000/calcular", json=payload)
resultado = response.json()
```

## Configuração de Parâmetros

### Parâmetros Principais

| Parâmetro | Tipo | Padrão | Descrição |
|-----------|------|--------|-----------|
| volume_mes | int | 150000 | Volume mensal de contatos |
| tmo_base | float | 240 | Tempo Médio de Operação (segundos) |
| sla_target | float | 0.80 | Meta de SLA (80%) |
| tempo_target | float | 20 | Tempo alvo de resposta (segundos) |
| max_overstaffing | float | 0.20 | Overstaffing máximo permitido (20%) |
| solver_mode | str | "heuristic" | Modo do solver ("heuristic", "exact", "hybrid") |
| erlang_mode | str | "erlang_c" | Modelo Erlang ("erlang_c", "erlang_a") |
| patience_time | float | 300 | Tempo de paciência em segundos (Erlang A) |

### Shrinkage (Pausas)

| Parâmetro | Padrão | Descrição |
|-----------|--------|-----------|
| absenteismo | 0.05 | Taxa de absenteísmo (5%) |
| ferias | 0.05 | Férias (5%) |
| treinamento | 0.03 | Treinamentos (3%) |
| reunioes | 0.02 | Reuniões (2%) |
| aderencia | 0.03 | Aderência/feedback (3%) |
| outras | 0.00 | Outras pausas (0%) |

## Arquitetura Técnica

### Camadas

1. **API Layer** (`app.py`): endpoints REST e interface web
2. **Engine Layer** (`engine.py`): orquestração do cálculo
3. **Solver Layer** (`solver.py`): otimização de escalas
4. **Math Layer** (`erlang.py`): fórmulas matemáticas
5. **Data Layer** (`models.py`): estruturas de dados
6. **IO Layer** (`demand.py`, `excel_export.py`): entrada/saída

### Fluxo de Cálculo

```
Input → Distribuição de Demanda → Curva Erlang → Solver → 
Cobertura → Indicadores por Intervalo → Agregação Diária → KPIs Mensais
```

## Testes

```bash
pytest tests/
```

## Licença

MIT
