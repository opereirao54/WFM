# Documentação da API WFM Engine

## Visão Geral

A WFM Engine é uma API REST para dimensionamento de equipes (Workforce Management) baseada nas fórmulas de Erlang C e Erlang A.

## Endpoints

### GET `/`

Retorna a interface web da aplicação.

**Resposta:** HTML

---

### POST `/calcular`

Executa o dimensionamento de equipes com base nos parâmetros fornecidos.

**Content-Type:** `application/json` ou `multipart/form-data`

#### Parâmetros Principais

| Campo | Tipo | Obrigatório | Padrão | Descrição |
|-------|------|-------------|--------|-----------|
| `volume_mes` | int | Não | 150000 | Volume mensal de contatos |
| `tmo_base` | float | Não | 240 | Tempo Médio de Operação (segundos) |
| `sla_target` | float | Não | 0.80 | Meta de SLA (0-1) |
| `tempo_target` | float | Não | 20 | Tempo alvo de resposta (segundos) |
| `mes` | int | Não | 5 | Mês de referência (1-12) |
| `ano` | int | Não | 2025 | Ano de referência |
| `sla_mode` | string | Não | "weighted" | Modo de cálculo do SLA ("weighted" ou "interval_floor") |
| `interval_floor_pct` | float | Não | 0.85 | Floor mínimo de intervalos OK |
| `max_overstaffing` | float | Não | 0.20 | Overstaffing máximo permitido |
| `max_horarios_entrada` | int | Não | 6 | Máximo de horários de entrada diferentes |
| `solver_mode` | string | Não | "heuristic" | Modo do solver ("heuristic", "exact", "hybrid") |
| `erlang_mode` | string | Não | "erlang_c" | Modelo Erlang ("erlang_c" ou "erlang_a") |
| `patience_time` | float | Não | 300.0 | Tempo de paciência em segundos (Erlang A) |
| `horario_abertura` | string | Não | "" | Horário de abertura (HH:MM) |
| `horario_fechamento` | string | Não | "" | Horário de fechamento (HH:MM) |
| `min_agentes` | int | Não | 0 | Mínimo de agentes por intervalo |
| `vol_sabado_pct` | float | Não | 0.35 | Volume de sábado como % do dia útil |
| `vol_domingo_pct` | float | Não | 0.15 | Volume de domingo como % do dia útil |
| `dias_funcionamento` | string | Não | "seg,ter,qua,qui,sex,sab,dom" | Dias de funcionamento |

#### Parâmetros de Shrinkage (Pausas)

| Campo | Tipo | Padrão | Descrição |
|-------|------|--------|-----------|
| `absenteismo` | float | 0.05 | Taxa de absenteísmo |
| `ferias` | float | 0.05 | Taxa de férias |
| `treinamento` | float | 0.03 | Taxa de treinamento |
| `reunioes` | float | 0.02 | Taxa de reuniões |
| `aderencia` | float | 0.03 | Taxa de aderência/feedback |
| `outras` | float | 0.00 | Outras pausas |

#### Upload de Curvas Excel

Para usar curvas personalizadas, envie via `multipart/form-data`:

```
curvas_excel: arquivo.xlsx
```

O arquivo deve conter as abas:
- `Curva_Semana`
- `Curva_Sabado`
- `Curva_Domingo`

Use `/templates` para baixar um modelo.

#### Resposta de Sucesso (200)

```json
{
  "status": "optimal",
  "mes": 5,
  "ano": 2025,
  "solver_mode": "heuristic",
  "erlang_mode": "erlang_c",
  "elapsed_sec": 0.523,
  "kpis": {
    "sla_ponderado_mes": 0.8523,
    "sla_target": 0.8,
    "overstaffing": 0.12,
    "max_overstaffing": 0.2,
    "ns_total_mes": 127500,
    "volume_total_mes": 150000,
    "intervalos_ok_pct": 0.89
  },
  "hc_fisico": {...},
  "turnos": [
    {
      "pool": "pool_812",
      "tipo": "8:12",
      "entrada": "08:00",
      "saida": "16:12",
      "agentes_bruto": 50,
      "agentes_liq": 41.9
    }
  ],
  "dias": [...],
  "alertas": [...]
}
```

#### Códigos de Status

- `optimal`: SLA atingido dentro dos limites
- `infeasible`: SLA não atingível com parâmetros atuais
- `constrained`: Restrições violadas (ex: overstaffing excedido)

#### Erros

**400 Bad Request**
```json
{"erro": "Mês inválido: 13."}
```

**500 Internal Server Error**
```json
{
  "erro": "Descrição do erro",
  "trace": "Stack trace completo"
}
```

---

### POST `/validar`

Valida um arquivo Excel de curvas antes do processamento.

**Content-Type:** `multipart/form-data`

#### Parâmetros

| Campo | Tipo | Descrição |
|-------|------|-----------|
| `curvas_excel` | file | Arquivo Excel para validar |

#### Resposta

```json
{
  "valido": true,
  "issues": []
}
```

Ou com problemas:

```json
{
  "valido": false,
  "issues": [
    {
      "sheet": "Curva_Semana",
      "row": 15,
      "col": "PESO_PCT",
      "severity": "erro",
      "message": "Valor negativo: -5.0"
    }
  ]
}
```

---

### GET `/templates`

Baixa um template Excel para preenchimento das curvas de demanda.

**Resposta:** Arquivo XLSX para download

---

### POST `/exportar`

Exporta os resultados do dimensionamento em formato Excel.

**Content-Type:** `application/json`

#### Corpo da Requisição

Envie o JSON completo retornado pelo endpoint `/calcular`.

#### Resposta

Arquivo XLSX para download com nome: `wfm_{ANO}{MES}_dimensionamento.xlsx`

---

### POST `/detectar_periodo`

Detecta automaticamente mês e ano de um arquivo Excel de curvas.

**Content-Type:** `multipart/form-data`

#### Parâmetros

| Campo | Tipo | Descrição |
|-------|------|-----------|
| `curvas_excel` | file | Arquivo Excel |

#### Resposta

```json
{
  "mes": 5,
  "ano": 2025
}
```

---

## Exemplos de Uso

### cURL

#### Calcular dimensionamento básico

```bash
curl -X POST http://localhost:5000/calcular \
  -H "Content-Type: application/json" \
  -d '{
    "volume_mes": 100000,
    "tmo_base": 240,
    "sla_target": 0.8,
    "tempo_target": 20
  }'
```

#### Com upload de curvas

```bash
curl -X POST http://localhost:5000/calcular \
  -F "volume_mes=100000" \
  -F "tmo_base=240" \
  -F "sla_target=0.8" \
  -F "tempo_target=20" \
  -F "curvas_excel=@minhas_curvas.xlsx"
```

#### Validar arquivo Excel

```bash
curl -X POST http://localhost:5000/validar \
  -F "curvas_excel=@minhas_curvas.xlsx"
```

#### Baixar template

```bash
curl -O http://localhost:5000/templates
```

### Python

```python
import requests

# Calcular
payload = {
    "volume_mes": 100000,
    "tmo_base": 240,
    "sla_target": 0.8,
    "tempo_target": 20,
    "mes": 6,
    "ano": 2025
}

response = requests.post("http://localhost:5000/calcular", json=payload)
resultado = response.json()

print(f"SLA: {resultado['kpis']['sla_ponderado_mes']*100:.1f}%")
print(f"Status: {resultado['status']}")

# Exportar
export_response = requests.post(
    "http://localhost:5000/exportar",
    json=resultado
)

with open("resultado.xlsx", "wb") as f:
    f.write(export_response.content)
```

### JavaScript (Fetch)

```javascript
const payload = {
  volume_mes: 100000,
  tmo_base: 240,
  sla_target: 0.8,
  tempo_target: 20
};

fetch('http://localhost:5000/calcular', {
  method: 'POST',
  headers: {'Content-Type': 'application/json'},
  body: JSON.stringify(payload)
})
.then(r => r.json())
.then(data => {
  console.log(`SLA: ${data.kpis.sla_ponderado_mes * 100}%`);
});
```

---

## Considerações

### Timeouts

O endpoint `/calcular` pode levar alguns segundos dependendo:
- Do volume de dados
- Do modo do solver (`exact` é mais lento que `heuristic`)
- Do número de restrições

### Limites

- Meses válidos: 1-12
- Anos válidos: 2000-2100
- SLA target: 0-1
- Shrinkage total máximo: 70%

### Performance

- **Heurístico**: < 2 segundos na maioria dos casos
- **MILP/Híbrido**: 2-30 segundos dependendo da complexidade
- **Exato**: Pode levar minutos para instâncias grandes
