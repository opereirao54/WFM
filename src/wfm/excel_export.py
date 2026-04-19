"""
Excel export and validation for WFM Engine.
"""
import io
from typing import Tuple, List
from .models import WFMOutput

# ── Validation ────────────────────────────────────────────────────────────
def validate_excel(file_bytes: bytes) -> Tuple[bool, List[dict]]:
    """
    Validate uploaded Excel.
    Returns (is_valid, list of issues).
    Each issue: {sheet, row, col, severity, message}
    """
    import openpyxl
    issues = []
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        return False, [{"sheet":"—","row":0,"col":"—","severity":"erro",
                        "message":f"Arquivo inválido ou corrompido: {e}"}]

    expected_sheets = ["Curva_Semana","Curva_Sabado","Curva_Domingo"]
    for sheet in expected_sheets:
        if sheet not in wb.sheetnames:
            issues.append({"sheet":sheet,"row":0,"col":"—","severity":"aviso",
                           "message":f"Aba '{sheet}' não encontrada — será usada curva padrão"})
            continue
        ws = wb[sheet]
        rows_data = []
        for row in ws.iter_rows(min_row=3, values_only=True):
            if row[0] is None: continue
            rows_data.append(row)

        if len(rows_data) != 48:
            issues.append({"sheet":sheet,"row":0,"col":"—","severity":"erro",
                           "message":f"Deve ter 48 linhas de dados, encontrou {len(rows_data)}"})
            continue

        pesos = []
        for i, row in enumerate(rows_data, 3):
            # peso_pct
            try:
                p = float(str(row[1]).replace(',','.')) if row[1] is not None else None
                if p is None:
                    issues.append({"sheet":sheet,"row":i,"col":"PESO_PCT","severity":"erro",
                                   "message":"Célula vazia"})
                elif p < 0:
                    issues.append({"sheet":sheet,"row":i,"col":"PESO_PCT","severity":"erro",
                                   "message":f"Valor negativo: {p}"})
                else:
                    pesos.append(p)
            except ValueError:
                issues.append({"sheet":sheet,"row":i,"col":"PESO_PCT","severity":"erro",
                               "message":f"Valor não numérico: '{row[1]}'"})
            # fator_tmo
            try:
                f = float(str(row[2]).replace(',','.')) if row[2] is not None else 1.0
                if f < 0.3 or f > 5.0:
                    issues.append({"sheet":sheet,"row":i,"col":"FATOR_TMO","severity":"aviso",
                                   "message":f"Valor fora do range esperado (0.3–5.0): {f}"})
            except ValueError:
                issues.append({"sheet":sheet,"row":i,"col":"FATOR_TMO","severity":"aviso",
                               "message":f"Valor não numérico: '{row[2]}' — usando 1.0"})

        if pesos:
            soma = sum(pesos)
            if soma < 98 or soma > 102:
                issues.append({"sheet":sheet,"row":0,"col":"PESO_PCT","severity":"erro",
                               "message":f"Soma dos pesos = {soma:.2f}% (deve ser 100%)"})
            elif soma < 99 or soma > 101:
                issues.append({"sheet":sheet,"row":0,"col":"PESO_PCT","severity":"aviso",
                               "message":f"Soma dos pesos = {soma:.2f}% (idealmente exatamente 100%)"})

    # Curva_Dias (optional)
    if "Curva_Dias" in wb.sheetnames:
        ws = wb["Curva_Dias"]
        tipos_validos = {"Util","Sabado","Domingo"}
        pesos_dias = []
        for i, row in enumerate(ws.iter_rows(min_row=3, values_only=True), 3):
            if row[0] is None: continue
            if str(row[1]).strip() not in tipos_validos:
                issues.append({"sheet":"Curva_Dias","row":i,"col":"TIPO","severity":"erro",
                               "message":f"Tipo inválido: '{row[1]}' — use Util, Sabado ou Domingo"})
            try:
                p = float(str(row[2]).replace(',','.')) if row[2] is not None else None
                if p is not None: pesos_dias.append(p)
            except ValueError:
                issues.append({"sheet":"Curva_Dias","row":i,"col":"PESO_PCT","severity":"erro",
                               "message":f"Valor não numérico: '{row[2]}'"})
        if pesos_dias:
            soma = sum(pesos_dias)
            if soma < 98 or soma > 102:
                issues.append({"sheet":"Curva_Dias","row":0,"col":"PESO_PCT","severity":"erro",
                               "message":f"Soma dos pesos dos dias = {soma:.2f}% (deve ser 100%)"})

    has_error = any(i["severity"]=="erro" for i in issues)
    return not has_error, issues


# ── Export ────────────────────────────────────────────────────────────────
def export_resultado_xlsx(out: WFMOutput) -> bytes:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()

    # Styles
    h_fill  = PatternFill("solid", fgColor="0e1117")
    h_font  = Font(bold=True, color="4f9cf9", size=10)
    ok_font = Font(color="a6e3a1", size=10)
    bad_font= Font(color="f38ba8", size=10)
    mut_font= Font(color="6c7a96", size=10)
    dat_font= Font(color="cdd6f4", size=10)
    tot_fill= PatternFill("solid", fgColor="1a2030")
    tot_font= Font(bold=True, color="f9e2af", size=10)
    bd = Side(style="thin", color="1e2535")
    bdr = Border(left=bd, right=bd, top=bd, bottom=bd)
    ctr = Alignment(horizontal="center")

    def hdr(cell, text):
        cell.value=text; cell.font=h_font; cell.fill=h_fill
        cell.alignment=ctr; cell.border=bdr

    def dat(cell, value, color=None):
        cell.value=value
        cell.font=Font(color=color or "cdd6f4", size=10)
        cell.alignment=ctr; cell.border=bdr

    # ── Sheet 1: Resumo Mensal ────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Resumo Mensal"
    ws1.sheet_view.showGridLines = False

    # Header info
    ws1['A1'] = f"WFM Engine — Dimensionamento {out.mes:02d}/{out.ano}"
    ws1['A1'].font = Font(bold=True, color="4f9cf9", size=12)
    ws1['A2'] = f"SLA Mensal: {out.sla_ponderado_mes*100:.1f}%  |  HC Bruto Total: {out.hc_fisico['total']}  |  Volume: {out.volume_total_mes:,.0f}  |  NS: {out.ns_total_mes:,.0f}"
    ws1['A2'].font = Font(color="f9e2af", size=10)

    # Pool summary
    ws1['A4'] = "POOL"; ws1['B4'] = "TIPO"; ws1['C4'] = "HC BRUTO"; ws1['D4'] = "HC LÍQUIDO REF"
    ws1['E4'] = "PL BASE"; ws1['F4'] = "PL EFETIVO"
    for col in "ABCDEF":
        hdr(ws1[f'{col}4'], ws1[f'{col}4'].value)
    rows_pool = [
        ("Pool 8:12","8:12",out.hc_fisico.get("pool_812",0),out.hc_liquido_ref.get("pool_812",0),out.pl_efetivo.get("8:12_base","—"),out.pl_efetivo.get("8:12","—")),
        ("Pool B-Sáb","6:20",out.hc_fisico.get("pool_b_sab",0),out.hc_liquido_ref.get("pool_b_sab",0),out.pl_efetivo.get("6:20_base","—"),out.pl_efetivo.get("6:20","—")),
        ("Pool B-Dom","6:20",out.hc_fisico.get("pool_b_dom",0),out.hc_liquido_ref.get("pool_b_dom",0),out.pl_efetivo.get("6:20_base","—"),out.pl_efetivo.get("6:20","—")),
        ("TOTAL","—",out.hc_fisico.get("total",0),out.hc_liquido_ref.get("total",0),"—","—"),
    ]
    for i, r in enumerate(rows_pool, 5):
        for j, v in enumerate(r, 1):
            c = ws1.cell(row=i, column=j, value=v)
            c.font = tot_font if r[0]=="TOTAL" else dat_font
            c.fill = tot_fill if r[0]=="TOTAL" else PatternFill("solid", fgColor="0e1117")
            c.alignment = ctr; c.border = bdr

    # Day summary table
    row_start = 11
    cols_day = ["DATA","DIA","TIPO","VOLUME","TMO(s)","HC LÍQ","HC BRUTO","SLA%","NS","TME(s)","OCUPAÇÃO%","FILA%","INTERVALOS OK%","STATUS"]
    for j, col in enumerate(cols_day, 1):
        hdr(ws1.cell(row=row_start, column=j), col)

    for i, d in enumerate(out.dias, row_start+1):
        vals = [d.data, d.dia_semana, d.tipo, d.volume_total, d.tmo_medio,
                d.hc_liq_max, d.hc_bruto_max,
                round(d.sla_ponderado*100,1), d.ns_total, d.tme_medio,
                round(d.ocupacao_media*100,1), round(d.fila_media*100,1),
                round(d.intervalos_ok_pct*100,1),
                "✓ OK" if d.status_sla=="ok" else "✗ Abaixo"]
        for j, v in enumerate(vals, 1):
            c = ws1.cell(row=i, column=j, value=v)
            c.alignment = ctr; c.border = bdr
            if j == 14:
                c.font = Font(color="a6e3a1" if d.status_sla=="ok" else "f38ba8", size=10)
            elif j == 8:
                c.font = Font(color="a6e3a1" if d.sla_ponderado >= out.sla_target else "f38ba8", size=10)
            else:
                c.font = dat_font

    # Total row
    tot_row = row_start + len(out.dias) + 1
    totals = ["TOTAL MÊS","—","—",
              out.volume_total_mes, "—", "—", "—",
              round(out.sla_ponderado_mes*100,1), out.ns_total_mes, "—","—","—",
              round((out.intervalos_ok_pct or 0)*100,1) if out.intervalos_ok_pct else "—",
              "✓ OK" if out.sla_ponderado_mes >= out.sla_target else "✗ Abaixo"]
    for j, v in enumerate(totals, 1):
        c = ws1.cell(row=tot_row, column=j, value=v)
        c.font = tot_font; c.fill = tot_fill; c.alignment = ctr; c.border = bdr

    # Auto width
    for col in ws1.columns:
        max_w = max((len(str(cell.value or "")) for cell in col), default=8)
        ws1.column_dimensions[get_column_letter(col[0].column)].width = min(max_w + 4, 30)

    # ── Sheet 2: Turnos ───────────────────────────────────────────────
    ws2 = wb.create_sheet("Turnos")
    ws2.sheet_view.showGridLines = False
    for j, h in enumerate(["POOL","TIPO","ENTRADA","SAÍDA","HC BRUTO","HC LÍQUIDO REF"], 1):
        hdr(ws2.cell(row=1, column=j), h)
    for i, t in enumerate(out.turnos, 2):
        for j, v in enumerate([t.pool,t.tipo,t.entrada,t.saida,t.agentes_bruto,t.agentes_liq], 1):
            c = ws2.cell(row=i, column=j, value=v); c.font=dat_font; c.alignment=ctr; c.border=bdr
    for col in ws2.columns:
        ws2.column_dimensions[get_column_letter(col[0].column)].width = 16

    # ── Sheet 3+: Detail per day ──────────────────────────────────────
    cols_iv = ["HORÁRIO","VOLUME","TMO(s)","HC LÍQUIDO","HC BRUTO","TRÁFEGO(Erl)","FILA(Pw%)","TME(s)","OCUPAÇÃO%","SLA%","NS","STATUS"]
    for dia in out.dias:
        if not dia.intervalos: continue
        ws = wb.create_sheet(f"{dia.data}")
        ws.sheet_view.showGridLines = False
        ws['A1'] = f"{dia.data} {dia.dia_semana} — Vol: {dia.volume_total:,.0f}  SLA: {dia.sla_ponderado*100:.1f}%  NS: {dia.ns_total:,.0f}"
        ws['A1'].font = Font(bold=True, color="4f9cf9", size=11)
        for j, h in enumerate(cols_iv, 1):
            hdr(ws.cell(row=2, column=j), h)
        for i, iv in enumerate(dia.intervalos, 3):
            ok = iv.sla_pct >= out.sla_target
            vals = [iv.horario, iv.volume, iv.tmo, iv.hc_liq, iv.hc_bruto,
                    iv.trafico_erl, round(iv.fila_pw*100,2), iv.tme_seg,
                    round(iv.ocupacao*100,2), round(iv.sla_pct*100,2), iv.ns,
                    "✓" if ok else "✗"]
            for j, v in enumerate(vals, 1):
                c = ws.cell(row=i, column=j, value=v)
                c.font = Font(color=("a6e3a1" if ok else "f38ba8") if j in (10,12) else "cdd6f4", size=10)
                c.alignment = ctr; c.border = bdr
        for col in ws.columns:
            ws.column_dimensions[get_column_letter(col[0].column)].width = 14

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
