import sys, os, io, json, re, time
sys.path.insert(0, os.path.dirname(__file__))

from flask import Flask, request, jsonify, send_file
from wfm.models import WFMInput, PausasAdicionais
from wfm.engine import run_engine
from wfm.demand import default_curva, get_templates_xlsx, parse_excel_upload, preview_excel_upload
from wfm.excel_export import validate_excel, export_resultado_xlsx
from wfm.forecast import get_forecast_template_xlsx, run_forecast, generate_forecast_xlsx

app = Flask(__name__)

BASE_DIR = os.path.dirname(__file__)
HTML_PATH = os.path.join(BASE_DIR, "..", "frontend", "index.html")
SAVED_DIR = os.path.abspath(os.path.join(BASE_DIR, "..", "saved"))
os.makedirs(SAVED_DIR, exist_ok=True)

_UNSAFE_NAME = re.compile(r"[\x00-\x1f/\\:*?\"<>|]+")
def _slug(s: str, fallback: str = "sem_nome") -> str:
    s = (s or "").strip()
    s = _UNSAFE_NAME.sub("", s).strip().replace(" ", "_")
    # remove leading dots e espaços
    s = s.lstrip(". ")
    return s[:60] or fallback

def _safe_saved_path(*parts: str) -> str:
    p = os.path.abspath(os.path.join(SAVED_DIR, *parts))
    if not p.startswith(SAVED_DIR + os.sep) and p != SAVED_DIR:
        raise ValueError("caminho fora do diretório de salvamentos")
    return p

with open(HTML_PATH, encoding="utf-8") as f:
    HTML = f.read()

@app.route("/")
def index(): return HTML

@app.route("/templates.xlsx")
def templates():
    return send_file(io.BytesIO(get_templates_xlsx()),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        download_name="wfm_curvas_template.xlsx", as_attachment=True)

@app.route("/validar", methods=["POST"])
def validar():
    try:
        f = request.files.get("curvas_excel")
        if not f or not f.filename:
            return jsonify({"valido": True, "issues": []})
        is_valid, issues = validate_excel(f.read())
        return jsonify({"valido": is_valid, "issues": issues})
    except Exception as e:
        return jsonify({"valido": False, "issues": [{"sheet":"—","row":0,"col":"—",
            "severity":"erro","message":str(e)}]})

@app.route("/exportar", methods=["POST"])
def exportar():
    try:
        data = request.json
        out = _rebuild_output(data)
        xlsx = export_resultado_xlsx(out)
        fname = f"wfm_{out.ano}{out.mes:02d}_dimensionamento.xlsx"
        return send_file(io.BytesIO(xlsx),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            download_name=fname, as_attachment=True)
    except Exception as e:
        import traceback
        return jsonify({"erro": str(e), "trace": traceback.format_exc()}), 500

def _rebuild_output(data):
    from wfm.models import WFMOutput, DiaOut, IntervaloOut, Turno, Alerta
    dias = []
    for d in data.get("dias",[]):
        ivs = [IntervaloOut(
            horario=iv["horario"], volume=iv["volume"], tmo=iv["tmo"],
            hc_liq=iv["hc_liq"], hc_bruto=iv["hc_bruto"],
            trafico_erl=iv["trafico_erl"], fila_pw=iv["fila_pw"]/100,
            tme_seg=iv["tme_seg"], ocupacao=iv["ocupacao"]/100,
            sla_pct=iv["sla_pct"]/100, ns=iv["ns"],
            p_abandon=iv.get("p_abandon",0)/100,
        ) for iv in d.get("intervalos",[])]
        dias.append(DiaOut(
            data=d["data"], dia_semana=d["dia_semana"], tipo=d["tipo"],
            peso_pct=d["peso_pct"], volume_total=d["volume_total"],
            tmo_medio=d["tmo_medio"], hc_liq_max=d["hc_liq_max"],
            hc_bruto_max=d["hc_bruto_max"],
            sla_ponderado=d["sla_ponderado"]/100, ns_total=d["ns_total"],
            tme_medio=d["tme_medio"], ocupacao_media=d["ocupacao_media"]/100,
            fila_media=d["fila_media"]/100,
            intervalos_ok_pct=d["intervalos_ok_pct"]/100,
            status_sla=d["status_sla"], intervalos=ivs,
        ))
    k = data.get("kpis",{})
    return WFMOutput(
        status=data.get("status","optimal"),
        mes=data.get("mes",5), ano=data.get("ano",2025),
        sla_ponderado_mes=k.get("sla_ponderado_mes",0),
        sla_target=k.get("sla_target",0.8),
        overstaffing=k.get("overstaffing",0),
        max_overstaffing=k.get("max_overstaffing",0.2),
        ns_total_mes=k.get("ns_total_mes",0),
        volume_total_mes=k.get("volume_total_mes",0),
        intervalos_ok_pct=k.get("intervalos_ok_pct"),
        interval_floor_target=k.get("interval_floor_target"),
        hc_fisico=data.get("hc_fisico",{}),
        hc_liquido_ref=data.get("hc_liquido_ref",{}),
        pl_efetivo=data.get("pl_efetivo",{}),
        turnos=[Turno(pool=t["pool"],tipo=t["tipo"],entrada=t["entrada"],
                      saida=t["saida"],agentes_bruto=t["agentes_bruto"],
                      agentes_liq=t["agentes_liq"]) for t in data.get("turnos",[])],
        dias=dias,
        alertas=[Alerta(codigo=a["codigo"],mensagem=a["mensagem"],
                        hc_adicional_necessario=a.get("hc_adicional_necessario",0))
                 for a in data.get("alertas",[])],
        solver_mode=data.get("solver_mode","heuristic"),
        erlang_mode=data.get("erlang_mode","erlang_c"),
        patience_time=data.get("patience_time",300.0),
        elapsed_sec=data.get("elapsed_sec",0),
        horario_abertura=data.get("horario_abertura",""),
        horario_fechamento=data.get("horario_fechamento",""),
        pausa_nr17_pct=data.get("pausa_nr17_pct",0),
        demanda_curves=data.get("demanda_curves",{}),
        cobertura_curves=data.get("cobertura_curves",{}),
    )

@app.route("/detectar_periodo", methods=["POST"])
def detectar_periodo():
    try:
        from wfm.demand import extract_mes_ano_from_xlsx
        import openpyxl, io
        f = request.files.get("curvas_excel")
        if not f or not f.filename:
            return jsonify({})
        wb = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True)
        result = extract_mes_ano_from_xlsx(wb)
        if result:
            return jsonify({"mes": result[0], "ano": result[1]})
        return jsonify({})
    except Exception:
        return jsonify({})

@app.route("/curva/preview", methods=["POST"])
def curva_preview():
    """Parse flexível do Excel e retorna preview com ajustes para aprovação."""
    try:
        f = request.files.get("curvas_excel")
        if not f or not f.filename:
            return jsonify({"erro": "Nenhum arquivo enviado."}), 400
        result = preview_excel_upload(f.read())
        return jsonify(result)
    except Exception as e:
        import traceback
        return jsonify({"erro": str(e), "trace": traceback.format_exc()}), 500


@app.route("/calcular", methods=["POST"])
def calcular():
    try:
        curva_sem = curva_sab = curva_dom = None
        dias_mes = []
        if request.content_type and "multipart" in request.content_type:
            d  = request.form
            fv = lambda k, v: d.get(k, v)
            f  = request.files.get("curvas_excel")
            if f and f.filename:
                curva_sem, curva_sab, curva_dom, dias_mes = parse_excel_upload(f.read())
        else:
            d  = request.json or {}
            fv = lambda k, v: d.get(k, v)

        mes_raw = int(fv("mes", 5))
        ano_raw = int(fv("ano", 2025))
        if not (1 <= mes_raw <= 12):
            return jsonify({"erro": f"Mês inválido: {mes_raw}."}), 400
        if not (2000 <= ano_raw <= 2100):
            return jsonify({"erro": f"Ano inválido: {ano_raw}."}), 400

        inp = WFMInput(
            volume_mes           = int(fv("volume_mes", 150000)),
            tmo_base             = float(fv("tmo_base", 240)),
            sla_target           = float(fv("sla_target", 0.80)),
            tempo_target         = float(fv("tempo_target", 20)),
            mes                  = mes_raw,
            ano                  = ano_raw,
            sla_mode             = fv("sla_mode", "weighted"),
            interval_floor_pct   = float(fv("interval_floor_pct", 0.85)),
            max_overstaffing     = float(fv("max_overstaffing", 0.20)),
            max_horarios_entrada = int(fv("max_horarios_entrada", 8)),
            solver_mode          = fv("solver_mode", "heuristic"),
            erlang_mode          = fv("erlang_mode", "erlang_c"),
            patience_time        = float(fv("patience_time", 300)),
            retry_rate           = float(fv("retry_rate", 0.30)),
            horario_abertura     = fv("horario_abertura", "").strip(),
            horario_fechamento   = fv("horario_fechamento", "").strip(),
            janela_entrada_inicio = fv("janela_entrada_inicio", "").strip(),
            janela_entrada_fim    = fv("janela_entrada_fim", "").strip(),
            janelas_bloqueadas    = [s.strip() for s in
                fv("janelas_bloqueadas", "").split(";") if s.strip()],
            min_agentes_intervalo = int(fv("min_agentes", 0)),
            pausas=PausasAdicionais(
                absenteismo = float(fv("absenteismo", 0.05)),
                ferias      = float(fv("ferias",      0.05)),
                treinamento = float(fv("treinamento", 0.03)),
                reunioes    = float(fv("reunioes",    0.02)),
                aderencia   = float(fv("aderencia",   0.03)),
                outras      = float(fv("outras",      0.00)),
            ),
            curva_semana  = curva_sem or default_curva("Util"),
            curva_sabado  = curva_sab or default_curva("Sabado"),
            curva_domingo = curva_dom or default_curva("Domingo"),
            dias_mes      = dias_mes or [],
            vol_sabado_pct  = float(fv("vol_sabado_pct", 0.35)),
            vol_domingo_pct = float(fv("vol_domingo_pct", 0.15)),
            dias_funcionamento = [s.strip().lower() for s in
                fv("dias_funcionamento", "seg,ter,qua,qui,sex,sab,dom").split(",")
                if s.strip()],
            allow_shift_620 = str(fv("allow_shift_620","true")).lower() in ("1","true","on","yes"),
            allow_shift_812 = str(fv("allow_shift_812","true")).lower() in ("1","true","on","yes"),
        )

        out = run_engine(inp)

        return jsonify({
            "status":             out.status,
            "mes":                out.mes,
            "ano":                out.ano,
            "solver_mode":        out.solver_mode,
            "erlang_mode":        out.erlang_mode,
            "patience_time":      out.patience_time,
            "elapsed_sec":        round(out.elapsed_sec, 3),
            "horario_abertura":   out.horario_abertura,
            "horario_fechamento": out.horario_fechamento,
            "kpis": {
                "sla_ponderado_mes":     round(out.sla_ponderado_mes, 4),
                "sla_target":            out.sla_target,
                "overstaffing":          round(out.overstaffing, 4),
                "max_overstaffing":      out.max_overstaffing,
                "ns_total_mes":          out.ns_total_mes,
                "volume_total_mes":      out.volume_total_mes,
                "intervalos_ok_pct":     round(out.intervalos_ok_pct,4) if out.intervalos_ok_pct else None,
                "interval_floor_target": out.interval_floor_target,
            },
            "hc_fisico":     out.hc_fisico,
            "hc_liquido_ref":out.hc_liquido_ref,
            "pl_efetivo":    out.pl_efetivo,
            "turnos": [{"pool":t.pool,"tipo":t.tipo,"entrada":t.entrada,"saida":t.saida,
                        "agentes_bruto":t.agentes_bruto,"agentes_liq":t.agentes_liq}
                       for t in out.turnos],
            "dias": [{
                "data":d.data,"dia_semana":d.dia_semana,"tipo":d.tipo,
                "peso_pct":d.peso_pct,"volume_total":d.volume_total,
                "tmo_medio":d.tmo_medio,"hc_liq_max":d.hc_liq_max,
                "hc_bruto_max":d.hc_bruto_max,
                "sla_ponderado":round(d.sla_ponderado*100,2),
                "ns_total":d.ns_total,"tme_medio":d.tme_medio,
                "ocupacao_media":round(d.ocupacao_media*100,2),
                "fila_media":round(d.fila_media*100,2),
                "intervalos_ok_pct":round(d.intervalos_ok_pct*100,2),
                "status_sla":d.status_sla,
                "intervalos":[{
                    "horario":iv.horario,"volume":iv.volume,"tmo":iv.tmo,
                    "hc_liq":iv.hc_liq,"hc_bruto":iv.hc_bruto,
                    "trafico_erl":iv.trafico_erl,
                    "fila_pw":round(iv.fila_pw*100,2),
                    "tme_seg":iv.tme_seg,
                    "ocupacao":round(iv.ocupacao*100,2),
                    "sla_pct":round(iv.sla_pct*100,2),
                    "ns":iv.ns,
                    "p_abandon":round(iv.p_abandon*100,2),
                } for iv in d.intervalos],
            } for d in out.dias],
            "alertas":[{"codigo":a.codigo,"mensagem":a.mensagem,
                        "hc_adicional_necessario":a.hc_adicional_necessario}
                       for a in out.alertas],
            "pausa_nr17_pct":   out.pausa_nr17_pct,
            "demanda_curves":   out.demanda_curves,
            "cobertura_curves": out.cobertura_curves,
        })
    except Exception as e:
        import traceback
        return jsonify({"erro": str(e), "trace": traceback.format_exc()}), 500

# ── Dimensionamentos salvos ───────────────────────────────────────────
@app.route("/saved", methods=["GET"])
def saved_list():
    """Lista a árvore de dimensionamentos salvos: operação → ano_mes → registros."""
    tree = {}
    try:
        for op in sorted(os.listdir(SAVED_DIR)):
            op_dir = _safe_saved_path(op)
            if not os.path.isdir(op_dir): continue
            tree[op] = {}
            for ym in sorted(os.listdir(op_dir)):
                ym_dir = _safe_saved_path(op, ym)
                if not os.path.isdir(ym_dir): continue
                regs = []
                for fn in sorted(os.listdir(ym_dir)):
                    if not fn.endswith(".json"): continue
                    fp = _safe_saved_path(op, ym, fn)
                    try:
                        st = os.stat(fp)
                        regs.append({
                            "nome": fn[:-5],
                            "rel_path": f"{op}/{ym}/{fn}",
                            "modified_at": int(st.st_mtime),
                            "size": st.st_size,
                        })
                    except OSError:
                        pass
                tree[op][ym] = regs
    except Exception as e:
        return jsonify({"erro": str(e)}), 500
    return jsonify({"tree": tree})


@app.route("/saved", methods=["POST"])
def saved_save():
    """Salva um dimensionamento em saved/<operacao>/<ano_mes>/<nome>.json"""
    try:
        data = request.get_json(force=True) or {}
        operacao = _slug(data.get("operacao"), "operacao")
        ano_mes  = _slug(data.get("ano_mes"), "000000")
        if not re.fullmatch(r"\d{6}", ano_mes):
            return jsonify({"erro": "ano_mes deve estar no formato AAAAMM (ex.: 202608)"}), 400
        nome = _slug(data.get("nome") or f"dim_{int(time.time())}", "dim")
        payload = data.get("payload")
        if not isinstance(payload, dict):
            return jsonify({"erro": "payload inválido"}), 400

        op_dir = _safe_saved_path(operacao, ano_mes)
        os.makedirs(op_dir, exist_ok=True)
        fp = _safe_saved_path(operacao, ano_mes, nome + ".json")

        to_write = {
            "operacao": operacao,
            "ano_mes": ano_mes,
            "nome": nome,
            "saved_at": int(time.time()),
            "fields": payload.get("fields") or {},
            "result": payload.get("result") or {},
        }
        with open(fp, "w", encoding="utf-8") as f:
            json.dump(to_write, f, ensure_ascii=False)
        return jsonify({"ok": True, "rel_path": f"{operacao}/{ano_mes}/{nome}.json"})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500


@app.route("/saved/load", methods=["GET"])
def saved_load():
    rel = request.args.get("path", "")
    try:
        parts = [p for p in rel.split("/") if p]
        if len(parts) != 3 or not parts[2].endswith(".json"):
            return jsonify({"erro": "caminho inválido"}), 400
        fp = _safe_saved_path(*parts)
        if not os.path.isfile(fp):
            return jsonify({"erro": "não encontrado"}), 404
        with open(fp, "r", encoding="utf-8") as f:
            return jsonify(json.load(f))
    except Exception as e:
        return jsonify({"erro": str(e)}), 500


@app.route("/saved", methods=["DELETE"])
def saved_delete():
    rel = request.args.get("path", "")
    try:
        parts = [p for p in rel.split("/") if p]
        if len(parts) != 3 or not parts[2].endswith(".json"):
            return jsonify({"erro": "caminho inválido"}), 400
        fp = _safe_saved_path(*parts)
        if os.path.isfile(fp):
            os.remove(fp)
        # cleanup empty dirs
        ym_dir = _safe_saved_path(parts[0], parts[1])
        if os.path.isdir(ym_dir) and not os.listdir(ym_dir):
            os.rmdir(ym_dir)
        op_dir = _safe_saved_path(parts[0])
        if os.path.isdir(op_dir) and not os.listdir(op_dir):
            os.rmdir(op_dir)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500


# ── Forecast endpoints ────────────────────────────────────────────────
@app.route("/forecast/template.xlsx")
def forecast_template():
    return send_file(io.BytesIO(get_forecast_template_xlsx()),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        download_name="wfm_forecast_template.xlsx", as_attachment=True)

@app.route("/forecast/calcular", methods=["POST"])
def forecast_calcular():
    try:
        f = request.files.get("historico_excel")
        if not f or not f.filename:
            return jsonify({"erro": "Nenhum arquivo de histórico enviado."}), 400
        mes = int(request.form.get("mes", 5))
        ano = int(request.form.get("ano", 2025))
        if not (1 <= mes <= 12):
            return jsonify({"erro": f"Mês inválido: {mes}"}), 400
        result = run_forecast(f.read(), ano, mes)
        # Remove internal keys before sending
        out = {k: v for k, v in result.items() if not k.startswith("_")}
        return jsonify(out)
    except Exception as e:
        import traceback
        return jsonify({"erro": str(e), "trace": traceback.format_exc()}), 500

@app.route("/forecast/exportar", methods=["POST"])
def forecast_exportar():
    try:
        f = request.files.get("historico_excel")
        if not f or not f.filename:
            return jsonify({"erro": "Nenhum arquivo de histórico enviado."}), 400
        mes = int(request.form.get("mes", 5))
        ano = int(request.form.get("ano", 2025))
        result = run_forecast(f.read(), ano, mes)
        xlsx = generate_forecast_xlsx(
            result["_curvas"]["sem"], result["_curvas"]["sab"], result["_curvas"]["dom"],
            result["_dias"], result["volume_mensal"], result["tmo_mensal"],
            result["_stats"]["sem"], result["_stats"]["sab"], result["_stats"]["dom"],
        )
        fname = f"wfm_forecast_{ano}{mes:02d}.xlsx"
        return send_file(io.BytesIO(xlsx),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            download_name=fname, as_attachment=True)
    except Exception as e:
        import traceback
        return jsonify({"erro": str(e), "trace": traceback.format_exc()}), 500


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="WFM Engine API Server")
    parser.add_argument("--port", type=int, default=5000, help="Port to run the server")
    parser.add_argument("--host", type=str, default="0.0.0.0", help="Host to bind the server")
    parser.add_argument("--debug", action="store_true", help="Run in debug mode")
    args = parser.parse_args()
    
    print("\n" + "="*50)
    print(f"  WFM ENGINE  ->  http://{args.host}:{args.port}")
    print("="*50 + "\n")
    app.run(debug=args.debug, host=args.host, port=args.port)
