"""Explorador local de pipelines. Usa el motor de CoverageLab y la biblioteca estándar."""
from __future__ import annotations

import argparse
import contextlib
import hashlib
import io
import json
import math
import mimetypes
import os
from pathlib import Path
import secrets
import shutil
import subprocess
import sys
import threading
from dataclasses import asdict
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, quote, urlsplit
import webbrowser

ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "explorer_web"
TEMPLATES = {"Absoluta", "relativa", "3", "4", "5", "6", "7", "8"}


class ExplorerHTTPServer(ThreadingHTTPServer):
    # En Windows HTTPServer permite dos listeners en el mismo puerto si se reutiliza.
    allow_reuse_address = False


def engine():
    import coverage_studio as core
    core.wait_for_heavy_modules()
    if not hasattr(core, "pd") or not hasattr(core, "Presentation"):
        raise RuntimeError("Faltan dependencias. Ejecuta: pip install -r requirements.txt")
    return core


def finite(value):
    try:
        number = float(value)
        return number if math.isfinite(number) else None
    except (TypeError, ValueError):
        return None


def fingerprint(path):
    with path.open("rb") as source:
        return hashlib.file_digest(source, "sha256").hexdigest()


def evolution_variation(series):
    previous = series.shift(12)
    # Las columnas de evolución del template usan 1 si la suma anterior es cero.
    return [finite(value) for value in (series / previous.where(previous != 0, 1) - 1) * 100]


def options_from_request(payload):
    core = engine()
    template = payload.get("template", "Absoluta")
    if template not in TEMPLATES:
        raise ValueError("Template desconocido.")
    options = core.ExecutionOptions.from_scenario(template) or core.ExecutionOptions(
        coverage_type=template, coverage_reason="Actualización periódica por contrato",
        trend_axis="simple", trend_granularity="monthly", include_english=False,
        round_coverage=False, auto_mode=True,
    )
    axis = payload.get("trend_axis", options.trend_axis)
    if axis not in {"simple", "doble"}:
        raise ValueError("Eje de tendencia inválido.")
    options.trend_axis = axis
    for key in ("include_english", "round_coverage", "animate_trend_pipeline"):
        if key in payload:
            if type(payload[key]) is not bool:
                raise ValueError(f"Opción inválida: {key}.")
            setattr(options, key, payload[key])
    return options


def serialize_options(options):
    return {key: value for key, value in asdict(options).items()
            if key not in {"pipeline_overrides", "sheet_metadata_overrides"}}


def monthly_problem(df, core):
    periods = core.pd.to_datetime(df[core.COL_DATA]).dt.to_period("M")
    if periods.duplicated().any():
        return "Hay meses duplicados. Corrige las fechas antes de comparar o generar."
    if not periods.is_monotonic_increasing:
        return "Las fechas no están ordenadas. Ordénalas antes de comparar o generar."
    if len(periods) > 1 and any(b.ordinal - a.ordinal != 1 for a, b in zip(periods, periods[1:])):
        return "Faltan meses en la serie. El desplazamiento requiere meses consecutivos."
    return None


def template_details(df, core):
    """Valores de las tablas del template, con las mismas ventanas y reglas de ceros."""
    specs = [
        ("MAT vs. MAT anterior", 12, 12, 0),
        ("Semestre vs. semestre anterior", 6, 6, 0),
        ("Trimestre vs. trimestre anterior", 3, 3, 0),
        ("Semestre vs. mismo semestre del año anterior", 6, 12, 0),
        ("Trimestre vs. mismo trimestre del año anterior", 3, 12, 0),
        ("MAT vs. hace dos años", 12, 24, 0),
        ("Semestre vs. hace dos años", 6, 24, 0),
        ("Trimestre vs. hace dos años", 3, 24, 0),
        ("MAT del año anterior vs. su año anterior", 12, 12, 12),
    ]
    dates = [date.isoformat() for date in df[core.COL_DATA]]
    n = len(df)

    def period(window, lag, offset, pipeline):
        end = n - offset - pipeline
        start = end - window
        prev_end = end - lag
        prev_start = prev_end - window
        if prev_start < 0:
            return None
        return {"current": [dates[start], dates[end - 1]],
                "previous": [dates[prev_start], dates[prev_end - 1]]}

    def variation(column, window, lag, offset=0, pipeline=0):
        end = n - offset - pipeline
        start = end - window
        prev_end = end - lag
        prev_start = prev_end - window
        if prev_start < 0:
            return None
        current, previous = df[column].iloc[start:end], df[column].iloc[prev_start:prev_end]
        if current.isna().any() or previous.isna().any() or current.eq(0).any() or previous.eq(0).any() or previous.sum() == 0:
            return None
        return finite((current.sum() / previous.sum() - 1) * 100)

    return {
        "averages": {key: finite(value) for key, value in core.compute_averages(df).items()},
        "variations": [{"label": label, "worldpanel": variation(core.COL_SELL_OUT, window, lag, offset),
                        "pipelines": [variation(core.COL_SELL_IN, window, lag, offset, p) for p in range(7)],
                        "wp_period": period(window, lag, offset, 0),
                        "pipeline_periods": [period(window, lag, offset, p) for p in range(7)]}
                       for label, window, lag, offset in specs],
        "correlations": [{"label": label, "pipelines": [finite(core.compute_pipeline_mat_correlation(df, p, window, offset)) for p in range(7)]}
                         for label, window, offset in [("Año actual · 12 meses", 12, 0), ("Año anterior · 12 meses", 12, 12), ("Últimos 24 meses", 24, 0)]],
    }


def dashboard_details(df, core, pipelines):
    """Panel inline del Excel, disponible sin escribir ningún entregable."""
    specs = [("purchase_average", core.COL_COMPRA_MEDIA, "Compra media"),
             ("purchase_occasion", core.COL_COMPRA_OCA, "Compra por ocasión"),
             ("frequency", core.COL_FREQ, "Frecuencia"),
             ("penetration", core.COL_PENET, "Penetración"),
             ("buyers_series", core.COL_BUYERS, "Compradores")]
    kpis = []
    for key, column, label in specs:
        recent = df[column].tail(12)
        latest, average = finite(recent.iloc[-1]), finite(recent.mean())
        change = finite(latest / average - 1) if average else 0.0
        amplitude = finite((recent.max() - recent.min()) / abs(average)) if average else 0.0
        kpis.append({"key": key, "label": label, "color": core.EXCEL_DASHBOARD_KPI_COLORS[column],
                     "latest": latest, "average": average, "change": change, "range": amplitude,
                     "minimum": finite(recent.min()), "maximum": finite(recent.max()),
                     "review": average == 0 or abs(change) > core.EXCEL_DASHBOARD_CHANGE_THRESHOLD or amplitude > core.EXCEL_DASHBOARD_RANGE_THRESHOLD})
    dates = list(df[core.COL_DATA])
    cutoff_month = dates[-1].month
    wp_mat = df[core.COL_SELL_OUT].rolling(12, min_periods=12).sum()
    rows, previous = [], {}
    for i, date in enumerate(dates):
        if date.month != cutoff_month or i < 11:
            continue
        wp = finite(wp_mat.iloc[i])
        values = [pipeline["sell_in_mat"][i] for pipeline in pipelines]
        prior = previous.get(date.year - 1)
        def growth(value, old):
            return finite((value / old - 1) * 100) if value is not None and old else None
        row = {"date": date.isoformat(), "worldpanel": wp,
               "worldpanel_yoy": growth(wp, prior["worldpanel"]) if prior else None,
               "sell_in": values,
               "sell_in_yoy": [growth(value, prior["sell_in"][p]) if prior else None for p, value in enumerate(values)],
               "coverage": [pipeline["coverage_series"][i] for pipeline in pipelines]}
        rows.append(row)
        previous[date.year] = row
    return {"kpis": kpis, "mat_periods": rows,
            "change_threshold": core.EXCEL_DASHBOARD_CHANGE_THRESHOLD,
            "range_threshold": core.EXCEL_DASHBOARD_RANGE_THRESHOLD}


def analyze_workbook(path, options):
    core = engine()
    categories = core.load_categories()
    country, basket, category, short_category, manufacturer = core.parse_file_metadata(path.name, categories)
    country_code, category_code, _ = core.parse_input_filename_parts(path.name)
    result = {"file": path.name, "fingerprint": fingerprint(path), "country": country,
              "category": category, "manufacturer": manufacturer,
              "country_code": country_code, "category_code": category_code, "basket": basket,
              "table_colors": {
                  "correlation": ["#" + color for color in core.EXCEL_CORRELATION_COLORS],
                  "positive": {"fill": "#" + core.EXCEL_VARIATION_POSITIVE_FILL, "ink": "#" + core.EXCEL_VARIATION_POSITIVE_FONT},
                  "negative": {"fill": "#" + core.EXCEL_VARIATION_NEGATIVE_FILL, "ink": "#" + core.EXCEL_VARIATION_NEGATIVE_FONT},
                  "coverage_current": "#" + core.EXCEL_COVERAGE_CURRENT_FILL,
                  "coverage_previous": "#" + core.EXCEL_COVERAGE_PREVIOUS_FILL,
              },
              "options": serialize_options(options), "sheets": [],
              "chart_colors": {
                  "worldpanel": core.COLOR_SELLOUT_TREND_LINE,
                  "sellin": core.COLOR_SELLIN_TREND_LINE,
                  "coverage": core.COLOR_COBERTURA_BAR,
                  "penetration": core.COLOR_PENETRACION_BAR,
                  "trend_fill": core.COLOR_TENDENCIA_FILL,
                  "comparison": "#008C95", "original": "#B0B0B0",
              }}
    result["selection_key"] = hashlib.sha256((path.name + ":" + result["fingerprint"]).encode("utf-8")).hexdigest()
    group = None
    inherited_category = None
    with core.pd.ExcelFile(path) as workbook:
        for name in workbook.sheet_names:
            captured = io.StringIO()
            with contextlib.redirect_stdout(captured):
                df, unit = core.load_and_preprocess_sheet(workbook, name)
            messages = core.re.sub(r"\x1b\[[0-9;]*m", "", captured.getvalue()).strip()
            sheet = {"name": name, "brand": core._clean_brand_name_from_sheet(name),
                     "warnings": [messages] if messages else [], "error": None}
            if df is None or df.empty:
                sheet["error"] = "La hoja no tiene una tabla válida con fechas y Sell-in."
                result["sheets"].append(sheet)
                continue
            group_title, group = core.build_metadata_group_for_sheet(name, group, manufacturer)
            metadata = core.resolve_sheet_bank_metadata(
                category_code=category_code, fabricante=manufacturer,
                marca_nombre_limpio=sheet["brand"], section_title=group_title,
                categories_df=categories, default_pais_nombre=country, default_cesta_nombre=basket,
                default_categoria_nombre=category, default_categoria_nombre_corto=short_category,
                sheet_metadata_hints=df.attrs.get("sheet_metadata_hints"),
                inherited_category_code=inherited_category,
            )
            if core._requires_metadata_category_resolution(category_code) and core.is_total_group_sheet(name):
                code = core._normalize_category_code(metadata.categoria_codigo)
                inherited_category = code if code and not core._requires_metadata_category_resolution(code) else None
            sheet.update(metadata=asdict(metadata), unit=unit, months=len(df),
                         category=metadata.categoria_nombre_corto or metadata.categoria_nombre)
            sheet["error"] = monthly_problem(df, core)
            if len(df) < 12:
                sheet["error"] = "Se requieren al menos 12 meses para calcular cobertura MAT."
            if sheet["error"]:
                result["sheets"].append(sheet)
                continue
            forced = core.extract_forced_pipeline_from_sheet_name(name)
            variations = core.compute_variations_dataframe(df)
            recommendation = core.select_optimal_pipeline(
                df, variations, forced_pipeline=forced, category_code=metadata.categoria_codigo,
                category_name=metadata.categoria_nombre, basket_name=metadata.cesta_nombre,
            )
            by_corr = core.select_correlation_pipeline(recommendation.candidates, forced)
            with contextlib.redirect_stdout(captured):
                coverage = core.compute_coverage_dataframe(df, country,
                    options.coverage_type, options.round_coverage, sheet["brand"])
            dates = [value.isoformat() for value in df[core.COL_DATA]]
            annual = variations.loc[variations["Tipo"] == "Anual"].iloc[0]
            wp_variation = finite(annual.get("WP by Numerator"))
            wp_mat = df[core.COL_SELL_OUT].rolling(12, min_periods=12).sum()
            sheet.update(dates=dates, sell_in=[finite(v) for v in df[core.COL_SELL_IN]],
                         sell_out=[finite(v) for v in df[core.COL_SELL_OUT]],
                         penetration=[finite(v) for v in df[core.COL_PENET]],
                         purchase_average=[finite(v) for v in df[core.COL_COMPRA_MEDIA]],
                         purchase_occasion=[finite(v) for v in df[core.COL_COMPRA_OCA]],
                         frequency=[finite(v) for v in df[core.COL_FREQ]],
                         buyers_series=[finite(v) for v in df[core.COL_BUYERS]],
                         template=template_details(df, core),
                         wp_mat=[finite(v) for v in wp_mat], pipelines=[],
                         wp_yoy=evolution_variation(wp_mat),
                         recommended=recommendation.pipeline, reason=recommendation.reason,
                         correlation_pipeline=by_corr.pipeline, forced=forced)
            candidates = {candidate.pipeline: candidate for candidate in recommendation.candidates}
            for pipeline in range(7):
                shifted = df[core.COL_SELL_IN].shift(pipeline)
                shifted_mat = shifted.rolling(12, min_periods=12).sum()
                candidate = candidates.get(pipeline)
                client_var = finite(annual.get(f"Cliente P{pipeline}"))
                cov = coverage[f"P{pipeline}"]
                current_cov = finite(cov.iloc[-1])
                prev_cov = finite(cov.iloc[-13]) if len(cov) >= 13 else None
                sheet["pipelines"].append({
                    "pipeline": pipeline, "available": current_cov is not None,
                    "correlation": finite(core.compute_pipeline_current_year_correlation(df, pipeline)),
                    "previous_correlation": finite(core.compute_pipeline_mat_correlation(df, pipeline, end_offset=12)),
                    "client_variation": client_var * 100 if client_var is not None else None,
                    "wp_variation": wp_variation * 100 if wp_variation is not None else None,
                    "gap": abs(client_var - wp_variation) * 100 if client_var is not None and wp_variation is not None else None,
                    "coverage": current_cov, "previous_coverage": prev_cov,
                    "stability": current_cov - prev_cov if current_cov is not None and prev_cov is not None else None,
                    "outlier": candidate.recent_shipment_outlier if candidate else False,
                    "sell_in": [finite(v) for v in shifted],
                    "sell_in_mat": [finite(v) for v in shifted_mat],
                    "yoy": evolution_variation(shifted_mat),
                    "coverage_series": [finite(v) for v in cov],
                })
            available = [p["pipeline"] for p in sheet["pipelines"] if p["available"]]
            sheet["dashboard"] = dashboard_details(df, core, sheet["pipelines"])
            sheet["default_pipeline"] = next((p for p in [forced, recommendation.pipeline, 0] if p in available), available[0])
            buyers = finite(df[core.COL_BUYERS].tail(12).mean())
            sheet["buyers"] = buyers
            if buyers is not None and buyers < core.LOW_PENETRATION_BUYERS_THRESHOLD:
                sheet["warnings"].append("Menos de 200 compradores promedio: resultados para uso interno.")
            if len(df) < 30:
                sheet["warnings"].append("La historia puede ser insuficiente para correlación o variación MAT en algunos pipelines. — significa sin base comparable.")
            for issue in core.detect_brand_data_issues(df, window=0):
                sheet["warnings"].append("Hay ceros o datos faltantes tratados como cero." if issue == "zero_dash" else "Hay valores negativos en la serie.")
            result["sheets"].append(sheet)
    if fingerprint(path) != result["fingerprint"]:
        raise ValueError("El Excel cambió durante el análisis. Vuelve a cargarlo.")
    return result


class ExplorerState:
    def __init__(self, root=ROOT):
        self.root = Path(root).resolve()
        self.token = secrets.token_urlsafe(32)
        self.lock = threading.RLock()
        self.analysis_lock = threading.Lock()
        self.analyses = {}
        self.jobs = {}

    def files(self):
        return sorted(path.name for path in self.root.glob("*.xlsx")
                      if not path.name.startswith("~$") and path.name != "file_temp_coverage.xlsx")

    def file_catalog(self):
        with self.analysis_lock, contextlib.redirect_stdout(io.StringIO()):
            core = engine()
            categories = core.load_categories()
            details = []
            for name in self.files():
                entry = {"name": name, "error": None}
                try:
                    path = self.source(name)
                    country, basket, category, short, manufacturer = core.parse_file_metadata(name, categories)
                    country_code, category_code, _ = core.parse_input_filename_parts(name)
                    info = path.stat()
                    entry.update(country=country, country_code=country_code, basket=basket,
                                 category=category, short_category=short, category_code=category_code,
                                 manufacturer=manufacturer, descriptor=core.extract_input_filename_descriptor(name),
                                 per_sheet=core._requires_metadata_category_resolution(category_code),
                                 size=info.st_size, modified=datetime.fromtimestamp(info.st_mtime, timezone.utc).isoformat())
                except (ValueError, OSError) as error:
                    entry["error"] = str(error)
                details.append(entry)
            return details

    def source(self, name):
        if not isinstance(name, str) or name not in self.files():
            raise ValueError("Selecciona un Excel de la carpeta del proyecto.")
        path = (self.root / name).resolve()
        if path.parent != self.root:
            raise ValueError("El archivo debe estar dentro del proyecto.")
        return path

    def analyze(self, payload):
        path = self.source(payload.get("file"))
        with self.analysis_lock:
            options = options_from_request(payload)
            key = path.name + fingerprint(path) + json.dumps(serialize_options(options), sort_keys=True)
            if key not in self.analyses:
                analysis = analyze_workbook(path, options)
                analysis["id"] = secrets.token_urlsafe(16)
                saved = self.root / "outputs" / "pipeline_explorer" / "selections" / f"{analysis['selection_key']}.json"
                analysis["saved_selection"] = json.loads(saved.read_text(encoding="utf-8")) if saved.exists() else None
                # Retain a small number of workbooks, never every historical preview.
                if len(self.analyses) >= 4:
                    self.analyses.pop(next(iter(self.analyses)))
                self.analyses[key] = analysis
            return self.analyses[key]

    def reviewed_selection(self, payload):
        analysis = next((value for value in self.analyses.values() if value["id"] == payload.get("analysis_id")), None)
        if analysis is None:
            raise ValueError("Carga y revisa el archivo antes de guardar o generar.")
        source = self.source(analysis["file"])
        if fingerprint(source) != analysis["fingerprint"]:
            raise ValueError("El Excel cambió desde la revisión. Vuelve a cargarlo antes de generar.")
        selections = payload.get("selections")
        if not isinstance(selections, dict) or not selections:
            raise ValueError("Selecciona al menos una hoja.")
        sheets = {sheet["name"]: sheet for sheet in analysis["sheets"]}
        for name, pipeline in selections.items():
            sheet = sheets.get(name)
            if sheet is None or sheet["error"] or type(pipeline) is not int or pipeline not in range(7):
                raise ValueError(f"Selección inválida para {name}.")
            if not sheet["pipelines"][pipeline]["available"]:
                raise ValueError(f"{name} no tiene historia suficiente para P{pipeline}.")
        return analysis, source, selections

    def save_selection(self, payload):
        with self.analysis_lock:
            analysis, _, selections = self.reviewed_selection(payload)
            selection = {"file": analysis["file"], "fingerprint": analysis["fingerprint"],
                         "options": analysis["options"], "selections": selections}
            directory = self.root / "outputs" / "pipeline_explorer" / "selections"
            directory.mkdir(parents=True, exist_ok=True)
            target = directory / f"{analysis['selection_key']}.json"
            tmp = target.with_suffix(".tmp")
            tmp.write_text(json.dumps(selection, ensure_ascii=False, indent=2), encoding="utf-8")
            tmp.replace(target)
            analysis["saved_selection"] = selection
            return {"message": "Selección guardada para este archivo y esta versión de sus datos."}

    def export(self, payload):
        with self.analysis_lock, self.lock:
            if any(job["status"] == "running" for job in self.jobs.values()):
                raise ValueError("Ya hay una generación en curso. Espera a que termine.")
            analysis, source, selections = self.reviewed_selection(payload)
            job_id = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + secrets.token_hex(3)
            workdir = self.root / "outputs" / "pipeline_explorer" / job_id
            workdir.mkdir(parents=True)
            shutil.copy2(source, workdir / source.name)
            if fingerprint(workdir / source.name) != analysis["fingerprint"]:
                raise ValueError("El Excel cambió al copiarlo. Vuelve a cargarlo.")
            shutil.copy2(self.root / "Modelo_PPT.pptx", workdir / "Modelo_PPT.pptx")
            config = {"file": source.name, "fingerprint": analysis["fingerprint"],
                      "options": analysis["options"], "selections": selections,
                      "metadata": {sheet["name"]: sheet["metadata"] for sheet in analysis["sheets"] if sheet["name"] in selections}}
            config_path = workdir / "selection.json"
            config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
            job = {"id": job_id, "status": "running", "directory": str(workdir), "outputs": [], "error": None}
            self.jobs[job_id] = job
            threading.Thread(target=self.run_export, args=(job_id, config_path), daemon=True).start()
            return dict(job)

    def run_export(self, job_id, config_path):
        workdir = config_path.parent
        try:
            env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8", SHOW_CAT_MSG="0")
            with (workdir / "generation.log").open("w", encoding="utf-8") as log:
                completed = subprocess.run([sys.executable, "-u", str(Path(__file__).resolve()), "--worker", str(config_path)],
                    cwd=str(workdir), env=env, stdin=subprocess.DEVNULL, stdout=log, stderr=subprocess.STDOUT,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            result_path = workdir / "result.json"
            if completed.returncode or not result_path.exists():
                raise RuntimeError("La generación falló. Consulta el registro para ver la causa.")
            result = json.loads(result_path.read_text(encoding="utf-8"))
            if result.get("error"):
                raise RuntimeError(result["error"])
            outputs = [{"label": label, "name": Path(path).name, "path": path,
                        "url": f"/api/download?job={quote(job_id)}&output={quote(label)}"}
                       for label, path in result["outputs"].items() if path]
            with self.lock:
                self.jobs[job_id].update(status="complete", outputs=outputs)
        except Exception as exc:
            with self.lock:
                self.jobs[job_id].update(status="failed", error=str(exc))

    def job(self, job_id):
        with self.lock:
            if job_id not in self.jobs:
                raise ValueError("Generación desconocida.")
            job = dict(self.jobs[job_id])
        path = Path(job["directory"]) / "generation.log"
        if path.exists():
            with path.open("rb") as log:
                log.seek(max(0, path.stat().st_size - 6000))
                job["log"] = log.read().decode("utf-8", errors="replace")
        else:
            job["log"] = "Preparando generación…"
        return job


def run_worker(config_path):
    config_path = Path(config_path).resolve()
    core = engine()
    config = json.loads(config_path.read_text(encoding="utf-8"))
    result = {}
    try:
        options = core.ExecutionOptions(**config["options"])
        options.pipeline_overrides = config["selections"]
        options.sheet_metadata_overrides = {name: core.SheetBankMetadata(**values) for name, values in config["metadata"].items()}
        app = core.CoverageStudioUltraApp()
        app.root_dir = str(config_path.parent)
        os.chdir(app.root_dir)
        outputs = app.process_file(config["file"], options, 1, 1)
        if not outputs or any(not path or not Path(path).is_file() for path in outputs.values()):
            raise RuntimeError("No se generaron todos los entregables. Revisa generation.log.")
        result = {"outputs": outputs}
        core.cleanup_temp_dir(app.root_dir)
        # Keep only deliverables, manifest and log in the execution directory.
        (config_path.parent / config["file"]).unlink(missing_ok=True)
        (config_path.parent / "Modelo_PPT.pptx").unlink(missing_ok=True)
    except BaseException as exc:
        result = {"error": str(exc) or "El motor interrumpió la generación."}
    (config_path.parent / "result.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return 1 if "error" in result else 0


def make_handler(state):
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *args):
            pass

        def send_bytes(self, data, content_type, status=200, attachment=None):
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.send_header("Content-Security-Policy", "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'")
            if attachment:
                self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(attachment)}")
            self.end_headers()
            self.wfile.write(data)

        def send_json(self, value, status=200):
            self.send_bytes(json.dumps(value, ensure_ascii=False, allow_nan=False).encode("utf-8"), "application/json; charset=utf-8", status)

        def allowed(self, api=False):
            host = self.headers.get("Host", "")
            expected = f"127.0.0.1:{self.server.server_port}"
            if host not in {expected, f"localhost:{self.server.server_port}"}:
                self.send_json({"error": "Host no permitido."}, 403)
                return False
            origin = self.headers.get("Origin")
            if origin and origin not in {f"http://{expected}", f"http://localhost:{self.server.server_port}"}:
                self.send_json({"error": "Origen no permitido."}, 403)
                return False
            if api:
                query = parse_qs(urlsplit(self.path).query)
                token = self.headers.get("X-Explorer-Token", query.get("token", [""])[0])
                if not secrets.compare_digest(token, state.token):
                    self.send_json({"error": "Sesión local inválida. Recarga la pantalla."}, 403)
                    return False
            return True

        def do_GET(self):
            path = urlsplit(self.path).path
            if not self.allowed(path.startswith("/api/")):
                return
            try:
                if path == "/":
                    html = (ASSETS / "index.html").read_text(encoding="utf-8").replace("__EXPLORER_TOKEN__", state.token)
                    self.send_bytes(html.encode("utf-8"), "text/html; charset=utf-8")
                elif path in {"/app.js", "/dashboard.js", "/style.css"}:
                    file = ASSETS / path[1:]
                    self.send_bytes(file.read_bytes(), mimetypes.guess_type(str(file))[0] + "; charset=utf-8")
                elif path == "/api/files":
                    self.send_json({"files": state.files(), "details": state.file_catalog()})
                elif path == "/api/job":
                    self.send_json(state.job(parse_qs(urlsplit(self.path).query).get("id", [""])[0]))
                elif path == "/api/jobs":
                    with state.lock:
                        latest = next(reversed(state.jobs), None)
                    self.send_json({"latest": state.job(latest) if latest else None})
                elif path == "/api/download":
                    params = parse_qs(urlsplit(self.path).query)
                    job = state.job(params.get("job", [""])[0])
                    label = params.get("output", [""])[0]
                    if label == "log":
                        file = Path(job["directory"]) / "generation.log"
                    elif label == "selection":
                        file = Path(job["directory"]) / "selection.json"
                    else:
                        output = next((out for out in job["outputs"] if out["label"] == label), None)
                        if output is None:
                            raise ValueError("Archivo de salida desconocido.")
                        file = Path(output["path"])
                    if not file.resolve().is_relative_to(Path(job["directory"]).resolve()):
                        raise ValueError("Archivo fuera de la ejecución.")
                    self.send_bytes(file.read_bytes(), mimetypes.guess_type(str(file))[0] or "application/octet-stream", attachment=file.name)
                else:
                    self.send_json({"error": "Ruta desconocida."}, 404)
            except (ValueError, FileNotFoundError) as exc:
                self.send_json({"error": str(exc)}, 400)
            except Exception as exc:
                self.send_json({"error": str(exc)}, 500)

        def do_POST(self):
            if not self.allowed(True):
                return
            try:
                length = int(self.headers.get("Content-Length", "0"))
                if not 0 < length <= 1024 * 1024:
                    raise ValueError("Solicitud inválida o demasiado grande.")
                payload = json.loads(self.rfile.read(length))
                if not isinstance(payload, dict):
                    raise ValueError("Solicitud inválida.")
                routes = {"/api/analyze": state.analyze, "/api/save": state.save_selection, "/api/export": state.export}
                operation = routes.get(urlsplit(self.path).path)
                if operation is None:
                    self.send_json({"error": "Ruta desconocida."}, 404)
                else:
                    self.send_json(operation(payload))
            except (ValueError, OSError) as exc:
                self.send_json({"error": str(exc)}, 400)
            except Exception as exc:
                self.send_json({"error": str(exc)}, 500)
    return Handler


def main(argv=None):
    parser = argparse.ArgumentParser(description="Explorar pipelines de CoverageLab en el navegador local.")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true")
    parser.add_argument("--worker", help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    if args.worker:
        return run_worker(args.worker)
    if not 0 <= args.port <= 65535:
        parser.error("El puerto debe estar entre 0 y 65535.")
    state = ExplorerState()
    try:
        server = ExplorerHTTPServer(("127.0.0.1", args.port), make_handler(state))
    except OSError as exc:
        parser.error(f"No se pudo abrir el puerto {args.port}: {exc}. Prueba --port 8766.")
    url = f"http://127.0.0.1:{server.server_port}/"
    print(f"CoverageLab · Explorador de pipelines\n{url}\nCtrl+C para cerrar la pantalla local.", flush=True)
    if not args.no_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nExplorador cerrado. Las generaciones ya iniciadas continúan en su proceso local.")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
