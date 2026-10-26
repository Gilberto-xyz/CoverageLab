"""Coordinación local de entradas, cobertura e investigación 5W1H."""
from __future__ import annotations

from datetime import datetime
import json
import os
from pathlib import Path
import secrets
import subprocess
import sys
import threading
import contextlib
from export_progress import ExportProgressReporter, RESEARCH_STAGES, historical_estimate, elapsed_seconds, snapshot as progress_snapshot

from studio_desktop import open_folder
from delivery_files import auxiliary, manifests, manifest_directory, organize_auxiliaries
from project_paths import external_research_directory, state_directory


class StudioWorkflow:
    def __init__(self, explorer, research_root=None):
        self.explorer = explorer
        self.root = explorer.root
        self.research_root = Path(research_root) if research_root else external_research_directory(self.root)
        self._inputs = None
        self._research = None
        self._desktop = None
        self._drafts = None
        self.lock = threading.RLock()
        self.research_jobs = {}
        self.research_analyses = {}
        self.restore_research_jobs()

    @property
    def inputs(self):
        with self.lock:
            if self._inputs is None:
                from studio_inputs import InputStudio
                self._inputs = InputStudio(self.root, research_root=self.research_root)
            return self._inputs

    @property
    def research(self):
        with self.lock:
            if self._research is None:
                from studio_5w1h import ResearchStudio
                self._research = ResearchStudio(self.root, external_dir=self.research_root)
            return self._research

    @staticmethod
    def coordinates(payload):
        return {key: int(payload.get(key, default)) for key, default in
                (("row", 1), ("limit", 60), ("column", 1), ("columns", 32))}

    def workbook(self, payload):
        return self.inputs.workbook(payload.get("file"), kind=payload.get("kind", "coverage"),
                                    sheet=payload.get("sheet") or None, **self.coordinates(payload))

    @property
    def drafts(self):
        with self.lock:
            if self._drafts is None:
                from studio_drafts import StudioDrafts
                self._drafts = StudioDrafts(self.inputs)
            return self._drafts

    def input_draft(self, payload):
        with self.explorer.analysis_lock:
            return self.drafts.get(payload.get("file"), payload.get("kind", "coverage"))

    def save_draft(self, payload):
        with self.explorer.analysis_lock:
            return self.drafts.save(payload)

    def discard_draft(self, payload):
        with self.explorer.analysis_lock:
            return self.drafts.discard(payload.get("file"), payload.get("kind", "coverage"))

    @property
    def desktop(self):
        with self.lock:
            if self._desktop is None:
                from studio_desktop import DesktopStudio
                self._desktop = DesktopStudio(self.inputs)
            return self._desktop

    def open_excel(self, payload):
        with self.explorer.analysis_lock:
            return self.invalidate(self.desktop.open(payload))

    def reload_excel(self, payload):
        with self.explorer.analysis_lock:
            return self.invalidate(self.desktop.reload(payload))

    def invalidate(self, result):
        if result.get("kind") == "coverage":
            for key, analysis in list(self.explorer.analyses.items()):
                if analysis.get("file") == result.get("file"):
                    del self.explorer.analyses[key]
        else:
            self.research_analyses.clear()
        return result

    def create_input(self, payload):
        with self.explorer.analysis_lock:
            return self.invalidate(self.inputs.create(payload))

    def save_input(self, payload):
        with self.explorer.analysis_lock:
            result = self.inputs.save(payload.get("file"), kind=payload.get("kind", "coverage"),
                                      fingerprint=payload.get("fingerprint"), edits=payload.get("edits"),
                                      number_locale=payload.get("number_locale", "auto"),
                                      number_locale_hint=payload.get("number_locale_hint", "decimal-dot"),
                                      preserve_text=payload.get("preserve_text", False))
            return self.invalidate(result)

    def add_sheet(self, payload):
        with self.explorer.analysis_lock:
            result = self.inputs.add_sheet(payload.get("file"), kind=payload.get("kind", "coverage"),
                                           fingerprint=payload.get("fingerprint"), name=payload.get("name"),
                                           source_sheet=payload.get("source_sheet") or None)
            return self.invalidate(result)

    def import_input(self, payload):
        with self.explorer.analysis_lock:
            return self.invalidate(self.inputs.import_workbook(payload))

    def research_analyze(self, payload):
        from pipeline_explorer import fingerprint
        file_id = payload.get("file")
        with self.lock:
            source = self.research.source(file_id)
            before = fingerprint(source)
            options = {key: payload[key] for key in ("mode", "start", "end", "elements", "pipeline") if key in payload}
            result = self.research.analyze(file_id, payload.get("sheet"), **options)
            if fingerprint(source) != before:
                raise ValueError("Los datos cambiaron durante el análisis. Vuelve a cargar la investigación.")
            result.update(id=secrets.token_urlsafe(16), fingerprint=before, source_id=file_id)
            self.research_analyses[result["id"]] = {"result": result, "options": options,
                                                   "file": file_id, "sheet": payload.get("sheet")}
            while len(self.research_analyses) > 8:
                del self.research_analyses[next(iter(self.research_analyses))]
            return result

    def export_research(self, payload):
        from pipeline_explorer import fingerprint
        with self.lock:
            if not self.explorer.output_root.resolve().is_relative_to(self.root):
                raise ValueError("La carpeta Output debe estar dentro del proyecto.")
            reviewed = self.research_analyses.get(payload.get("analysis_id"))
            if reviewed is None or not reviewed["result"].get("ready"):
                raise ValueError("Carga y revisa datos válidos de 5W1H antes de generar el informe.")
            source = self.research.source(reviewed["file"])
            if fingerprint(source) != reviewed["result"]["fingerprint"]:
                raise ValueError("El archivo cambió desde la revisión. Vuelve a analizar 5W1H.")
            if any(job["status"] == "running" for job in self.research_jobs.values()):
                raise ValueError("Ya hay un informe 5W1H en curso.")
            started = datetime.now().astimezone()
            job_id = "research_" + started.strftime("%Y%m%d_%H%M%S") + "_" + secrets.token_hex(3)
            # La vista puede filtrar meses/elementos; el informe conserva las hojas completas.
            options = {key: value for key, value in reviewed["options"].items() if key in {"mode", "pipeline"}}
            options.update(sheet=reviewed["sheet"], fingerprint=reviewed["result"]["fingerprint"])
            if "sheets" in payload:
                options["sheets"] = payload["sheets"]
            if "size_threshold" in payload:
                options["size_threshold"] = payload["size_threshold"]
            overrides = payload.get("pipeline_overrides", {})
            if not isinstance(overrides, dict) or any(not isinstance(k, str) or type(v) is not int or v not in range(7)
                                                      for k, v in overrides.items()):
                raise ValueError("Selección de pipelines inválida.")
            options["pipeline_overrides"] = overrides
            info = self.research.inspect(reviewed['file'])
            selected = set(options.get('sheets') or [s['name'] for s in info['sheets'] if s['has_data']])
            work_units = sum(max(1, s['rows'] / 36) * max(1, s['columns'] / 12)
                             for s in info['sheets'] if s['name'] in selected)
            estimate = historical_estimate(self.research_jobs.values(), work_units)
            progress_dir = state_directory(self.root) / '5w1h' / job_id
            progress_dir.mkdir(parents=True, exist_ok=True)
            progress_path = progress_dir / 'progress.json'
            ExportProgressReporter(progress_path, started.isoformat(), estimate, stages=RESEARCH_STAGES).update('starting')
            job = {"id": job_id, "kind": "research", "status": "running", "file": reviewed["file"],
                   "fingerprint": options["fingerprint"], "started_at": started.isoformat(),
                   "title": source.stem, "directory": None, "outputs": [], "error": None,
                   "options": options, "log": "Generación del informe completo 5W1H iniciada.",
                   'work_units': work_units, 'estimate_seconds': estimate, 'progress_path': str(progress_path),
                   'progress': progress_snapshot({}, 'running', started.isoformat())}
            self.research_jobs[job_id] = job
            threading.Thread(target=self.run_research_export, args=(job_id,), daemon=True).start()
            return dict(job)

    def run_research_export(self, job_id):
        job = self.research_jobs[job_id]
        try:
            result = self.research.export(job["file"], options=job["options"], destination=self.explorer.output_root,
                                          progress={'path': job['progress_path'], 'started_at': job['started_at'],
                                                    'estimate_seconds': job.get('estimate_seconds')})
            directory = Path(result["directory"]).resolve()
            if not directory.is_relative_to(self.explorer.output_root.resolve()) or directory == self.explorer.output_root.resolve():
                raise ValueError("El informe debe quedar en una carpeta propia dentro de Output.")
            outputs = []
            for label, name in result["outputs"].items():
                path = Path(name).resolve()
                if not path.is_relative_to(directory) or not path.is_file():
                    raise ValueError("El motor devolvió un resultado fuera de la entrega.")
                outputs.append({"label": label, "name": path.name, "path": str(path)})
            if not outputs:
                raise ValueError("El motor no generó archivos de salida.")
            with self.lock:
                progress_path = Path(job['progress_path'])
                raw = self.read_research_progress(job)
                saved_progress = auxiliary(directory, 'export-progress.json', write=True)
                if progress_path.is_file():
                    saved_progress.write_text(json.dumps(raw, ensure_ascii=False), encoding='utf-8')
                job.update(status="complete", title=result.get("title", directory.name), directory=str(directory),
                           outputs=outputs, log="Informe completo generado con los datos y pipelines revisados.",
                           finished_at=datetime.now().astimezone().isoformat(), elapsed_seconds=elapsed_seconds(job['started_at']),
                           progress_path=str(saved_progress), progress=progress_snapshot(raw, 'complete', job['started_at']))
                auxiliary(directory, "research-job.json", write=True).write_text(json.dumps(job, ensure_ascii=False, indent=2), encoding="utf-8")
                with contextlib.suppress(OSError):
                    progress_path.unlink(missing_ok=True)
                    progress_path.parent.rmdir()
        except Exception as exc:
            with self.lock:
                raw = self.read_research_progress(job)
                job.update(status="failed", error=str(exc), log=str(exc),
                           finished_at=datetime.now().astimezone().isoformat(), elapsed_seconds=elapsed_seconds(job['started_at']),
                           progress=progress_snapshot(raw, 'failed', job['started_at']))

    def read_research_progress(self, job):
        path = Path(job.get('progress_path') or state_directory(self.root)/'5w1h'/job['id']/'progress.json')
        if path.resolve().is_relative_to(self.root):
            with contextlib.suppress(OSError, ValueError):
                raw = json.loads(path.read_text(encoding='utf-8'))
                if isinstance(raw, dict):
                    return raw
        return job.get('progress', {})

    def restore_research_jobs(self):
        for manifest in manifests(self.explorer.output_root, "research-job.json"):
            try:
                job = json.loads(manifest.read_text(encoding="utf-8"))
                directory = Path(job["directory"]).resolve()
                if directory != manifest_directory(manifest).resolve() or not directory.is_relative_to(self.explorer.output_root.resolve()):
                    continue
                if job["kind"] != "research" or job["status"] != "complete" or not isinstance(job["file"], str):
                    continue
                if not isinstance(job["id"], str) or not isinstance(job.get("fingerprint"), str) or len(job["fingerprint"]) != 64:
                    continue
                datetime.fromisoformat(job["started_at"])
                for output in job["outputs"]:
                    old_path = Path(output["path"]).resolve()
                    if not old_path.is_relative_to(directory):
                        raise ValueError("Resultado fuera de la entrega.")
                    name = {"Provenance": "investigacion_5w1h.json", "Registro": "generacion_5w1h.log"}.get(output.get("label"))
                    if name and not old_path.is_file():
                        output["path"] = str(auxiliary(directory, name))
                if not job["outputs"] or any(not Path(out["path"]).resolve().is_relative_to(directory)
                                             or not Path(out["path"]).is_file() for out in job["outputs"]):
                    continue
                moved = organize_auxiliaries(directory, ("research-job.json", "_result.json", "generacion_5w1h.log", "investigacion_5w1h.json"), input_copy=True)
                for output in job["outputs"]:
                    output["path"] = moved.get(output["path"], output["path"])
                if moved:
                    try:
                        auxiliary(directory, "research-job.json", write=True).write_text(json.dumps(job, ensure_ascii=False, indent=2), encoding="utf-8")
                    except OSError:
                        pass
                self.research_jobs[job["id"]] = job
            except (KeyError, TypeError, ValueError, OSError):
                continue

    def research_job(self, job_id=None, file_id=None):
        with self.lock:
            if not job_id:
                matching = [(key, job) for key, job in self.research_jobs.items() if job["file"] == file_id]
                job_id = max(matching, key=lambda item: datetime.fromisoformat(item[1].get("started_at", "1970-01-01T00:00:00+00:00")).timestamp())[0] if matching else None
            if job_id not in self.research_jobs:
                if not job_id:
                    return None
                raise ValueError("Informe 5W1H desconocido.")
            job = dict(self.research_jobs[job_id])
            raw = self.read_research_progress(job) if job['status'] == 'running' else job.get('progress', {})
            if job['status'] == 'running': self.research_jobs[job_id]['progress'] = raw
            job['progress'] = progress_snapshot(raw, job['status'], job.get('started_at'))
            job["available_outputs"] = [out["label"] for out in job["outputs"] if Path(out["path"]).is_file()]
            try:
                from pipeline_explorer import fingerprint
                job["source_changed"] = fingerprint(self.research.source(job["file"])) != job["fingerprint"]
            except (ValueError, OSError):
                job["source_changed"] = True
            return job

    def open_research_output(self, payload):
        job = self.research_job(payload.get("id"))
        if not job or job["status"] != "complete" or payload.get("file") != job["file"]:
            raise ValueError("Selecciona una entrega completada de este archivo.")
        directory = Path(job["directory"]).resolve()
        if not directory.is_relative_to(self.explorer.output_root.resolve()) or not directory.is_dir():
            raise ValueError("La carpeta no pertenece a una entrega 5W1H local.")
        label = payload.get("output")
        if label == "folder":
            return open_folder(directory)
        else:
            output = next((out for out in job["outputs"] if out["label"] == label), None)
            if output is None:
                raise ValueError("Resultado desconocido.")
            target = Path(output["path"]).resolve()
            if not target.is_relative_to(directory) or target.suffix.lower() not in {".xlsx", ".pptx", ".json", ".log", ".txt"} or not target.is_file():
                raise ValueError("Resultado inválido o no disponible.")
        if sys.platform == "win32":
            try:
                os.startfile(str(target))
            except OSError as exc:
                if getattr(exc, "winerror", None) != 1155 or target.suffix.lower() not in {".json", ".log", ".txt"}:
                    raise
                subprocess.Popen(["notepad.exe", str(target)])
        else:
            subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(target)],
                             stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return {"message": "Resultado 5W1H abierto en su aplicación local."}
