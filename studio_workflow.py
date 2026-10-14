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


class StudioWorkflow:
    def __init__(self, explorer, research_root=None):
        self.explorer = explorer
        self.root = explorer.root
        self.research_root = Path(research_root) if research_root else self.root.parent / "5W1H"
        self._inputs = None
        self._research = None
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
                                      number_locale=payload.get("number_locale", "auto"))
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
            job = {"id": job_id, "kind": "research", "status": "running", "file": reviewed["file"],
                   "fingerprint": options["fingerprint"], "started_at": started.isoformat(),
                   "title": source.stem, "directory": None, "outputs": [], "error": None,
                   "options": options, "log": "Generación del informe completo 5W1H iniciada."}
            self.research_jobs[job_id] = job
            threading.Thread(target=self.run_research_export, args=(job_id,), daemon=True).start()
            return dict(job)

    def run_research_export(self, job_id):
        job = self.research_jobs[job_id]
        try:
            result = self.research.export(job["file"], options=job["options"], destination=self.explorer.output_root)
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
                job.update(status="complete", title=result.get("title", directory.name), directory=str(directory),
                           outputs=outputs, log="Informe completo generado con los datos y pipelines revisados.")
                (directory / "research-job.json").write_text(json.dumps(job, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as exc:
            with self.lock:
                job.update(status="failed", error=str(exc), log=str(exc))

    def restore_research_jobs(self):
        for manifest in sorted(self.explorer.output_root.glob("*/research-job.json")):
            try:
                job = json.loads(manifest.read_text(encoding="utf-8"))
                directory = Path(job["directory"]).resolve()
                if directory != manifest.parent.resolve() or not directory.is_relative_to(self.explorer.output_root.resolve()):
                    continue
                if job["kind"] != "research" or job["status"] != "complete" or not isinstance(job["file"], str):
                    continue
                if not isinstance(job["id"], str) or not isinstance(job.get("fingerprint"), str) or len(job["fingerprint"]) != 64:
                    continue
                datetime.fromisoformat(job["started_at"])
                if not job["outputs"] or any(not Path(out["path"]).resolve().is_relative_to(directory)
                                             or not Path(out["path"]).is_file() for out in job["outputs"]):
                    continue
                self.research_jobs[job["id"]] = job
            except (KeyError, TypeError, ValueError, OSError):
                continue

    def research_job(self, job_id=None, file_id=None):
        with self.lock:
            if not job_id:
                job_id = next((key for key, job in reversed(self.research_jobs.items()) if job["file"] == file_id), None)
            if job_id not in self.research_jobs:
                if not job_id:
                    return None
                raise ValueError("Informe 5W1H desconocido.")
            job = dict(self.research_jobs[job_id])
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
            target = directory
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
