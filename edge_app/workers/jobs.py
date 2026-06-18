"""Fila assíncrona leve para processamentos pesados.

Usa ThreadPoolExecutor para funcionar sem Redis em instalações pequenas. Em produção
maior, pode ser substituída por Celery/RQ mantendo a mesma interface de status.
"""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable, Optional
import os
import secrets
import threading
import traceback


@dataclass
class Job:
    id: str
    title: str
    status: str = "queued"  # queued | running | finished | failed
    progress: int = 0
    message: str = "Aguardando processamento..."
    result_path: Optional[str] = None
    download_name: Optional[str] = None
    error: Optional[str] = None
    created_at: datetime = field(default_factory=datetime.utcnow)
    updated_at: datetime = field(default_factory=datetime.utcnow)


# Compatibilidade com versões anteriores que importavam JobState de
# edge_app.async_jobs. O estado atual do job é representado pela dataclass Job.
JobState = Job


class JobManager:
    def __init__(self, storage_dir: str, max_workers: int = 2, ttl_hours: int = 12):
        self.storage_dir = Path(storage_dir)
        self.storage_dir.mkdir(parents=True, exist_ok=True)
        self.executor = ThreadPoolExecutor(max_workers=max_workers)
        self.ttl = timedelta(hours=ttl_hours)
        self._jobs: dict[str, Job] = {}
        self._lock = threading.Lock()

    def create(self, title: str, func: Callable[[Callable[[int, str], None]], tuple[str, str]]) -> Job:
        self.cleanup()
        job = Job(id=secrets.token_urlsafe(12), title=title)
        with self._lock:
            self._jobs[job.id] = job
        self.executor.submit(self._run, job.id, func)
        return job

    def get(self, job_id: str) -> Optional[Job]:
        with self._lock:
            return self._jobs.get(job_id)

    def update(self, job_id: str, **kwargs) -> None:
        with self._lock:
            job = self._jobs.get(job_id)
            if not job:
                return
            for key, value in kwargs.items():
                setattr(job, key, value)
            job.updated_at = datetime.utcnow()

    def _run(self, job_id: str, func: Callable[[Callable[[int, str], None]], tuple[str, str]]) -> None:
        def progress(percent: int, message: str) -> None:
            self.update(job_id, progress=max(0, min(100, int(percent))), message=message)

        self.update(job_id, status="running", progress=5, message="Processamento iniciado...")
        try:
            result_path, download_name = func(progress)
            if not result_path or not os.path.exists(result_path):
                raise FileNotFoundError("O processamento terminou sem gerar arquivo de saída.")
            self.update(
                job_id,
                status="finished",
                progress=100,
                message="Arquivo pronto para download.",
                result_path=str(result_path),
                download_name=download_name or Path(result_path).name,
            )
        except Exception as exc:  # log detalhado fica no campo interno, mensagem ao usuário é limpa
            self.update(
                job_id,
                status="failed",
                progress=100,
                message="O processamento falhou. Revise os arquivos enviados e tente novamente.",
                error=f"{exc}\n{traceback.format_exc(limit=8)}",
            )

    def cleanup(self) -> None:
        cutoff = datetime.utcnow() - self.ttl
        with self._lock:
            expired = [job_id for job_id, job in self._jobs.items() if job.created_at < cutoff]
            for job_id in expired:
                job = self._jobs.pop(job_id, None)
                if job and job.result_path:
                    try:
                        Path(job.result_path).unlink(missing_ok=True)
                    except Exception:
                        pass
