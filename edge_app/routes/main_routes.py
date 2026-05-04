"""Rotas principais, dashboard, health check e status de jobs."""
from edge_app import application as app_module

healthz = app_module.healthz
home = app_module.home
job_page = app_module.job_page
job_status = app_module.job_status
job_download = app_module.job_download
