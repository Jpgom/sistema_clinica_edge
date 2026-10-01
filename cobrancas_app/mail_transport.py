"""Transporte SMTP com resultados explícitos de aceite, teste e incerteza."""
from __future__ import annotations

import html
import re
import smtplib
from pathlib import Path
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr, formatdate, make_msgid


class DeliveryUncertain(RuntimeError):
    """A conexão caiu durante o envio: o servidor pode ter aceitado a mensagem."""


def validate_config(cfg, valid_email):
    if not cfg.get("username") or not cfg.get("password") or not valid_email(cfg.get("sender_email")):
        raise ValueError("Configure usuário, senha de app e remetente válido do Gmail antes de enviar.")
    if cfg.get("host") != "smtp.gmail.com":
        raise ValueError("Esta versão está configurada para envio somente pelo Gmail.")
    if cfg.get("security") not in {"ssl", "starttls"}:
        raise ValueError("Selecione SSL ou STARTTLS nas configurações do Gmail.")
    if cfg.get("test_mode") and not valid_email(cfg.get("test_email")):
        raise ValueError("Modo de teste ativo sem e-mail de teste válido.")


def normalize_cc(valid_email, to_email, values, strict=True):
    """Uma união determinística compartilhada pela prévia e pelo transporte."""
    cc_list = []
    for raw in values:
        for value in re.split(r"[;,/\n\r]+", str(raw or "")):
            value = value.strip().lower()
            if not value:
                continue
            if strict and not valid_email(value):
                raise ValueError(f"E-mail de cópia inválido: {value}")
            if value not in cc_list and value != str(to_email or "").strip().lower():
                cc_list.append(value)
    return sorted(cc_list)


def send_email(cfg, valid_email, to_email, cc_value, subject, html_body, text_body, attachments=None):
    validate_config(cfg, valid_email)
    original_to = str(to_email or "").strip()
    original_cc = str(cc_value or "")
    if not valid_email(original_to):
        raise ValueError("Empresa sem e-mail principal válido.")
    cc_list = normalize_cc(valid_email, original_to, [original_cc])
    if cfg["test_mode"]:
        to_list, cc_list = [cfg["test_email"].strip()], []
        subject = f"[TESTE → {original_to}] {subject}"
        banner = f"MODO DE TESTE — Destino original: {original_to}; CC: {original_cc or '—'}"
        html_body = f"<div style='padding:10px;background:#fff3cd;margin-bottom:15px'>{html.escape(banner)}</div>" + html_body
        text_body = banner + "\n\n" + text_body
    else:
        to_list = [original_to]
    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"] = formataddr((cfg.get("sender_name", ""), cfg["sender_email"]))
    msg["To"] = ", ".join(to_list)
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=cfg["sender_email"].split("@")[-1])
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(text_body, "plain", "utf-8"))
    alt.attach(MIMEText(html_body, "html", "utf-8"))
    msg.attach(alt)
    for attachment in attachments or []:
        filename = attachment.get('filename') or Path(attachment.get('path', '')).name
        if attachment.get('generated') == 'complementary_report':
            data = attachment.get('data')
        else:
            path = Path(attachment["path"])
            if not path.is_file():
                raise ValueError(f"Anexo não encontrado: {filename}")
            data = attachment.get("data")
            if data is None:
                data = path.read_bytes()
        if not data:
            raise ValueError(f"Anexo vazio: {filename}")
        part = MIMEBase(*attachment.get('mime_type', 'application/octet-stream').split('/', 1))
        part.set_payload(data)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", "attachment", filename=filename)
        msg.attach(part)
    recipients = to_list + cc_list
    serialized = msg.as_string()
    if len(serialized.encode("utf-8")) > 25 * 1024 * 1024:
        raise ValueError("A mensagem com os anexos ultrapassa 25 MB. Reduza os arquivos antes de enviar pelo Gmail.")
    server = None
    try:
        if cfg["security"] == "ssl":
            server = smtplib.SMTP_SSL(cfg["host"], 465, timeout=45)
        else:
            server = smtplib.SMTP(cfg["host"], 587, timeout=45)
            server.ehlo()
            server.starttls()
            server.ehlo()
        server.login(cfg["username"], cfg["password"])
        try:
            refused = server.sendmail(cfg["sender_email"], recipients, serialized) or {}
        except (smtplib.SMTPRecipientsRefused, smtplib.SMTPSenderRefused, smtplib.SMTPDataError):
            # Recusa explícita antes do aceite: repetir é seguro.
            raise
        except Exception as exc:
            raise DeliveryUncertain("Conexão interrompida durante o envio. Confira a pasta Enviados antes de reenviar; o aceite pode ter ocorrido.") from exc
        return {
            "test_mode": bool(cfg["test_mode"]), "to_email": to_list[0], "subject": subject,
            "accepted": [r for r in recipients if r not in refused],
            "refused": {str(r): str(reason) for r, reason in refused.items()},
            "primary_accepted": to_list[0] not in refused, "message_id": msg["Message-ID"],
        }
    finally:
        if server is not None:
            try:
                server.quit()
            except Exception:
                try:
                    server.close()
                except Exception:
                    pass
