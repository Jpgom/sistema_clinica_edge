# Sistema Interno EDGE

Build Command:
pip install -r requirements.txt

Start Command:
gunicorn app:app

Funcionalidades:
- Relatórios
- Encaminhamentos
- Renumerador
- Recibo eSocial
- Físico e Mental
- Anamnese ocupacional
- Envio periódicos
- Separar exames
- Exames a prazo
- PGR / SST
- Recibos para papel pré-impresso (`/recibos`)

Observações:
- Encaminhamentos no servidor saem em .docx.
- Físico e Mental usa banco SQLite local (`fisico_mental.db`).
- Para PDF no Físico e Mental, o sistema tenta usar LibreOffice/soffice.
- Recibos gera a prévia e o PDF a partir do mesmo desenho. O PDF de impressão
  contém somente os dados, para uso no papel pré-impresso. A calibração é salva
  por usuário no diretório de dados do servidor. Veja `recibo_app/README.md`.
