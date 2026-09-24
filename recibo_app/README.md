# Recibos no EDGE

O módulo web fica em `/recibos`, registrado como Blueprint em `edge_app/application.py`.
Ele usa a autenticação, as permissões e a proteção CSRF do EDGE.

`core.py` calcula os valores e desenha os campos em um PDF de **144,018 × 105,41 mm**.
A prévia é rasterizada a partir desse mesmo PDF com `modelo_preview.png` ao fundo.
O PDF de impressão contém somente os dados para o papel pré-impresso.

O arquivo `posicionamento.json` é o perfil inicial vindo do aplicativo original.
Os ajustes feitos no site são salvos por usuário em `DATA_DIR/recibos` (ou
`RENDER_DISK_PATH/recibos` em produção). O deslocamento de 35 mm da antiga
impressão direta do Windows não é aplicado ao PDF.

`app.py` é o aplicativo Tkinter original, preservado como referência; a versão
web usa `web.py` e `core.py`.

Ao imprimir, escolha o papel personalizado e **tamanho real / 100%**. A
alimentação e o driver da impressora podem deslocar fisicamente o papel;
confirme o alinhamento com um teste impresso antes de usar em produção.
