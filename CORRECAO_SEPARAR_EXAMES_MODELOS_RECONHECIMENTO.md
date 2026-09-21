# Correção — Separar exames: modelos e reconhecimento

Ajustes aplicados:

1. Corrigido o erro ao cadastrar modelos quando o Render usa Persistent Disk em `/var/data`.
   - Antes o sistema tentava gravar o caminho do modelo como relativo à pasta do código.
   - Como os PDFs modelo ficam em `/var/data/separar_exames/modelos`, isso gerava erro de `subpath`.
   - Agora os modelos são gravados com caminho absoluto seguro e continuam persistentes após deploy.

2. O processamento foi alterado para modo de confiabilidade máxima.
   - A triagem rápida foi desativada para evitar que páginas de exames sejam descartadas antes da leitura completa.
   - O sistema continua lendo todo o lote, sem parar quando parece que terminou.

3. Reduzidos falsos `NÃO ENCONTRADO`.
   - Quando a página tem funcionário e exame esperado, mas a evidência do tipo ficou baixa, ela vai para conferência em vez de desaparecer.
   - A varredura de segurança passou a aceitar mais casos com identidade forte do funcionário/empresa.

4. Mantida integração com o site EDGE.
