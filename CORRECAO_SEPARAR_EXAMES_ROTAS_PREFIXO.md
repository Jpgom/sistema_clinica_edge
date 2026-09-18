# Correção Separar exames - rotas internas

Corrigido erro 404 ao clicar nas abas internas da função Separar exames quando o módulo está montado em `/separar-exames`.

Ajustes:
- Links internos passaram a usar `url_for`.
- Chamadas JS `/api/...`, `/arquivo`, `/modelos`, `/download` agora respeitam o prefixo `/separar-exames`.
- Adicionado `EDGE_BASE_PATH` e helper `edgeUrl()` no `common.js`.
