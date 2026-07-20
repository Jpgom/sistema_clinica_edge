# Atualização - Tipos de riscos editáveis

Incluído gerenciamento de tipos de riscos no módulo SST/PGR.

## O que mudou

- Nova tela: `/pgr/tipos-risco`
- Permite cadastrar novos tipos de risco.
- Permite editar nomes, ordem e cor hexadecimal no padrão `#RRGGBB`.
- A cor escolhida é usada como fundo das células nos laudos gerados.
- Ao alterar o nome de um tipo de risco, os riscos já cadastrados com o nome antigo são atualizados para o novo nome.
- A tela de cadastro/edição de riscos passa a carregar os tipos cadastrados no banco.
- Os tipos ergonômicos personalizados que tenham “ERGONÔMICO” no nome continuam sendo tratados como ergonômicos nas rotinas da AET.
