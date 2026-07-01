# Automação SST - PGR, PCMSO e LTCAT

Sistema Flask para cadastro de riscos, setores/cargos, empresas, exames e geração de laudos em Word.

## Funcionalidades

- Cadastro de riscos ocupacionais com dados para PGR, PCMSO, Plano de Ação e LTCAT.
- Importação de riscos em massa por modelo Excel.
- Cadastro de setores/cargos organizados por grupos.
- Importação de setores/cargos em massa por modelo Excel, vinculando os setores ao grupo escolhido.
- Cadastro completo de empresas para preencher a Identificação da Empresa nos laudos.
- Cadastro de exames para o PCMSO.
- Geração de PGR completo.
- Geração de PCMSO completo.
- Geração de LTCAT completo automática: todos os riscos ambientais (FÍSICO, QUÍMICO e BIOLÓGICO) selecionados nos setores entram no LTCAT.
- Exportações avulsas: Plano de Ação, Inventário PGR, Riscos/Exames PCMSO, Relação Função x Atividade e Descritivo Setor.

## Rodar localmente

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Acesse:

```text
http://127.0.0.1:5000
```

## Deploy no Render

O projeto está pronto para Render com PostgreSQL.

Arquivos importantes:

- `render.yaml`
- `Procfile`
- `runtime.txt`
- `requirements.txt`

No Render, use Blueprint apontando para o repositório GitHub. O sistema usa `DATABASE_URL` quando existir; localmente usa SQLite automático em `instance/sst_riscos.db`.

## Atualizações

Para atualizar no Render, substitua os arquivos, depois rode:

```bash
git add .
git commit -m "atualizacao do sistema"
git push
```

O banco existente é preservado. As novas tabelas/colunas são criadas automaticamente ao iniciar o sistema.

## Atualização V20

- Cadastro rápido de risco dentro da tela **Gerar laudos**, sem recarregar a página.
- O novo risco é salvo via rota `/api/risco/novo` e aparece imediatamente nas listas de riscos por setor.
- Opção para marcar automaticamente o novo risco nos setores já selecionados.

## V21 - Grupos de riscos

- Nova aba **Grupos de riscos**.
- Permite criar pacotes de riscos, por exemplo: LIMPEZA, MANUTENÇÃO, ADMINISTRATIVO.
- Na tela **Gerar laudos**, cada setor agora permite aplicar um ou mais grupos de riscos.
- Ao marcar um grupo, todos os riscos daquele grupo são selecionados automaticamente para o setor.
- Mesmo usando um grupo, ainda é possível marcar riscos isolados ou desmarcar riscos específicos.
- A LTCAT continua automática: usa apenas os riscos ambientais selecionados no setor, ou seja, FÍSICO, QUÍMICO e BIOLÓGICO.


## V23 - Relatório Psicossocial preservado no Word

A junção PGR + AET + Relatório Psicossocial agora insere o PDF psicossocial como páginas-imagem dentro do DOCX final. Essa abordagem evita as quebras de layout da conversão PDF -> Word por texto/tabelas, preservando gráficos, tabelas, cores e espaçamentos exatamente como no PDF original. O conteúdo do psicossocial fica visualmente fiel, porém não editável como texto.

## V24 - Otimização da junção de arquivos

- A conversão visual do PDF psicossocial agora usa JPEG otimizado em DPI moderado, reduzindo o tamanho do DOCX final e o tempo de processamento.
- Os arquivos temporários da junção são apagados automaticamente após a geração, evitando acúmulo no disco do Render.
- O `Procfile` foi ajustado com timeout maior para evitar reinício do worker durante junções mais pesadas.
- A aba **Juntar arquivos** mostra aviso de processamento e bloqueia clique duplo enquanto o Word é gerado.
- Variáveis opcionais no Render:
  - `PSICOSSOCIAL_RENDER_DPI` padrão `135`.
  - `PSICOSSOCIAL_JPEG_QUALITY` padrão `82`.
  - `PSICOSSOCIAL_MAX_PAGES` padrão `80`.
  - `MAX_UPLOAD_MB` padrão `1024`.

## V25 - Junção sem recarregar e modelos atualizados

- A aba **Juntar arquivos** agora gera o Word por download assíncrono, mantendo a página aberta para juntar outros arquivos em seguida.
- A opção **PGR + AET + Psicossocial** dentro da tela **Gerar laudos** também mantém a página e as marcações preenchidas após o download.
- Os modelos completos de PGR, PCMSO e LTCAT foram substituídos pelos modelos enviados na atualização atual.

## V26 - melhorias adicionadas

- Módulo inicial de Ergonomia/AET: gera uma AET editável a partir da empresa, setores, cargos e riscos ergonômicos/psicossociais selecionados.
- Importação inteligente de laudos antigos: nova aba para enviar PGR/PCMSO/LTCAT antigo em DOCX ou PDF, extrair empresa, CNPJ, setores, riscos e exames e salvar no sistema após revisão.
- Tela de geração em etapas: wizard com barra lateral de progresso para Empresa, Dados, Setores, Riscos, Exames e Gerar.
- Motor de regras de exames: botão para sugerir exames automaticamente conforme riscos selecionados por setor.
- Pacote ZIP da empresa: gera PGR, PCMSO, LTCAT, AET e resumo da geração em um único ZIP.

## V28 - AET técnica e plano de ação agrupado

- AET com campos técnicos adicionais: tipo de documento, motivo da análise, condição ergonômica geral, fontes de dados, limitações da análise e diagnóstico por setor.
- Formulário da AET com opções de seleção por setor, reduzindo preenchimento manual.
- Presets por CNAE/atividade para pré-marcar respostas de AET, editáveis antes da geração.
- Plano de ação do PGR agora agrupa o mesmo risco em uma única linha e lista todos os setores na coluna GES.

## Correção V31 - Request Entity Too Large

A versão V31 aumenta os limites internos do Flask/Werkzeug para formulários grandes da tela Gerar Laudos.

Variáveis opcionais no Render:

- `MAX_UPLOAD_MB=1024`
- `MAX_FORM_MEMORY_MB=80`
- `MAX_FORM_PARTS=50000`

Use `MAX_FORM_MEMORY_MB` maior se houver muitas empresas/setores/riscos/exames/AET no mesmo envio.

## V32 - Importação inteligente e limite do formulário

- Importação de PGR/PCMSO/LTCAT antigo mais precisa para arquivos Word: extrai identificação da empresa, setores/cargos/CBO/funcionários/descrições e riscos do inventário.
- A tela de revisão da importação permite editar os dados antes de salvar.
- Limites padrão aumentados para evitar `Request Entity Too Large` ao gerar laudos com muitos setores, riscos, exames e PDF psicossocial.

Variáveis recomendadas no Render:

```text
MAX_UPLOAD_MB=1024
MAX_FORM_MEMORY_MB=512
MAX_FORM_PARTS=500000
```

## V34 - Modelos reutilizáveis de laudos antigos

- A importação inteligente agora salva um modelo reutilizável com setores, cargos, riscos e exames extraídos do laudo antigo.
- O modelo não fica preso à empresa original: ele pode ser aplicado em qualquer empresa cadastrada na tela Gerar laudos.
- Ao aplicar o modelo em uma empresa, o sistema cria/atualiza os setores, cargos, riscos e exames necessários e cria uma configuração de geração já marcada para revisão.

## V39 - Importação por banco e modelos importados

- A leitura de laudos antigos agora salva um rascunho no banco, evitando enviar JSON gigante em campos ocultos.
- A revisão da importação envia apenas o ID do rascunho, reduzindo erro de formulário grande.
- Modelos importados podem ser excluídos individualmente ou todos de uma vez.
- Ao aplicar modelo importado em uma empresa, o formulário é compactado no navegador e envia apenas empresa/modelo/grupo.
- A sugestão de exames também envia somente setores e riscos selecionados.


## V42 - CNAEs e modelos padrão de AET

- Adicionada a tela **Modelos AET/CNAEs** com catálogo de CNAEs de referência e modelos técnicos de AET por atividade.
- A etapa AET da tela **Gerar laudos** agora possui link para consultar os modelos disponíveis.
- O botão **Aplicar sugestões do CNAE** usa CNAE principal, CNAE secundário, descrição da atividade e nome da empresa para pré-preencher o formulário de AET.
- Foram adicionados modelos para comércio/vestuário, mercados, administrativo, limpeza, condomínios/portaria, restaurantes, clínicas, oficinas, construção civil, logística/transporte, escolas, estética, postos e atividades rurais.


## V44 - Recibos eSocial

- Nova aba **Recibos eSocial**.
- Converte a planilha RELFUNCGERAL (.xls exportado como HTML ou .xlsx) em PDF A4 paisagem.
- O PDF mantém todas as colunas principais na mesma folha e usa somente: EVENTO, empresa, NOME, CPF, TIPO, STATUS, DATA, Recibo eSocial e Recibo Sefaz.
- A coluna unidade e demais colunas extras são ignoradas automaticamente.


## V45 - Correção Recibos eSocial

- Corrigido erro: `'Page' object has no attribute 'get_text_length'`.
- A conversão RELFUNCGERAL para PDF em A4 paisagem foi ajustada para usar medição de texto compatível com PyMuPDF.
- Mantidas somente as colunas EVENTO, empresa, NOME, CPF, TIPO, STATUS, DATA, Recibo eSocial e Recibo Sefaz.


## V46 - Recibos eSocial robusto

- Corrige leitura de planilhas RELFUNCGERAL exportadas como Excel HTML com frameset.
- Quando o .xls só aponta para a pasta `_arquivos/sheet001.htm`, o sistema exibe orientação clara em vez de falhar silenciosamente.
- A aba Recibos eSocial agora aceita `.zip` contendo o `.xls` e a pasta `_arquivos`, além de `.xls` completo e `.xlsx`.
- Adiciona suporte a `.xls` binário via `xlrd`.

## V47 - Correção do Inventário de Riscos

- Corrigida a última parte da tabela do Inventário de Riscos no PGR.
- As linhas fixas "CONTROLES EXISTENTES NO GES E SUA EFICÁCIA" e "Monitoramento da saúde do trabalhador através de exames ocupacionais." não são duplicadas.
- A frase "NENHUM FATOR DE RISCO PSICOSSOCIAL FOI IDENTIFICADO PARA A SETOR" aparece apenas quando não houver risco psicossocial e fica como última linha da tabela.
