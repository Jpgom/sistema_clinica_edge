# Integração Sistema EDGE + PGR/SST no Render

Este pacote já une os dois projetos em um único Web Service Flask para o Render.

## Como ficou

- Sistema EDGE principal: `/`
- Módulo PGR/SST: `/pgr/`
- O card **PGR / SST** foi adicionado à página inicial e aos menus laterais do Sistema EDGE.
- O módulo PGR/SST usa o mesmo login do Sistema EDGE. Quem não estiver logado será enviado para `/login`.
- Usuário com cargo `visualizador` pode abrir o módulo PGR/SST, mas não pode cadastrar, editar ou gerar documentos por POST.
- O banco PostgreSQL do Render continua sendo usado pela variável `DATABASE_URL`. As tabelas do PGR/SST são criadas no mesmo banco, sem apagar as tabelas do Sistema EDGE.

## Arquivos principais alterados

- `app.py`: agora monta o Sistema EDGE em `/` e o PGR/SST em `/pgr`.
- `requirements.txt`: dependências dos dois projetos foram unificadas.
- `render.yaml`: timeout ajustado para geração de documentos maiores.
- `pgr_app/`: contém o projeto PGR/SST movido para dentro do Sistema EDGE.
- `edge_app/templates/*.html`: inclusão do link/card **PGR / SST**.

## Deploy no Render

1. Substitua o conteúdo do repositório `sistema_clinica_edge` por este pacote integrado.
2. Faça commit e push para o GitHub.
3. No Render, abra o Web Service `sistema_clinica_edge`.
4. Clique em **Manual Deploy** > **Deploy latest commit**.
5. Após o deploy, acesse:
   - `https://seu-site.onrender.com/` para o Sistema EDGE
   - `https://seu-site.onrender.com/pgr/` para o PGR/SST

## Variáveis de ambiente recomendadas no Render

Mantenha as variáveis que já existem no seu serviço. Para o módulo PGR/SST, estas ajudam em uploads e geração de documentos grandes:

```text
MAX_UPLOAD_MB=1024
MAX_FORM_MEMORY_MB=512
MAX_FORM_PARTS=500000
```

A variável `DATABASE_URL` precisa continuar apontando para o PostgreSQL do Render. A variável `SECRET_KEY` também deve permanecer igual, pois ela mantém o login funcionando entre os dois módulos.

## Observação sobre o banco gratuito

No print do Render aparece aviso de expiração do banco gratuito. Antes de colocar em produção de forma definitiva, faça backup ou atualize o banco para evitar exclusão automática no prazo informado pelo Render.
