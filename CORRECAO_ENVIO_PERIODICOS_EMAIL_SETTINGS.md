# Correção - Envio Periódicos / Configuração de Gmail

Ajustes aplicados:

- Corrigida a tela `/envio-periodicos/settings` para evitar Internal Server Error ao salvar Gmail/SMTP.
- Validação robusta da porta SMTP.
- Normalização da segurança SMTP (`starttls`, `ssl` ou `none`).
- Senha de app do Google salva sem espaços digitados por cópia/cola.
- Mensagem de erro amigável na própria tela caso algum campo esteja inválido.
- Tratamento de erro 500 específico para a tela de e-mail.

Configuração recomendada:

- Servidor: `smtp.gmail.com`
- Porta: `587`
- Segurança: `STARTTLS`
- Usuário Gmail: e-mail completo
- Senha: senha de app do Google
- E-mail remetente: mesmo Gmail usado no usuário
