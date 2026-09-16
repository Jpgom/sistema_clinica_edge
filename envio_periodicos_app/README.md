# EDGE - Convocação de Exames Periódicos V5

Sistema web para organizar competências de exames periódicos por unidade, enviar convocações pelo Gmail, anexar encaminhamentos e comparar comparecimento.

## V5 — proteção operacional e acompanhamento em tempo real

A V5 foi preparada para reduzir erros de operação por clique duplo, atualização da página, fechamento da aba, reenvio acidental, upload repetido e interrupções durante o envio.

### Envio de e-mails em segundo plano
- ao clicar em **Enviar pendentes** ou **Enviar lembretes**, o sistema cria uma fila persistente no banco;
- o envio continua no servidor mesmo se o usuário fechar a janela de progresso, atualizar a página ou fechar a aba;
- ao voltar para a competência, um envio ainda ativo é detectado e o andamento pode ser reaberto;
- somente um processo de envio pode ficar ativo por competência;
- clique duplo ou repetição da mesma solicitação durante um envio reaproveita o processo já existente em vez de criar outro;
- empresas que já possuem envio confirmado não são reenviadas automaticamente.

### Tela de progresso
Durante o envio, a tela mostra:
- estado do processo;
- percentual concluído;
- quantidade de e-mails enviados;
- quantidade de empresas atendidas;
- erros ou envios que precisam de revisão;
- destinatário/grupo que está sendo processado naquele momento;
- eventos recentes do processo.

### Interrupção crítica do SMTP
Se o servidor for interrompido exatamente durante a comunicação com o Gmail, o sistema não tenta reenviar automaticamente um e-mail cujo resultado possa ser incerto. Esse grupo fica marcado como **REVISAR** e o usuário deve conferir a pasta **Enviados** do Gmail antes de optar por um reenvio manual.

### Uploads idempotentes
- subir exatamente a mesma base da competência novamente não duplica colaboradores;
- a mesma planilha de comparecimento não é processada duas vezes;
- o mesmo encaminhamento não é armazenado duas vezes para a mesma empresa/competência;
- a importação da base e do controle de comparecimento usa transações: ou termina de forma consistente, ou é revertida;
- ZIPs são validados antes de serem processados e possuem limites de quantidade/tamanho descompactado para evitar arquivos anormais.

### Proteção durante envio
Enquanto uma competência está enviando e-mails, alterações que poderiam mudar o conteúdo do envio ficam bloqueadas, incluindo base, empresas, colaboradores, encaminhamentos e comparecimento. As configurações do Gmail também não podem ser alteradas durante um envio ativo.

### Cliques repetidos e navegação
- formulários POST bloqueiam o segundo clique enquanto a primeira operação está sendo executada;
- uploads e salvamentos mostram uma tela de processamento;
- voltar pelo navegador reabilita corretamente os controles;
- ações de ativar/desativar usam o estado desejado, em vez de simplesmente inverter, evitando que uma repetição reverta a ação;
- criação de competência possui proteção contra duas requisições simultâneas para o mesmo mês/unidade.

### Recuperação de ações destrutivas
Antes de excluir dados importantes, o sistema exige a criação bem-sucedida de um backup automático do SQLite. Se o backup falhar, a exclusão é cancelada.

Encaminhamentos removidos vão primeiro para uma **lixeira interna** e são mantidos por até 30 dias. O sistema conserva os 30 backups automáticos mais recentes do banco em `data/backups`.

### SQLite mais tolerante a concorrência
A V5 usa WAL, `busy_timeout` e transações de escrita nas operações críticas para reduzir falhas do tipo `database is locked` e manter consistência quando mais de uma ação ocorre próxima no tempo.

## Regras de duplicidade mantidas

### Para convocação/e-mail
- o mesmo colaborador aparece uma única vez por empresa;
- a relação enviada no e-mail contém somente **COLABORADOR**.

### Para a planilha de encaminhamentos
- mesmo NOME + mesmo CARGO aparece uma única vez;
- o mesmo NOME em CARGOS diferentes permanece em linhas separadas;
- a coluna **COMPLEMENTARES** fica em branco.

## Recursos mantidos
- painel por unidade;
- BELÉM e MACAPÁ criadas automaticamente;
- cadastro de novas unidades;
- competências separadas por unidade;
- edição de competência;
- cadastro/edição e exclusão múltipla de empresas;
- Gmail via SMTP e senha de app;
- empresas com o mesmo e-mail agrupadas em uma única mensagem;
- encaminhamentos associados por CNPJ;
- renomeação automática para `ENCAMINHAMENTO PARA EXAMES (COMPETÊNCIA) - EMPRESA - CNPJ.zip`;
- upload manual de encaminhamento;
- conferência do dia 20 e lembretes de faltantes;
- histórico de envios simplificado.

## Atualizar sem perder dados
1. Feche a versão anterior.
2. Faça uma cópia de segurança da pasta `data` atual.
3. Extraia os arquivos da V5.
4. Mantenha a pasta `data` da instalação anterior.
5. Inicie a V5 normalmente.

Na primeira abertura, as novas tabelas e índices operacionais são criados automaticamente. Empresas, Gmail, competências, histórico e encaminhamentos existentes são preservados.

## Executar no Windows
Dê dois cliques em `INICIAR_SISTEMA.bat`.

A primeira execução instala as dependências e abre o sistema em:

`http://127.0.0.1:5000`

## Render
A configuração incluída usa um único processo Gunicorn com múltiplas threads, de modo que exista somente um worker interno responsável pela fila persistente de e-mails:

`gunicorn --workers 1 --threads 4 --timeout 180 app:app`

Mantenha um disco persistente montado em `/var/data`.

## Gmail
- servidor: `smtp.gmail.com`
- porta: `587`
- segurança: `STARTTLS`
- usuário: Gmail completo
- senha: senha de app do Google

Use o modo de teste antes do primeiro disparo real de cada configuração.

## Observação importante
Nenhum sistema pode prometer falha operacional absolutamente zero. A V5 foi desenhada para tornar as principais ações repetíveis e recuperáveis e, principalmente, evitar duplicidade de e-mail quando o resultado de um envio é incerto.


## Ajustes V5.1
- A criação e a adição de base em uma competência não baixam mais automaticamente a planilha para encaminhamentos. O download ocorre somente pelo botão **Planilha para encaminhamentos** / **Baixar planilha para encaminhamentos**.
- A competência pode ser **excluída definitivamente** em Opções avançadas. A ação exige digitar `EXCLUIR`, apaga também histórico de envios e arquivos de encaminhamento daquela competência e não utiliza a lixeira interna. O cadastro geral das empresas não é apagado.
