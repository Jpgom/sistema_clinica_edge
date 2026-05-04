# Correção do Renumerador de Recibos

## Problema corrigido

Alguns arquivos DOCX reais possuem o número do recibo em estruturas diferentes do padrão. O caso mais crítico encontrado foi quando o número aparece grudado ao texto do total, por exemplo:

```text
TOTAL R$: 150,00TOTAL R$: 150,00142
```

Nesse caso, a versão anterior não detectava o número `142`, pois esperava que o número estivesse sozinho em um parágrafo.

## O que foi alterado

- Detecção de recibos após cada `NOTA DE BALCÃO`.
- Suporte a múltiplos recibos dentro do mesmo DOCX.
- Suporte a números isolados, como `00142`.
- Suporte a números grudados ao total, como `TOTAL R$: 150,00142`.
- Substituição segura por intervalo de texto dentro do XML do Word.
- Preservação máxima de formatação, tabelas e runs do Word.
- Relatório com avisos quando algo inconsistente for detectado.
- Validação para comparar quantidade de datas alteradas com recibos renumerados.

## Arquivo principal alterado

- `edge_app/application.py`

## Resultado esperado

Os recibos enviados em lote devem ser detectados e renumerados mesmo quando:

- há mais de um recibo no mesmo arquivo DOCX;
- o número está separado em vários runs do Word;
- o número está grudado ao texto do total;
- o documento possui tabelas, quebras ou formatação interna diferente.

## Observação importante

O sistema agora evita substituição global de números. Ele só altera o número localizado dentro do bloco de recibo identificado após `NOTA DE BALCÃO`, reduzindo o risco de alterar CNPJ, datas, valores ou telefones.
