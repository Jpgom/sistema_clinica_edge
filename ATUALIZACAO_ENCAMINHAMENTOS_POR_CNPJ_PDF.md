# Atualização - Encaminhamentos por CNPJ e saída em PDF

## Alterações

- A tela principal de Encaminhamentos agora permite escolher o formato de saída: PDF ou Word (.docx).
- O formato padrão selecionado na tela é PDF.
- O sistema passa a gerar um ZIP principal chamado `encaminhamentos.zip`.
- Dentro do ZIP principal, cada empresa sai em um ZIP separado, nomeado somente com os números do CNPJ da empresa.
- Dentro de cada ZIP de empresa, os encaminhamentos ficam dentro de uma pasta também nomeada somente com o CNPJ.
- A base de encaminhamentos deve conter as colunas EMPRESA, CNPJ, NOME, CARGO e COMPLEMENTARES.
- Os nomes e dados principais são padronizados em maiúsculo.
- O PDF é gerado de forma nativa pelo sistema, sem depender de conversão externa por LibreOffice.

## Exemplo de estrutura

```text
encaminhamentos.zip
├── 07426369000100.zip
│   └── 07426369000100/
│       ├── ENCAMINHAMENTO NOME DO FUNCIONARIO.pdf
│       └── ENCAMINHAMENTO OUTRO FUNCIONARIO.pdf
├── 17973836000175.zip
│   └── 17973836000175/
│       └── ENCAMINHAMENTO NOME DO FUNCIONARIO.pdf
└── RESUMO_ENCAMINHAMENTOS.txt
```
