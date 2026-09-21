# Correção - Separar exames - download do Arquivo de Exames

Ajuste aplicado no download em ZIP do **Arquivo de Exames**.

## Mudança

As pastas das empresas dentro do ZIP agora são criadas no padrão:

```text
CNPJ - NOME DA EMPRESA
```

O CNPJ é salvo somente com números para evitar que barras (`/`) criem subpastas indevidas dentro do ZIP.

Exemplo:

```text
ARQUIVOS SEPARADOS 09-2026/
└── 12345678000199 - EMPRESA EXEMPLO LTDA/
    └── AUDIOMETRIA/
        └── exame.pdf
```

Mesmo quando o download tiver somente uma empresa, o ZIP manterá a pasta da empresa para facilitar a conferência.
