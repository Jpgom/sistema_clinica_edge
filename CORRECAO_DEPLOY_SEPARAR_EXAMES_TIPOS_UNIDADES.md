# Correção de deploy - Separar exames

Correção aplicada em `separar_exames_app/app.py`.

## Problema
O deploy falhava na inicialização do Gunicorn por erro de sintaxe em uma f-string:

```python
result["competency"] = f"{unit["name"]} - {month:02d}/{year}"
```

## Correção
A string foi ajustada para não conflitar as aspas internas:

```python
result["competency"] = f"{unit['name']} - {month:02d}/{year}"
```

## Validação
Executado `python -m compileall -q .` para conferir sintaxe dos arquivos Python do projeto.
