# Correção Exames a Prazo - seleção de guias

Ajuste realizado para evitar o erro:

"Uma ou mais guias selecionadas não pertencem às bases carregadas. Recarregue as bases e tente novamente."

Causa identificada: algumas abas da planilha base possuem espaço invisível no final do nome, por exemplo "JANEIRO.2025 OK ". A tela exibia a guia sem esse espaço e, ao enviar o formulário, a validação por texto exato recusava a guia.

Correção:
- os nomes das guias são limpos antes de ir para o formulário;
- a validação aceita equivalência por texto normalizado e mês/ano;
- a leitura final continua resolvendo a guia real dentro da planilha base.
