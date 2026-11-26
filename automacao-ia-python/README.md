# Classificador de observações com IA (Python + Claude)

Transforma **observações em texto livre** (o que o técnico de campo escreveu sobre um equipamento) em **dados estruturados**: categoria padronizada, data prometida, se precisa de ação e qual, resumo e nível de confiança.

> Problema real: centenas de observações por semana, cada uma escrita de um jeito ("cliente não estava", "ninguém atendeu", "volto dia 28"). Ler e classificar uma por uma toma horas. Com a classificação automática, o time só olha o que tem **"Precisa de ação = Sim"** ou **confiança baixa**.

## Exemplo

| Observação (entrada) | Categoria | Data prometida | Precisa de ação |
|---|---|---|---|
| cliente nao estava em casa, vou tentar de novo dia 28/11 | Cliente não localizado | 28/11/2026 | Sim |
| Agendado retirada para 02/12 com o cliente | Aguardando retirada | 02/12/2026 | Não |
| Cliente disse que o modem foi roubado junto com a TV | Extraviado / perdido | | Sim |
| devolvido no almoxarifado com caixa e fonte | Entregue no almoxarifado | | Não |

Planilha completa de exemplo: [`exemplo/observacoes_tecnicos.xlsx`](exemplo/observacoes_tecnicos.xlsx) → [`exemplo/observacoes_tecnicos_classificado_DEMO.xlsx`](exemplo/observacoes_tecnicos_classificado_DEMO.xlsx) (abas `CLASSIFICADO` e `RESUMO`).

## Como funciona

1. Lê a planilha e a coluna de texto livre.
2. Envia as observações **em lotes** para o Claude, com um prompt que fixa as categorias permitidas ([`categorias.json`](categorias.json)) e exige resposta em JSON.
3. **Valida** a resposta: se a IA devolver uma categoria fora da lista ou um JSON inválido, a linha cai nas regras de segurança e fica marcada na coluna `Origem`.
4. Gera um Excel com destaque para "Precisa de ação" e confiança abaixo de 0,6, e uma aba de resumo.

Sem chave de API, roda em **modo demonstração**, com regras por palavra-chave, para qualquer pessoa testar.

## Rodando

```bash
pip install -r requirements.txt

# modo demonstração (sem chave)
python classificar.py exemplo/observacoes_tecnicos.xlsx

# com IA
export ANTHROPIC_API_KEY="sua-chave"        # Windows (PowerShell): $env:ANTHROPIC_API_KEY="sua-chave"
python classificar.py exemplo/observacoes_tecnicos.xlsx

# na sua planilha
python classificar.py minha_planilha.xlsx --coluna "Observação" --saida resultado.xlsx
```

Para adaptar a outro processo, basta trocar a lista em `categorias.json` (por exemplo, motivos de chamado, tipos de reclamação ou status de pedido).

## Por que IA, e não só regras?

A planilha `_DEMO` foi gerada **pelas regras** (sem IA) e mostra o limite delas:

| Observação | Regras | O que a IA entende |
|---|---|---|
| "equipamento está comigo no carro, **devolvo amanhã** na base" | Entregue no almoxarifado ❌ | Com técnico (a devolução ainda **vai** acontecer) |
| "cliente mudou de endereço" | Outros ❌ | Cliente não localizado |

Regras olham palavras; a IA entende **tempo verbal e contexto**. As regras continuam no código como rede de segurança, e a coluna `Origem` mostra quem classificou cada linha.

## Cuidados

- **Dados sensíveis:** antes de enviar dados reais a uma API externa, confirme a política da empresa. Remova nomes, CPF e endereços; o classificador só precisa do texto da observação.
- **Revisão humana:** linhas com confiança baixa ou categoria "Outros" devem ser revisadas.

---
Autor: **Renan Vicente Braga** · [voltar ao portfólio](../README.md)
