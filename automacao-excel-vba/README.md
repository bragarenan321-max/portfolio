# Consolidador de Relatórios (Excel/VBA)

Junta várias extrações de sistema que chegam **com colunas em ordens diferentes e formatos bagunçados** e gera um **relatório único e padronizado**, com uma aba de conferência que mostra o que deu certo e o que precisa de atenção.

> Inspirado numa rotina real de trabalho: todo dia, várias extrações precisavam ser unidas no Power Query, coladas num painel, separadas e conferidas à mão. Esta é uma versão **genérica**, com dados fictícios, que qualquer empresa pode adaptar mudando apenas as abas de configuração.

## O problema

| Na prática, as extrações chegam assim | O consolidador entrega |
|---|---|
| Colunas em ordens diferentes em cada arquivo | Colunas casadas **pelo nome**, sem acento e sem diferenciar maiúsculas |
| Códigos como texto: `000000000041001196` | Número: `41001196`, que casa com as tabelas de apoio |
| Datas como texto com hora: `31/10/2026 18:48:46` | Data real: `31/10/2026` |
| Status escritos de 5 jeitos (`ENTREGUE `, `entregue`, `Entregue c/ atraso`) | Valor padronizado pela aba **REGRAS** |
| Valor do item em outra planilha | Busca automática na aba **BASE_PRECOS** |
| Conferência feita no olho | Aba **RESULTADO** com OK / ATENÇÃO / ERRO e a lista de códigos sem cadastro |

## Como funciona

```
ENTRADAS\*.xlsx  ──►  lê todos os arquivos  ──►  casa colunas pelo nome (aba CONFIG)
                                                    │
                     converte tipos (texto, número, data)
                     padroniza textos (aba REGRAS)
                     busca valores (abas BASE_*)
                     aplica fórmulas (ex.: Quantidade × Valor unitário)
                                                    │
SAIDAS\RELATORIO CONSOLIDADO - 26_11_2026 - 17h40.xlsx   +   aba RESULTADO (conferência)
```

Tudo é configurado em abas, **sem mexer no código**:

| Aba | Para que serve |
|---|---|
| `PARAMETROS` | pastas de entrada e saída, linha do cabeçalho, nome do relatório |
| `CONFIG` | quais colunas o relatório tem e de onde cada uma vem |
| `REGRAS` | tabela "de → para" para padronizar textos |
| `BASE_PRECOS` | exemplo de tabela de apoio para buscas |

Tipos aceitos na `CONFIG`: `TEXTO`, `NUMERO`, `DATA`, `REGRA`, `BUSCA`, `FIXO`, `ARQUIVO` (nome do arquivo de origem) e `FORMULA`.

## Teste em 2 minutos

1. Baixe a pasta [`exemplo`](exemplo/). Ela já traz 3 extrações "bagunçadas" de propósito (colunas trocadas, códigos com zeros, datas como texto, uma linha em branco, uma quantidade escrita "dez", um código sem cadastro e um status sem regra).
2. Abra `Consolidador.xlsx`, pressione **Alt+F11**, vá em **Arquivo → Importar arquivo** e escolha [`modConsolidador.bas`](modConsolidador.bas).
3. Salve como **Pasta de Trabalho Habilitada para Macro (.xlsm)** na mesma pasta.
4. Rode a macro **Diagnosticar** (Alt+F8) e veja a aba `RESULTADO`. Depois rode **Consolidar**.

O relatório sai em `SAIDAS`, com uma aba `DADOS` (tabela formatada) e uma aba `RESUMO` (linhas e valores por status). Na aba `RESULTADO`, os problemas plantados no exemplo aparecem sinalizados.

## Decisões de projeto

- **Leitura em memória (arrays), não célula a célula.** Milhares de linhas são processadas em segundos.
- **Casamento de colunas pelo nome normalizado.** Se o sistema mudar a ordem das colunas amanhã, nada quebra.
- **Chaves sem zeros à esquerda** nas buscas (`0000136988` = `136988`), o erro mais comum de PROCV entre sistemas.
- **Valores inválidos não somem.** São mantidos como vieram e contados na conferência, para alguém olhar.
- **Saída com data e hora no nome.** Nunca sobrescreve um relatório anterior.

## Uso de IA no desenvolvimento

O código foi desenvolvido com apoio de IA generativa (Claude). Eu defini o problema, as regras de negócio e os casos de borda a partir da rotina real, testei com os dados e corrigi o que deu errado (por exemplo, a conversão de datas em texto e os zeros à esquerda que quebravam as buscas). A IA acelerou a escrita do código; a especificação, os testes e as decisões são meus.

---
Autor: **Renan Vicente Braga** · [voltar ao portfólio](../README.md)
