# VarejoMax Brasil — Inteligência e Análise de Dados para Varejo e E-commerce

> Projeto Prático **MA2026** · FATEC SP · Orientação: Prof. Veríssimo

Plataforma de análise de dados construída de forma evolutiva, que parte de uma planilha bruta e inconsistente de vendas e chega a um **dashboard interativo no Power BI**, passando por Excel avançado, Power Query, Power Pivot/DAX, VBA, SQL e consumo de APIs REST.

> **Nota:** a VarejoMax Brasil é uma empresa **fictícia** e todos os dados são **simulados**.

---

## Sumário

- [Contexto](#contexto)
- [Problema e pergunta central](#problema-e-pergunta-central)
- [Objetivos](#objetivos)
- [Arquitetura e pipeline](#arquitetura-e-pipeline)
- [Fontes de dados](#fontes-de-dados)
- [Modelo dimensional](#modelo-dimensional)
- [Indicadores (KPIs)](#indicadores-kpis)
- [Premissas de tratamento](#premissas-de-tratamento)
- [Estrutura do repositório](#estrutura-do-repositório)
- [Cronograma e entregas](#cronograma-e-entregas)
- [Resultados esperados](#resultados-esperados)
- [Status do projeto](#status-do-projeto)
- [Equipe](#equipe)

---

## Contexto

A **VarejoMax Brasil** é uma rede omnichannel que vende pelo site, pelo aplicativo, por marketplace e em lojas físicas (São Paulo, Rio de Janeiro, Belo Horizonte, Curitiba, Porto Alegre e Brasília).

| Item | Detalhe |
|---|---|
| Catálogo | 49 produtos em 6 categorias |
| Volume | ~9 mil pedidos (~13,7 mil linhas de itens, 30 colunas) |
| Período | janeiro/2025 a agosto/2026 |
| Abrangência | clientes de 24 cidades brasileiras |

Os dados chegam em **uma única planilha bruta** exportada do sistema, misturando pedido, cliente, produto, pagamento e logística em cada linha. Não há padronização, e os problemas incluem:

- datas gravadas como texto;
- categorias e cidades escritas de formas diferentes;
- linhas duplicadas;
- valores que não batem com preço × quantidade;
- CEPs sem o zero à esquerda;
- registros de teste.

Como consequência, cada área chega a números diferentes e não existe uma visão consolidada por canal, região, produto ou transportadora.

## Problema e pergunta central

A gestão não consegue cruzar, de forma confiável, **vendas, logística e comportamento dos clientes**.

> **Como os dados de vendas, logística e clientes podem apoiar decisões no varejo e no e-commerce?**

### Perguntas de apoio

- Quais produtos, categorias e canais concentram faturamento e margem, e como isso evolui ao longo do tempo?
- Onde e com quais transportadoras o prazo real excede o previsto, e qual o efeito sobre avaliação e devoluções?
- Como se comportam os clientes novos e os recorrentes (ticket médio, frequência e recompra)?
- Câmbio e clima se relacionam com as vendas de determinadas categorias e com os atrasos de entrega?

### Hipóteses a validar

1. O atraso na entrega reduz a nota do cliente.
2. Regiões mais distantes do centro de distribuição concentram os maiores atrasos.
3. O desempenho varia entre transportadoras e entre canais.
4. Picos de demanda, como a Black Friday, pioram o prazo.

## Objetivos

**Geral:** construir uma plataforma de inteligência e análise de dados para a VarejoMax Brasil, conectando fontes heterogêneas (planilhas, APIs externas e banco relacional) a um dashboard interativo no Power BI.

**Específicos:**

1. **Excel avançado:** estruturar, padronizar e auditar a planilha `Vendas_Bruto` com PROCX, validações e colunas de auditoria, sem erros de fórmula, registrando cada inconsistência e a decisão de tratamento.
2. **Power Query (M):** limpar, transformar e consolidar as fontes (tipos, textos, datas, duplicatas, registros de teste e valores ausentes).
3. **Power Pivot e DAX:** modelar um Star Schema e criar KPIs (faturamento, ticket médio, margem, prazo, atraso e recompra).
4. **VBA:** automatizar relatórios operacionais, como o resumo periódico de vendas por categoria e o relatório de pedidos atrasados por transportadora.
5. **SQL:** modelar e implementar um banco relacional (DDL e DML) para consultas analíticas.
6. **APIs REST:** enriquecer a base com câmbio e clima, relacionando fatores externos às vendas e aos atrasos.
7. **Power BI:** desenvolver um dashboard com navegação interativa e filtros dinâmicos para apoio à decisão.

## Arquitetura e pipeline

```text
Dado Bruto → Excel Avançado → Power Query → Power Pivot/DAX → VBA → SQL → API/Web → Power BI
```

| Camada | Papel no projeto |
|---|---|
| Excel Avançado | Auditoria, validações e registro de inconsistências da planilha bruta |
| Power Query | Limpeza, padronização e consolidação (linguagem M) |
| Power Pivot / DAX | Modelo estrela e medidas analíticas |
| VBA | Macros de relatórios operacionais |
| SQL | Persistência relacional e consultas de agregação |
| API / Web | Enriquecimento com dados externos (câmbio e clima) |
| Power BI | Dashboard final para tomada de decisão |

## Fontes de dados

| Tipo | Origem | Finalidade |
|---|---|---|
| **Dataset principal** | Planilha bruta simulada (`.xlsx`), aba `Vendas_Bruto` (~13,7 mil linhas, 30 colunas, uma linha por item de pedido). Aba de apoio `Ref_Cidades` (região e coordenadas de 24 cidades). | Histórico de pedidos com cliente, produto, valores, pagamento, entrega, status e avaliação |
| **API: câmbio** | [PTAX do Banco Central do Brasil](https://dadosabertos.bcb.gov.br/) (serviço Olinda), retorno em JSON | Relacionar o dólar às vendas, ao ticket médio e à margem de eletrônicos |
| **API: clima** | [Open-Meteo Historical Weather API](https://open-meteo.com/en/docs/historical-weather-api), retorno em JSON | Cruzar precipitação por cidade e dia com o atraso das entregas |
| **Banco relacional** | PostgreSQL *(a confirmar)* | Persistência (DDL) e consultas analíticas (DML) |

**Por que câmbio?** Eletrônicos concentram grande parte do faturamento e têm a menor margem do catálogo, o que os torna sensíveis ao dólar. A cotação é ligada à `Dim_Calendario` pela data.

**Por que clima?** Para testar se a chuva ajuda a explicar atrasos (data real menos data prevista) por região. A precipitação é buscada pelas coordenadas de `Ref_Cidades`.

**Complementares:** a API de Localidades do IBGE pode padronizar cidade, UF e região. Como os CEPs da base são simulados, o ViaCEP servirá apenas para demonstrar o consumo de uma API, com CEPs reais de teste.

## Modelo dimensional

Modelo estrela com uma tabela fato e seis dimensões.

**Tabela fato — `Fato_Vendas`** (uma linha por item de pedido: `id_pedido` + item)
quantidade, desconto, valor do item, custo, frete, datas de pedido, envio, entrega prevista e entrega real, status e avaliação.

**Dimensões**

| Dimensão | Conteúdo |
|---|---|
| `Dim_Produto` | produto, categoria e marca |
| `Dim_Cliente` | cadastro e cidade |
| `Dim_Cidade` | UF, região e coordenadas |
| `Dim_Calendario` | atributos de data |
| `Dim_Canal` | canal e loja |
| `Dim_Transportadora` | transportadora |

## Indicadores (KPIs)

Implementados em **DAX** e conferidos contra consultas **SQL**.

| Indicador | Definição |
|---|---|
| Faturamento | Soma do valor dos itens dos pedidos não cancelados |
| Quantidade de vendas | Unidades vendidas e número de pedidos distintos |
| Ticket médio | Faturamento ÷ contagem distinta de pedidos |
| Margem bruta | Valor dos itens − (custo unitário × quantidade) |
| Produtos mais vendidos | Ranking por quantidade e por faturamento |
| Participação percentual | Peso de cada categoria e de cada canal sobre o total |
| Prazo de entrega | Média de dias entre envio e entrega real; % de entregas atrasadas (real > prevista) |
| Cancelamento e devolução | % de pedidos cancelados e devolvidos |
| Clientes | Novos e recorrentes, taxa de recompra e avaliação média (nota de 1 a 5) |

## Premissas de tratamento

- Pedidos **cancelados** não entram no faturamento.
- **Devoluções** são analisadas à parte.
- O **frete** é lançado apenas no item 1 de cada pedido e tratado separadamente do faturamento de produtos.
- **Nota de avaliação em branco** significa que o cliente não avaliou.

## Estrutura do repositório

```text
MA2026-projeto-pratico/
├── MA2026-entrega-parcial/
│   ├── 01-excel/            # Planilha auditada, PROCX e log de inconsistências
│   ├── 02-powerquery/       # Consultas em M e fontes tratadas
│   ├── 03-powerpivot/       # Modelo estrela e medidas DAX
│   └── 04-documentacao/     # Relatório parcial
└── MA2026-entrega-final/
    ├── 01-vba/              # Macros (.bas)
    ├── 02-bancodedados-sql/ # Scripts DDL e consultas DML
    ├── 03-api-web/          # Consumo das APIs de câmbio e clima
    └── 04-powerbi-dashboard/# Dashboard (.pbix)
```

## Cronograma e entregas

| Etapa | Foco técnico | Entregável principal | Aula final |
|---|---|---|---|
| Etapa 1 | Excel avançado, PROCX, auditoria e Power Query | Planilha limpa e tratada (`01-excel`, `02-powerquery`) com registro dos problemas e decisões | Aula 5 |
| Etapa 2 | Power Pivot (Star Schema) e DAX inicial | Modelo dimensional e medidas iniciais (`03-powerpivot`) | Aula 7 |
| **Parcial 1** | Defesa parcial e relatório | Apresentação oral (10–15 min), repositório congelado e relatório em `04-documentacao` | Aula 8 |
| Etapa 3 | Automação VBA | Macros de resumo de vendas e pedidos atrasados (`01-vba`) | Aula 11 |
| Etapa 4 | Modelagem de BD (DDL) e consultas (DML) | Script `.sql` com tabelas e consultas analíticas (`02-bancodedados-sql`) | Aula 13 |
| Etapa 5 | Consumo de API REST | Conexão documentada com câmbio e clima (`03-api-web`) | Aula 15 |
| Checkpoint | Validação técnica | Feedback do docente, sem nota isolada | Aula 16 |
| Etapa 6 | Dashboard Power BI e relatório | Painel interativo (`04-powerbi-dashboard`) | Aula 18 |
| **Final** | Entrega final e demonstração ao vivo | Defesa completa das 4 ferramentas ao vivo | Aula 20 |

## Resultados esperados

- Planilha auditada e tratada, com log dos problemas encontrados e das decisões de tratamento.
- Modelo estrela com uma tabela fato e ao menos cinco dimensões, com medidas DAX conferidas contra consultas SQL.
- Ao menos três macros VBA funcionais e documentadas.
- Banco relacional com script DDL e no mínimo dez consultas DML analíticas.
- Consumo de ao menos uma API pública, integrada ao modelo e documentada (a segunda API é meta opcional).
- Dashboard no Power BI com ao menos cinco páginas: Visão Executiva, Produtos e Categorias, Clientes, Logística e Canais.

## Status do projeto

- [ ] Etapa 1 — Excel avançado e Power Query
- [ ] Etapa 2 — Power Pivot e DAX
- [ ] Parcial 1 — Defesa e relatório parcial
- [ ] Etapa 3 — Macros VBA
- [ ] Etapa 4 — Banco de dados SQL
- [ ] Etapa 5 — APIs (câmbio e clima)
- [ ] Etapa 6 — Dashboard Power BI
- [ ] Entrega final

## Equipe

| Nome | Responsabilidade |
|---|---|
| _Nome do integrante_ | _Papel / etapas_ |
| _Nome do integrante_ | _Papel / etapas_ |
| _Nome do integrante_ | _Papel / etapas_ |

**Orientador / Docente:** Prof. Veríssimo

---

<sub>Projeto acadêmico da FATEC SP. Empresa e dados fictícios, sem relação com organizações reais.</sub>
