# Duas versões geradas automaticamente

Ao rodar o sales_dashboard_v10.py, agora saem DOIS arquivos:

1. dashboard_sales_insightsv10.html          → COMPLETO (analista, você)
2. dashboard_sales_insightsv10_diretoria.html → DIRETORIA (resumido p/ compartilhar)

## O que muda na versão DIRETORIA
- Selo "VERSÃO DIRETORIA" no cabeçalho.
- SEM botões de exportação (PowerPoint, Excel, One-Pager, Insights).
- SEM Decomposição Preço × Volume × Mix.
- SEM Diagnóstico por Canal (granular). A Evolução Systane É MANTIDA.
- Insights: mantém o panorama MACRO (Visão Brasil); esconde o detalhamento por franquia.
- Plano de Ação: mantém o resumo executivo; esconde o detalhamento operacional.
- Gráficos e KPIs: MANTIDOS (evolução, franquia, tipo de cliente, targets).
- Menu: links das seções ocultas são removidos automaticamente.

## Como funciona
Um flag no HTML (var MODO_DIRETORIA) controla tudo por CSS. O gerador Python
liga o flag na 2ª saída — mesma base de dados, sem duplicar manutenção.

## Uso
Rode o Python normalmente (um run só gera os dois). Compartilhe o arquivo
"_diretoria" com a diretoria; use o completo no seu dia a dia.
