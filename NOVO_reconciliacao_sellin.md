# Reconciliação do sell-in (por que CL parecia incompleto)

## Diagnóstico
O sell-in NÃO é filtrado nem cortado: o ler_sellin carrega 100% do f_SELLIN.xlsx.
Com 1 aba, tudo entra no total. O filtro "farmacia" é só do SELL-OUT.
Logo, o total bruto do sell-in (inclusive CL) está completo no dashboard.

O "CL não trouxe 100%" quase sempre é CRUZAMENTO: produto de CL sem par no
DePara_Produtos.xlsx não casa com sell-out/target, então aparece incompleto nas
seções cruzadas (targets, SI×SO, Histórico SKU com SO) — mas está no total de sell-in.

## O que foi adicionado
Ao rodar o Python (GERAR_RECONCILIACAO = True), sai um Excel:
  sellin_reconciliacao.xlsx
com 4 abas:
  1. Por Franquia      -> total (Unid/BRL/USD) por franquia + TOTAL GERAL.
                          CONFIRA O CL AQUI: compare com a sua fonte. Se bater, o
                          carregamento está 100% e o gap é de cruzamento (aba 3).
  2. CL - Produtos     -> detalhe dos produtos de CL com valores.
  3. Sem DePara (incluir) -> produtos de sell-in SEM par no DePara. É o "faltante":
                          a coluna ONDE_INCLUIR diz como adicionar no DePara_Produtos.xlsx.
  4. Linhas sem ANO-MES -> linhas que não somam por período (se houver).

## Como usar
1. Rode o Python. Abra sellin_reconciliacao.xlsx.
2. Veja "Por Franquia": o total de CL bate com a sua fonte?
   - BATE  -> carregamento ok; vá na aba "Sem DePara" e inclua os produtos de CL
             no DePara_Produtos.xlsx para eles cruzarem com sell-out/target.
   - NÃO BATE -> me diga o número que falta; investigamos a origem (formato do arquivo).

## Config
GERAR_RECONCILIACAO = True         # liga/desliga
PATH_RECONCILIACAO = "sellin_reconciliacao.xlsx"
