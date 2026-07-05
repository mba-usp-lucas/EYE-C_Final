# Novo: Histórico por SKU (Excel) — tabelona

## Onde
Botão "Histórico SKU (Excel)" na barra de exportação (ao lado de Exportar Insights).

## O que gera
Um Excel com UMA LINHA POR SKU (Franquia × Produto) e as colunas:
- Franquia, Produto
- Sell-out mês a mês dos últimos 3 anos (SO jan/24 ... SO dez/26)
- Sell-in mês a mês dos últimos 3 anos (SI jan/24 ... SI dez/26)
- Target 2026 (soma do ano)
- Realizado SI 2026 (acumulado YTD)
- YTGO (Saldo p/ Target) = Target 2026 − Realizado
- Atingimento %

## Detalhes
- A métrica segue o seletor do dashboard: Unidades ou Reais (o nome da aba e do
  arquivo indica qual). Para a outra métrica, troque no dashboard e exporte de novo.
- Sell-out casa com sell-in/target pelo nome normalizado do produto (ignora sufixos
  como "(NVR)" e parênteses), igual ao resto do dashboard.
- Meses futuros (sem dado) ficam em branco.
- Já vem com AutoFiltro e larguras de coluna ajustadas. (O congelamento de painel
  não é suportado na gravação via navegador; se quiser, use Exibir > Congelar Painéis
  na coluna C para travar Franquia/Produto.)
- Respeita o filtro de canal de sell-out ativo no dashboard.

## Observação
Se aparecer aviso de "SheetJS não carregada", é o bloqueio de CDN do PC corporativo:
use o xlsx.mini.min.js da pasta libs_local/ ao gerar o HTML (igual aos outros exports Excel).
