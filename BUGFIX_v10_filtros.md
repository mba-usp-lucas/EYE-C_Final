# v10.15 - "Produtos Maior Gap SI×SO" agora cruza (match fuzzy)

## Problema
No Plano de Recuperação, a seção "Produtos · Maior Gap SI × SO" mostrava
"Sem produtos com sell-in e sell-out cruzáveis" mesmo havendo produtos.
Causa: cruzava SI×SO por nome EXATO normalizado; nos dados reais os nomes do
sell-out (IQVIA) e sell-in (IRIS) diferem (sufixos, apresentação), então não batia.

## Correção
Passou a usar o mesmo match flexível do resto do dashboard:
1) nome exato → 2) um nome contém o outro → 3) mesma 1ª palavra (4+ letras).
Assim DURYSTA cruza com "DURYSTA IMPLANTE 10MCG", SYSTANE ULTRA com
"SYSTANE ULTRA 15ML", etc. Vale no HTML e no PowerPoint (mesma função).

## Observação
É match aproximado (a mesma lógica que o dashboard já usa quando o
DePara_Produtos.xlsx não cobre 100%). Para cruzamento exato, mantenha o
DePara_Produtos.xlsx atualizado.
