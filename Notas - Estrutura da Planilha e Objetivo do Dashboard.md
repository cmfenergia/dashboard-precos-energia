# Notas do projeto — Dashboard de Preços de Energia (CMF Energia)

> Backup do entendimento consolidado com o Claude sobre a planilha base, para não se perder caso a conversa/IDE reinicie.

## Objetivo do dashboard

Construir um dashboard que mostra a evolução dos preços de energia no mercado livre ao longo do tempo, voltado a consumidores livres (ou em processo de migração), priorizando gráficos/visualizações em vez de tabelas de números.

Fonte de dados base: `Arqquivo Base\Curva Forward - CMF Energia.xlsx`

---

## Aba "Forward"

Dados começam na linha 86 (~01/01/2022). Coluna C = data de referência (mês/ano) de cada linha.

- **Colunas A/B** — IGPM e IPCA do mês da linha.
- **Colunas K:V** — preço da energia negociada para o ano do cabeçalho (K1=2022 ... V1=2033), considerando os meses *restantes* daquele ano a partir da data da linha.
  - Ex.: K86 (01/01/2022) = 234 → preço do ano cheio jan-dez/2022.
  - K90 (01/05/2022) = 170 → preço de mai-dez/22.
  - K97 (01/12/2022) = 91 → preço de só dez/22.
  - A partir do mês em que o ano de referência já é passado, a célula fica vazia.
- **Coluna W** — "Preços Mercado Livre (I50% - 36 meses futuros)": média dos preços dos 36 meses seguintes à data da linha.
- **Coluna X** — média histórica da série W desde o início da amostra até a linha atual.
- **Colunas Y/Z** — mesmo conceito de W/X, mas para 60 meses futuros (5 anos).
- **Colunas AA/AB/AC** — ágio da energia no cativo nas bandeiras Amarela, Vermelha I e Vermelha II (acréscimo sobre o preço cativo tradicional/bandeira verde).
- **Colunas AD:AG** — break even do mercado livre vs. cativo em cada bandeira (Verde, Amarela, Vermelha I, Vermelha II). É o preço em que o consumidor livre não tem nem economia nem prejuízo frente ao cativo naquela bandeira. Pagar menos = economia; mais = prejuízo.
- **Coluna AH** — "% Economia (Livre x Cativo)" = (AD − W) / AD (usa o break even da bandeira Verde, mais conservador, vs. média dos 36 meses futuros).
- **Colunas AI:AM** — preço do ano corrente relativo (A+0 a A+4) na data da linha. A+0 = ano atual (dinâmico, hoje 2026), A+1 = atual+1, etc. Os headers da planilha se atualizam sozinhos (ex. "A+0 (atual 2026)").
  - Ex.: AK111 (linha com data 01/02/2024, ano base 2024, A+2 = 2026) = preço de 2026 negociado em 01/02/2024 = 207.

---

## Aba "Atac-I50-SECOeS"

Repete a mesma dinâmica de preços ano a ano da aba "Forward", mas aplicando correção monetária pelo IPCA (coluna B da própria aba), trazendo tudo para o poder de compra da data-base mais recente com IPCA preenchido. A coluna A (IGPM) existe mas **não** é usada em nenhuma fórmula dessa aba.

Fórmula confirmada (ex. célula R122): `=Forward!R122*(1+SOMA($B122:$B$228))`.

A data-base do IPCA é dinâmica: conforme o usuário preenche novos meses na coluna B, todos os valores da aba se atualizam automaticamente. Em 02/07/2026, o último IPCA real preenchido era 01/05/2026 (linha 138); linhas seguintes (jun-ago/26) ficam com 0,00% como placeholder até serem preenchidas.

Propósito: comparar preços de datas diferentes em moeda constante (R$ de hoje).

### Os 4 tipos de gráfico da aba (a partir da coluna AS)

1. **Anexo 1 — "Mercado Cativo x Mercado Livre"**: compara as 4 curvas de break-even do cativo (Verde/Amarela/Vermelha I/Vermelha II) com o preço do mercado livre. Linha preta pontilhada = coluna Y (média 60 meses futuros); linha preta contínua = coluna Z (média histórica da série Y). Há um 2º gráfico logo abaixo (séries "Economia"/"Prejuízo", colunas AK/AL) e blocos de comentário laranja que devem ser **ignorados** (desatualizados).

2. **Anexo 2 — "Ano A+0 (Atual 2026)" até "Ano A+4 (Atual 2030)"** (5 gráficos): preço nominal x indexado por IPCA, por posição **relativa** ao ano corrente (A+0..A+4) — títulos e dados se deslocam automaticamente ano a ano.

3. **Anexo 3 — "Preço do 2026" até "Preço do 2030"** (5 gráficos): mesma lógica nominal x indexado IPCA, mas para **anos fixos/específicos**.

4. **Anexo 4 — "Modalidade Atacadista - Incentivada 50% - SE/CO"**: o gráfico mais importante, mais visto pelos clientes. Evolução do preço de cada ano específico (2022 a 2031, colunas K:T) ao longo do tempo, corrigido por IPCA. Anos não mais negociados (hoje: 2022-2025) aparecem tracejados; anos ainda em negociação aparecem em linha contínua.

---

## Próximos passos combinados
Seguir explicando as demais abas da planilha (`Varej-I50-SECOeS`, `Varej-I50-NE`, `Varej-Conv-SECOeS`, `Varej-Conv-N`, etc.) antes de definir a arquitetura do dashboard.