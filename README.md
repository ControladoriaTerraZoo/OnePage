# One Page

Dashboard "One Page" (Consolidado / Varejo / Atacado) ao vivo, no mesmo
padrão do **Forecast 2026** e do **Painel de Tarefas**: uma planilha Google
Sheets (DRE Banco) alimenta uma API feita em Google Apps Script, e uma
página estática (`index.html`, publicada via GitHub Pages) consome essa API
e monta o relatório no navegador — sem precisar exportar/subir arquivo
nenhum a cada período.

```
Google Sheets (DRE Banco)  →  Apps Script (API, doGet)  →  index.html (GitHub Pages)
```

## 1. Formato da planilha DRE Banco

A planilha precisa ter estas abas:

| Aba | Conteúdo |
|---|---|
| `2026` | Igual ao arquivo atual: `Mês Ano \| Loja \| Canal \| DRE \| -2026 \| Orçado -2026`, com todas as unidades, **incluindo o Atacado** |
| `2025` | Mesmo formato, para `-2025 \| Orçado -2025` |
| `Atacado` | Base **Lynkz Gerencial Consolidado**, já totalizada por período: `DRE Consolidado \| Real <período> \| AV \| Orçado <período> \| AV \| AH \| Real <período-1> \| AV \| AH` (mesmo formato do arquivo que você já usa) |
| `USUARIOS` | `nome \| email \| senha \| papel \| ativo` — mesma aba de login usada no Forecast 2026 |

> **Atacado é pré-totalizado, não mensal.** Diferente das abas `2026`/`2025`
> (que têm uma linha por mês e por isso permitem qualquer recorte — 3M, 6M,
> 9M...), a aba `Atacado` guarda um total fixo por período nas colunas
> "Real ...". **A cada fechamento (6M, 9M, 12M etc.) atualize os valores
> dessa aba** com o novo período do Lynkz — o dashboard identifica a coluna
> certa pelo ano no cabeçalho ("Real 6M/2026" → ano 2026). Se a aba não
> tiver uma coluna "Real" para o ano selecionado no filtro, o card do
> Atacado avisa que os dados precisam ser atualizados em vez de mostrar
> número errado.

## 2. Publicar a API (Google Apps Script)

1. Abra a planilha DRE Banco → **Extensões > Apps Script**.
2. Apague o conteúdo padrão e cole o conteúdo de [`apps-script/Codigo.gs`](apps-script/Codigo.gs) deste repositório.
3. Confira a constante `SPREADSHEET_ID` no topo do arquivo — já vem preenchida com o ID da planilha DRE Banco atual; troque só se for usar outra planilha.
4. **Implantar > Nova implantação > Tipo: App da Web**
   - Executar como: **Eu (sua conta)**
   - Quem pode acessar: **Qualquer pessoa com o link**
5. Copie a URL gerada (termina em `/exec`).

A chave de acesso (`CHAVE_ACESSO`) já vem preenchida igual nos dois arquivos
(`Codigo.gs` e `index.html`) — troque as duas juntas se quiser gerar uma
nova.

## 3. Publicar o dashboard (GitHub Pages)

1. Cole a URL copiada no passo anterior na constante `API_URL` no topo do
   `<script>` de [`index.html`](index.html) (procure por
   `COLE_AQUI_A_URL_DO_APPS_SCRIPT`), commite e dê push.
2. No GitHub: **Settings > Pages > Source: Deploy from a branch**, branch
   `main`, pasta `/ (root)`. O GitHub publica em
   `https://controladoriaterrazoo.github.io/OnePage/`.

Pronto: a partir daí, atualizar o relatório é só **editar os valores na
planilha** — o dashboard sempre busca os dados ao vivo. Não é mais
necessário exportar `.xlsx` nem rodar nada localmente para o dia a dia.

## Como o dashboard calcula os números

- **Receita Bruta**, **Lucro Bruto Ajustado**: soma das linhas de topo da DRE Banco (Consolidado = todos os canais, Varejo = Canal `VAREJO`)
- **Margem Bruta** = Lucro Bruto Ajustado / Receita Líquida
- **EBITDA**: soma da linha `EBITDA`
- **Margem EBITDA** = EBITDA / Receita Líquida
- **Atacado**: lido diretamente da aba `Atacado` (Lynkz), não da DRE Banco — os dois divergem de propósito, por isso a nota "Atacado: Lynkz Gerencial Consolidado" no card.
- Cada métrica é comparada com o período selecionado em "Comparar com" (variação % e R$, e p.p. para as margens). Verde = variação positiva, vermelho = negativa.
- O filtro "Período" (3M, 6M, 9M...) é gerado automaticamente a partir de quantos meses existem na aba do ano mais recente.

## Login

Mesmo modelo do Forecast 2026: e-mail/senha validados no navegador contra a
aba `USUARIOS`. Não é uma autenticação forte (a senha trafega em texto na
planilha), só um controle de acesso simples — mantenha a URL do App da Web
e do GitHub Pages fora de divulgação pública se os dados forem sensíveis.

## Modo offline (sem depender da API)

Este repositório também mantém um gerador em Python que lê arquivos
`.xlsx` exportados manualmente e gera um `.html` estático — útil para
arquivar um período fechado ou rodar sem depender do Apps Script. Veja
[`scripts/generate_dre_report.py`](scripts/generate_dre_report.py) e o
relatório de exemplo em `reports/one-page-6M-2026.html`.

```bash
pip install openpyxl
python3 scripts/generate_dre_report.py \
  --file data/DRE_Aberta_Mensal.xlsx \
  --atacado-file data/DRE_Atacado.xlsx \
  --year 2026 --months 1-6 \
  --compare-year 2025 --compare-months 1-6 \
  --label "6 M/2026" \
  --output reports/one-page-6M-2026.html
```

## Estrutura

```
index.html                       # dashboard (GitHub Pages)
apps-script/Codigo.gs            # API (Google Apps Script, publicado a partir da planilha)
scripts/generate_dre_report.py   # gerador offline (.xlsx -> .html)
data/                             # .xlsx exportados para o modo offline (não versionado)
reports/                          # relatórios .html gerados no modo offline
```
