# One Page

Gera o relatório "One Page" (Consolidado / Varejo / Atacado) a partir da
planilha **DRE Banco** (Google Sheets), no mesmo formato do modelo usado nas
apresentações mensais.

Planilha de origem: [DRE Banco](https://docs.google.com/spreadsheets/d/1qYlMyLFWbtqkGw7ingnFFAilDDYxDwzun0TGs8lBMr0)

## Como funciona

O script lê a aba "DRE Aberta Mensal" (uma aba por ano, ex: `2025`, `2026`,
...) com as colunas `Mês Ano | Loja | Canal | DRE | -<ano> | Orçado -<ano>` e
calcula, para o período informado:

- **Receita Bruta**
- **Lucro Bruto Ajustado**
- **Margem Bruta** = Lucro Bruto Ajustado / Receita Líquida
- **EBITDA**
- **Margem EBITDA** = EBITDA / Receita Líquida

para três recortes:

| Coluna do relatório | Filtro na DRE Banco |
|---|---|
| Consolidado | soma de todos os canais (Varejo, Atacado, Corporativo, CD-Logística, Produção) |
| Varejo | Canal = `VAREJO` |
| Atacado | Canal = `ATACADO` |

Cada métrica é comparada com o mesmo período do ano anterior (variação % e
R$, e variação em p.p. para as margens). Valores positivos aparecem em verde
e negativos em vermelho, igual ao modelo original.

> **Nota sobre o Atacado:** o One Page oficial usa como fonte do Atacado o
> **Lynkz Gerencial Consolidado**, que pode divergir da aba `ATACADO` da DRE
> Banco (contábil). O script usa a DRE Banco por ser a fonte disponível; se
> a base do Lynkz também estiver disponível em planilha, ela pode ser
> plugada no lugar do filtro `Canal = ATACADO` em `SEGMENTS` no script.

## Uso

1. No Google Sheets, exporte a planilha: `Arquivo > Fazer download >
   Microsoft Excel (.xlsx)` e salve em `data/DRE_Aberta_Mensal.xlsx`
   (a pasta `data/` não é versionada no git, pois contém dados financeiros
   por loja).

2. Rode o script informando o período atual e o período de comparação:

   ```bash
   python3 scripts/generate_dre_report.py \
     --file data/DRE_Aberta_Mensal.xlsx \
     --year 2026 --months 1-6 \
     --compare-year 2025 --compare-months 1-6 \
     --label "6 M/2026" \
     --slide-number 3 \
     --output reports/one-page-6M-2026.html
   ```

3. Abra o `.html` gerado no navegador para conferir, e commite o arquivo em
   `reports/` se quiser manter o histórico dos períodos.

### Gerando para um novo período (ex: 9M/2026)

Basta trocar `--months` e `--label`:

```bash
python3 scripts/generate_dre_report.py \
  --file data/DRE_Aberta_Mensal.xlsx \
  --year 2026 --months 1-9 \
  --compare-year 2025 --compare-months 1-9 \
  --label "9 M/2026" \
  --output reports/one-page-9M-2026.html
```

## Requisitos

```bash
pip install openpyxl
```

## Estrutura

```
scripts/generate_dre_report.py   # gerador do relatório
data/                             # coloque aqui o .xlsx exportado (não versionado)
reports/                          # relatórios .html gerados
```
