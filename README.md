# One Page

Gera o relatório "One Page" (Consolidado / Varejo / Atacado) a partir da
planilha **DRE Banco** (Google Sheets) e da **DRE do Atacado** (Lynkz
Gerencial Consolidado), no mesmo formato do modelo usado nas apresentações
mensais.

Planilha de origem (Varejo/Consolidado): [DRE Banco](https://docs.google.com/spreadsheets/d/1qYlMyLFWbtqkGw7ingnFFAilDDYxDwzun0TGs8lBMr0)

## Como funciona

O script lê a aba "DRE Aberta Mensal" da DRE Banco (uma aba por ano, ex:
`2025`, `2026`, ...) com as colunas
`Mês Ano | Loja | Canal | DRE | -<ano> | Orçado -<ano>` e calcula, para o
período informado:

- **Receita Bruta**
- **Lucro Bruto Ajustado**
- **Margem Bruta** = Lucro Bruto Ajustado / Receita Líquida
- **EBITDA**
- **Margem EBITDA** = EBITDA / Receita Líquida

para três recortes:

| Coluna do relatório | Fonte |
|---|---|
| Consolidado | soma de todos os canais da DRE Banco (Varejo, Atacado, Corporativo, CD-Logística, Produção) |
| Varejo | Canal = `VAREJO` na DRE Banco |
| Atacado | DRE do Atacado (Lynkz Gerencial Consolidado) — arquivo `--atacado-file` |

Cada métrica é comparada com o mesmo período do ano anterior (variação % e
R$, e variação em p.p. para as margens). Valores positivos aparecem em verde
e negativos em vermelho, igual ao modelo original.

> **Nota sobre o Atacado:** o One Page oficial usa como fonte do Atacado o
> **Lynkz Gerencial Consolidado**, que diverge da aba `ATACADO` da DRE Banco
> (contábil). Por isso o script lê o Atacado de um arquivo separado
> (`--atacado-file`), já totalizado no período (ex: "Real 6M/2026" vs "Real
> 6M/2025"). Sem esse arquivo, o script cai de volta para o Canal=`ATACADO`
> da DRE Banco, como aproximação.

## Uso

1. No Google Sheets, exporte as planilhas:
   `Arquivo > Fazer download > Microsoft Excel (.xlsx)`
   - DRE Banco → salve em `data/DRE_Aberta_Mensal.xlsx`
   - DRE do Atacado (Lynkz) → salve em `data/DRE_Atacado.xlsx`

   (a pasta `data/` não é versionada no git, pois contém dados financeiros
   por loja).

2. Rode o script informando o período atual e o período de comparação:

   ```bash
   python3 scripts/generate_dre_report.py \
     --file data/DRE_Aberta_Mensal.xlsx \
     --atacado-file data/DRE_Atacado.xlsx \
     --year 2026 --months 1-6 \
     --compare-year 2025 --compare-months 1-6 \
     --label "6 M/2026" \
     --slide-number 3 \
     --output reports/one-page-6M-2026.html
   ```

3. Abra o `.html` gerado no navegador para conferir, e commite o arquivo em
   `reports/` se quiser manter o histórico dos períodos.

### Gerando para um novo período (ex: 9M/2026)

Troque `--months` e `--label`, e **reexporte a DRE do Atacado do Lynkz para
o novo período** (o arquivo do Atacado já vem totalizado, então não dá para
recalcular outro intervalo de meses a partir do mesmo arquivo — diferente
da DRE Banco, que é mensal e permite qualquer recorte):

```bash
python3 scripts/generate_dre_report.py \
  --file data/DRE_Aberta_Mensal.xlsx \
  --atacado-file data/DRE_Atacado_9M.xlsx \
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
data/                             # coloque aqui os .xlsx exportados (não versionado)
reports/                          # relatórios .html gerados
```
