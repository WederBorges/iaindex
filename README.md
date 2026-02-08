# TCC – Índice de Ênfase em Inteligência Artificial

Este repositório contém o código utilizado na análise quantitativa do TCC
sobre a presença e evolução de termos relacionados à Inteligência Artificial
em relatórios corporativos (2023–2025).

## Estrutura
- `data/` – dados tratados (Excel de saída da análise)
- `src/` – scripts de processamento e análise
  - `analisar_pdfs.py` – análise de termos em PDFs (IA vs Dados/BI)
  - `listar_empresas.py` – lista empresas na pasta de PDFs
- `notebooks/` – análises exploratórias
- `requirements.txt` – dependências

## Como executar
Na raiz do projeto:
```bash
python src/analisar_pdfs.py
python src/listar_empresas.py
```
O Excel gerado é salvo em `data/analise_termos3.xlsx`. Ajuste `PASTA_RAIZ` em `src/analisar_pdfs.py` para a pasta onde estão os PDFs.

## Metodologia
- Contagem de frequência de termos
- Agregação anual
- Cálculo de variação (Δ)
- Construção de índice de ênfase em IA

## Autor
Weder
