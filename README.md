# 📊 Análise de Termos em PDFs - IA vs Dados/BI

Script Python para varrer PDFs, contar termos relacionados a **IA/LLM** e **Dados/BI**, e gerar relatórios detalhados em Excel.

## 🎯 Objetivo

Analisar documentos PDF (relatórios financeiros, apresentações, etc.) para identificar e quantificar menções a:
- **IA/LLM**: Inteligência Artificial, Machine Learning, LLMs, IA Generativa, etc.
- **Dados/BI**: Business Intelligence, Analytics, Engenharia de Dados, Bancos de Dados, etc.

## ✨ Características

- ✅ **Contagem precisa** com regex e word boundaries
- ✅ **Tratamento inteligente de falsos positivos**:
  - Rejeita "IA" quando faz parte de "IAS" (International Accounting Standards)
  - Rejeita "BI" quando se refere a "Bilhões" (contexto numérico/monetário)
  - Aceita "Ia generativa" (primeira minúscula, resto maiúsculo)
- ✅ **Normalização de texto** (minúsculo, sem acentos, espaços normalizados)
- ✅ **Progresso visual** com tqdm
- ✅ **Tratamento robusto de erros** (continua mesmo se um PDF falhar)
- ✅ **Código em português**, fácil de entender e modificar

## 📦 Instalação

### Pré-requisitos

- Python 3.8 ou superior
- pip

### Passos

1. Clone o repositório:
```bash
git clone https://github.com/seu-usuario/iaindex.git
cd iaindex
```

2. Instale as dependências:
```bash
pip install -r requirements.txt
```

## ⚙️ Configuração

1. Abra o arquivo `analisar_pdfs.py`
2. Ajuste a variável `PASTA_RAIZ` para apontar para a pasta onde estão seus PDFs:
```python
PASTA_RAIZ = r"C:\caminho\para\seus\pdfs"
```

3. (Opcional) Configure o filtro de empresa:
```python
EMPRESA_FILTRO = "AMERICANAS"  # None = processa todas as empresas
```

4. (Opcional) Configure se deve incluir PDFs sem ocorrências:
```python
INCLUIR_PDFS_SEM_OCORRENCIAS = False  # True = inclui PDFs com zero ocorrências
```

### Estrutura de Pastas Esperada

```
PASTA_RAIZ/
  ├── Empresa1/
  │   ├── 2023/
  │   │   └── arquivo1.pdf
  │   ├── 2024/
  │   │   └── arquivo2.pdf
  │   └── 2025/
  │       └── arquivo3.pdf
  └── Empresa2/
      └── ...
```

## 🚀 Execução

```bash
python analisar_pdfs.py
```

O script irá:
1. Varrer recursivamente a pasta raiz
2. Processar todos os PDFs encontrados
3. Contar termos por grupo (IA_LLM e DADOS_BI)
4. Gerar arquivo Excel com análises detalhadas

## 📊 Saída

O script gera um arquivo Excel (`analise_termos.xlsx`) com as seguintes abas:

### Abas Analíticas por Ano
- **analitico_2023, analitico_2024, analitico_2025**: Dados detalhados por ano
  - Cada linha = 1 PDF + 1 grupo (mesmo PDF pode ter 2 linhas: IA_LLM e DADOS_BI)
  - Colunas: ano, empresa, pdf_nome, pdf_caminho, total_paginas, total_palavras_pdf, grupo, ocorrencias_total_grupo, termos_encontrados, ocorrencias_por_termo

### Aba Agregada
- **analitico_todos**: Todos os dados agregados (concatenação dos 3 anos)

### Aba de Resumo
- **resumo_empresas**: Resumo por empresa, ano e grupo
  - pdfs_com_ocorrencia (nunique)
  - ocorrencias_total (sum)

### Aba de Evolução
- **evolucao**: Evolução temporal por empresa + grupo
  - ocorr_2023, ocorr_2024, ocorr_2025
  - delta_24_23, delta_25_24
  - pct_24_23, pct_25_24

### Aba de Auditoria
- **parametros**: Lista completa de termos utilizados por grupo (rastreabilidade)

## 🔍 Grupos de Termos

### IA_LLM
- Inteligência Artificial, IA Generativa, Machine Learning, Deep Learning
- LLM, NLP, Transformers, RAG, Prompt Engineering
- GPT, ChatGPT, Gemini, Claude, etc.

### DADOS_BI
- Business Intelligence, Analytics, Data Science
- Engenharia de Dados, ETL, Data Warehouse, Data Lake
- SQL, Power BI, Tableau, Snowflake, etc.

Os termos podem ser facilmente editados nos dicionários no topo do arquivo `analisar_pdfs.py`.

## 🛠️ Scripts Auxiliares

- `listar_empresas.py`: Lista empresas disponíveis na pasta de PDFs

## 📝 Exemplo de Uso

```python
# Processar apenas uma empresa
EMPRESA_FILTRO = "AMERICANAS"
python analisar_pdfs.py

# Processar todas as empresas
EMPRESA_FILTRO = None
python analisar_pdfs.py
```

## 🐛 Tratamento de Falsos Positivos

O script implementa várias estratégias para evitar falsos positivos:

1. **Siglas curtas (IA, BI, LLM)**: 
   - Só conta quando isoladas com delimitadores (espaço, pontuação, etc.)
   - Verifica se está em maiúsculo no texto original

2. **Rejeição de padrões conhecidos**:
   - "IA" em "IAS" (International Accounting Standards) → ❌ Rejeita
   - "BI" em contexto numérico ("R$ 1,5 BI") → ❌ Rejeita (é Bilhões)
   - "BI" em contexto de tecnologia ("Power BI") → ✅ Aceita

3. **Word boundaries**: Usa regex com word boundaries para evitar capturar termos dentro de palavras maiores

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo LICENSE para mais detalhes.

## 🤝 Contribuindo

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests.

## 📧 Contato

Para dúvidas ou sugestões, abra uma issue no repositório.
