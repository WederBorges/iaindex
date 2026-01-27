"""
Script para varrer PDFs, contar termos por grupos (IA vs Dados/BI) e gerar Excel com análises.
"""

import os
import re
import json
from pathlib import Path
from typing import Dict, List, Tuple, Optional

import pdfplumber
import pandas as pd
from tqdm import tqdm
import unicodedata

# ============================================================================
# CONFIGURAÇÕES
# ============================================================================

PASTA_RAIZ = r"C:\Users\weder\Downloads\01 - AMOSTRA VALIDADA (APÓS DATA REDUCTION)-20260126T164115Z-3-001\01 - AMOSTRA VALIDADA (APÓS DATA REDUCTION)"  # Ajuste conforme necessário
ARQUIVO_EXCEL_SAIDA = "analise_termos.xlsx"
INCLUIR_PDFS_SEM_OCORRENCIAS = False  # Se True, inclui PDFs com zero ocorrências
EMPRESA_FILTRO = "AMERICANAS"  # None = processa todas as empresas, ou nome da empresa (ex: "AMERICANAS")

# ============================================================================
# DICIONÁRIOS DE TERMOS
# ============================================================================

TERMOS_IA_LLM = {
    "IA_LLM": [
        # Guarda-chuva
        "inteligencia artificial", "artificial intelligence",
        "ia generativa", "generative ai", "genai",
        "modelo generativo", "modelos generativos",
        
        # Machine Learning / Deep Learning
        "machine learning", "aprendizado de maquina", "aprendizagem de maquina",
        "deep learning", "aprendizado profundo",
        "rede neural", "redes neurais", "neural network", "neural networks",
        
        # NLP / Linguagem
        "processamento de linguagem natural", "nlp", "natural language processing",
        "modelo de linguagem", "modelos de linguagem",
        "modelo de linguagem grande", "modelos de linguagem grandes",
        "large language model", "large language models",
        "llm", "llms",
        
        # Transformers e técnicas modernas
        "transformer", "transformers", "attention mechanism", "self attention",
        "embeddings", "vetor de embeddings", "vector embedding",
        "fine tuning", "finetuning", "ajuste fino",
        "instruction tuning", "rlhf", "reinforcement learning from human feedback",
        "prompt engineering", "engenharia de prompt", "prompting",
        "rag", "retrieval augmented generation", "retrieval-augmented generation",
        "vector database", "banco de vetores", "base vetorial",
        
        # Agentes / Copilotos / Chatbots
        "agentes de ia", "ai agents", "agentes autonomos", "autonomous agents",
        "chatbot", "chatbots", "assistente virtual", "assistentes virtuais",
        "copilot", "copiloto",
        
        # Visão / fala
        "computer vision", "visao computacional", "visão computacional",
        "reconhecimento de fala", "speech recognition",
        "reconhecimento de imagem", "image recognition",
        
        # Modelos/Plataformas
        "gpt", "chatgpt", "openai",
        "gemini", "claude", "llama", "mistral"
    ],
    "SIGLAS_SENSIVEIS": ["IA", "AI", "LLM"]
}

TERMOS_DADOS_BI = {
    "DADOS_BI": [
        # Conceitos gerais de dados
        "dados", "data", "data driven", "data-driven", "orientado a dados",
        "analise de dados", "análise de dados", "data analytics", "analytics",
        "cientista de dados", "data scientist", "data science", "ciencia de dados", "ciência de dados",
        
        # BI e visualização
        # "bi" NÃO está aqui: contamos só via SIGLAS_SENSIVEIS + verificar_bi_bilhoes
        # (evita contar "R$ 22,8 bi" = bilhões como Business Intelligence)
        "business intelligence", "inteligencia de negocios", "inteligência de negócios",
        "dashboard", "dashboards", "painel", "kpi", "kpis", "indicador", "indicadores",
        "visualizacao de dados", "visualização de dados", "data visualization",
        
        # Engenharia de dados / pipelines
        "engenharia de dados", "data engineering",
        "etl", "elt", "pipeline de dados", "data pipeline", "integracao de dados", "integração de dados",
        "orquestracao de dados", "orquestração de dados", "airflow",
        
        # Armazenamento / arquiteturas
        "data warehouse", "dw", "warehouse",
        "data lake", "datalake", "lakehouse", "datamart",
        "big data", "hadoop", "spark",
        
        # Banco de dados / linguagens
        "banco de dados", "database", "db", "dbms", "sgbd",
        "sql", "nosql", "query", "consultas sql",
        "modelagem de dados", "data modeling",
        
        # Plataformas e ferramentas
        "power bi", "tableau", "qlik", "looker",
        "snowflake", "bigquery", "redshift", "databricks",
        "postgresql", "mysql", "oracle", "sql server", "mongodb",
        
        # Governança / qualidade / segurança
        "governanca de dados", "governança de dados",
        "qualidade de dados", "data quality",
        "catalogo de dados", "catálogo de dados", "data catalog",
        "linhagem de dados", "data lineage",
        "metadados", "metadata",
        "lgpd", "privacidade de dados", "data privacy"
    ],
    "SIGLAS_SENSIVEIS": ["BI", "DW", "ETL", "ELT", "SQL", "KPI"]
}

# ============================================================================
# FUNÇÕES AUXILIARES
# ============================================================================

def remover_acentos(texto: str) -> str:
    """Remove acentos de uma string."""
    nfkd = unicodedata.normalize('NFKD', texto)
    return ''.join([c for c in nfkd if not unicodedata.combining(c)])

def normalizar_texto(texto: str) -> str:
    """
    Normaliza texto: minúsculo, sem acento, espaços normalizados.
    """
    texto = texto.lower()
    texto = remover_acentos(texto)
    texto = re.sub(r'\s+', ' ', texto)  # Normaliza espaços
    return texto.strip()

def contar_palavras_aproximado(texto: str) -> int:
    """Conta palavras aproximadas usando regex."""
    palavras = re.findall(r'\b\w+\b', texto)
    return len(palavras)

def criar_regex_termo(termo: str, usar_word_boundary: bool = True) -> re.Pattern:
    """
    Cria regex para buscar um termo com word boundaries.
    Aceita variações com espaço/hífen.
    Para termos curtos (possíveis siglas), usa padrão mais rigoroso.
    """
    # Escapa caracteres especiais do regex
    termo_escaped = re.escape(termo)
    
    # Permite espaço ou hífen entre palavras do termo composto
    termo_escaped = termo_escaped.replace(r'\ ', r'[\s\-]+')
    
    # Termos muito curtos (2-3 letras, apenas letras) podem ser siglas
    # Usar padrão mais rigoroso para evitar falsos positivos
    termo_limpo = termo.replace(' ', '').replace('-', '')
    is_sigla_curta = len(termo_limpo) <= 3 and termo_limpo.isalpha()
    
    if usar_word_boundary:
        if is_sigla_curta:
            # Para siglas curtas, usar delimitadores mais rigorosos
            # Aceita: espaço, pontuação, início/fim de linha
            delimitador_antes = r'(?:^|[\s([{.,;:!?\-])'
            delimitador_depois = r'(?=[\s)\].,;:!?\-]|$)'
            pattern = delimitador_antes + termo_escaped + delimitador_depois
        else:
            # Para termos longos, word boundary padrão é suficiente
            pattern = r'\b' + termo_escaped + r'\b'
    else:
        pattern = termo_escaped
    
    return re.compile(pattern, re.IGNORECASE)

def verificar_bi_bilhoes(pos_inicio_match: int, texto: str) -> bool:
    """
    Verifica se "BI" está em contexto de "Bilhões" (numérico/monetário).
    Retorna True se deve rejeitar (é bilhões), False se deve aceitar (é Business Intelligence).
    """
    # Verificar contexto antes de "BI" (últimos 20 caracteres)
    contexto_antes = texto[max(0, pos_inicio_match - 20):pos_inicio_match]
    
    # Padrões que indicam "Bilhões":
    # - Números antes: "1,5 BI", "R$ 2 BI", "2.5 BI"
    # - Símbolos monetários: "R$", "$", "€"
    # - Palavras relacionadas: "milhões", "mil", "reais", "dólares"
    
    # Verificar se há números próximos antes (padrões: "1,5", "2.3", "R$ 1,5", "2.500")
    # Buscar números com vírgula ou ponto, possivelmente precedidos por símbolo monetário
    tem_numero_antes = bool(re.search(r'[\d]{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?\s*$', contexto_antes))
    
    # Verificar se há símbolos monetários seguidos de números
    tem_simbolo_monetario = bool(re.search(r'[R$€£]\s*[\d]{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?\s*$', contexto_antes, re.IGNORECASE))
    
    # Verificar se há palavras relacionadas a valores
    palavras_valores = ['milhões', 'mil', 'reais', 'dólares', 'euros', 'valor', 'total', 'receita', 
                      'vendas', 'lucro', 'prejuízo', 'patrimônio', 'ativo', 'passivo']
    tem_palavra_valor = any(palavra in contexto_antes.lower() for palavra in palavras_valores)
    
    # Verificar contexto depois (próximos 10 caracteres)
    pos_fim_match = pos_inicio_match + 2  # "BI" tem 2 caracteres
    contexto_depois = texto[pos_fim_match:min(len(texto), pos_fim_match + 10)]
    
    # Se há número antes OU símbolo monetário OU palavra de valor, provavelmente é "Bilhões"
    if tem_numero_antes or tem_simbolo_monetario or tem_palavra_valor:
        return True  # Rejeitar: é "Bilhões"
    
    # Verificar se está seguido de palavras que indicam valor
    palavras_depois_valor = ['reais', 'dólares', 'euros', 'em', 'de', 'no', 'na']
    if any(contexto_depois.lower().startswith(palavra) for palavra in palavras_depois_valor):
        return True  # Rejeitar: provavelmente é "Bilhões"
    
    return False  # Aceitar: provavelmente é "Business Intelligence"

def verificar_data_eh_data(pos_inicio: int, pos_fim: int, texto_norm: str) -> bool:
    """
    Verifica se "data" está em contexto de DATA (date) em relatórios.
    Retorna True se deve REJEITAR (é date: "data do balanço", "data de divulgação").
    Assim evitamos contar "data" = data analytics vs "data" = data.
    """
    tam = len(texto_norm)
    ctx_antes = texto_norm[max(0, pos_inicio - 25):pos_inicio]
    ctx_depois = texto_norm[pos_fim:min(tam, pos_fim + 25)]
    # Rejeitar se padrões típicos de "data" = date
    antes_date = [
        "em data", "a data", "ate data", "à data",
        "dia data", "na data", "pela data", "por data", "ate a data",
        "da data", "das data",
    ]
    depois_date = [
        " do ", " da ", " de ", " do balanco", " base", " de divulgacao",
        " de publicacao", " de referencia", " de corte", " de fechamento",
        " limite", " valor", " vencimento",
    ]
    if any(p in ctx_antes for p in antes_date):
        return True
    if any(ctx_depois.startswith(p) for p in depois_date):
        return True
    # "data" seguido de " do/da/de balanço|divulgação|referência|..."
    if re.match(r'\s+(do|da|de)\s+(balanco|divulgacao|referencia|corte|fechamento|valor|vencimento)\b', ctx_depois):
        return True
    return False

def verificar_indicador_financeiro(pos_inicio: int, pos_fim: int, texto_norm: str) -> bool:
    """
    Verifica se "indicador" / "indicadores" está em contexto puramente financeiro.
    Retorna True se deve REJEITAR (ex.: "indicadores financeiros", "indicadores econômicos").
    """
    tam = len(texto_norm)
    ctx_depois = texto_norm[pos_fim:min(tam, pos_fim + 30)]
    # Rejeitar quando seguido de "financeiro(s)", "econômico(s)", "contábil(eis)"
    if re.match(r'\s*(financeiro|financeiros|economico|economicos|contabil|contabeis)\b', ctx_depois):
        return True
    return False

# Termos que passam por verificação de contexto (evitar falsos positivos em relatórios)
# Função recebe (pos_inicio, pos_fim, texto_normalizado) e retorna True para REJEITAR o match.
# Ex.: "data" = date (data do balanço) vs data analytics; "indicador" = financeiro vs BI.
VERIFICACOES_CONTEXTO = {
    "data": verificar_data_eh_data,
    "indicador": verificar_indicador_financeiro,
    "indicadores": verificar_indicador_financeiro,
}

def buscar_sigla_no_texto_original(texto_original: str, sigla: str) -> Tuple[int, List[str]]:
    """
    Busca sigla curta no texto original com padrões rigorosos.
    Aceita:
    - Siglas totalmente maiúsculas: "IA", "LLM", "BI"
    - Siglas com primeira minúscula (início de frase): "Ia generativa", "Bi é importante"
    
    Rejeita:
    - Siglas minúsculas dentro de palavras: "eu ia", "via"
    - Siglas entre hífens dentro de palavras: "DIA-IA-DIA", "DIA-IA", "IA-DIA"
    - "IA" quando faz parte de "IAS" (International Accounting Standards)
    
    Exemplos:
    - " IA " -> conta (espaço antes e depois, maiúsculo)
    - "Ia generativa" -> conta (primeira minúscula, resto maiúsculo, seguido de espaço)
    - "eu ia" -> NÃO conta (minúsculo)
    - "via" -> NÃO conta (dentro de palavra)
    - ".IA " -> conta (pontuação antes, espaço depois)
    - "(IA)" -> conta (parênteses)
    - "IAS 8" -> NÃO conta (IA faz parte de IAS - contabilidade)
    - "DIA-IA-DIA" -> NÃO conta (entre hífens dentro de palavra)
    - "-IA-" -> NÃO conta (entre hífens, pode ser parte de palavra composta)
    """
    sigla_escaped = re.escape(sigla)
    
    # Lista de padrões a rejeitar (siglas que contêm a sigla procurada)
    # Exemplo: "IAS" contém "IA", então rejeitar "IA" quando está em "IAS"
    padroes_rejeitar = {
        "IA": ["IAS"],  # Rejeitar IA quando faz parte de IAS (International Accounting Standards)
        "AI": ["AIS", "AID", "AIM"],  # Possíveis falsos positivos
    }
    
    # "IA" como sufixo: tesourar-ia, econom-ia, etc. (quebra de linha pode gerar "tesourar IA ")
    # Rejeitar quando a palavra antes é "tesourar" ou outros radicais que + "ia" formam palavra.
    sufixos_ia_rejeitar = ("tesourar", "econom", "burgues", "demonstr", "secretar")
    
    # Padrão unificado: aceita sigla totalmente maiúscula OU primeira minúscula + resto maiúsculo
    # Delimitadores antes: espaço, início de linha, pontuação, parênteses
    # Delimitadores depois: espaço, fim de linha, pontuação, parênteses
    # NOTA: hífen será tratado separadamente para evitar falsos positivos
    delimitador_antes = r'(?:^|[\s([{.,;:!?])'
    delimitador_depois = r'(?=[\s)\].,;:!?]|$)'
    
    # Padrão 1: Sigla totalmente maiúscula
    pattern1 = delimitador_antes + sigla_escaped + delimitador_depois
    
    # Padrão 2: Sigla com primeira minúscula (ex: "Ia", "Bi", "Llm")
    if len(sigla) > 1:
        sigla_primeira_minuscula = sigla[0].lower() + sigla[1:].upper()
        sigla_primeira_minuscula_escaped = re.escape(sigla_primeira_minuscula)
        pattern2 = delimitador_antes + sigla_primeira_minuscula_escaped + delimitador_depois
    else:
        pattern2 = None
    
    # Padrão 3: Sigla entre hífens (precisa verificação especial)
    # Aceita apenas se NÃO estiver dentro de palavra composta
    pattern3_hifen = r'[A-Za-z]*-' + sigla_escaped + r'-[A-Za-z]*'  # Entre hífens com letras antes e depois
    
    # Buscar todos os padrões
    count = 0
    exemplos = []  # Lista de até 3 exemplos de contexto
    
    # Buscar padrão 1 (totalmente maiúscula, sem hífen problemático)
    for match in re.finditer(pattern1, texto_original, re.MULTILINE):
        match_text = match.group(0)
        sigla_match = re.search(sigla_escaped, match_text)
        if sigla_match and sigla_match.group(0).isupper():
            pos_inicio_match = match.start()
            pos_fim_match = match.end()
            
            # Verificação especial para "BI" (Business Intelligence vs Bilhões)
            if sigla == "BI":
                if verificar_bi_bilhoes(pos_inicio_match, texto_original):
                    continue  # Rejeitar: é "Bilhões", não "Business Intelligence"
            
            # "IA" como sufixo (ex.: "tesouraria" → "tesourar IA " por quebra de linha)
            if sigla == "IA":
                # Calcular posição exata da sigla no texto original
                # match.start() é início do match completo (inclui delimitadores)
                # sigla_match.start() é offset da sigla dentro do match_text
                if sigla_match:
                    pos_sigla_inicio = match.start() + sigla_match.start()
                    ctx_antes = texto_original[max(0, pos_sigla_inicio - 15):pos_sigla_inicio]
                    # Verificar se termina com radical que forma palavra com "ia"
                    ctx_limpo = ctx_antes.rstrip().lower()
                    if any(ctx_limpo.endswith(r) for r in sufixos_ia_rejeitar):
                        continue  # Rejeitar: é sufixo ("tesourar**ia**", "econom**ia**", etc.)
            
            # Verificar se não faz parte de um padrão a rejeitar (ex: "IA" em "IAS")
            if sigla in padroes_rejeitar:
                deve_rejeitar = False
                for padrao_rejeitar in padroes_rejeitar[sigla]:
                    # Usar posição exata do fim da sigla (não do match completo com delimitadores)
                    if sigla_match:
                        pos_sigla_fim = match.start() + sigla_match.end()
                    else:
                        pos_sigla_fim = pos_fim_match
                    
                    # Verificar se após a sigla vem o restante do padrão (ex: "S" após "IA" = "IAS")
                    if pos_sigla_fim < len(texto_original):
                        resto_padrao = padrao_rejeitar[len(sigla):]  # Ex: "S" para "IAS"
                        if len(resto_padrao) > 0:
                            # Verificar se o texto após a sigla forma o padrão completo
                            texto_apos = texto_original[pos_sigla_fim:pos_sigla_fim + len(resto_padrao)]
                            if texto_apos.upper() == resto_padrao.upper():
                                # Verificar se o padrão completo está isolado (não é parte de palavra maior)
                                pos_apos_padrao = pos_sigla_fim + len(resto_padrao)
                                if pos_apos_padrao < len(texto_original):
                                    char_apos_padrao = texto_original[pos_apos_padrao]
                                    # Se não é letra, o padrão está isolado - REJEITAR
                                    if not char_apos_padrao.isalpha():
                                        deve_rejeitar = True
                                        break
                
                if deve_rejeitar:
                    continue  # Rejeitar: faz parte de padrão maior (ex: "IAS")
            
            count += 1
            # Capturar exemplo de contexto (até 3)
            if len(exemplos) < 3:
                if sigla_match:
                    pos_sigla_inicio = match.start() + sigla_match.start()
                    pos_sigla_fim = match.start() + sigla_match.end()
                else:
                    pos_sigla_inicio = pos_inicio_match
                    pos_sigla_fim = pos_fim_match
                ctx_antes = texto_original[max(0, pos_sigla_inicio - 30):pos_sigla_inicio]
                ctx_depois = texto_original[pos_sigla_fim:min(len(texto_original), pos_sigla_fim + 30)]
                sigla_encontrada = texto_original[pos_sigla_inicio:pos_sigla_fim]
                exemplo = f"...{ctx_antes}**{sigla_encontrada}**{ctx_depois}..."
                exemplos.append(exemplo)
    
    # Buscar padrão 2 (primeira minúscula)
    if pattern2:
        for match in re.finditer(pattern2, texto_original, re.MULTILINE):
            match_text = match.group(0)
            sigla_match = re.search(sigla_primeira_minuscula_escaped, match_text)
            if sigla_match:
                sigla_encontrada = sigla_match.group(0)
                if len(sigla_encontrada) > 1 and sigla_encontrada[0].islower() and sigla_encontrada[1:].isupper():
                    pos_inicio_match = match.start()
                    pos_fim_match = match.end()
                    
                    # Verificação especial para "BI" (Business Intelligence vs Bilhões)
                    if sigla == "BI":
                        if verificar_bi_bilhoes(pos_inicio_match, texto_original):
                            continue  # Rejeitar: é "Bilhões", não "Business Intelligence"
                    
                    # "IA" como sufixo (ex.: "tesouraria" → "tesourar Ia " por quebra de linha)
                    if sigla == "IA":
                        if sigla_match:
                            pos_sigla_inicio = match.start() + sigla_match.start()
                            ctx_antes = texto_original[max(0, pos_sigla_inicio - 15):pos_sigla_inicio]
                            ctx_limpo = ctx_antes.rstrip().lower()
                            if any(ctx_limpo.endswith(r) for r in sufixos_ia_rejeitar):
                                continue  # Rejeitar: é sufixo
                    
                    # Verificar se não faz parte de um padrão a rejeitar (similar ao padrão 1)
                    if sigla in padroes_rejeitar:
                        deve_rejeitar = False
                        for padrao_rejeitar in padroes_rejeitar[sigla]:
                            # Usar posição exata do fim da sigla
                            if sigla_match:
                                pos_sigla_fim = match.start() + sigla_match.end()
                            else:
                                pos_sigla_fim = pos_fim_match
                            
                            resto_padrao = padrao_rejeitar[len(sigla):]
                            if len(resto_padrao) > 0 and pos_sigla_fim < len(texto_original):
                                texto_apos = texto_original[pos_sigla_fim:pos_sigla_fim + len(resto_padrao)]
                                if texto_apos.upper() == resto_padrao.upper():
                                    pos_apos_padrao = pos_sigla_fim + len(resto_padrao)
                                    if pos_apos_padrao < len(texto_original):
                                        char_apos_padrao = texto_original[pos_apos_padrao]
                                        if not char_apos_padrao.isalpha():
                                            deve_rejeitar = True
                                            break
                    
                    if deve_rejeitar:
                        continue  # Rejeitar: faz parte de padrão maior
                    
                    count += 1
                    # Capturar exemplo de contexto (até 3)
                    if len(exemplos) < 3:
                        if sigla_match:
                            pos_sigla_inicio = match.start() + sigla_match.start()
                            pos_sigla_fim = match.start() + sigla_match.end()
                        else:
                            pos_sigla_inicio = pos_inicio_match
                            pos_sigla_fim = pos_fim_match
                        ctx_antes = texto_original[max(0, pos_sigla_inicio - 30):pos_sigla_inicio]
                        ctx_depois = texto_original[pos_sigla_fim:min(len(texto_original), pos_sigla_fim + 30)]
                        sigla_encontrada = texto_original[pos_sigla_inicio:pos_sigla_fim]
                        exemplo = f"...{ctx_antes}**{sigla_encontrada}**{ctx_depois}..."
                        exemplos.append(exemplo)
    
    # Buscar padrão 3 (entre hífens) - mas rejeitar se está dentro de palavra composta
    # Exemplo: "DIA-IA-DIA" -> NÃO conta (IA está dentro de palavra composta)
    # Exemplo: "-IA-" no início/fim -> pode contar se não estiver em palavra composta
    matches_hifen = list(re.finditer(r'-' + sigla_escaped + r'-', texto_original, re.IGNORECASE))
    for match_hifen in matches_hifen:
        pos_inicio = match_hifen.start()  # Posição do primeiro hífen
        pos_fim = match_hifen.end()  # Posição após o segundo hífen
        
        # Verificar contexto: há letras antes do primeiro hífen E depois do segundo hífen?
        char_antes_hifen1 = texto_original[pos_inicio - 1] if pos_inicio > 0 else ''
        char_depois_hifen2 = texto_original[pos_fim] if pos_fim < len(texto_original) else ''
        
        # Se há letras antes E depois, é palavra composta - REJEITAR
        # Exemplo: "DIA-IA-DIA" -> char_antes_hifen1='A', char_depois_hifen2='D'
        if char_antes_hifen1.isalpha() and char_depois_hifen2.isalpha():
            continue  # Rejeitar: está dentro de palavra composta como "DIA-IA-DIA"
        
        # Se não há letras antes OU depois, pode ser sigla isolada
        # Verificar se a sigla está em maiúsculo (ou primeira minúscula + resto maiúsculo)
        sigla_no_match = texto_original[pos_inicio + 1:pos_fim - 1]  # Extrair sigla entre hífens
        
        # Verificação especial para "BI" (Business Intelligence vs Bilhões)
        if sigla == "BI":
            if verificar_bi_bilhoes(pos_inicio + 1, texto_original):  # +1 para posição da sigla (após o hífen)
                continue  # Rejeitar: é "Bilhões", não "Business Intelligence"
        
        # Aceitar se totalmente maiúscula OU primeira minúscula + resto maiúsculo
        if sigla_no_match.isupper():
            count += 1
            # Capturar exemplo (até 3)
            if len(exemplos) < 3:
                ctx_antes = texto_original[max(0, pos_inicio - 20):pos_inicio + 1]
                ctx_depois = texto_original[pos_fim - 1:min(len(texto_original), pos_fim + 20)]
                exemplo = f"...{ctx_antes}**{sigla_no_match}**{ctx_depois}..."
                exemplos.append(exemplo)
        elif len(sigla_no_match) > 1 and sigla_no_match[0].islower() and sigla_no_match[1:].isupper():
            count += 1
            # Capturar exemplo (até 3)
            if len(exemplos) < 3:
                ctx_antes = texto_original[max(0, pos_inicio - 20):pos_inicio + 1]
                ctx_depois = texto_original[pos_fim - 1:min(len(texto_original), pos_fim + 20)]
                exemplo = f"...{ctx_antes}**{sigla_no_match}**{ctx_depois}..."
                exemplos.append(exemplo)
    
    return count, exemplos

def contar_termos_no_texto(
    texto_original: str,
    texto_normalizado: str,
    termos: List[str],
    siglas_sensiveis: List[str]
) -> Tuple[Dict[str, int], List[str], Dict[str, List[str]]]:
    """
    Conta ocorrências de termos no texto e captura exemplos de contexto.
    Retorna: (dicionário termo -> contagem, lista de termos encontrados, exemplos_contexto)
    Termos em VERIFICACOES_CONTEXTO usam finditer e checagem de contexto.
    """
    ocorrencias = {}
    termos_encontrados = []
    exemplos_contexto = {}  # termo -> lista de até 3 exemplos
    
    # Conta termos normais (no texto normalizado)
    for termo in termos:
        regex = criar_regex_termo(termo, usar_word_boundary=True)
        verificar = VERIFICACOES_CONTEXTO.get(termo)
        exemplos = []
        
        if verificar is None:
            matches = regex.findall(texto_normalizado)
            count = len(matches)
            # Capturar exemplos do texto original (para mostrar contexto real)
            # Buscar diretamente no texto original com regex case-insensitive
            regex_original = criar_regex_termo(termo, usar_word_boundary=True)
            for i, m_orig in enumerate(regex_original.finditer(texto_original)):
                if i >= 3:
                    break
                ctx_antes = texto_original[max(0, m_orig.start() - 30):m_orig.start()]
                ctx_depois = texto_original[m_orig.end():min(len(texto_original), m_orig.end() + 30)]
                termo_real = m_orig.group()
                exemplo = f"...{ctx_antes}**{termo_real}**{ctx_depois}..."
                exemplos.append(exemplo)
        else:
            count = 0
            for m in regex.finditer(texto_normalizado):
                if not verificar(m.start(), m.end(), texto_normalizado):
                    count += 1
                    # Capturar exemplos do texto original (primeiros 3)
                    if len(exemplos) < 3:
                        # Buscar no texto original usando regex
                        regex_original = criar_regex_termo(termo, usar_word_boundary=True)
                        encontrados_orig = list(regex_original.finditer(texto_original))
                        if len(encontrados_orig) > len(exemplos):
                            m_orig = encontrados_orig[len(exemplos)]
                            ctx_antes = texto_original[max(0, m_orig.start() - 30):m_orig.start()]
                            ctx_depois = texto_original[m_orig.end():min(len(texto_original), m_orig.end() + 30)]
                            termo_real = m_orig.group()
                            exemplo = f"...{ctx_antes}**{termo_real}**{ctx_depois}..."
                            exemplos.append(exemplo)
        
        if count > 0:
            ocorrencias[termo] = count
            termos_encontrados.append(termo)
            exemplos_contexto[termo] = exemplos
    
    # Conta siglas sensíveis (no texto original, apenas maiúsculas)
    for sigla in siglas_sensiveis:
        count, exemplos_sigla = buscar_sigla_no_texto_original(texto_original, sigla)
        if count > 0:
            ocorrencias[sigla] = count
            termos_encontrados.append(sigla)
            exemplos_contexto[sigla] = exemplos_sigla
    
    return ocorrencias, termos_encontrados, exemplos_contexto

# ============================================================================
# FUNÇÕES DE PROCESSAMENTO
# ============================================================================

def extrair_texto_pdf(caminho_pdf: str) -> Tuple[str, int]:
    """
    Extrai texto de um PDF usando pdfplumber.
    Retorna: (texto_completo, total_paginas)
    """
    texto_completo = []
    total_paginas = 0
    
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            total_paginas = len(pdf.pages)
            for pagina in pdf.pages:
                texto_pagina = pagina.extract_text()
                if texto_pagina:
                    texto_completo.append(texto_pagina)
    except Exception as e:
        raise Exception(f"Erro ao extrair texto do PDF: {e}")
    
    texto_final = "\n".join(texto_completo)
    return texto_final, total_paginas

def processar_pdf(
    caminho_pdf: str,
    empresa: str,
    ano: str
) -> Optional[List[Dict]]:
    """
    Processa um único PDF e retorna lista de dicionários com resultados.
    Cada dicionário representa um grupo (IA_LLM ou DADOS_BI).
    Retorna None se houver erro.
    """
    try:
        # Extrair texto
        texto_original, total_paginas = extrair_texto_pdf(caminho_pdf)
        texto_normalizado = normalizar_texto(texto_original)
        total_palavras = contar_palavras_aproximado(texto_normalizado)
        
        pdf_nome = os.path.basename(caminho_pdf)
        
        resultados = []
        
        # Processar grupo IA_LLM
        ocorrencias_ia, termos_ia, exemplos_ia = contar_termos_no_texto(
            texto_original,
            texto_normalizado,
            TERMOS_IA_LLM["IA_LLM"],
            TERMOS_IA_LLM["SIGLAS_SENSIVEIS"]
        )
        total_ia = sum(ocorrencias_ia.values())
        
        # Criar string com exemplos de contexto (até 3 por termo)
        exemplos_texto_ia = []
        for termo in termos_ia:
            if termo in exemplos_ia and exemplos_ia[termo]:
                exemplos_termo = " | ".join(exemplos_ia[termo])
                exemplos_texto_ia.append(f"{termo}: {exemplos_termo}")
        exemplos_str_ia = " || ".join(exemplos_texto_ia) if exemplos_texto_ia else ""
        
        if total_ia > 0 or INCLUIR_PDFS_SEM_OCORRENCIAS:
            resultados.append({
                "ano": ano,
                "empresa": empresa,
                "pdf_nome": pdf_nome,
                "pdf_caminho": caminho_pdf,
                "total_paginas": total_paginas,
                "total_palavras_pdf": total_palavras,
                "grupo": "IA_LLM",
                "ocorrencias_total_grupo": total_ia,
                "termos_encontrados": ", ".join(termos_ia) if termos_ia else "",
                "ocorrencias_por_termo": json.dumps(ocorrencias_ia, ensure_ascii=False),
                "exemplos_contexto": exemplos_str_ia
            })
        
        # Processar grupo DADOS_BI
        ocorrencias_bi, termos_bi, exemplos_bi = contar_termos_no_texto(
            texto_original,
            texto_normalizado,
            TERMOS_DADOS_BI["DADOS_BI"],
            TERMOS_DADOS_BI["SIGLAS_SENSIVEIS"]
        )
        total_bi = sum(ocorrencias_bi.values())
        
        # Criar string com exemplos de contexto (até 3 por termo)
        exemplos_texto_bi = []
        for termo in termos_bi:
            if termo in exemplos_bi and exemplos_bi[termo]:
                exemplos_termo = " | ".join(exemplos_bi[termo])
                exemplos_texto_bi.append(f"{termo}: {exemplos_termo}")
        exemplos_str_bi = " || ".join(exemplos_texto_bi) if exemplos_texto_bi else ""
        
        if total_bi > 0 or INCLUIR_PDFS_SEM_OCORRENCIAS:
            resultados.append({
                "ano": ano,
                "empresa": empresa,
                "pdf_nome": pdf_nome,
                "pdf_caminho": caminho_pdf,
                "total_paginas": total_paginas,
                "total_palavras_pdf": total_palavras,
                "grupo": "DADOS_BI",
                "ocorrencias_total_grupo": total_bi,
                "termos_encontrados": ", ".join(termos_bi) if termos_bi else "",
                "ocorrencias_por_termo": json.dumps(ocorrencias_bi, ensure_ascii=False),
                "exemplos_contexto": exemplos_str_bi
            })
        
        return resultados
        
    except Exception as e:
        print(f"\nERRO ao processar {caminho_pdf}: {e}")
        return None

def varrer_pastas() -> List[Dict]:
    """
    Varre recursivamente a pasta raiz e processa todos os PDFs.
    Retorna lista de dicionários com resultados.
    """
    pasta_raiz = Path(PASTA_RAIZ)
    
    if not pasta_raiz.exists():
        raise FileNotFoundError(f"Pasta raiz não encontrada: {PASTA_RAIZ}")
    
    todos_resultados = []
    erros = []
    
    # Encontrar todos os PDFs
    pdfs = list(pasta_raiz.rglob("*.pdf"))
    
    if not pdfs:
        print(f"Nenhum PDF encontrado em {PASTA_RAIZ}")
        return []
    
    print(f"Encontrados {len(pdfs)} PDFs para processar.\n")
    
    # Processar cada PDF
    for caminho_pdf in tqdm(pdfs, desc="Processando PDFs"):
        # Identificar empresa (pasta imediatamente abaixo da raiz)
        partes = caminho_pdf.relative_to(pasta_raiz).parts
        if len(partes) < 2:
            erros.append(f"PDF fora da estrutura esperada: {caminho_pdf}")
            continue
        
        empresa = partes[0]
        
        # Filtrar por empresa se especificado
        if EMPRESA_FILTRO is not None and empresa != EMPRESA_FILTRO:
            continue  # Pular esta empresa
        
        # Identificar ano (pasta 2023/2024/2025)
        ano = None
        for parte in partes[1:]:
            if parte in ["2023", "2024", "2025"]:
                ano = parte
                break
        
        # Se não encontrou ano na estrutura de pastas, tentar extrair do nome do arquivo
        if ano is None:
            nome_arquivo = partes[-1]  # Nome do PDF
            # Tentar encontrar ano no nome do arquivo (2023, 2024, 2025)
            ano_match = re.search(r'(202[3-5])', nome_arquivo)
            if ano_match:
                ano = ano_match.group(1)
            else:
                # Se não encontrou, usar "DESCONHECIDO" para não perder o PDF
                ano = "DESCONHECIDO"
                # Opcional: comentar a linha abaixo se quiser pular PDFs sem ano identificado
                # erros.append(f"Ano não identificado para: {caminho_pdf}")
                # continue
        
        # Processar PDF
        resultado = processar_pdf(str(caminho_pdf), empresa, ano)
        
        if resultado:
            todos_resultados.extend(resultado)
    
    if erros:
        print(f"\n{len(erros)} erros encontrados durante o processamento.")
        for erro in erros[:10]:  # Mostrar apenas os primeiros 10
            print(f"  - {erro}")
        if len(erros) > 10:
            print(f"  ... e mais {len(erros) - 10} erros.")
    
    return todos_resultados

# ============================================================================
# FUNÇÕES DE GERAÇÃO DE EXCEL
# ============================================================================

def gerar_aba_analitica_por_ano(df_completo: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """
    Gera abas analíticas separadas por ano.
    """
    abas = {}
    anos = sorted(df_completo["ano"].unique())
    
    for ano in anos:
        df_ano = df_completo[df_completo["ano"] == ano].copy()
        abas[f"analitico_{ano}"] = df_ano
    
    return abas

def gerar_aba_resumo(df_completo: pd.DataFrame) -> pd.DataFrame:
    """
    Gera aba de resumo agrupada por empresa, ano e grupo.
    """
    resumo = df_completo.groupby(["empresa", "ano", "grupo"]).agg({
        "pdf_nome": "nunique",  # PDFs únicos com ocorrência
        "ocorrencias_total_grupo": "sum"  # Total de ocorrências
    }).reset_index()
    
    resumo.columns = ["empresa", "ano", "grupo", "pdfs_com_ocorrencia", "ocorrencias_total"]
    resumo = resumo.sort_values(["empresa", "ano", "grupo"])
    
    return resumo

def gerar_aba_evolucao(df_completo: pd.DataFrame) -> pd.DataFrame:
    """
    Gera aba de evolução por empresa e grupo com deltas e percentuais.
    """
    # Criar pivot table: empresa + grupo como índice, ano como coluna
    pivot = df_completo.groupby(["empresa", "grupo", "ano"])["ocorrencias_total_grupo"].sum().reset_index()
    pivot = pivot.pivot_table(
        index=["empresa", "grupo"],
        columns="ano",
        values="ocorrencias_total_grupo",
        fill_value=0
    ).reset_index()
    
    # Garantir que temos colunas para 2023, 2024, 2025 (mesmo que não existam dados)
    anos_esperados = ["2023", "2024", "2025"]
    for ano in anos_esperados:
        if ano not in pivot.columns:
            pivot[ano] = 0
    
    # Renomear colunas de ano (apenas as que existem)
    renomear = {}
    for ano in anos_esperados:
        if ano in pivot.columns:
            renomear[ano] = f"ocorr_{ano}"
    pivot = pivot.rename(columns=renomear)
    
    # Garantir que as colunas renomeadas existam (criar com 0 se não existirem)
    for col in ["ocorr_2023", "ocorr_2024", "ocorr_2025"]:
        if col not in pivot.columns:
            pivot[col] = 0
    
    # Calcular deltas
    pivot["delta_24_23"] = pivot["ocorr_2024"] - pivot["ocorr_2023"]
    pivot["delta_25_24"] = pivot["ocorr_2025"] - pivot["ocorr_2024"]
    
    # Calcular percentuais (evitar divisão por zero)
    pivot["pct_24_23"] = pivot.apply(
        lambda row: (row["ocorr_2024"] / row["ocorr_2023"] * 100) if row["ocorr_2023"] > 0 else 0,
        axis=1
    )
    pivot["pct_25_24"] = pivot.apply(
        lambda row: (row["ocorr_2025"] / row["ocorr_2024"] * 100) if row["ocorr_2024"] > 0 else 0,
        axis=1
    )
    
    # Arredondar percentuais
    pivot["pct_24_23"] = pivot["pct_24_23"].round(2)
    pivot["pct_25_24"] = pivot["pct_25_24"].round(2)
    
    return pivot

def gerar_aba_auditoria() -> pd.DataFrame:
    """
    Gera aba de auditoria com lista de termos por grupo.
    """
    dados = []
    
    # Grupo IA_LLM
    dados.append({
        "grupo": "IA_LLM",
        "tipo": "Termos",
        "lista": ", ".join(TERMOS_IA_LLM["IA_LLM"])
    })
    dados.append({
        "grupo": "IA_LLM",
        "tipo": "Siglas Sensíveis",
        "lista": ", ".join(TERMOS_IA_LLM["SIGLAS_SENSIVEIS"])
    })
    
    # Grupo DADOS_BI
    dados.append({
        "grupo": "DADOS_BI",
        "tipo": "Termos",
        "lista": ", ".join(TERMOS_DADOS_BI["DADOS_BI"])
    })
    dados.append({
        "grupo": "DADOS_BI",
        "tipo": "Siglas Sensíveis",
        "lista": ", ".join(TERMOS_DADOS_BI["SIGLAS_SENSIVEIS"])
    })
    
    return pd.DataFrame(dados)

def gerar_excel(resultados: List[Dict]):
    """
    Gera arquivo Excel com todas as abas solicitadas.
    """
    if not resultados:
        print("Nenhum resultado para gerar Excel.")
        return
    
    # Criar DataFrame completo
    df_completo = pd.DataFrame(resultados)
    
    print(f"\nGerando Excel com {len(df_completo)} registros...")
    
    # Criar writer Excel
    with pd.ExcelWriter(ARQUIVO_EXCEL_SAIDA, engine='openpyxl') as writer:
        # Abas analíticas por ano
        abas_analiticas = gerar_aba_analitica_por_ano(df_completo)
        for nome_aba, df_aba in abas_analiticas.items():
            df_aba.to_excel(writer, sheet_name=nome_aba, index=False)
        
        # Aba agregada (todos os anos)
        df_completo.to_excel(writer, sheet_name="analitico_todos", index=False)
        
        # Aba de resumo
        df_resumo = gerar_aba_resumo(df_completo)
        df_resumo.to_excel(writer, sheet_name="resumo_empresas", index=False)
        
        # Aba de evolução
        df_evolucao = gerar_aba_evolucao(df_completo)
        df_evolucao.to_excel(writer, sheet_name="evolucao", index=False)
        
        # Aba de auditoria
        df_auditoria = gerar_aba_auditoria()
        df_auditoria.to_excel(writer, sheet_name="parametros", index=False)
    
    print(f"\n✓ Excel gerado com sucesso: {ARQUIVO_EXCEL_SAIDA}")
    print(f"  Total de registros: {len(df_completo)}")
    print(f"  Total de PDFs únicos: {df_completo['pdf_nome'].nunique()}")

# ============================================================================
# MAIN
# ============================================================================

def main():
    """Função principal."""
    print("=" * 70)
    print("ANÁLISE DE TERMOS EM PDFs - IA vs Dados/BI")
    print("=" * 70)
    print(f"Pasta raiz: {PASTA_RAIZ}")
    print(f"Incluir PDFs sem ocorrências: {INCLUIR_PDFS_SEM_OCORRENCIAS}")
    if EMPRESA_FILTRO:
        print(f"Filtro de empresa: {EMPRESA_FILTRO} (apenas esta empresa será processada)")
    else:
        print("Filtro de empresa: Nenhum (todas as empresas serão processadas)")
    print("=" * 70)
    
    try:
        # Varrer pastas e processar PDFs
        resultados = varrer_pastas()
        
        if resultados:
            # Gerar Excel
            gerar_excel(resultados)
        else:
            print("\nNenhum resultado encontrado.")
    
    except Exception as e:
        print(f"\nERRO CRÍTICO: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
