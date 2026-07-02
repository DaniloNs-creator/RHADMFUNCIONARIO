# -*- coding: utf-8 -*-
"""
==================================================================================
SPED STUDIO — Plataforma Corporativa de Leitura, Validação, Correção e Exportação
de Arquivos SPED (EFD ICMS/IPI e EFD Contribuições)
==================================================================================

Arquivo único (app.py), conforme requisito de manutenção do usuário.

O arquivo está organizado em SEÇÕES que representam, logicamente, os módulos que
existiriam em um projeto multi-arquivo (parser/ services/ rules/ validators/ ui/
exports/ utils/). Cada seção é isolada por um cabeçalho de comentário para
facilitar navegação, manutenção e uma eventual quebra futura em múltiplos
arquivos, caso o projeto cresça.

ÍNDICE DE SEÇÕES
  1. Configuração geral / constantes
  2. Layouts oficiais (subconjunto) dos principais registros SPED
  3. Utilitários (parsing numérico, decimal, datas)
  4. PARSER — leitura de arquivo SPED texto (|delimitado|)
  5. Identificação de tipo de arquivo / empresa / período
  6. MOTOR DE REGRAS TRIBUTÁRIAS (configurável)
  7. VALIDADORES — detecção de inconsistências fiscais
  8. SERVIÇOS — edição, correção em massa, desfazer, auditoria
  9. IMPORTAÇÃO DE CT-e (XML) PARA BLOCO D — exclusivo EFD Contribuições
 10. EXPORTAÇÃO — TXT SPED, Excel multi-abas, CSV
 11. INTERFACE (Streamlit) — dashboard, navegação, telas
 12. Bootstrap / main()

Requisitos: ver requirements.txt sugerido ao final deste arquivo (comentário).
==================================================================================
"""

from __future__ import annotations

import io
import re
import copy
import uuid
import base64
import zipfile
import hashlib
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Optional
import xml.etree.ElementTree as ET

import numpy as np
import pandas as pd
import streamlit as st

try:
    import openpyxl  # noqa: F401  (garante disponibilidade do engine do pandas)
except ImportError:
    pass


# ==================================================================================
# 1. CONFIGURAÇÃO GERAL / CONSTANTES
# ==================================================================================

APP_TITLE = "SPED Studio — Auditoria & Correção Fiscal"
APP_ICON = "🧾"

TIPO_ICMS_IPI = "EFD ICMS/IPI"
TIPO_CONTRIBUICOES = "EFD Contribuições"
TIPO_DESCONHECIDO = "Desconhecido"

# Blocos "assinatura" — presença forte indica o layout do arquivo
BLOCOS_ASSINATURA_CONTRIB = {"M", "F", "P"}       # M100/M105/M200... só existem em Contribuições
BLOCOS_ASSINATURA_ICMS_IPI = {"H", "K", "G"}      # Inventário, Produção e Estoque, Controle de ICMS ST

COLUNA_STATUS_ORIGINAL = "original"
COLUNA_STATUS_EDITADO = "editado"
COLUNA_STATUS_NOVO = "novo (importado)"

CORES = {
    "primaria": "#0B3D2E",
    "secundaria": "#134E36",
    "destaque": "#C9A24B",
    "fundo_card": "#F4F6F5",
    "erro": "#B3261E",
    "alerta": "#B7791F",
    "ok": "#1E7A4C",
}


# ==================================================================================
# 2. LAYOUTS OFICIAIS (SUBCONJUNTO) DOS PRINCIPAIS REGISTROS SPED
# ------------------------------------------------------------------------------
# Observação de engenharia: o leiaute completo da EFD ICMS/IPI e da EFD
# Contribuições possui centenas de registros e é atualizado periodicamente pelo
# Guia Prático (SPED Fiscal / SPED Contribuições). Mapeamos aqui os registros
# de maior relevância operacional para auditoria (abertura, documentos, itens,
# apuração e totalizadores). Registros não mapeados continuam sendo lidos e
# preservados normalmente (modo genérico campo_1..campo_N), garantindo que
# NENHUM dado seja perdido na leitura/gravação, mesmo sem leiaute nomeado.
# Ao evoluir o projeto, novos registros podem ser adicionados a este dicionário
# sem alterar nenhuma outra parte do sistema.
# ==================================================================================

REGISTRO_LAYOUTS: dict[str, list[str]] = {
    # --- Bloco 0 (comum) ---
    "0000": ["COD_VER", "COD_FIN", "DT_INI", "DT_FIN", "NOME", "CNPJ", "CPF",
              "UF", "IE", "COD_MUN", "IM", "SUFRAMA", "IND_PERFIL", "IND_ATIV"],
    "0001": ["IND_MOV"],
    "0150": ["COD_PART", "NOME", "COD_PAIS", "CNPJ", "CPF", "IE", "COD_MUN",
              "SUFRAMA", "ENDERECO", "NUM", "COMPL", "BAIRRO"],
    "0200": ["COD_ITEM", "DESCR_ITEM", "COD_BARRA", "COD_ANT_ITEM", "UNID_INV",
              "TIPO_ITEM", "COD_NCM", "EX_IPI", "COD_GEN", "COD_LST", "ALIQ_ICMS"],

    # --- Bloco C (Documentos Fiscais — Mercadorias) — comum às duas EFDs ---
    "C001": ["IND_MOV"],
    "C100": ["IND_OPER", "IND_EMIT", "COD_PART", "COD_MOD", "COD_SIT", "SER",
              "NUM_DOC", "CHV_NFE", "DT_DOC", "DT_E_S", "VL_DOC", "IND_PGTO",
              "VL_DESC", "VL_ABAT_NT", "VL_MERC", "IND_FRT", "VL_FRT",
              "VL_SEG", "VL_OUT_DA", "VL_BC_ICMS", "VL_ICMS", "VL_BC_ICMS_ST",
              "VL_ICMS_ST", "VL_IPI", "VL_PIS", "VL_COFINS", "VL_PIS_ST", "VL_COFINS_ST"],
    "C170": ["NUM_ITEM", "COD_ITEM", "DESCR_COMPL", "QTD", "UNID", "VL_ITEM",
              "VL_DESC", "IND_MOV", "CST_ICMS", "CFOP", "COD_NAT", "VL_BC_ICMS",
              "ALIQ_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", "ALIQ_ST", "VL_ICMS_ST",
              "IND_APUR", "CST_IPI", "COD_ENQ", "VL_BC_IPI", "ALIQ_IPI", "VL_IPI",
              "CST_PIS", "VL_BC_PIS", "ALIQ_PIS", "VL_PIS", "CST_COFINS",
              "VL_BC_COFINS", "ALIQ_COFINS", "VL_COFINS", "COD_CTA", "VL_ABAT_NAO_TRIB"],
    "C190": ["CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_OPR", "VL_BC_ICMS", "VL_ICMS",
              "VL_BC_ICMS_ST", "VL_ICMS_ST", "VL_RED_BC", "VL_IPI", "COD_OBS"],
    "C500": ["IND_OPER", "IND_EMIT", "COD_PART", "COD_MOD", "COD_SIT", "SER",
              "SUB", "NUM_DOC", "DT_DOC", "DT_E_S", "VL_DOC", "VL_DESC",
              "VL_FORN", "VL_SERV_NT", "VL_TERC", "VL_DA", "VL_BC_ICMS", "VL_ICMS"],

    # --- Bloco D (Documentos Fiscais — Serviços de Transporte / CT-e) ---
    "D001": ["IND_MOV"],
    "D100": ["IND_OPER", "IND_EMIT", "COD_PART", "COD_MOD", "COD_SIT", "SER",
              "NUM_DOC", "CHV_CTE", "DT_DOC", "DT_A_P", "TP_CTE", "CHV_CTE_REF",
              "VL_DOC", "VL_DESC", "IND_FRT", "VL_SERV", "VL_BC_ICMS", "VL_ICMS",
              "VL_NT", "COD_INF", "COD_CTA"],
    "D101": ["VL_BC_PIS", "ALIQ_PIS", "VL_PIS", "COD_CTA"],
    "D105": ["VL_BC_COFINS", "ALIQ_COFINS", "VL_COFINS", "COD_CTA"],
    "D190": ["CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_OPR", "VL_BC_ICMS", "VL_ICMS",
              "VL_RED_BC", "COD_OBS"],

    # --- Bloco E (Apuração ICMS/IPI — EFD ICMS/IPI) ---
    "E001": ["IND_MOV"],
    "E110": ["VL_TOT_DEBITOS", "VL_AJ_DEBITOS", "VL_TOT_AJ_DEBITOS", "VL_ESTORNOS_CRED",
              "VL_TOT_CREDITOS", "VL_AJ_CREDITOS", "VL_TOT_AJ_CREDITOS", "VL_ESTORNOS_DEB",
              "VL_SLD_CREDOR_ANT", "VL_SLD_APURADO", "VL_TOT_DED", "VL_ICMS_RECOLHER",
              "VL_SLD_CREDOR_TRANSPORTAR", "DEB_ESP"],

    # --- Bloco M (Apuração PIS/COFINS — EFD Contribuições) ---
    "M001": ["IND_MOV"],
    "M100": ["COD_CRED", "IND_CRED_ORI", "VL_BC_PIS", "ALIQ_PIS", "VL_CRED_PIS",
              "VL_AJUS_ACRES", "VL_AJUS_REDUC", "VL_CRED_DIF", "VL_CRED_DISP",
              "PER_DE_CRED", "VL_CRED_DESC", "VL_CRED_OUT", "COD_CTA"],
    "M105": ["NAT_BC_CRED", "VL_BC_PIS_TOT", "VL_BC_PIS_CUM", "VL_BC_PIS_NC",
              "VL_BC_PIS", "VL_CRED_PIS_TOT", "VL_CRED_PIS_NC"],
    "M200": ["VL_TOT_CONT_NC_PER", "VL_TOT_CRED_DESC", "VL_TOT_CRED_DESC_ANT",
              "VL_TOT_CONT_NC_DEV", "VL_RET_NC", "VL_OUT_DED_NC", "VL_CONT_NC_REC",
              "VL_TOT_CONT_CUM_PER", "VL_RET_CUM", "VL_OUT_DED_CUM", "VL_CONT_CUM_REC",
              "VL_TOT_CONT_REC"],
    "M210": ["COD_CONT", "VL_REC_BRT", "VL_BC_CONT", "ALIQ_PIS", "QUANT_BC_PIS",
              "ALIQ_PIS_QUANT", "VL_CONT_APUR", "VL_AJUS_ACRES", "VL_AJUS_REDUC",
              "VL_CONT_DIFER", "VL_CONT_DIFER_ANT", "VL_CONT_PER"],

    # --- Encerramento ---
    "9001": ["IND_MOV"],
    "9900": ["REG_BLC", "QTD_REG_BLC"],
    "9990": ["QTD_LIN_9"],
    "9999": ["QTD_LIN"],
}

# Registros de item usados na detecção de inconsistências, por tipo de arquivo
REGISTRO_ITEM_POR_TIPO = {
    TIPO_ICMS_IPI: "C170",
    TIPO_CONTRIBUICOES: "C170",
}


# ==================================================================================
# 3. UTILITÁRIOS
# ==================================================================================

def dec(valor, default="0") -> Decimal:
    """Converte string SPED (vírgula decimal) em Decimal, com fallback seguro."""
    if valor is None:
        valor = default
    valor = str(valor).strip()
    if valor == "":
        valor = default
    valor = valor.replace(".", "").replace(",", ".") if "," in valor else valor
    try:
        return Decimal(valor)
    except InvalidOperation:
        try:
            return Decimal(default)
        except InvalidOperation:
            return Decimal("0")


def dec_to_sped(valor: Decimal, casas=2) -> str:
    """Formata Decimal no padrão SPED (vírgula decimal, sem separador de milhar)."""
    quant = Decimal("1." + ("0" * casas)) if casas > 0 else Decimal("1")
    valor = valor.quantize(quant, rounding=ROUND_HALF_UP)
    s = f"{valor:.{casas}f}"
    return s.replace(".", ",")


def safe_get(lista: list, idx: int, default=""):
    return lista[idx] if idx is not None and 0 <= idx < len(lista) else default


def novo_id() -> str:
    return uuid.uuid4().hex[:12]


def agora_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def gerar_chave_hash(*partes) -> str:
    h = hashlib.sha1("|".join(str(p) for p in partes).encode("utf-8", "ignore"))
    return h.hexdigest()[:16]


# ==================================================================================
# 4. PARSER — LEITURA DE ARQUIVO SPED (TXT, DELIMITADO POR "|")
# ------------------------------------------------------------------------------
# Implementado manualmente (sem dependência de terceiros) para garantir
# robustez, previsibilidade e total controle sobre reconstrução do arquivo na
# exportação. Preserva 100% dos campos originais, mesmo de registros não
# mapeados em REGISTRO_LAYOUTS.
# ==================================================================================

@dataclass
class RegistroSped:
    idx: int                     # posição sequencial no arquivo (ordem original)
    bloco: str
    registro: str
    campos: list                 # lista de strings, na ordem original (sem REG)
    origem: str = "sped"         # "sped" | "cte_import"
    status: str = COLUNA_STATUS_ORIGINAL
    uid: str = field(default_factory=novo_id)


def parse_sped(conteudo: str) -> list[RegistroSped]:
    """Faz o parsing de um arquivo SPED texto delimitado por pipe."""
    linhas = conteudo.splitlines()
    registros: list[RegistroSped] = []
    for i, linha in enumerate(linhas):
        linha = linha.strip("\r\n")
        if not linha.strip():
            continue
        # remove pipe inicial/final se existir: |REG|C1|C2|...|
        partes = linha.split("|")
        if partes and partes[0] == "":
            partes = partes[1:]
        if partes and partes[-1] == "":
            partes = partes[:-1]
        if not partes:
            continue
        reg = partes[0].strip().upper()
        campos = partes[1:]
        bloco = reg[0] if reg else "?"
        registros.append(RegistroSped(idx=i, bloco=bloco, registro=reg, campos=campos))
    return registros


def registros_para_dataframe(registros: list[RegistroSped]) -> pd.DataFrame:
    """DataFrame 'mestre' leve, usado para navegação, dashboard e filtros."""
    dados = [{
        "uid": r.uid, "idx": r.idx, "bloco": r.bloco, "registro": r.registro,
        "n_campos": len(r.campos), "origem": r.origem, "status": r.status,
    } for r in registros]
    df = pd.DataFrame(dados)
    return df


def registro_para_dict_nomeado(r: RegistroSped) -> dict:
    """Expande os campos posicionais em um dicionário nomeado, se houver leiaute
    conhecido para o registro. Caso contrário, usa nomes genéricos campo_N."""
    layout = REGISTRO_LAYOUTS.get(r.registro)
    d = {"uid": r.uid, "idx": r.idx, "bloco": r.bloco, "registro": r.registro,
         "status": r.status, "origem": r.origem}
    if layout:
        for i, nome in enumerate(layout):
            d[nome] = safe_get(r.campos, i)
    else:
        for i, val in enumerate(r.campos):
            d[f"campo_{i+1}"] = val
    return d


def dataframe_detalhado(registros: list[RegistroSped], registro_tipo: str) -> pd.DataFrame:
    """Monta um DataFrame detalhado (colunas nomeadas) para um tipo de registro
    específico (ex.: 'C170'), usado nas telas de Visão por Registros / Itens."""
    filtrados = [r for r in registros if r.registro == registro_tipo]
    if not filtrados:
        return pd.DataFrame()
    linhas = [registro_para_dict_nomeado(r) for r in filtrados]
    return pd.DataFrame(linhas)


# ==================================================================================
# 5. IDENTIFICAÇÃO DE TIPO DE ARQUIVO / EMPRESA / PERÍODO
# ==================================================================================

def identificar_tipo_arquivo(registros: list[RegistroSped]) -> str:
    blocos_presentes = {r.bloco for r in registros}
    registros_presentes = {r.registro for r in registros}
    if {"M100", "M105", "M200", "M210"} & registros_presentes:
        return TIPO_CONTRIBUICOES
    if blocos_presentes & BLOCOS_ASSINATURA_ICMS_IPI:
        return TIPO_ICMS_IPI
    if blocos_presentes & BLOCOS_ASSINATURA_CONTRIB:
        return TIPO_CONTRIBUICOES
    # fallback: EFD Contribuições costuma ter apenas A/C/D/F/I/M/P/1/9
    if blocos_presentes <= {"A", "C", "D", "F", "I", "M", "P", "1", "9", "0"}:
        return TIPO_CONTRIBUICOES
    return TIPO_ICMS_IPI


def extrair_info_empresa(registros: list[RegistroSped]) -> dict:
    zero = next((r for r in registros if r.registro == "0000"), None)
    if not zero:
        return {}
    d = registro_para_dict_nomeado(zero)
    return {
        "razao_social": d.get("NOME", ""),
        "cnpj": d.get("CNPJ", ""),
        "uf": d.get("UF", ""),
        "ie": d.get("IE", ""),
        "dt_ini": d.get("DT_INI", ""),
        "dt_fin": d.get("DT_FIN", ""),
        "cod_ver": d.get("COD_VER", ""),
    }


# ==================================================================================
# 6. MOTOR DE REGRAS TRIBUTÁRIAS (CONFIGURÁVEL)
# ------------------------------------------------------------------------------
# A tabela de regras é mantida em st.session_state como DataFrame editável.
# Cada linha define, para uma combinação CST + CFOP + tipo de operação:
#   - se exige base de cálculo, alíquota e valor do imposto
#   - alíquota padrão sugerida (usada quando a alíquota estiver ausente)
#   - base de cálculo padrão: "VL_ITEM" (valor do item) ou "VL_DOC" (valor doc.)
# A fórmula de cálculo do imposto é sempre: imposto = base * aliquota / 100,
# mas fica isolada em uma função (calcular_imposto) para permitir substituição
# futura por fórmulas mais complexas (ex.: MVA/ICMS-ST, redução de base etc.)
# ==================================================================================

def regras_padrao() -> pd.DataFrame:
    dados = [
        # CST(ICMS/PIS-COFINS simplificado) | CFOP | tipo_operacao | exige_base | exige_aliquota | exige_imposto | aliquota_padrao | base_padrao
        {"cst": "000", "cfop_prefixo": "5", "tipo_operacao": "Saída", "tributo": "ICMS",
         "exige_base": True, "exige_aliquota": True, "exige_imposto": True,
         "aliquota_padrao": "18,00", "base_padrao": "VL_ITEM"},
        {"cst": "060", "cfop_prefixo": "5", "tipo_operacao": "Saída", "tributo": "ICMS",
         "exige_base": False, "exige_aliquota": False, "exige_imposto": False,
         "aliquota_padrao": "0,00", "base_padrao": "VL_ITEM"},
        {"cst": "040", "cfop_prefixo": "5", "tipo_operacao": "Saída", "tributo": "ICMS",
         "exige_base": False, "exige_aliquota": False, "exige_imposto": False,
         "aliquota_padrao": "0,00", "base_padrao": "VL_ITEM"},
        {"cst": "01", "cfop_prefixo": "5", "tipo_operacao": "Saída", "tributo": "PIS",
         "exige_base": True, "exige_aliquota": True, "exige_imposto": True,
         "aliquota_padrao": "1,65", "base_padrao": "VL_ITEM"},
        {"cst": "01", "cfop_prefixo": "5", "tipo_operacao": "Saída", "tributo": "COFINS",
         "exige_base": True, "exige_aliquota": True, "exige_imposto": True,
         "aliquota_padrao": "7,60", "base_padrao": "VL_ITEM"},
        {"cst": "04", "cfop_prefixo": "5", "tipo_operacao": "Saída", "tributo": "PIS",
         "exige_base": False, "exige_aliquota": False, "exige_imposto": False,
         "aliquota_padrao": "0,00", "base_padrao": "VL_ITEM"},
        {"cst": "04", "cfop_prefixo": "5", "tipo_operacao": "Saída", "tributo": "COFINS",
         "exige_base": False, "exige_aliquota": False, "exige_imposto": False,
         "aliquota_padrao": "0,00", "base_padrao": "VL_ITEM"},
    ]
    df = pd.DataFrame(dados)
    df.insert(0, "regra_id", [novo_id() for _ in range(len(df))])
    df["ativo"] = True
    return df


def buscar_regra(regras: pd.DataFrame, cst: str, cfop: str, tributo: str) -> Optional[dict]:
    if regras is None or regras.empty:
        return None
    cst = (cst or "").strip()
    cfop = (cfop or "").strip()
    candidatos = regras[
        (regras["tributo"] == tributo)
        & (regras["ativo"])
        & (regras["cst"].astype(str).str.strip() == cst)
        & (cfop.startswith(regras["cfop_prefixo"].astype(str)) if cfop else True)
    ]
    if candidatos.empty:
        return None
    return candidatos.iloc[0].to_dict()


def calcular_imposto(base: Decimal, aliquota_pct: Decimal) -> Decimal:
    """Fórmula central e única do motor: imposto = base * aliquota / 100."""
    return (base * aliquota_pct / Decimal("100"))


# ==================================================================================
# 7. VALIDADORES — DETECÇÃO DE INCONSISTÊNCIAS FISCAIS
# ==================================================================================

def detectar_inconsistencias(df_c170: pd.DataFrame, regras: pd.DataFrame,
                              tipo_arquivo: str) -> pd.DataFrame:
    """Varre os itens (C170) e aponta inconsistências críticas, com base no
    motor de regras configurável. Retorna um DataFrame de achados, um por
    combinação item + tributo com problema."""
    if df_c170 is None or df_c170.empty:
        return pd.DataFrame(columns=[
            "uid", "idx", "tributo", "cst", "cfop", "problema", "severidade",
            "vl_item", "base_atual", "aliquota_atual", "imposto_atual",
            "base_sugerida", "aliquota_sugerida", "imposto_sugerido",
        ])

    achados = []
    tributos = ["ICMS", "PIS", "COFINS"] if tipo_arquivo == TIPO_CONTRIBUICOES else ["ICMS", "IPI"]
    # ICMS sempre relevante nas duas EFDs; para Contribuições focamos PIS/COFINS
    # mas mantemos ICMS por consistência de leitura do C170.
    mapa_campos = {
        "ICMS": ("CST_ICMS", "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS"),
        "IPI":  ("CST_IPI", "VL_BC_IPI", "ALIQ_IPI", "VL_IPI"),
        "PIS":  ("CST_PIS", "VL_BC_PIS", "ALIQ_PIS", "VL_PIS"),
        "COFINS": ("CST_COFINS", "VL_BC_COFINS", "ALIQ_COFINS", "VL_COFINS"),
    }

    for _, row in df_c170.iterrows():
        vl_item = dec(row.get("VL_ITEM", "0"))
        cfop = str(row.get("CFOP", "")).strip()
        for tributo in tributos:
            campo_cst, campo_base, campo_aliq, campo_imp = mapa_campos[tributo]
            if campo_cst not in row:
                continue
            cst = str(row.get(campo_cst, "")).strip()
            if cst == "":
                continue
            base_atual = row.get(campo_base, "")
            aliq_atual = row.get(campo_aliq, "")
            imp_atual = row.get(campo_imp, "")

            regra = buscar_regra(regras, cst, cfop, tributo)
            exige_base = exige_aliq = exige_imp = None
            aliquota_padrao = Decimal("0")
            if regra:
                exige_base = bool(regra["exige_base"])
                exige_aliq = bool(regra["exige_aliquota"])
                exige_imp = bool(regra["exige_imposto"])
                aliquota_padrao = dec(regra.get("aliquota_padrao", "0"))
            else:
                # sem regra cadastrada: heurística conservadora — se já existe
                # valor de imposto lançado, exige-se consistência dos 3 campos.
                tem_algum_valor = any(str(v).strip() not in ("", "0", "0,00")
                                       for v in (base_atual, aliq_atual, imp_atual))
                exige_base = exige_aliq = exige_imp = tem_algum_valor

            base_vazia = str(base_atual).strip() in ("", "0", "0,00")
            aliq_vazia = str(aliq_atual).strip() in ("", "0", "0,00")
            imp_vazio = str(imp_atual).strip() in ("", "0", "0,00")

            problemas = []
            if exige_base and base_vazia:
                problemas.append("Base de cálculo ausente")
            if exige_aliq and aliq_vazia:
                problemas.append("Alíquota ausente")
            if exige_imp and imp_vazio:
                problemas.append("Valor do imposto ausente")

            # inconsistência de coerência: imposto informado mas não bate com base*aliq
            if not base_vazia and not aliq_vazia and not imp_vazio:
                esperado = calcular_imposto(dec(base_atual), dec(aliq_atual))
                informado = dec(imp_atual)
                if abs(esperado - informado) > Decimal("0.05"):
                    problemas.append(
                        f"Imposto divergente do cálculo (esperado {dec_to_sped(esperado)})")

            if not problemas:
                continue

            base_sug = vl_item if base_vazia else dec(base_atual)
            aliq_sug = aliquota_padrao if aliq_vazia else dec(aliq_atual)
            imp_sug = calcular_imposto(base_sug, aliq_sug)

            severidade = "Crítica" if (exige_base and exige_aliq and exige_imp) else "Atenção"

            achados.append({
                "uid": row.get("uid"), "idx": row.get("idx"), "tributo": tributo,
                "cst": cst, "cfop": cfop, "problema": "; ".join(problemas),
                "severidade": severidade, "vl_item": dec_to_sped(vl_item),
                "base_atual": base_atual, "aliquota_atual": aliq_atual,
                "imposto_atual": imp_atual,
                "base_sugerida": dec_to_sped(base_sug),
                "aliquota_sugerida": dec_to_sped(aliq_sug),
                "imposto_sugerido": dec_to_sped(imp_sug),
                "campo_base": campo_base, "campo_aliq": campo_aliq, "campo_imp": campo_imp,
            })

    return pd.DataFrame(achados)


def validar_integridade_blocos(registros: list[RegistroSped]) -> list[dict]:
    """Verifica abertura/fechamento de blocos e presença de totalizador 9900."""
    problemas = []
    blocos_abertos = {}
    for r in registros:
        if r.registro.endswith("001"):
            blocos_abertos[r.bloco] = r.idx
    contagem_por_registro = {}
    for r in registros:
        contagem_por_registro[r.registro] = contagem_por_registro.get(r.registro, 0) + 1

    reg_9900 = {r.registro: safe_get(r.campos, 0) for r in registros if r.registro == "9900"}
    # reg_9900 fica sobrescrito por chave repetida; usar lista real:
    linhas_9900 = [(safe_get(r.campos, 0), safe_get(r.campos, 1))
                   for r in registros if r.registro == "9900"]
    contagem_9900 = {reg: int(qtd) for reg, qtd in linhas_9900 if str(qtd).isdigit()}

    for registro, qtd_real in contagem_por_registro.items():
        if registro in contagem_9900 and contagem_9900[registro] != qtd_real:
            problemas.append({
                "tipo": "Totalizador 9900 divergente",
                "registro": registro,
                "detalhe": f"9900 informa {contagem_9900[registro]} ocorrências, "
                           f"arquivo contém {qtd_real}.",
            })
    return problemas


# ==================================================================================
# 8. SERVIÇOS — EDIÇÃO, CORREÇÃO EM MASSA, AUDITORIA
# ==================================================================================

def registrar_auditoria(uid: str, registro: str, campo: str, valor_anterior,
                         valor_novo, motivo: str, regra_aplicada: str = ""):
    st.session_state.audit_log.append({
        "data_hora": agora_str(),
        "usuario": st.session_state.get("usuario_atual", "analista.fiscal"),
        "uid_registro": uid,
        "registro": registro,
        "campo": campo,
        "valor_anterior": valor_anterior,
        "valor_novo": valor_novo,
        "regra_aplicada": regra_aplicada,
        "motivo": motivo,
    })


def get_registro_por_uid(uid: str) -> Optional[RegistroSped]:
    return st.session_state.registros_map.get(uid)


def atualizar_campo_registro(uid: str, nome_campo: str, novo_valor: str, motivo: str,
                              regra_aplicada: str = ""):
    r = get_registro_por_uid(uid)
    if r is None:
        return False
    layout = REGISTRO_LAYOUTS.get(r.registro)
    if not layout or nome_campo not in layout:
        return False
    pos = layout.index(nome_campo)
    while len(r.campos) <= pos:
        r.campos.append("")
    valor_anterior = r.campos[pos]
    if str(valor_anterior) == str(novo_valor):
        return True
    r.campos[pos] = novo_valor
    r.status = COLUNA_STATUS_EDITADO
    registrar_auditoria(uid, r.registro, nome_campo, valor_anterior, novo_valor,
                         motivo, regra_aplicada)
    return True


def aplicar_correcao_massa(uids: list[str], campo_base: str, campo_aliq: str,
                            campo_imp: str, base_valor: Optional[str],
                            aliquota_valor: Optional[str], regra_nome: str = "Correção em massa"):
    aplicados = 0
    for uid in uids:
        r = get_registro_por_uid(uid)
        if r is None:
            continue
        layout = REGISTRO_LAYOUTS.get(r.registro)
        if not layout:
            continue
        d = registro_para_dict_nomeado(r)
        base_final = dec(base_valor) if base_valor not in (None, "") else dec(d.get(campo_base, "0"))
        aliq_final = dec(aliquota_valor) if aliquota_valor not in (None, "") else dec(d.get(campo_aliq, "0"))
        imposto_final = calcular_imposto(base_final, aliq_final)

        if base_valor not in (None, ""):
            atualizar_campo_registro(uid, campo_base, dec_to_sped(base_final),
                                      "Correção em massa - base", regra_nome)
        if aliquota_valor not in (None, ""):
            atualizar_campo_registro(uid, campo_aliq, dec_to_sped(aliq_final),
                                      "Correção em massa - alíquota", regra_nome)
        atualizar_campo_registro(uid, campo_imp, dec_to_sped(imposto_final),
                                  "Correção em massa - imposto recalculado", regra_nome)
        aplicados += 1
    return aplicados


def desfazer_ultima_alteracao_uid(uid: str):
    """Restaura o registro ao seu estado original, com base no snapshot inicial."""
    original = st.session_state.registros_originais_map.get(uid)
    atual = get_registro_por_uid(uid)
    if not original or not atual:
        return False
    campos_antes = list(atual.campos)
    atual.campos = list(original.campos)
    atual.status = COLUNA_STATUS_ORIGINAL if atual.origem == "sped" else COLUNA_STATUS_NOVO
    registrar_auditoria(uid, atual.registro, "(registro completo)",
                         "|".join(campos_antes), "|".join(atual.campos),
                         "Restauração ao valor original")
    return True


# ==================================================================================
# 9. IMPORTAÇÃO DE CT-e (XML) PARA BLOCO D — EXCLUSIVO EFD CONTRIBUIÇÕES
# ------------------------------------------------------------------------------
# Gera, a partir de cada XML de CT-e, os registros:
#   D100 — dados gerais do conhecimento de transporte
#   D101 — informações de crédito de PIS sobre o serviço de transporte
#   D105 — informações de crédito de COFINS sobre o serviço de transporte
# A base de cálculo do PIS/COFINS é sugerida como o valor da prestação do
# serviço (vTPrest) e a alíquota vem do motor de regras (tributo="PIS"/"COFINS",
# CST configurável na tela de importação). Tudo fica marcado como "novo
# (importado)" e sujeito a REVISÃO MANUAL antes da exportação, exatamente como
# a lógica de correção assistida do restante do sistema.
# ==================================================================================

NS_CTE_CANDIDATAS = [
    "{http://www.portalfiscal.inf.br/cte}",
    "",  # fallback sem namespace
]


def _find(elem: ET.Element, caminho: str):
    for ns in NS_CTE_CANDIDATAS:
        tag_path = "/".join(f"{ns}{p}" for p in caminho.split("/"))
        achado = elem.find(tag_path)
        if achado is not None:
            return achado
    return None


def _text(elem: ET.Element, caminho: str, default="") -> str:
    achado = _find(elem, caminho)
    return achado.text.strip() if (achado is not None and achado.text) else default


def parse_cte_xml(conteudo_bytes: bytes, nome_arquivo: str) -> dict:
    """Extrai os campos essenciais de um XML de CT-e (evento cteProc/CTe)."""
    try:
        root = ET.fromstring(conteudo_bytes)
    except ET.ParseError as e:
        return {"erro": f"XML inválido em {nome_arquivo}: {e}"}

    inf_cte = None
    for ns in NS_CTE_CANDIDATAS:
        inf_cte = root.find(f".//{ns}infCte")
        if inf_cte is not None:
            break
    if inf_cte is None:
        return {"erro": f"Não foi encontrado nó infCte em {nome_arquivo}."}

    chave = inf_cte.attrib.get("Id", "").replace("CTe", "").strip()
    ide = _find(inf_cte, "ide")
    emit = _find(inf_cte, "emit")
    dest = _find(inf_cte, "dest")
    vprest = _find(inf_cte, "vPrest")
    icms_root = _find(inf_cte, "imp/ICMS")

    icms_vals = {"vBC": "0", "pICMS": "0", "vICMS": "0", "CST": ""}
    if icms_root is not None:
        for filho in list(icms_root):
            for ns in NS_CTE_CANDIDATAS:
                tag = filho.tag.replace(ns, "")
            # filho é algo como ICMS00, ICMS20, ICMS45, ICMS60, ICMS90...
            icms_vals["CST"] = _text(filho, "CST", icms_vals["CST"]) or _text(filho, "CST")
            icms_vals["vBC"] = _text(filho, "vBC", icms_vals["vBC"])
            icms_vals["pICMS"] = _text(filho, "pICMS", icms_vals["pICMS"])
            icms_vals["vICMS"] = _text(filho, "vICMS", icms_vals["vICMS"])

    dados = {
        "arquivo": nome_arquivo,
        "chave_cte": chave,
        "cfop": _text(ide, "CFOP"),
        "nat_op": _text(ide, "natOp"),
        "serie": _text(ide, "serie"),
        "num_doc": _text(ide, "nCT"),
        "dt_emi": _text(ide, "dhEmi")[:10].replace("-", "") if _text(ide, "dhEmi") else "",
        "mod": _text(ide, "mod", "57"),
        "tp_cte": _text(ide, "tpCTe", "0"),
        "emit_cnpj": _text(emit, "CNPJ"),
        "emit_nome": _text(emit, "xNome"),
        "dest_cnpj": _text(dest, "CNPJ"),
        "dest_nome": _text(dest, "xNome"),
        "vl_tprest": _text(vprest, "vTPrest", "0"),
        "vl_rec": _text(vprest, "vRec", "0"),
        "icms_cst": icms_vals["CST"],
        "icms_vbc": icms_vals["vBC"],
        "icms_p": icms_vals["pICMS"],
        "icms_v": icms_vals["vICMS"],
    }
    return dados


def gerar_registros_d_a_partir_de_cte(cte: dict, regras: pd.DataFrame,
                                       cst_pis="01", cst_cofins="01",
                                       cod_cta="") -> list[RegistroSped]:
    """Converte os dados extraídos de um CT-e em registros D100 / D101 / D105,
    prontos para inserção no bloco D do SPED Contribuições."""
    vl_serv = dec(cte.get("vl_tprest", "0"))

    campos_d100 = [
        "0",                              # IND_OPER (0=Aquisição/Entrada — ajustar se necessário)
        "0",                              # IND_EMIT (0=Emissão de terceiros)
        "",                                # COD_PART (vincular ao cadastro 0150, se aplicável)
        cte.get("mod", "57"),             # COD_MOD
        "00",                              # COD_SIT
        cte.get("serie", ""),             # SER
        cte.get("num_doc", ""),           # NUM_DOC
        cte.get("chave_cte", ""),         # CHV_CTE
        cte.get("dt_emi", ""),            # DT_DOC
        cte.get("dt_emi", ""),            # DT_A_P
        cte.get("tp_cte", "0"),           # TP_CTE
        "",                                 # CHV_CTE_REF
        dec_to_sped(vl_serv),             # VL_DOC
        "0,00",                            # VL_DESC
        "0",                                # IND_FRT
        dec_to_sped(vl_serv),             # VL_SERV
        dec_to_sped(dec(cte.get("icms_vbc", "0"))),   # VL_BC_ICMS
        dec_to_sped(dec(cte.get("icms_v", "0"))),     # VL_ICMS
        "0,00",                            # VL_NT
        "",                                 # COD_INF
        cod_cta,                           # COD_CTA
    ]
    r_d100 = RegistroSped(idx=-1, bloco="D", registro="D100", campos=campos_d100,
                           origem="cte_import", status=COLUNA_STATUS_NOVO)

    regra_pis = buscar_regra(regras, cst_pis, cte.get("cfop", ""), "PIS")
    aliq_pis = dec(regra_pis["aliquota_padrao"]) if regra_pis else Decimal("1.65")
    vl_bc_pis = vl_serv
    vl_pis = calcular_imposto(vl_bc_pis, aliq_pis)
    campos_d101 = [dec_to_sped(vl_bc_pis), dec_to_sped(aliq_pis), dec_to_sped(vl_pis), cod_cta]
    r_d101 = RegistroSped(idx=-1, bloco="D", registro="D101", campos=campos_d101,
                           origem="cte_import", status=COLUNA_STATUS_NOVO)

    regra_cofins = buscar_regra(regras, cst_cofins, cte.get("cfop", ""), "COFINS")
    aliq_cofins = dec(regra_cofins["aliquota_padrao"]) if regra_cofins else Decimal("7.60")
    vl_bc_cofins = vl_serv
    vl_cofins = calcular_imposto(vl_bc_cofins, aliq_cofins)
    campos_d105 = [dec_to_sped(vl_bc_cofins), dec_to_sped(aliq_cofins), dec_to_sped(vl_cofins), cod_cta]
    r_d105 = RegistroSped(idx=-1, bloco="D", registro="D105", campos=campos_d105,
                           origem="cte_import", status=COLUNA_STATUS_NOVO)

    return [r_d100, r_d101, r_d105]


def importar_ctes(arquivos_upload, cst_pis: str, cst_cofins: str, cod_cta: str) -> dict:
    """Processa uma lista de uploads (XML avulsos ou um .zip contendo XMLs)."""
    resultado = {"importados": 0, "erros": [], "resumo": []}
    regras = st.session_state.regras_tributarias

    def _processar_um(nome, conteudo_bytes):
        cte = parse_cte_xml(conteudo_bytes, nome)
        if "erro" in cte:
            resultado["erros"].append(cte["erro"])
            return
        novos = gerar_registros_d_a_partir_de_cte(cte, regras, cst_pis, cst_cofins, cod_cta)
        proximo_idx = max((r.idx for r in st.session_state.registros), default=0) + 1
        for i, novo in enumerate(novos):
            novo.idx = proximo_idx + i
            st.session_state.registros.append(novo)
            st.session_state.registros_map[novo.uid] = novo
            st.session_state.registros_originais_map[novo.uid] = copy.deepcopy(novo)
            registrar_auditoria(novo.uid, novo.registro, "(criação via CT-e)",
                                 "", "|".join(novo.campos),
                                 f"Importação CT-e {cte.get('chave_cte','')}",
                                 "Importação Bloco D - CT-e")
        resultado["importados"] += 1
        resultado["resumo"].append({
            "arquivo": nome, "chave_cte": cte.get("chave_cte", ""),
            "cfop": cte.get("cfop", ""), "vl_prestacao": cte.get("vl_tprest", "0"),
            "emitente": cte.get("emit_nome", ""),
        })

    for up in arquivos_upload:
        nome = up.name
        conteudo = up.read()
        if nome.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(conteudo)) as z:
                for info in z.infolist():
                    if info.filename.lower().endswith(".xml"):
                        with z.open(info) as f:
                            _processar_um(info.filename, f.read())
        elif nome.lower().endswith(".xml"):
            _processar_um(nome, conteudo)
        else:
            resultado["erros"].append(f"Arquivo ignorado (formato não suportado): {nome}")

    st.session_state.registros_df = registros_para_dataframe(st.session_state.registros)
    return resultado


# ==================================================================================
# 10. EXPORTAÇÃO — TXT SPED, EXCEL MULTI-ABAS, CSV
# ==================================================================================

def reconstruir_linha_sped(r: RegistroSped) -> str:
    return "|" + "|".join([r.registro] + [str(c) for c in r.campos]) + "|"


def exportar_txt_sped(registros: list[RegistroSped]) -> bytes:
    ordenados = sorted(registros, key=lambda r: (r.idx if r.idx >= 0 else 10**9))
    linhas = [reconstruir_linha_sped(r) for r in ordenados]
    conteudo = "\r\n".join(linhas) + "\r\n"
    return conteudo.encode("latin-1", errors="replace")


def montar_excel_relatorio(df_inconsistencias: pd.DataFrame,
                            registros_alterados: pd.DataFrame,
                            df_regras: pd.DataFrame,
                            df_auditoria: pd.DataFrame,
                            info_empresa: dict) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        resumo = pd.DataFrame([{
            "Empresa": info_empresa.get("razao_social", ""),
            "CNPJ": info_empresa.get("cnpj", ""),
            "Período": f"{info_empresa.get('dt_ini','')} a {info_empresa.get('dt_fin','')}",
            "Total de inconsistências": len(df_inconsistencias) if df_inconsistencias is not None else 0,
            "Registros alterados": len(registros_alterados) if registros_alterados is not None else 0,
            "Gerado em": agora_str(),
        }])
        resumo.to_excel(writer, sheet_name="Resumo Gerencial", index=False)
        (df_inconsistencias if df_inconsistencias is not None else pd.DataFrame()).to_excel(
            writer, sheet_name="Inconsistencias", index=False)
        (registros_alterados if registros_alterados is not None else pd.DataFrame()).to_excel(
            writer, sheet_name="Registros Alterados", index=False)
        (df_regras if df_regras is not None else pd.DataFrame()).to_excel(
            writer, sheet_name="Regras Aplicadas", index=False)
        (df_auditoria if df_auditoria is not None else pd.DataFrame()).to_excel(
            writer, sheet_name="Log de Auditoria", index=False)
    return buffer.getvalue()


def link_download(dados: bytes, nome_arquivo: str, rotulo: str, mime: str):
    b64 = base64.b64encode(dados).decode()
    href = f'<a download="{nome_arquivo}" href="data:{mime};base64,{b64}">{rotulo}</a>'
    st.markdown(href, unsafe_allow_html=True)


# ==================================================================================
# 11. INTERFACE (STREAMLIT)
# ==================================================================================

def inicializar_estado():
    defaults = {
        "registros": [],
        "registros_map": {},
        "registros_originais_map": {},
        "registros_df": pd.DataFrame(),
        "tipo_arquivo": TIPO_DESCONHECIDO,
        "info_empresa": {},
        "regras_tributarias": regras_padrao(),
        "audit_log": [],
        "usuario_atual": "analista.fiscal",
        "arquivo_carregado": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def aplicar_estilo():
    st.markdown(f"""
    <style>
        .stApp {{ background-color: #FFFFFF; }}
        .metric-card {{
            background-color: {CORES['fundo_card']};
            border-left: 4px solid {CORES['primaria']};
            padding: 14px 16px; border-radius: 6px; margin-bottom: 8px;
        }}
        section[data-testid="stSidebar"] {{
            background-color: {CORES['primaria']};
        }}
        section[data-testid="stSidebar"] * {{ color: #F4F6F5 !important; }}
        h1, h2, h3 {{ color: {CORES['primaria']}; }}
        .badge-critica {{ background:{CORES['erro']}; color:white; padding:2px 8px; border-radius:10px; font-size:0.75em; }}
        .badge-atencao {{ background:{CORES['alerta']}; color:white; padding:2px 8px; border-radius:10px; font-size:0.75em; }}
        .badge-ok {{ background:{CORES['ok']}; color:white; padding:2px 8px; border-radius:10px; font-size:0.75em; }}
    </style>
    """, unsafe_allow_html=True)


def carregar_arquivo(conteudo_texto: str):
    registros = parse_sped(conteudo_texto)
    st.session_state.registros = registros
    st.session_state.registros_map = {r.uid: r for r in registros}
    st.session_state.registros_originais_map = {r.uid: copy.deepcopy(r) for r in registros}
    st.session_state.registros_df = registros_para_dataframe(registros)
    st.session_state.tipo_arquivo = identificar_tipo_arquivo(registros)
    st.session_state.info_empresa = extrair_info_empresa(registros)
    st.session_state.audit_log = []
    st.session_state.arquivo_carregado = True


# ---------- Páginas ----------

def pagina_upload():
    st.header("📤 Upload do Arquivo SPED")
    st.write("Envie um arquivo SPED (.txt) da **EFD ICMS/IPI** ou da **EFD Contribuições**.")
    up = st.file_uploader("Arquivo SPED", type=["txt"])
    if up is not None:
        conteudo = up.read().decode("latin-1", errors="replace")
        with st.spinner("Lendo e estruturando o arquivo..."):
            carregar_arquivo(conteudo)
        st.success(f"Arquivo lido com sucesso: {len(st.session_state.registros)} registros.")
        st.info(f"Tipo identificado: **{st.session_state.tipo_arquivo}**")
        info = st.session_state.info_empresa
        c1, c2, c3 = st.columns(3)
        c1.metric("Empresa", info.get("razao_social", "—"))
        c2.metric("CNPJ", info.get("cnpj", "—"))
        c3.metric("Período", f"{info.get('dt_ini','—')} a {info.get('dt_fin','—')}")

    if st.session_state.arquivo_carregado:
        st.divider()
        problemas = validar_integridade_blocos(st.session_state.registros)
        if problemas:
            st.warning(f"{len(problemas)} divergência(s) de totalizador encontradas (ver detalhes).")
            st.dataframe(pd.DataFrame(problemas), use_container_width=True)
        else:
            st.success("Nenhuma divergência de totalizador (9900) encontrada.")


def pagina_dashboard():
    st.header("📊 Dashboard")
    if not st.session_state.arquivo_carregado:
        st.info("Faça upload de um arquivo SPED na aba **Upload do Arquivo** para começar.")
        return

    df = st.session_state.registros_df
    tipo = st.session_state.tipo_arquivo
    reg_item = REGISTRO_ITEM_POR_TIPO.get(tipo, "C170")
    df_itens = dataframe_detalhado(st.session_state.registros, reg_item)
    inconsistencias = detectar_inconsistencias(df_itens, st.session_state.regras_tributarias, tipo)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total de registros", len(df))
    c2.metric("Blocos distintos", df["bloco"].nunique() if not df.empty else 0)
    c3.metric("Itens (C170)", len(df_itens))
    c4.metric("Inconsistências", len(inconsistencias))

    st.markdown("#### Registros por bloco")
    if not df.empty:
        contagem_bloco = df.groupby("bloco").size().reset_index(name="quantidade").sort_values("bloco")
        st.bar_chart(contagem_bloco.set_index("bloco"))

    if not df_itens.empty and "CFOP" in df_itens.columns:
        st.markdown("#### Itens por CFOP")
        cont_cfop = df_itens.groupby("CFOP").size().reset_index(name="quantidade").sort_values(
            "quantidade", ascending=False).head(15)
        st.bar_chart(cont_cfop.set_index("CFOP"))

    if not inconsistencias.empty:
        st.markdown("#### Inconsistências por tributo")
        cont_trib = inconsistencias.groupby("tributo").size().reset_index(name="quantidade")
        st.bar_chart(cont_trib.set_index("tributo"))

        st.markdown("#### Resumo de inconsistências críticas")
        criticas = inconsistencias[inconsistencias["severidade"] == "Crítica"]
        st.write(f"**{len(criticas)}** inconsistência(s) crítica(s) — exigem base, alíquota e imposto "
                 "simultaneamente, conforme regra tributária aplicável.")
    else:
        st.success("Nenhuma inconsistência crítica detectada com as regras atuais.")


def pagina_blocos():
    st.header("🧱 Visão por Blocos")
    if not st.session_state.arquivo_carregado:
        st.info("Nenhum arquivo carregado.")
        return
    df = st.session_state.registros_df
    blocos = sorted(df["bloco"].unique())
    bloco_sel = st.selectbox("Selecione o bloco", blocos)
    df_bloco = df[df["bloco"] == bloco_sel]
    st.write(f"**{len(df_bloco)}** registros no bloco `{bloco_sel}`.")
    cont = df_bloco.groupby("registro").size().reset_index(name="quantidade").sort_values(
        "quantidade", ascending=False)
    st.dataframe(cont, use_container_width=True, hide_index=True)


def pagina_registros():
    st.header("📋 Visão por Registros")
    if not st.session_state.arquivo_carregado:
        st.info("Nenhum arquivo carregado.")
        return
    tipos_disponiveis = sorted(st.session_state.registros_df["registro"].unique())
    reg_sel = st.selectbox("Tipo de registro", tipos_disponiveis)
    df_det = dataframe_detalhado(st.session_state.registros, reg_sel)
    if df_det.empty:
        st.warning("Sem registros deste tipo.")
        return
    st.caption(f"{len(df_det)} registro(s) — leiaute "
               f"{'reconhecido' if reg_sel in REGISTRO_LAYOUTS else 'genérico (campo_N)'}.")
    st.dataframe(df_det.drop(columns=["uid"]), use_container_width=True, hide_index=True)


def pagina_notas_fiscais():
    st.header("🧾 Visão por Notas Fiscais / Documentos")
    if not st.session_state.arquivo_carregado:
        st.info("Nenhum arquivo carregado.")
        return
    aba1, aba2 = st.tabs(["C100 — Documentos (Mercadorias)", "D100 — Documentos (Transporte/CT-e)"])
    with aba1:
        df_c100 = dataframe_detalhado(st.session_state.registros, "C100")
        if df_c100.empty:
            st.info("Nenhum registro C100 no arquivo.")
        else:
            st.dataframe(df_c100.drop(columns=["uid"]), use_container_width=True, hide_index=True)
    with aba2:
        df_d100 = dataframe_detalhado(st.session_state.registros, "D100")
        if df_d100.empty:
            st.info("Nenhum registro D100 no arquivo.")
        else:
            st.dataframe(df_d100.drop(columns=["uid"]), use_container_width=True, hide_index=True)


def pagina_itens():
    st.header("📦 Visão por Itens (C170)")
    if not st.session_state.arquivo_carregado:
        st.info("Nenhum arquivo carregado.")
        return
    df_itens = dataframe_detalhado(st.session_state.registros, "C170")
    if df_itens.empty:
        st.info("Nenhum registro C170 no arquivo.")
        return
    colunas_disponiveis = [c for c in df_itens.columns if c not in ("uid",)]
    filtro_cfop = st.text_input("Filtrar por CFOP (prefixo)")
    df_filtrado = df_itens
    if filtro_cfop:
        df_filtrado = df_filtrado[df_filtrado["CFOP"].astype(str).str.startswith(filtro_cfop)]
    st.dataframe(df_filtrado[colunas_disponiveis], use_container_width=True, hide_index=True)


def pagina_inconsistencias():
    st.header("🚨 Inconsistências Fiscais")
    if not st.session_state.arquivo_carregado:
        st.info("Nenhum arquivo carregado.")
        return
    tipo = st.session_state.tipo_arquivo
    df_itens = dataframe_detalhado(st.session_state.registros, "C170")
    inconsistencias = detectar_inconsistencias(df_itens, st.session_state.regras_tributarias, tipo)
    if inconsistencias.empty:
        st.success("Nenhuma inconsistência encontrada com as regras vigentes.")
        return

    c1, c2 = st.columns(2)
    tributo_sel = c1.multiselect("Tributo", sorted(inconsistencias["tributo"].unique()),
                                  default=list(inconsistencias["tributo"].unique()))
    severidade_sel = c2.multiselect("Severidade", sorted(inconsistencias["severidade"].unique()),
                                     default=list(inconsistencias["severidade"].unique()))
    filtrado = inconsistencias[
        inconsistencias["tributo"].isin(tributo_sel) & inconsistencias["severidade"].isin(severidade_sel)
    ]
    st.write(f"**{len(filtrado)}** ocorrência(s) encontradas.")
    st.dataframe(filtrado.drop(columns=["campo_base", "campo_aliq", "campo_imp"]),
                 use_container_width=True, hide_index=True)
    st.session_state["_ultima_lista_inconsistencias"] = filtrado


def pagina_correcoes_massa():
    st.header("🛠️ Correções em Massa")
    if not st.session_state.arquivo_carregado:
        st.info("Nenhum arquivo carregado.")
        return
    tipo = st.session_state.tipo_arquivo
    df_itens = dataframe_detalhado(st.session_state.registros, "C170")
    inconsistencias = detectar_inconsistencias(df_itens, st.session_state.regras_tributarias, tipo)
    if inconsistencias.empty:
        st.success("Nenhuma inconsistência pendente de correção.")
        return

    st.markdown("#### 1. Filtre os itens a corrigir")
    c1, c2, c3 = st.columns(3)
    tributo = c1.selectbox("Tributo", sorted(inconsistencias["tributo"].unique()))
    cst_opts = sorted(inconsistencias[inconsistencias["tributo"] == tributo]["cst"].unique())
    cst_sel = c2.multiselect("CST", cst_opts, default=cst_opts)
    cfop_prefixo = c3.text_input("Prefixo do CFOP (opcional)")

    filtrado = inconsistencias[
        (inconsistencias["tributo"] == tributo) & (inconsistencias["cst"].isin(cst_sel))
    ]
    if cfop_prefixo:
        filtrado = filtrado[filtrado["cfop"].astype(str).str.startswith(cfop_prefixo)]

    st.write(f"**{len(filtrado)}** item(ns) selecionado(s) para correção.")
    st.dataframe(filtrado[["uid", "cst", "cfop", "problema", "base_atual", "aliquota_atual",
                            "imposto_atual", "base_sugerida", "aliquota_sugerida",
                            "imposto_sugerido"]], use_container_width=True, hide_index=True)

    st.markdown("#### 2. Defina a correção (prévia antes de gravar)")
    c1, c2 = st.columns(2)
    usar_base_item = c1.checkbox("Preencher base com o valor sugerido (VL_ITEM)", value=True)
    aliquota_manual = c2.text_input("Alíquota a aplicar (%) — em branco usa a sugerida", value="")

    if st.button("Aplicar correção em massa a todos os itens filtrados", type="primary",
                  disabled=filtrado.empty):
        uids = filtrado["uid"].tolist()
        campo_base = filtrado.iloc[0]["campo_base"]
        campo_aliq = filtrado.iloc[0]["campo_aliq"]
        campo_imp = filtrado.iloc[0]["campo_imp"]
        base_valor = None if usar_base_item else "0,00"
        aliquota_valor = aliquota_manual.strip() or None
        # quando base deve usar VL_ITEM, passamos None para que a função use o
        # valor já existente por linha; para forçar o VL_ITEM explicitamente,
        # aplicamos por registro individualmente:
        aplicados = 0
        for _, row in filtrado.iterrows():
            atualizar_campo_registro(row["uid"], campo_base,
                                      row["base_sugerida"] if usar_base_item else row["base_atual"],
                                      "Correção em massa - base", "Correção em massa")
            aliq_final = aliquota_valor or row["aliquota_sugerida"]
            atualizar_campo_registro(row["uid"], campo_aliq, aliq_final,
                                      "Correção em massa - alíquota", "Correção em massa")
            imposto_calc = calcular_imposto(dec(row["base_sugerida"] if usar_base_item else row["base_atual"]),
                                             dec(aliq_final))
            atualizar_campo_registro(row["uid"], campo_imp, dec_to_sped(imposto_calc),
                                      "Correção em massa - imposto recalculado", "Correção em massa")
            aplicados += 1
        st.success(f"Correção aplicada a {aplicados} item(ns). Consulte o Log de Auditoria.")


def pagina_editor_manual():
    st.header("✏️ Editor Manual Avançado")
    if not st.session_state.arquivo_carregado:
        st.info("Nenhum arquivo carregado.")
        return
    tipos_editaveis = sorted([t for t in st.session_state.registros_df["registro"].unique()
                               if t in REGISTRO_LAYOUTS])
    if not tipos_editaveis:
        st.warning("Nenhum registro com leiaute nomeado disponível para edição assistida.")
        return
    reg_sel = st.selectbox("Registro a editar", tipos_editaveis)
    df_det = dataframe_detalhado(st.session_state.registros, reg_sel)
    st.caption("Edite diretamente na grade. Campos alterados serão registrados na auditoria.")

    df_editavel = df_det.drop(columns=["idx", "bloco", "registro", "status", "origem"])
    df_editado = st.data_editor(df_editavel, use_container_width=True, hide_index=True,
                                 key=f"editor_{reg_sel}", disabled=["uid"])

    if st.button("💾 Salvar alterações", type="primary"):
        alterados = 0
        df_original_idx = df_editavel.set_index("uid")
        df_novo_idx = df_editado.set_index("uid")
        for uid in df_novo_idx.index:
            for coluna in df_novo_idx.columns:
                antigo = df_original_idx.loc[uid, coluna]
                novo = df_novo_idx.loc[uid, coluna]
                if str(antigo) != str(novo):
                    if atualizar_campo_registro(uid, coluna, novo, "Edição manual via grade"):
                        alterados += 1
        st.success(f"{alterados} campo(s) atualizado(s).")

    st.divider()
    st.markdown("#### Restaurar registro ao original")
    uid_restaurar = st.selectbox("UID do registro", df_det["uid"].tolist())
    if st.button("↩️ Restaurar valor original deste registro"):
        if desfazer_ultima_alteracao_uid(uid_restaurar):
            st.success("Registro restaurado ao valor original.")
        else:
            st.error("Não foi possível restaurar (registro não encontrado).")


def pagina_regras():
    st.header("⚙️ Motor de Regras Tributárias")
    st.caption("Cadastre, edite e ative/desative regras por CST + prefixo de CFOP + tributo. "
               "A fórmula do motor é sempre: imposto = base × alíquota / 100.")
    df_regras = st.data_editor(
        st.session_state.regras_tributarias, use_container_width=True, hide_index=True,
        num_rows="dynamic", key="editor_regras",
        column_config={
            "tributo": st.column_config.SelectboxColumn(options=["ICMS", "IPI", "PIS", "COFINS"]),
            "tipo_operacao": st.column_config.SelectboxColumn(options=["Entrada", "Saída"]),
        }
    )
    if st.button("💾 Salvar regras"):
        for i, row in df_regras.iterrows():
            if not row.get("regra_id"):
                df_regras.at[i, "regra_id"] = novo_id()
        st.session_state.regras_tributarias = df_regras
        st.success("Regras atualizadas.")


def pagina_importar_cte():
    st.header("🚚 Importar CT-e (XML) para o Bloco D")
    st.caption("Disponível **somente para EFD Contribuições** — gera os registros "
               "D100 (documento), D101 (crédito PIS) e D105 (crédito COFINS) a partir "
               "do XML do CT-e, com valores sujeitos a revisão manual antes da exportação.")

    if not st.session_state.arquivo_carregado:
        st.info("Carregue primeiro um arquivo SPED na aba **Upload do Arquivo**.")
        return

    if st.session_state.tipo_arquivo != TIPO_CONTRIBUICOES:
        st.error(
            f"O arquivo carregado foi identificado como **{st.session_state.tipo_arquivo}**. "
            "A importação de CT-e para o Bloco D está habilitada apenas quando o arquivo "
            f"carregado é **{TIPO_CONTRIBUICOES}**."
        )
        return

    c1, c2, c3 = st.columns(3)
    cst_pis = c1.text_input("CST PIS a aplicar no crédito", value="01")
    cst_cofins = c2.text_input("CST COFINS a aplicar no crédito", value="01")
    cod_cta = c3.text_input("COD_CTA (plano de contas, opcional)", value="")

    st.info("A base de cálculo do PIS/COFINS é sugerida como o **valor da prestação de "
            "serviço (vTPrest)** do CT-e. A alíquota vem da regra cadastrada para o CST/CFOP "
            "informados (ou 1,65% / 7,60% como padrão de referência do regime não-cumulativo, "
            "caso não haja regra específica). **Revise antes de exportar.**")

    arquivos = st.file_uploader("XML(s) de CT-e ou um .zip contendo os XMLs",
                                 type=["xml", "zip"], accept_multiple_files=True)

    if arquivos and st.button("📥 Importar CT-e(s) para o Bloco D", type="primary"):
        with st.spinner("Processando XML(s)..."):
            resultado = importar_ctes(arquivos, cst_pis, cst_cofins, cod_cta)
        if resultado["importados"]:
            st.success(f"{resultado['importados']} CT-e(s) importado(s) com sucesso.")
            st.dataframe(pd.DataFrame(resultado["resumo"]), use_container_width=True, hide_index=True)
        if resultado["erros"]:
            st.warning("Alguns arquivos apresentaram problemas:")
            for e in resultado["erros"]:
                st.write(f"- {e}")

    st.divider()
    st.markdown("#### Registros D100/D101/D105 importados nesta sessão")
    df_d100 = dataframe_detalhado(st.session_state.registros, "D100")
    if not df_d100.empty:
        df_import = df_d100[df_d100["origem"] == "cte_import"]
        st.dataframe(df_import.drop(columns=["uid"]), use_container_width=True, hide_index=True)
    else:
        st.caption("Nenhum registro D100 no arquivo ainda.")


def pagina_exportacao():
    st.header("📦 Exportação")
    if not st.session_state.arquivo_carregado:
        st.info("Nenhum arquivo carregado.")
        return

    tipo = st.session_state.tipo_arquivo
    df_itens = dataframe_detalhado(st.session_state.registros, "C170")
    inconsistencias = detectar_inconsistencias(df_itens, st.session_state.regras_tributarias, tipo)
    df_master = st.session_state.registros_df
    registros_alterados = df_master[df_master["status"] != COLUNA_STATUS_ORIGINAL]
    df_auditoria = pd.DataFrame(st.session_state.audit_log)

    st.markdown("#### 1. Arquivo SPED corrigido (.txt)")
    txt_bytes = exportar_txt_sped(st.session_state.registros)
    st.download_button("⬇️ Baixar SPED corrigido (.txt)", data=txt_bytes,
                        file_name="sped_corrigido.txt", mime="text/plain")

    st.markdown("#### 2. Relatório Excel (multi-abas)")
    excel_bytes = montar_excel_relatorio(inconsistencias, registros_alterados,
                                          st.session_state.regras_tributarias,
                                          df_auditoria, st.session_state.info_empresa)
    st.download_button("⬇️ Baixar relatório Excel", data=excel_bytes,
                        file_name="relatorio_auditoria_sped.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.markdown("#### 3. CSV de inconsistências")
    csv_bytes = inconsistencias.to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ Baixar CSV de inconsistências", data=csv_bytes,
                        file_name="inconsistencias.csv", mime="text/csv")

    st.markdown("#### Resumo antes de exportar")
    c1, c2, c3 = st.columns(3)
    c1.metric("Registros totais", len(df_master))
    c2.metric("Registros alterados/importados", len(registros_alterados))
    c3.metric("Inconsistências remanescentes", len(inconsistencias))


def pagina_auditoria():
    st.header("🕵️ Log de Auditoria")
    if not st.session_state.audit_log:
        st.info("Nenhuma alteração registrada nesta sessão.")
        return
    df_log = pd.DataFrame(st.session_state.audit_log)
    st.dataframe(df_log.sort_values("data_hora", ascending=False), use_container_width=True,
                 hide_index=True)


# ==================================================================================
# 12. BOOTSTRAP / MAIN
# ==================================================================================

PAGINAS = {
    "Dashboard": pagina_dashboard,
    "Upload do Arquivo": pagina_upload,
    "Visão por Blocos": pagina_blocos,
    "Visão por Registros": pagina_registros,
    "Visão por Notas Fiscais": pagina_notas_fiscais,
    "Visão por Itens": pagina_itens,
    "Inconsistências": pagina_inconsistencias,
    "Correções em Massa": pagina_correcoes_massa,
    "Editor Manual": pagina_editor_manual,
    "Motor de Regras": pagina_regras,
    "Importar CT-e (Bloco D)": pagina_importar_cte,
    "Exportação": pagina_exportacao,
    "Log de Auditoria": pagina_auditoria,
}


def main():
    st.set_page_config(page_title=APP_TITLE, page_icon=APP_ICON, layout="wide")
    inicializar_estado()
    aplicar_estilo()

    with st.sidebar:
        st.markdown(f"## {APP_ICON} SPED Studio")
        if st.session_state.arquivo_carregado:
            st.caption(f"Tipo: **{st.session_state.tipo_arquivo}**")
            st.caption(st.session_state.info_empresa.get("razao_social", ""))
        pagina_sel = st.radio("Navegação", list(PAGINAS.keys()), label_visibility="collapsed")

    PAGINAS[pagina_sel]()


if __name__ == "__main__":
    main()


# ==================================================================================
# requirements.txt (sugerido)
# ----------------------------------------------------------------------------------
# streamlit>=1.33
# pandas>=2.0
# numpy>=1.24
# openpyxl>=3.1
#
# Execução local:
#   pip install -r requirements.txt
#   streamlit run app.py
#
# Sugestões de evolução futura:
#   - Suporte a ECD/ECF (arquitetura de REGISTRO_LAYOUTS já preparada para extensão)
#   - Persistência em banco (SQLite/Postgres) em vez de session_state, para
#     sessões colaborativas multiusuário e histórico entre sessões
#   - Autenticação real de usuário (hoje o "usuário" é um mock de sessão)
#   - Recalcularização automática dos totalizadores 9900/9999 na exportação
#   - Suporte a múltiplos CT-e/NF-e simultâneos com deduplicação por chave
#   - Regras tributárias versionadas (histórico de alterações do motor)
#   - Validações cruzadas C100/C170/C190 e D100/D101/D105/D190 completas
#   - Importação de CT-e também para EFD ICMS/IPI (Bloco D sem PIS/COFINS),
#     hoje deliberadamente restrita à EFD Contribuições conforme solicitado
# ==================================================================================
