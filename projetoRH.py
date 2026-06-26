# =============================================================================
# SPED STUDIO v2.0 — ARQUIVO ÚNICO
# Plataforma profissional de auditoria, edição e validação de arquivos SPED
# EFD ICMS/IPI | EFD Contribuições | CT-e XML → Bloco D | C500 | F100
#
# Para executar:
#   pip install streamlit pandas numpy openpyxl lxml pydantic python-dateutil xlsxwriter
#   streamlit run app.py
# =============================================================================

from __future__ import annotations
import re
import io
import os
import sys
import json
import logging
import xml.etree.ElementTree as ET
from copy import deepcopy
from dataclasses import dataclass, field, asdict
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from typing import Optional

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

logging.basicConfig(level=logging.WARNING)
logger = logging.getLogger(__name__)

# =============================================================================
# ░░░░░░░░░░░░░░░░░░  BLOCO 1 — PARSER SPED  ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
# =============================================================================

CAMPOS_REGISTRO: dict[str, list[str]] = {
    "0000": ["REG","COD_VER","TIPO_ESCRIT","IND_SIT_ESP","NUM_REC_ANTERIOR",
             "DT_INI","DT_FIN","NOME","CNPJ","CPF","UF","IE","COD_MUN","IM",
             "SUFRAMA","IND_PERFIL","IND_ATIV"],
    "0001": ["REG","IND_MOV"],
    "0005": ["REG","FANTASIA","CEP","END","NUM","COMPL","BAIRRO","FONE","FAX","EMAIL"],
    "0150": ["REG","COD_PART","NOME","COD_PAIS","CNPJ","CPF","IE","COD_MUN","SUFRAMA","END","NUM","COMPL","BAIRRO"],
    "0190": ["REG","UNID","DESCR"],
    "0200": ["REG","COD_ITEM","DESCR_ITEM","COD_BARRA","COD_ANT_ITEM","UNID_INV",
             "TIPO_ITEM","COD_NCM","EX_IPI","COD_GEN","COD_LST","ALIQ_ICMS"],
    "0990": ["REG","QTD_LIN_0"],
    "C001": ["REG","IND_MOV"],
    "C100": ["REG","IND_OPER","IND_EMIT","COD_PART","COD_MOD","COD_SIT","SER","NUM_DOC",
             "CHV_NFE","DT_DOC","DT_E_S","VL_DOC","IND_PGTO","VL_DESC","VL_ABAT_NT",
             "VL_MERC","IND_FRT","VL_FRT","VL_SEG","VL_OUT_DA","VL_BC_ICMS","VL_ICMS",
             "VL_BC_ICMS_ST","VL_ICMS_ST","VL_IPI","VL_PIS","VL_COFINS","VL_PIS_ST","VL_COFINS_ST"],
    "C101": ["REG","VL_FCP_UF_DEST","VL_ICMS_UF_DEST","VL_ICMS_UF_REM"],
    "C110": ["REG","COD_INF","TXT_COMPL"],
    "C120": ["REG","COD_DOC_IMP","NUM_DOC__IMP","PIS_IMP","COFINS_IMP","VL_DESPESAS_ADU"],
    "C170": ["REG","NUM_ITEM","COD_ITEM","DESCR_COMPL","QTD","UNID","VL_ITEM","VL_DESC",
             "IND_MOV","CST_ICMS","CFOP","COD_NAT","VL_BC_ICMS","ALIQ_ICMS","VL_ICMS",
             "VL_BC_ICMS_ST","ALIQ_ST","VL_ICMS_ST","IND_APUR","CST_IPI","COD_ENQ",
             "VL_BC_IPI","ALIQ_IPI","VL_IPI","CST_PIS","VL_BC_PIS","ALIQ_PIS","QUANT_BC_PIS",
             "ALIQ_PIS_REAIS","VL_PIS","CST_COFINS","VL_BC_COFINS","ALIQ_COFINS",
             "QUANT_BC_COFINS","ALIQ_COFINS_REAIS","VL_COFINS","COD_CTA","VL_ABAT_NT"],
    "C190": ["REG","CST_ICMS","CFOP","ALIQ_ICMS","VL_OPR","VL_BC_ICMS","VL_ICMS",
             "VL_BC_ICMS_ST","VL_ICMS_ST","VL_RED_BC","VL_IPI","COD_OBS"],
    "C195": ["REG","COD_OBS","TXT_COMPL"],
    "C400": ["REG","COD_MOD","ECF_MOD","ECF_FAB","ECF_CX"],
    "C405": ["REG","DT_DOC","CRO","CRZ","NUM_COO_FIN","GT_FIN","VL_BRT"],
    "C425": ["REG","COD_ITEM","QTD","UNID","VL_ITEM","VL_PIS","VL_COFINS"],
    "C460": ["REG","COD_MOD","COD_SIT","NUM_DOC","DT_DOC","VL_DOC","VL_PIS","VL_COFINS"],
    "C470": ["REG","COD_ITEM","QTD","UNID","CST_ICMS","CFOP","VL_BC_ICMS","ALIQ_ICMS",
             "VL_ICMS","VL_BC_IPI","ALIQ_IPI","VL_IPI","CST_PIS","VL_BC_PIS","ALIQ_PIS",
             "VL_PIS","CST_COFINS","VL_BC_COFINS","ALIQ_COFINS","VL_COFINS"],
    "C490": ["REG","CST_ICMS","CFOP","ALIQ_ICMS","VL_OPR","VL_BC_ICMS","VL_ICMS","VL_RED_BC","COD_OBS"],
    # ---- C500 Energia Elétrica ----
    "C500": ["REG","IND_OPER","IND_EMIT","COD_PART","COD_MOD","COD_SIT","SER","SUB","NUM_DOC",
             "DT_DOC","DT_E_S","VL_DOC","VL_DESC","VL_FORN","VL_SERV_NT","VL_TERC","VL_DA",
             "VL_BC_ICMS","VL_ICMS","VL_BC_ICMS_ST","VL_ICMS_ST","COD_INF","VL_PIS","VL_COFINS"],
    "C501": ["REG","CST_PIS","VL_ITEM","NAT_BC_CRED","VL_BC_PIS","ALIQ_PIS","VL_PIS"],
    "C505": ["REG","CST_COFINS","VL_ITEM","NAT_BC_CRED","VL_BC_COFINS","ALIQ_COFINS","VL_COFINS"],
    "C590": ["REG","CST_ICMS","CFOP","ALIQ_ICMS","VL_OPR","VL_BC_ICMS","VL_ICMS",
             "VL_BC_ICMS_ST","VL_ICMS_ST","VL_RED_BC","COD_OBS"],
    "C990": ["REG","QTD_LIN_C"],
    # ---- BLOCO D ----
    "D001": ["REG","IND_MOV"],
    "D100": ["REG","IND_OPER","IND_EMIT","COD_PART","COD_MOD","COD_SIT","SER","SUB","NUM_DOC",
             "CHV_CTE","DT_DOC","DT_A_P","TP_CT_E","CHV_CTE_REF","VL_DOC","VL_DESC",
             "IND_FRT","VL_SERV","VL_BC_ICMS","VL_ICMS","VL_NT","COD_INF","VL_PIS","VL_COFINS"],
    "D101": ["REG","IND_NAT_FRT","VL_ITEM","CST_PIS","NAT_BC_CRED","VL_BC_PIS","ALIQ_PIS","VL_PIS","COD_CTA"],
    "D105": ["REG","IND_NAT_FRT","VL_ITEM","CST_COFINS","NAT_BC_CRED","VL_BC_COFINS","ALIQ_COFINS","VL_COFINS","COD_CTA"],
    "D190": ["REG","CST_ICMS","CFOP","ALIQ_ICMS","VL_OPR","VL_BC_ICMS","VL_ICMS","VL_RED_BC","COD_OBS"],
    "D195": ["REG","COD_OBS","TXT_COMPL"],
    "D500": ["REG","IND_OPER","IND_EMIT","COD_PART","COD_MOD","COD_SIT","SER","SUB","NUM_DOC",
             "DT_DOC","DT_A_P","VL_DOC","VL_DESC","VL_SERV","VL_SERV_NT","VL_TERC","VL_DA",
             "VL_BC_ICMS","VL_ICMS","COD_INF","VL_PIS","VL_COFINS"],
    "D990": ["REG","QTD_LIN_D"],
    # ---- BLOCO E ----
    "E001": ["REG","IND_MOV"],
    "E100": ["REG","DT_INI","DT_FIN"],
    "E110": ["REG","VL_TOT_DEBITOS","VL_AJ_DEBITOS","VL_TOT_AJ_DEBITOS","VL_ESTORNOS_CRED",
             "VL_TOT_CREDITOS","VL_AJ_CREDITOS","VL_TOT_AJ_CREDITOS","VL_ESTORNOS_DEB",
             "VL_SLD_CREDOR_ANT","VL_SLD_APURADO","VL_TOT_DED","VL_ICMS_RECOLHER",
             "VL_SLD_CREDOR_TRANSPORTAR","DEB_ESP"],
    "E116": ["REG","COD_OR","VL_OR","DT_VCTO","COD_REC","NUM_PROC","IND_PROC","PROC","TXT_COMPL","MES_REF"],
    "E200": ["REG","UF","DT_INI","DT_FIN"],
    "E210": ["REG","IND_MOV_ST","VL_SLD_CRED_ANT_ST","VL_DEBITOS_ST","VL_OUTROS_DEB_ST",
             "VL_AJ_DEBITOS_ST","VL_ESTORNO_CRED_ST","VL_CREDITOS_ST","VL_OUTROS_CRED_ST",
             "VL_AJ_CREDITOS_ST","VL_ESTORNO_DEB_ST","VL_SLD_DEVEDOR_ST","VL_DEDUCOES_ST",
             "VL_ICMS_RECOL_ST","VL_ICMS_RECOL_EX_ST","VL_SLD_CRED_ST_TRANSPORTAR","DEB_ESP_ST"],
    "E990": ["REG","QTD_LIN_E"],
    # ---- BLOCO F (EFD Contribuições) ----
    "F001": ["REG","IND_MOV"],
    "F100": ["REG","IND_OPER","COD_PART","COD_ITEM","DT_OPER","VL_OPER","CST_PIS","VL_BC_PIS",
             "ALIQ_PIS","VL_PIS","CST_COFINS","VL_BC_COFINS","ALIQ_COFINS","VL_COFINS",
             "NAT_BC_CRED","IND_ORIG_CRED","COD_CTA","COD_CCUS","DESC_DOC_OPR"],
    "F111": ["REG","NUM_PROC","IND_PROC"],
    "F120": ["REG","NAT_BC_CRED","IDENT_BEM_IMOB","IND_ORIG_CRED","IND_UTIL_BEM_IMOB",
             "VL_OPER_DEP","PARC_OPER_NAO_BC_CRED","CST_PIS","VL_BC_PIS","ALIQ_PIS","VL_PIS",
             "CST_COFINS","VL_BC_COFINS","ALIQ_COFINS","VL_COFINS","COD_CTA","COD_CCUS"],
    "F200": ["REG","IND_OPER","COD_PART","COD_ITEM","DT_OPER","VL_OPER","CST_PIS","VL_BC_PIS",
             "ALIQ_PIS","VL_PIS","CST_COFINS","VL_BC_COFINS","ALIQ_COFINS","VL_COFINS",
             "NAT_BC_CRED","IND_ORIG_CRED","COD_CTA","VL_TOT_PARC","VL_PARC_PASS",
             "VL_BC_CRED_PARC","CST_PIS_PARC","VL_BC_PIS_PARC","ALIQ_PIS_PARC","VL_PIS_PARC",
             "CST_COFINS_PARC","VL_BC_COFINS_PARC","ALIQ_COFINS_PARC","VL_COFINS_PARC",
             "DT_AQUIS_IMOB","VL_TOT_IMOB","VL_BC_IMOB","VL_BC_IMOB_AV","IND_CONT_IMOB","IDENT_IMOB"],
    "F500": ["REG","VL_REC_CAIXA","CST_PIS","VL_DESC_PIS","VL_BC_PIS","ALIQ_PIS","VL_PIS",
             "CST_COFINS","VL_DESC_COFINS","VL_BC_COFINS","ALIQ_COFINS","VL_COFINS","COD_MOD","COD_CTA","INFO_COMPL"],
    "F600": ["REG","IND_NAT_RET","DT_RET","VL_BC_RET","VL_RET","COD_REC","IND_NAT_REC",
             "CNPJ","VL_RET_PIS","VL_RET_COFINS","IND_DEC"],
    "F990": ["REG","QTD_LIN_F"],
    # ---- BLOCO H ----
    "H001": ["REG","IND_MOV"],
    "H005": ["REG","DT_INV","VL_INV","MOT_INV"],
    "H010": ["REG","COD_ITEM","UNID","QTD","VL_UNIT","VL_ITEM","IND_PROP","COD_PART",
             "TXT_COMPL","COD_CTA","VL_ITEM_IR"],
    "H020": ["REG","CST_ICMS","BC_ICMS","VL_ICMS"],
    "H990": ["REG","QTD_LIN_H"],
    # ---- BLOCO 1 ----
    "1001": ["REG","IND_MOV"],
    "1010": ["REG","IND_EXP","IND_CCRF","IND_COMB","IND_USINA","IND_VA","IND_EE","IND_CART","IND_FORM","IND_AER"],
    "1200": ["REG","COD_AJ_APUR","SLD_INICIAL","CREDITO_APR","CREDITO_RECEB","DEBITO_APR",
             "OUTROS_DEB","SLD_CREDOR_ANT","OUTROS_CRED","SLD_APURADO","DEDUCOES","SLD_REPOR"],
    "1600": ["REG","COD_PART","TOT_CREDITOS","TOT_DEBITOS"],
    "1990": ["REG","QTD_LIN_1"],
    # ---- BLOCO M (EFD Contribuições) ----
    "M001": ["REG","IND_MOV"],
    "M100": ["REG","COD_CRED","IND_CRED_ORI","VL_BC_PIS","ALIQ_PIS","QUANT_BC_PIS","ALIQ_PIS_REAIS",
             "VL_CRED","VL_CRED_EXC","VL_CRED_DIS","VL_CRED_EFE","VL_CRED_PER","VL_CRED_DEV",
             "VL_OUT_DEV","VL_CRED_AUT","IND_ORI_CRED","VL_CRED_DESC_PA_ANT","COD_REC","VL_CRED_DESC","DT_RECOL"],
    "M200": ["REG","VL_TOT_CONT_NC_PER","VL_TOT_CRED_DESC","VL_TOT_CRED_DESC_ANT","VL_TOT_CRED_ACUM",
             "VL_TOT_NRE","VL_TOT_DESC_PAG","VL_TOT_CONT_NC_DEV","VL_CONT_NC_REC"],
    "M210": ["REG","COD_CONT","VL_REC_BRT","VL_BC_CONT","VL_AJUS_ACRES_BC_PIS","VL_AJUS_REDUC_BC_PIS",
             "VL_BC_CONT_AJUS","ALIQ_PIS","QUANT_BC_PIS","ALIQ_PIS_REAIS","VL_CONT_APUR",
             "VL_AJUS_ACRES","VL_AJUS_REDUC","VL_CONT_DIFER","VL_CONT_DIFER_ANT","VL_CONT_PER"],
    "M400": ["REG","CST_PIS","VL_TOT_REC","COD_CTA","DESC_COMPL"],
    "M500": ["REG","COD_CRED","IND_CRED_ORI","VL_BC_COFINS","ALIQ_COFINS","QUANT_BC_COFINS",
             "ALIQ_COFINS_REAIS","VL_CRED","VL_CRED_EXC","VL_CRED_DIS","VL_CRED_EFE",
             "VL_CRED_PER","VL_CRED_DEV","VL_OUT_DEV","VL_CRED_AUT","IND_ORI_CRED",
             "VL_CRED_DESC_PA_ANT","COD_REC","VL_CRED_DESC","DT_RECOL"],
    "M600": ["REG","VL_TOT_CONT_CUM_PER","VL_TOT_CONT_CUM_DES","VL_TOT_CONT_CUM_ANT",
             "VL_TOT_CONT_CUM","VL_RET_NC","VL_OUT_DED","VL_CONT_SOL_PAG","VL_CONT_APUR_DEV",
             "VL_OUT_ACRES","VL_CONT_PER"],
    "M610": ["REG","COD_CONT","VL_REC_BRT","VL_BC_CONT","VL_AJUS_ACRES_BC_COFINS",
             "VL_AJUS_REDUC_BC_COFINS","VL_BC_CONT_AJUS","ALIQ_COFINS","QUANT_BC_COFINS",
             "ALIQ_COFINS_REAIS","VL_CONT_APUR","VL_AJUS_ACRES","VL_AJUS_REDUC","VL_CONT_DIFER",
             "VL_CONT_DIFER_ANT","VL_CONT_PER"],
    "M800": ["REG","CST_COFINS","VL_TOT_REC","COD_CTA","DESC_COMPL"],
    "M990": ["REG","QTD_LIN_M"],
    # ---- BLOCO P ----
    "P001": ["REG","IND_MOV"],
    "P100": ["REG","DT_INI","DT_FIN","VL_REC_TOT_EST","COD_ATIV_ECON","VL_REC_ATIV_ESTAB",
             "VL_EXC","VL_BC_CONT","ALIQ_CONT","VL_CONT_APU","COD_CTA","INFO_COMPL"],
    "P200": ["REG","PER_REF","VL_TOT_CONT_APU","VL_RED_BC","VL_BC_CONT_RED","VL_RETENCOES",
             "VL_OUT_DED","VL_CONT_SOL_PAG","VL_CONT_APUR_DEV","VL_OUT_ACRES","VL_CONT_PER","DT_RECOL"],
    "P990": ["REG","QTD_LIN_P"],
    # ---- BLOCO 9 ----
    "9001": ["REG","IND_MOV"],
    "9900": ["REG","REG_BLC","QTD_REG_BLC"],
    "9990": ["REG","QTD_LIN_9"],
    "9999": ["REG","QTD_LIN"],
    # ---- EFD Contribuições Bloco A ----
    "A001": ["REG","IND_MOV"],
    "A010": ["REG","CNPJ"],
    "A100": ["REG","IND_OPER","IND_EMIT","COD_PART","COD_SIT","SER","SUB","NUM_DOC","CHV_NFSE",
             "DT_DOC","DT_EXE_SERV","VL_DOC","VL_DESC","VL_BC_PIS","ALIQ_PIS","VL_PIS",
             "VL_BC_COFINS","ALIQ_COFINS","VL_COFINS","COD_MUN","COD_SERV","VL_DES_PERD"],
    "A170": ["REG","NUM_ITEM","COD_ITEM","DESCR_COMPL","VL_ITEM","VL_DESC","NAT_BC_CRED",
             "IND_ORIG_CRED","CST_PIS","VL_BC_PIS","ALIQ_PIS","VL_PIS","CST_COFINS",
             "VL_BC_COFINS","ALIQ_COFINS","VL_COFINS","COD_CTA"],
    "A990": ["REG","QTD_LIN_A"],
}


@dataclass
class LinhaRaw:
    numero: int
    conteudo: str
    registro: str
    campos: list
    bloco: str


@dataclass
class ResultadoParsing:
    tipo_arquivo: str
    cnpj: str
    nome_empresa: str
    ie: str
    uf: str
    dt_ini: str
    dt_fin: str
    cod_ver: str
    linhas_raw: list
    df_registros: pd.DataFrame
    blocos_encontrados: list
    erros_parsing: list = field(default_factory=list)


def _normalizar_decimal(valor: str) -> str:
    if not valor:
        return valor
    return valor.strip().replace(",", ".")


def _parse_linha(numero: int, linha: str) -> LinhaRaw:
    linha = linha.rstrip("\r\n")
    if linha.startswith("|"):
        linha = linha[1:]
    if linha.endswith("|"):
        linha = linha[:-1]
    campos = linha.split("|")
    registro = campos[0].strip() if campos else ""
    bloco = registro[0] if registro else ""
    return LinhaRaw(numero=numero, conteudo=linha, registro=registro,
                    campos=campos, bloco=bloco)


def _identificar_tipo(df: pd.DataFrame) -> str:
    rows_0000 = df[df["REG"] == "0000"]
    if rows_0000.empty:
        return "DESCONHECIDO"
    blocos = df["BLOCO"].unique().tolist()
    if "M" in blocos or "F" in blocos or "A" in blocos:
        return "EFD_CONTRIBUICOES"
    return "EFD_ICMS_IPI"


def _linha_para_dict(linha: LinhaRaw) -> dict:
    reg = linha.registro
    nomes = CAMPOS_REGISTRO.get(reg, [])
    resultado: dict = {"LINHA_NUM": linha.numero, "REG": reg, "BLOCO": linha.bloco}
    for i, campo in enumerate(linha.campos):
        chave = nomes[i] if i < len(nomes) else f"CAMPO_{i:03d}"
        resultado[chave] = _normalizar_decimal(campo)
    return resultado


def parse_arquivo_sped(conteudo) -> ResultadoParsing:
    erros: list = []
    if isinstance(conteudo, bytes):
        for enc in ("utf-8", "latin-1", "cp1252"):
            try:
                texto = conteudo.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            texto = conteudo.decode("latin-1", errors="replace")
            erros.append("Encoding não identificado; usado latin-1.")
    else:
        texto = conteudo

    linhas_raw: list = []
    registros_dicts: list = []
    blocos_set: set = set()

    for num, linha in enumerate(texto.splitlines(), start=1):
        linha = linha.strip()
        if not linha:
            continue
        lr = _parse_linha(num, linha)
        linhas_raw.append(lr)
        blocos_set.add(lr.bloco)
        registros_dicts.append(_linha_para_dict(lr))

    df = pd.DataFrame(registros_dicts)
    if df.empty:
        return ResultadoParsing(
            tipo_arquivo="DESCONHECIDO", cnpj="", nome_empresa="", ie="", uf="",
            dt_ini="", dt_fin="", cod_ver="", linhas_raw=[], df_registros=df,
            blocos_encontrados=[], erros_parsing=["Arquivo vazio."]
        )

    tipo = _identificar_tipo(df)
    row0 = df[df["REG"] == "0000"].iloc[0] if not df[df["REG"] == "0000"].empty else pd.Series()
    g = lambda k: str(row0.get(k, "")).strip() if isinstance(row0, pd.Series) and k in row0 else ""

    return ResultadoParsing(
        tipo_arquivo=tipo, cnpj=g("CNPJ"), nome_empresa=g("NOME"),
        ie=g("IE"), uf=g("UF"), dt_ini=g("DT_INI"), dt_fin=g("DT_FIN"),
        cod_ver=g("COD_VER"), linhas_raw=linhas_raw, df_registros=df,
        blocos_encontrados=sorted(list(blocos_set)), erros_parsing=erros,
    )


def reconstruir_txt(linhas_raw: list, df_editado: pd.DataFrame) -> str:
    edicoes: dict = {}
    for _, row in df_editado.iterrows():
        lnum = int(row.get("LINHA_NUM", -1))
        if lnum > 0:
            edicoes[lnum] = row.to_dict()

    linhas_saida: list = []
    for lr in linhas_raw:
        if lr.numero in edicoes:
            dados = edicoes[lr.numero]
            reg = lr.registro
            nomes = CAMPOS_REGISTRO.get(reg, [])
            novos_campos = []
            for i in range(len(lr.campos)):
                chave = nomes[i] if i < len(nomes) else f"CAMPO_{i:03d}"
                if chave in dados and chave not in ("LINHA_NUM", "BLOCO"):
                    novos_campos.append(str(dados[chave]) if dados[chave] is not None else "")
                else:
                    novos_campos.append(lr.campos[i])
            linhas_saida.append("|" + "|".join(novos_campos) + "|")
        else:
            linhas_saida.append("|" + "|".join(lr.campos) + "|")
    return "\n".join(linhas_saida) + "\n"


# =============================================================================
# ░░░░░░░░░░░░░░░░░░  BLOCO 2 — PARSER CT-e XML  ░░░░░░░░░░░░░░░░░░░░░░░░░░░
# =============================================================================

def _get_text_xml(element, path: str, ns: dict, default: str = "") -> str:
    try:
        found = element.find(path, ns)
        return found.text.strip() if found is not None and found.text else default
    except Exception:
        return default


def _get_decimal_xml(element, path: str, ns: dict, default: str = "0") -> str:
    val = _get_text_xml(element, path, ns, default)
    try:
        return str(Decimal(val.replace(",", ".")))
    except InvalidOperation:
        return default


def _formatar_data_xml(data_xml: str) -> str:
    if not data_xml:
        return ""
    data_xml = data_xml[:10]
    partes = data_xml.split("-")
    if len(partes) == 3:
        return partes[2] + partes[1] + partes[0]
    return data_xml


def parse_xml_cte(conteudo_xml) -> dict:
    erros: list = []
    try:
        if isinstance(conteudo_xml, str):
            conteudo_xml = conteudo_xml.encode("utf-8")
        root = ET.fromstring(conteudo_xml)
    except ET.ParseError as e:
        return {"d100": {}, "d101": None, "d105": None, "chave": "", "erros": [f"XML inválido: {e}"]}

    tag_root = root.tag
    ns = {}
    if "portalfiscal" in tag_root:
        uri = tag_root.split("}")[0].replace("{", "")
        ns = {"cte": uri}
    else:
        ns = {"cte": "http://www.portalfiscal.inf.br/cte"}

    infCte = root.find(".//cte:infCte", ns)
    if infCte is None:
        infCte = root.find(".//infCte")
        if infCte is None:
            return {"d100": {}, "d101": None, "d105": None, "chave": "", "erros": ["infCte não encontrado"]}
        ns = {}

    chave = infCte.get("Id", "").replace("CTe", "")

    ide = infCte.find("cte:ide", ns) or infCte.find("ide") or ET.Element("ide")
    cod_mod = _get_text_xml(ide, "cte:mod", ns) or _get_text_xml(ide, "mod", {})
    ser     = _get_text_xml(ide, "cte:serie", ns) or _get_text_xml(ide, "serie", {})
    num_doc = _get_text_xml(ide, "cte:nCT", ns) or _get_text_xml(ide, "nCT", {})
    dt_doc  = _formatar_data_xml(_get_text_xml(ide, "cte:dhEmi", ns) or _get_text_xml(ide, "dhEmi", {}))
    cfop    = _get_text_xml(ide, "cte:CFOP", ns) or _get_text_xml(ide, "CFOP", {})
    ind_oper = "1" if cfop.startswith(("1","2","3")) else "0"

    emit = infCte.find("cte:emit", ns) or infCte.find("emit")
    cnpj_emit = nome_emit = ""
    if emit is not None:
        cnpj_emit = _get_text_xml(emit, "cte:CNPJ", ns) or _get_text_xml(emit, "CNPJ", {})
        nome_emit = _get_text_xml(emit, "cte:xNome", ns) or _get_text_xml(emit, "xNome", {})

    vPrest = infCte.find("cte:vPrest", ns) or infCte.find("vPrest")
    vl_serv = vl_doc = "0"
    if vPrest is not None:
        vl_serv = _get_decimal_xml(vPrest, "cte:vTPrest", ns) or _get_decimal_xml(vPrest, "vTPrest", {})
        vl_rec  = _get_decimal_xml(vPrest, "cte:vRec", ns) or _get_decimal_xml(vPrest, "vRec", {})
        try:
            vl_doc = str(Decimal(vl_rec) if vl_rec != "0" else Decimal(vl_serv))
        except Exception:
            vl_doc = vl_serv

    imposto = infCte.find("cte:imp", ns) or infCte.find("imp")
    vl_bc_icms = vl_icms = "0"
    vl_pis = vl_bc_pis = aliq_pis = "0"; cst_pis = "99"
    vl_cofins = vl_bc_cofins = aliq_cofins = "0"; cst_cofins = "99"

    if imposto is not None:
        icms = imposto.find("cte:ICMS", ns) or imposto.find("ICMS")
        if icms is not None:
            for filho in icms:
                bc_c  = filho.find("cte:vBC", ns) or filho.find("vBC")
                vl_c  = filho.find("cte:vICMS", ns) or filho.find("vICMS")
                if bc_c is not None and bc_c.text: vl_bc_icms = bc_c.text.strip()
                if vl_c is not None and vl_c.text:  vl_icms    = vl_c.text.strip()
                break

        def _extrai_contrib(elem_name, cst_default):
            elem = imposto.find(f"cte:{elem_name}", ns) or imposto.find(elem_name)
            bc, aliq, vl, cst = "0","0","0", cst_default
            if elem is not None:
                for filho in elem:
                    bc_c  = filho.find("cte:vBC", ns) or filho.find("vBC")
                    al_c  = filho.find("cte:pPIS", ns) or filho.find("pPIS") or \
                            filho.find("cte:pCOFINS", ns) or filho.find("pCOFINS")
                    vl_c  = filho.find("cte:vPIS", ns) or filho.find("vPIS") or \
                            filho.find("cte:vCOFINS", ns) or filho.find("vCOFINS")
                    cs_c  = filho.find("cte:CST", ns) or filho.find("CST")
                    if bc_c is not None and bc_c.text:  bc   = bc_c.text.strip()
                    if al_c is not None and al_c.text:  aliq = al_c.text.strip()
                    if vl_c is not None and vl_c.text:  vl   = vl_c.text.strip()
                    if cs_c is not None and cs_c.text:  cst  = cs_c.text.strip()
                    break
            return bc, aliq, vl, cst

        vl_bc_pis, aliq_pis, vl_pis, cst_pis         = _extrai_contrib("PIS", "99")
        vl_bc_cofins, aliq_cofins, vl_cofins, cst_cofins = _extrai_contrib("COFINS", "99")

    d100 = {
        "REG":"D100","IND_OPER":ind_oper,"IND_EMIT":"0","COD_PART":cnpj_emit,
        "COD_MOD":cod_mod or "57","COD_SIT":"00","SER":ser,"SUB":"","NUM_DOC":num_doc,
        "CHV_CTE":chave,"DT_DOC":dt_doc,"DT_A_P":dt_doc,"TP_CT_E":"","CHV_CTE_REF":"",
        "VL_DOC":vl_doc,"VL_DESC":"0","IND_FRT":"0","VL_SERV":vl_serv,
        "VL_BC_ICMS":vl_bc_icms,"VL_ICMS":vl_icms,"VL_NT":"","COD_INF":"",
        "VL_PIS":vl_pis,"VL_COFINS":vl_cofins,
        "CNPJ_EMIT":cnpj_emit,"NOME_EMIT":nome_emit,"_CFOP":cfop,
    }

    nao_credito = {"06","07","08","09","49","50","51","52","53","54","55","56","98","99"}
    d101 = None if cst_pis in nao_credito else {
        "REG":"D101","IND_NAT_FRT":"07","VL_ITEM":vl_serv,"CST_PIS":cst_pis,
        "NAT_BC_CRED":"17","VL_BC_PIS":vl_bc_pis,"ALIQ_PIS":aliq_pis,"VL_PIS":vl_pis,"COD_CTA":"",
    }
    d105 = None if cst_cofins in nao_credito else {
        "REG":"D105","IND_NAT_FRT":"07","VL_ITEM":vl_serv,"CST_COFINS":cst_cofins,
        "NAT_BC_CRED":"17","VL_BC_COFINS":vl_bc_cofins,"ALIQ_COFINS":aliq_cofins,"VL_COFINS":vl_cofins,"COD_CTA":"",
    }
    return {"d100":d100,"d101":d101,"d105":d105,"chave":chave,"erros":erros}


# =============================================================================
# ░░░░░░░░░░░░░░░░░░  BLOCO 3 — MOTOR DE REGRAS  ░░░░░░░░░░░░░░░░░░░░░░░░░░░
# =============================================================================

@dataclass
class RegraCST:
    tributo: str
    cst: str
    descricao: str
    exige_base: bool
    exige_aliquota: bool
    exige_imposto: bool
    formula: str
    observacao: str = ""


def _to_decimal(valor, default: Decimal = Decimal("0")) -> Decimal:
    if valor is None or str(valor).strip() in ("", "None", "nan"):
        return default
    try:
        return Decimal(str(valor).replace(",", ".").strip())
    except InvalidOperation:
        return default


def calcular_imposto(regra: RegraCST, bc: str, aliq: str,
                     qtd: str = "0", aliq_reais: str = "0") -> Decimal:
    bc_d   = _to_decimal(bc)
    aliq_d = _to_decimal(aliq)
    qtd_d  = _to_decimal(qtd)
    ar_d   = _to_decimal(aliq_reais)
    formula = (regra.formula or "0").lower().strip()
    try:
        if formula == "0":
            return Decimal("0")
        elif formula == "bc * aliq / 100":
            return (bc_d * aliq_d / Decimal("100")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        elif formula == "qtd * aliq_reais":
            return (qtd_d * ar_d).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        else:
            env = {"bc":bc_d,"aliq":aliq_d,"qtd":qtd_d,"aliq_reais":ar_d,"Decimal":Decimal}
            return Decimal(str(eval(formula, {"__builtins__":{}}, env))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except Exception:
        return Decimal("0")


def buscar_regra(regras: list, tributo: str, cst: str) -> Optional[RegraCST]:
    cst_norm = str(cst).strip().zfill(2)
    for r in regras:
        if r.tributo.upper() == tributo.upper() and r.cst == cst_norm:
            return r
    for r in regras:
        if r.tributo.upper() == tributo.upper() and r.cst.strip() == str(cst).strip():
            return r
    return None


def regras_para_df(regras: list) -> pd.DataFrame:
    return pd.DataFrame([asdict(r) for r in regras])


def df_para_regras(df: pd.DataFrame) -> list:
    out = []
    for _, row in df.iterrows():
        try:
            out.append(RegraCST(
                tributo=str(row.get("tributo","")).strip(),
                cst=str(row.get("cst","")).strip(),
                descricao=str(row.get("descricao","")).strip(),
                exige_base=bool(row.get("exige_base",False)),
                exige_aliquota=bool(row.get("exige_aliquota",False)),
                exige_imposto=bool(row.get("exige_imposto",False)),
                formula=str(row.get("formula","0")).strip(),
                observacao=str(row.get("observacao","")).strip(),
            ))
        except Exception:
            pass
    return out


# Tabela de regras padrão ICMS
_REGRAS_ICMS = [
    ("00","Tributada integralmente",True,True,True,"bc * aliq / 100"),
    ("10","Tributada com ST",True,True,True,"bc * aliq / 100"),
    ("20","Redução de BC",True,True,True,"bc * aliq / 100"),
    ("30","Isenta/NT com ST",False,False,False,"0"),
    ("40","Isenta",False,False,False,"0"),
    ("41","Não tributada",False,False,False,"0"),
    ("50","Suspensão",False,False,False,"0"),
    ("51","Diferimento",True,True,True,"bc * aliq / 100"),
    ("60","Cobrado por ST anteriormente",False,False,False,"0"),
    ("70","Redução BC com ST",True,True,True,"bc * aliq / 100"),
    ("90","Outras",False,False,False,"0"),
    ("101","SN com crédito",False,False,False,"0"),
    ("102","SN sem crédito",False,False,False,"0"),
    ("103","SN isenção faixa RB",False,False,False,"0"),
    ("201","SN crédito com ST",False,False,True,"0"),
    ("202","SN sem crédito com ST",False,False,True,"0"),
    ("300","Imune",False,False,False,"0"),
    ("400","NT pelo SN",False,False,False,"0"),
    ("500","ICMS cobrado por ST/antecipação SN",False,False,False,"0"),
    ("900","Outros SN",False,False,False,"0"),
]

# Tabela de regras padrão PIS/COFINS
_REGRAS_PC = [
    ("01","Trib. base = valor operação alíquota normal",True,True,True,"bc * aliq / 100"),
    ("02","Trib. base = valor operação alíquota diferenciada",True,True,True,"bc * aliq / 100"),
    ("03","Trib. base = quantidade × alíquota unidade",True,True,True,"qtd * aliq_reais"),
    ("04","Monofásica – revenda alíquota zero",False,False,False,"0"),
    ("05","Substituição tributária",False,False,False,"0"),
    ("06","Alíquota zero",False,False,False,"0"),
    ("07","Isenta",False,False,False,"0"),
    ("08","Sem incidência",False,False,False,"0"),
    ("09","Suspensão",False,False,False,"0"),
    ("49","Outras saídas",False,False,False,"0"),
    ("50","Crédito – receita trib. mercado interno",True,True,True,"bc * aliq / 100"),
    ("51","Crédito – receita não-trib. mercado interno",True,True,True,"bc * aliq / 100"),
    ("52","Crédito – receita exportação",True,True,True,"bc * aliq / 100"),
    ("53","Crédito – receitas trib. e não-trib. MI",True,True,True,"bc * aliq / 100"),
    ("54","Crédito – trib. MI e exportação",True,True,True,"bc * aliq / 100"),
    ("55","Crédito – não-trib. MI e exportação",True,True,True,"bc * aliq / 100"),
    ("56","Crédito – trib. não-trib. MI e exportação",True,True,True,"bc * aliq / 100"),
    ("60","Crédito presumido – trib. MI",True,True,True,"bc * aliq / 100"),
    ("61","Crédito presumido – não-trib. MI",True,True,True,"bc * aliq / 100"),
    ("62","Crédito presumido – exportação",True,True,True,"bc * aliq / 100"),
    ("63","Crédito presumido – trib. e não-trib. MI",True,True,True,"bc * aliq / 100"),
    ("64","Crédito presumido – trib. MI e exportação",True,True,True,"bc * aliq / 100"),
    ("65","Crédito presumido – não-trib. MI e exportação",True,True,True,"bc * aliq / 100"),
    ("66","Crédito presumido – todos",True,True,True,"bc * aliq / 100"),
    ("67","Crédito presumido – outras operações",True,True,True,"bc * aliq / 100"),
    ("70","Aquisição sem crédito",False,False,False,"0"),
    ("71","Aquisição isenta",False,False,False,"0"),
    ("72","Aquisição suspensão",False,False,False,"0"),
    ("73","Aquisição alíquota zero",False,False,False,"0"),
    ("74","Aquisição sem incidência",False,False,False,"0"),
    ("75","Aquisição substituição tributária",False,False,False,"0"),
    ("98","Outras entradas",False,False,False,"0"),
    ("99","Outras operações",False,False,False,"0"),
]

TODAS_REGRAS_PADRAO: list = (
    [RegraCST("ICMS", cst, desc, eb, ea, ei, f) for cst,desc,eb,ea,ei,f in _REGRAS_ICMS] +
    [RegraCST("PIS",  cst, desc, eb, ea, ei, f) for cst,desc,eb,ea,ei,f in _REGRAS_PC] +
    [RegraCST("COFINS",cst,desc.replace("PIS","COFINS"),eb,ea,ei,f) for cst,desc,eb,ea,ei,f in _REGRAS_PC]
)


def avaliar_linha_sped(row: pd.Series, regras: list) -> list:
    incs = []
    reg = str(row.get("REG","")).strip()
    ln  = int(row.get("LINHA_NUM", 0))
    bl  = str(row.get("BLOCO","")).strip()

    def _add(tipo, trib, campo, desc, atual, sug, regra_nome):
        incs.append({
            "LINHA_NUM":ln,"REGISTRO":reg,"BLOCO":bl,"TIPO":tipo,
            "TRIBUTO":trib,"CAMPO":campo,"DESCRICAO":desc,
            "VALOR_ATUAL":str(atual),"VALOR_SUGERIDO":str(sug),
            "REGRA":regra_nome,
        })

    if reg == "C170":
        # ICMS
        cst = str(row.get("CST_ICMS","")).strip()
        vl_item = _to_decimal(row.get("VL_ITEM"))
        bc   = _to_decimal(row.get("VL_BC_ICMS"))
        aliq = _to_decimal(row.get("ALIQ_ICMS"))
        vl   = _to_decimal(row.get("VL_ICMS"))
        r = buscar_regra(regras,"ICMS",cst)
        if r:
            if r.exige_base and bc == 0:
                _add("CRITICO","ICMS","VL_BC_ICMS",f"CST {cst} exige base de cálculo",bc,vl_item,f"CST-ICMS-{cst}")
            if r.exige_aliquota and aliq == 0:
                _add("CRITICO","ICMS","ALIQ_ICMS",f"CST {cst} exige alíquota",aliq,"",f"CST-ICMS-{cst}")
            if r.exige_imposto and vl == 0:
                sug = calcular_imposto(r,str(bc),str(aliq))
                _add("CRITICO","ICMS","VL_ICMS",f"CST {cst} exige ICMS",vl,sug,f"CST-ICMS-{cst}")
            if not r.exige_base and bc > 0:
                _add("AVISO","ICMS","VL_BC_ICMS",f"CST {cst} não deveria ter base",bc,"0",f"CST-ICMS-{cst}")
            if not r.exige_imposto and vl > 0:
                _add("AVISO","ICMS","VL_ICMS",f"CST {cst} não deveria ter ICMS",vl,"0",f"CST-ICMS-{cst}")
        # PIS
        cst_p = str(row.get("CST_PIS","")).strip()
        bc_p  = _to_decimal(row.get("VL_BC_PIS"))
        al_p  = _to_decimal(row.get("ALIQ_PIS"))
        vl_p  = _to_decimal(row.get("VL_PIS"))
        rp = buscar_regra(regras,"PIS",cst_p)
        if rp and rp.exige_imposto and vl_p == 0 and vl_item > 0:
            sug = calcular_imposto(rp,str(bc_p or vl_item),str(al_p))
            _add("CRITICO","PIS","VL_PIS",f"CST PIS {cst_p} exige PIS",vl_p,sug,f"CST-PIS-{cst_p}")
        # COFINS
        cst_c = str(row.get("CST_COFINS","")).strip()
        bc_c  = _to_decimal(row.get("VL_BC_COFINS"))
        al_c  = _to_decimal(row.get("ALIQ_COFINS"))
        vl_c  = _to_decimal(row.get("VL_COFINS"))
        rc = buscar_regra(regras,"COFINS",cst_c)
        if rc and rc.exige_imposto and vl_c == 0 and vl_item > 0:
            sug = calcular_imposto(rc,str(bc_c or vl_item),str(al_c))
            _add("CRITICO","COFINS","VL_COFINS",f"CST COFINS {cst_c} exige COFINS",vl_c,sug,f"CST-COFINS-{cst_c}")

    elif reg == "C190":
        cst = str(row.get("CST_ICMS","")).strip()
        aliq = _to_decimal(row.get("ALIQ_ICMS"))
        vl_opr = _to_decimal(row.get("VL_OPR"))
        bc   = _to_decimal(row.get("VL_BC_ICMS"))
        vl   = _to_decimal(row.get("VL_ICMS"))
        r = buscar_regra(regras,"ICMS",cst)
        if r and r.exige_base and bc == 0 and vl_opr > 0:
            _add("CRITICO","ICMS","VL_BC_ICMS",f"C190 CST {cst}: base vazia",bc,vl_opr,f"CST-ICMS-{cst}")
        if r and r.exige_imposto and vl == 0 and bc > 0:
            sug = calcular_imposto(r,str(bc),str(aliq))
            _add("CRITICO","ICMS","VL_ICMS",f"C190 CST {cst}: ICMS vazio",vl,sug,f"CST-ICMS-{cst}")

    elif reg in ("D101","D105"):
        tributo = "PIS" if reg == "D101" else "COFINS"
        cst  = str(row.get(f"CST_{tributo}","")).strip()
        bc   = _to_decimal(row.get(f"VL_BC_{tributo}"))
        aliq = _to_decimal(row.get(f"ALIQ_{tributo}"))
        vl   = _to_decimal(row.get(f"VL_{tributo}"))
        vl_s = _to_decimal(row.get("VL_ITEM"))
        r = buscar_regra(regras,tributo,cst)
        if r and r.exige_imposto and vl == 0 and vl_s > 0:
            sug = calcular_imposto(r,str(bc or vl_s),str(aliq))
            _add("CRITICO",tributo,f"VL_{tributo}",f"CT-e {reg} CST {cst} exige {tributo}",vl,sug,f"CST-{tributo}-{cst}")

    elif reg in ("C501","C505"):
        tributo = "PIS" if reg == "C501" else "COFINS"
        cst  = str(row.get(f"CST_{tributo}","")).strip()
        bc   = _to_decimal(row.get(f"VL_BC_{tributo}"))
        aliq = _to_decimal(row.get(f"ALIQ_{tributo}"))
        vl   = _to_decimal(row.get(f"VL_{tributo}"))
        vl_i = _to_decimal(row.get("VL_ITEM"))
        r = buscar_regra(regras,tributo,cst)
        if r and r.exige_imposto and vl == 0 and vl_i > 0:
            sug = calcular_imposto(r,str(bc or vl_i),str(aliq))
            _add("CRITICO",tributo,f"VL_{tributo}",f"Energia {reg} CST {cst} exige {tributo}",vl,sug,f"CST-{tributo}-{cst}")

    elif reg == "F100":
        vl_op = _to_decimal(row.get("VL_OPER"))
        for tributo in ("PIS","COFINS"):
            cst  = str(row.get(f"CST_{tributo}","")).strip()
            bc   = _to_decimal(row.get(f"VL_BC_{tributo}"))
            aliq = _to_decimal(row.get(f"ALIQ_{tributo}"))
            vl   = _to_decimal(row.get(f"VL_{tributo}"))
            r = buscar_regra(regras,tributo,cst)
            if r and r.exige_imposto and vl == 0 and vl_op > 0:
                sug = calcular_imposto(r,str(bc or vl_op),str(aliq))
                _add("CRITICO",tributo,f"VL_{tributo}",f"F100 CST {cst} exige {tributo}",vl,sug,f"CST-{tributo}-{cst}")
    return incs


def avaliar_dataframe(df: pd.DataFrame, regras: list) -> pd.DataFrame:
    alvos = ["C170","C190","D100","D101","D105","C501","C505","F100","C500"]
    df_alvo = df[df["REG"].isin(alvos)] if "REG" in df.columns else pd.DataFrame()
    todas: list = []
    for _, row in df_alvo.iterrows():
        todas.extend(avaliar_linha_sped(row, regras))
    cols = ["LINHA_NUM","REGISTRO","BLOCO","TIPO","TRIBUTO","CAMPO","DESCRICAO","VALOR_ATUAL","VALOR_SUGERIDO","REGRA"]
    return pd.DataFrame(todas, columns=cols) if todas else pd.DataFrame(columns=cols)


# =============================================================================
# ░░░░░░░░░░░░░░░░░░  BLOCO 4 — EXPORTAÇÃO  ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
# =============================================================================

_COR_AZUL   = "1E3A5F"
_COR_CINZA  = "F2F4F7"
_COR_VERM   = "FDECEA"
_COR_AMAR   = "FFF3CD"
_COR_BRNC   = "FFFFFF"


def _add_aba_excel(wb: Workbook, df: pd.DataFrame, nome: str, cor_cab: str = _COR_AZUL):
    ws = wb.create_sheet(title=nome[:31])
    if df.empty:
        ws.append(["Sem dados"])
        return
    ws.append(list(df.columns))
    fill_cab = PatternFill("solid", fgColor=cor_cab)
    font_cab = Font(bold=True, color="FFFFFF", size=10)
    aln_cab  = Alignment(horizontal="center", vertical="center")
    for c in range(1, len(df.columns) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = fill_cab; cell.font = font_cab; cell.alignment = aln_cab
    ws.freeze_panes = "A2"
    fill_alt  = PatternFill("solid", fgColor=_COR_CINZA)
    fill_crit = PatternFill("solid", fgColor=_COR_VERM)
    fill_avis = PatternFill("solid", fgColor=_COR_AMAR)
    tipo_idx  = list(df.columns).index("TIPO") if "TIPO" in df.columns else -1
    for i, row in enumerate(dataframe_to_rows(df, index=False, header=False), start=2):
        ws.append(row)
        fill = fill_alt if i % 2 == 0 else PatternFill("solid", fgColor=_COR_BRNC)
        if tipo_idx >= 0 and tipo_idx < len(row):
            tv = str(row[tipo_idx])
            if tv == "CRITICO": fill = fill_crit
            elif tv == "AVISO":  fill = fill_avis
        for c in range(1, len(row) + 1):
            cell = ws.cell(row=i, column=c)
            cell.fill = fill
            cell.font = Font(size=9)
            cell.alignment = Alignment(vertical="center")
    for col in ws.columns:
        ml = max((len(str(cell.value)) for cell in col if cell.value), default=8)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(ml + 4, 55)
    ws.auto_filter.ref = ws.dimensions


def exportar_excel_auditoria(df_incs, df_alt, df_log, df_regras, info) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    # Capa
    ws_c = wb.create_sheet("Capa")
    ws_c.sheet_view.showGridLines = False
    ws_c.column_dimensions["A"].width = 32
    ws_c.column_dimensions["B"].width = 50
    ws_c["A1"] = "SPED STUDIO — RELATÓRIO DE AUDITORIA FISCAL"
    ws_c["A1"].fill = PatternFill("solid", fgColor=_COR_AZUL)
    ws_c["A1"].font = Font(bold=True, color="FFFFFF", size=13)
    ws_c.merge_cells("A1:B1")
    ws_c["A1"].alignment = Alignment(horizontal="center")
    dados = [
        ("Empresa", info.get("nome_empresa","")),("CNPJ", info.get("cnpj","")),
        ("IE", info.get("ie","")),("UF", info.get("uf","")),
        ("Período", f"{info.get('dt_ini','')} a {info.get('dt_fin','')}"),
        ("Tipo SPED", info.get("tipo_arquivo","")),
        ("Gerado em", datetime.now().strftime("%d/%m/%Y %H:%M:%S")),
        ("Total inconsistências", str(len(df_incs))),
        ("Registros alterados", str(len(df_alt))),
    ]
    for r, (k, v) in enumerate(dados, start=3):
        ws_c.cell(row=r,column=1,value=k).font = Font(bold=True,size=10)
        ws_c.cell(row=r,column=2,value=v).font = Font(size=10)
    _add_aba_excel(wb, df_incs,   "Inconsistências",    _COR_AZUL)
    _add_aba_excel(wb, df_alt,    "Registros Alterados","2E6099")
    _add_aba_excel(wb, df_log,    "Log Auditoria",      "2E6099")
    _add_aba_excel(wb, df_regras, "Regras",             "2E7D32")
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def exportar_csv_incs(df_incs: pd.DataFrame) -> bytes:
    buf = io.StringIO()
    df_incs.to_csv(buf, index=False, sep=";", encoding="utf-8-sig")
    return buf.getvalue().encode("utf-8-sig")


# =============================================================================
# ░░░░░░░░░░░░░░░░░░  BLOCO 5 — STREAMLIT UI  ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
# =============================================================================

st.set_page_config(
    page_title="SPED Studio",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif;}
.stApp{background-color:#0F1621;color:#E2E8F0;}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#1A2535 0%,#111827 100%);border-right:1px solid #2D3748;}
[data-testid="stSidebar"] *{color:#CBD5E0 !important;}
.sped-header{background:linear-gradient(135deg,#1E3A5F 0%,#2D6099 50%,#1A4A7C 100%);padding:1.2rem 1.8rem;border-radius:10px;margin-bottom:1.2rem;border-left:4px solid #4299E1;box-shadow:0 4px 15px rgba(66,153,225,.2);}
.sped-header h1{color:#fff;font-size:1.5rem;font-weight:700;margin:0;}
.sped-header p{color:#BEE3F8;font-size:.82rem;margin:.3rem 0 0;}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(155px,1fr));gap:.8rem;margin-bottom:1.2rem;}
.kpi-card{background:#1A2535;border:1px solid #2D3748;border-radius:8px;padding:1rem 1.2rem;text-align:center;transition:border-color .2s;}
.kpi-card:hover{border-color:#4299E1;}
.kpi-card .kv{font-size:1.8rem;font-weight:700;color:#4299E1;line-height:1;}
.kpi-card .kl{font-size:.72rem;color:#718096;margin-top:.35rem;text-transform:uppercase;letter-spacing:.05em;}
.kpi-card.critico .kv{color:#FC8181;}
.kpi-card.aviso .kv{color:#F6AD55;}
.kpi-card.ok .kv{color:#68D391;}
.stTabs [data-baseweb="tab-list"]{background:#1A2535;border-radius:8px;}
.stTabs [data-baseweb="tab"]{color:#718096 !important;}
.stTabs [aria-selected="true"]{color:#4299E1 !important;border-bottom:2px solid #4299E1 !important;}
.stButton>button{background:linear-gradient(135deg,#2B6CB0,#2C5282);color:#fff;border:none;border-radius:6px;font-weight:500;transition:all .2s;}
.stButton>button:hover{background:linear-gradient(135deg,#3182CE,#2B6CB0);transform:translateY(-1px);}
.stTextInput input,.stSelectbox select,.stNumberInput input{background-color:#1A2535 !important;color:#E2E8F0 !important;border:1px solid #2D3748 !important;border-radius:6px !important;}
[data-testid="stFileUploader"]{background:#1A2535 !important;border:2px dashed #2D3748 !important;border-radius:10px !important;}
.stDataFrame{border-radius:8px;overflow:hidden;}
[data-testid="stDataFrame"] *{font-size:.8rem !important;}
::-webkit-scrollbar{width:6px;height:6px;}
::-webkit-scrollbar-track{background:#1A2535;}
::-webkit-scrollbar-thumb{background:#2D3748;border-radius:3px;}
.section-title{font-size:.75rem;font-weight:600;color:#718096;text-transform:uppercase;letter-spacing:.08em;margin:1rem 0 .4rem;padding-bottom:.3rem;border-bottom:1px solid #2D3748;}
</style>
""", unsafe_allow_html=True)

# ── Estado global ─────────────────────────────────────────────────────────────
def _init():
    defs = {
        "resultado":      None,
        "df_editado":     None,
        "df_original":    None,
        "regras":         list(TODAS_REGRAS_PADRAO),
        "df_incs":        pd.DataFrame(),
        "df_log":         pd.DataFrame(columns=["DATA_HORA","USUARIO","LINHA_NUM","REGISTRO",
                                                 "CAMPO","VALOR_ANTES","VALOR_DEPOIS","REGRA","MOTIVO"]),
        "pagina":         "upload",
        "usuario":        "analista_fiscal",
        "df_ctes_import": pd.DataFrame(),
    }
    for k, v in defs.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init()

# ── Helpers ───────────────────────────────────────────────────────────────────
def _log(ln, reg, campo, antes, depois, regra="", motivo=""):
    nova = {"DATA_HORA":datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "USUARIO":st.session_state.usuario,"LINHA_NUM":ln,"REGISTRO":reg,
            "CAMPO":campo,"VALOR_ANTES":antes,"VALOR_DEPOIS":depois,"REGRA":regra,"MOTIVO":motivo}
    st.session_state.df_log = pd.concat(
        [st.session_state.df_log, pd.DataFrame([nova])], ignore_index=True)


def _fmt_cnpj(c):
    c = re.sub(r"\D","",str(c))
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}" if len(c)==14 else c


def _fmt_data(d):
    d = re.sub(r"\D","",str(d))
    return f"{d[:2]}/{d[2:4]}/{d[4:]}" if len(d)==8 else d


def _kpi(v, l, cls=""):
    return f'<div class="kpi-card {cls}"><div class="kv">{v}</div><div class="kl">{l}</div></div>'


def _adicionar_reg(reg, bloco, valores):
    df = st.session_state.df_editado
    if df is None:
        df = pd.DataFrame()
    max_ln = int(df["LINHA_NUM"].max())+1 if not df.empty and "LINHA_NUM" in df.columns else 1
    novo = {"REG":reg,"BLOCO":bloco,"LINHA_NUM":max_ln}
    novo.update(valores)
    df_novo = pd.DataFrame([novo])
    if not df.empty:
        for col in df.columns:
            if col not in df_novo.columns:
                df_novo[col] = ""
        df_novo = df_novo.reindex(columns=df.columns, fill_value="")
    st.session_state.df_editado = pd.concat([df, df_novo], ignore_index=True)
    if st.session_state.df_original is not None and not df.empty:
        st.session_state.df_original = pd.concat(
            [st.session_state.df_original, df_novo], ignore_index=True)
    _log(max_ln, reg, "NOVO_REGISTRO", "", json.dumps(valores)[:100], "manual", f"{reg} adicionado")


# =============================================================================
# ── SIDEBAR ───────────────────────────────────────────────────────────────────
# =============================================================================
def _sidebar():
    with st.sidebar:
        st.markdown("""
        <div style="text-align:center;padding:1rem 0 1.5rem;">
            <div style="font-size:2rem;">🧾</div>
            <div style="font-size:1.1rem;font-weight:700;color:#4299E1;">SPED Studio</div>
            <div style="font-size:.7rem;color:#718096;">v2.0 · Auditoria Fiscal</div>
        </div>""", unsafe_allow_html=True)

        st.markdown('<div class="section-title">Tipo de Arquivo</div>', unsafe_allow_html=True)
        st.radio("", ["EFD ICMS/IPI","EFD Contribuições"], key="tipo_sped_aba", label_visibility="collapsed")

        st.markdown('<div class="section-title">Analista</div>', unsafe_allow_html=True)
        st.session_state.usuario = st.text_input("", value=st.session_state.usuario,
                                                   label_visibility="collapsed", placeholder="Nome do analista")

        st.markdown('<div class="section-title">Navegação</div>', unsafe_allow_html=True)
        menu = [
            ("upload",         "📤 Upload do Arquivo"),
            ("dashboard",      "📊 Dashboard"),
            ("blocos",         "📦 Blocos & Registros"),
            ("notas",          "🗒️ Notas Fiscais"),
            ("inconsistencias","⚠️ Inconsistências"),
            ("correcoes",      "🔧 Correções em Massa"),
            ("editor",         "✏️ Editor Manual"),
            ("cte",            "🚚 CT-e XML → Bloco D"),
            ("energia",        "⚡ Energia (C500/C501/C505)"),
            ("aluguel",        "🏢 Aluguel (F100)"),
            ("regras",         "⚙️ Motor de Regras"),
            ("exportacao",     "📥 Exportação"),
            ("log",            "🔍 Log de Auditoria"),
        ]
        for key, label in menu:
            if st.button(label, key=f"nav_{key}", use_container_width=True):
                st.session_state.pagina = key
                st.rerun()

        r = st.session_state.resultado
        if r:
            st.markdown('<div class="section-title">Arquivo Atual</div>', unsafe_allow_html=True)
            st.markdown(f"""
            <div style="font-size:.75rem;color:#68D391;background:#1C4532;padding:.5rem;border-radius:6px;">
                ✅ {r.tipo_arquivo}<br>
                📅 {_fmt_data(r.dt_ini)} → {_fmt_data(r.dt_fin)}<br>
                🏢 {(r.nome_empresa or "—")[:26]}<br>
                📋 {len(r.df_registros):,} linhas
            </div>""", unsafe_allow_html=True)


# =============================================================================
# ── PÁGINA: UPLOAD ────────────────────────────────────────────────────────────
# =============================================================================
def pg_upload():
    st.markdown("""<div class="sped-header">
        <h1>📤 Upload do Arquivo SPED</h1>
        <p>Selecione o arquivo TXT do SPED para iniciar a auditoria fiscal.</p>
    </div>""", unsafe_allow_html=True)

    col1, col2 = st.columns([2,1])
    with col1:
        arq = st.file_uploader("Arquivo SPED (.txt)", type=["txt"], key="up_sped")
        if arq:
            with st.spinner("🔄 Processando..."):
                resultado = parse_arquivo_sped(arq.read())
            st.session_state.resultado   = resultado
            st.session_state.df_editado  = resultado.df_registros.copy()
            st.session_state.df_original = resultado.df_registros.copy()
            st.session_state.df_incs     = pd.DataFrame()

            for e in resultado.erros_parsing:
                st.warning(f"⚠️ {e}")

            st.success(f"✅ **{arq.name}** carregado — {len(resultado.df_registros):,} registros")

            c1,c2,c3,c4 = st.columns(4)
            c1.metric("Tipo", resultado.tipo_arquivo.replace("_"," "))
            c2.metric("CNPJ", _fmt_cnpj(resultado.cnpj))
            c3.metric("Período", f"{_fmt_data(resultado.dt_ini)} → {_fmt_data(resultado.dt_fin)}")
            c4.metric("Blocos", " ".join(resultado.blocos_encontrados))

            with st.spinner("🔍 Validando..."):
                df_incs = avaliar_dataframe(resultado.df_registros, st.session_state.regras)
            st.session_state.df_incs = df_incs

            n_cr = len(df_incs[df_incs["TIPO"]=="CRITICO"]) if not df_incs.empty else 0
            n_av = len(df_incs[df_incs["TIPO"]=="AVISO"])   if not df_incs.empty else 0
            if n_cr > 0:
                st.error(f"🚨 **{n_cr}** inconsistências críticas — revise antes de exportar.")
            elif n_av > 0:
                st.warning(f"⚠️ **{n_av}** avisos encontrados.")
            else:
                st.success("✅ Nenhuma inconsistência crítica detectada.")

            if st.button("➡️ Ir para o Dashboard", type="primary"):
                st.session_state.pagina = "dashboard"
                st.rerun()

    with col2:
        st.markdown("""
        <div style="background:#1A2535;border:1px solid #2D3748;border-radius:8px;padding:1rem;">
            <div style="font-weight:600;color:#4299E1;margin-bottom:.8rem;">📌 Suporte</div>
            <div style="font-size:.82rem;color:#CBD5E0;line-height:1.9;">
                ✅ EFD ICMS/IPI<br>✅ EFD Contribuições<br>🔜 ECD<br>🔜 ECF<br><br>
                ✅ CT-e XML → D100<br>✅ Energia C500/C501/C505<br>✅ Aluguel F100<br>
                ✅ Motor de Regras CST<br>✅ Excel + TXT exportação
            </div>
        </div>""", unsafe_allow_html=True)


# =============================================================================
# ── PÁGINA: DASHBOARD ─────────────────────────────────────────────────────────
# =============================================================================
def pg_dashboard():
    r = st.session_state.resultado
    if r is None:
        st.info("Carregue um arquivo SPED primeiro.")
        return
    df      = st.session_state.df_editado
    df_incs = st.session_state.df_incs

    st.markdown(f"""<div class="sped-header">
        <h1>📊 Dashboard — {r.nome_empresa}</h1>
        <p>CNPJ: {_fmt_cnpj(r.cnpj)} · {r.uf} · {_fmt_data(r.dt_ini)} a {_fmt_data(r.dt_fin)} · {r.tipo_arquivo}</p>
    </div>""", unsafe_allow_html=True)

    n_c  = len(df_incs[df_incs["TIPO"]=="CRITICO"]) if not df_incs.empty else 0
    n_a  = len(df_incs[df_incs["TIPO"]=="AVISO"])   if not df_incs.empty else 0
    n170 = len(df[df["REG"]=="C170"]) if "REG" in df.columns else 0
    n100 = len(df[df["REG"]=="C100"]) if "REG" in df.columns else 0
    nd   = len(df[df["REG"]=="D100"]) if "REG" in df.columns else 0
    nf   = len(df[df["REG"]=="F100"]) if "REG" in df.columns else 0

    st.markdown('<div class="kpi-grid">'+
        _kpi(f"{len(df):,}","Registros")+
        _kpi(f"{n100:,}","NF-e (C100)")+
        _kpi(f"{n170:,}","Itens (C170)")+
        _kpi(f"{nd:,}","CT-e (D100)")+
        _kpi(f"{nf:,}","Aluguel (F100)")+
        _kpi(f"{n_c:,}","Críticos","critico")+
        _kpi(f"{n_a:,}","Avisos","aviso")+
        _kpi(f"{len(df_incs):,}","Total Inconsist.")+
    '</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 📦 Registros por Bloco")
        if "BLOCO" in df.columns:
            bc = df["BLOCO"].value_counts().reset_index()
            bc.columns = ["Bloco","Qtd"]
            st.bar_chart(bc.set_index("Bloco"), height=220)
    with c2:
        st.markdown("#### ⚠️ Inconsistências por Tributo")
        if not df_incs.empty and "TRIBUTO" in df_incs.columns:
            tc = df_incs["TRIBUTO"].value_counts().reset_index()
            tc.columns = ["Tributo","Qtd"]
            st.bar_chart(tc.set_index("Tributo"), height=220)

    c3, c4 = st.columns(2)
    with c3:
        st.markdown("#### 🗒️ Notas por CFOP")
        df_c100 = df[df["REG"]=="C100"] if "REG" in df.columns else pd.DataFrame()
        if not df_c100.empty and "CFOP" in df_c100.columns:
            st.dataframe(df_c100["CFOP"].value_counts().head(10).reset_index().rename(columns={"index":"CFOP","CFOP":"Qtd"}),
                         use_container_width=True, height=180)
    with c4:
        st.markdown("#### 🏷️ Itens por CST ICMS")
        df_c170 = df[df["REG"]=="C170"] if "REG" in df.columns else pd.DataFrame()
        if not df_c170.empty and "CST_ICMS" in df_c170.columns:
            st.dataframe(df_c170["CST_ICMS"].value_counts().head(10).reset_index().rename(columns={"index":"CST","CST_ICMS":"Qtd"}),
                         use_container_width=True, height=180)

    if not df_incs.empty:
        st.markdown("#### 🚨 Top Inconsistências Críticas")
        df_c = df_incs[df_incs["TIPO"]=="CRITICO"].head(20)
        if not df_c.empty:
            st.dataframe(df_c[["LINHA_NUM","REGISTRO","TRIBUTO","CAMPO","DESCRICAO","VALOR_ATUAL","VALOR_SUGERIDO"]],
                         use_container_width=True, height=240)


# =============================================================================
# ── PÁGINA: BLOCOS & REGISTROS ────────────────────────────────────────────────
# =============================================================================
def pg_blocos():
    df = st.session_state.df_editado
    if df is None:
        st.info("Carregue um arquivo SPED primeiro.")
        return
    st.markdown("""<div class="sped-header">
        <h1>📦 Blocos & Registros</h1>
        <p>Filtre e visualize qualquer bloco ou registro do arquivo.</p>
    </div>""", unsafe_allow_html=True)

    blocos = sorted(df["BLOCO"].dropna().unique().tolist()) if "BLOCO" in df.columns else []
    c1,c2,c3 = st.columns([1,1,2])
    with c1: bl = st.selectbox("Bloco",["(todos)"]+blocos)
    with c2:
        regs = df["REG"].dropna().unique().tolist()
        if bl != "(todos)": regs = df[df["BLOCO"]==bl]["REG"].dropna().unique().tolist()
        rg = st.selectbox("Registro",["(todos)"]+sorted(regs))
    with c3: busca = st.text_input("🔍 Buscar em qualquer campo","")

    df_f = df.copy()
    if bl != "(todos)": df_f = df_f[df_f["BLOCO"]==bl]
    if rg != "(todos)": df_f = df_f[df_f["REG"]==rg]
    if busca:
        mask = df_f.apply(lambda r: r.astype(str).str.contains(busca,case=False,na=False).any(), axis=1)
        df_f = df_f[mask]

    st.markdown(f"**{len(df_f):,} registros**")
    cols = ["LINHA_NUM","BLOCO","REG"] + [c for c in df_f.columns if c not in ("LINHA_NUM","BLOCO","REG")][:28]
    cols = [c for c in cols if c in df_f.columns]
    st.dataframe(df_f[cols], use_container_width=True, height=450)
    csv = df_f[cols].to_csv(index=False,sep=";",encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button("⬇️ CSV filtrado", csv, "registros.csv", "text/csv")


# =============================================================================
# ── PÁGINA: NOTAS FISCAIS ─────────────────────────────────────────────────────
# =============================================================================
def pg_notas():
    df = st.session_state.df_editado
    if df is None:
        st.info("Carregue um arquivo SPED primeiro.")
        return
    st.markdown("""<div class="sped-header">
        <h1>🗒️ Notas Fiscais</h1>
        <p>Visão de NF-e (C100), Energia (C500) e CT-e (D100) com drill-down.</p>
    </div>""", unsafe_allow_html=True)

    t1,t2,t3 = st.tabs(["NF-e (C100)","Energia (C500)","CT-e (D100)"])

    def _show_tab(reg, cols_pref, drill_reg=None, drill_src_cols=None):
        df_r = df[df["REG"]==reg].copy() if "REG" in df.columns else pd.DataFrame()
        if df_r.empty:
            st.info(f"Sem registros {reg}.")
            return
        cols = [c for c in cols_pref if c in df_r.columns]
        st.dataframe(df_r[cols], use_container_width=True, height=280)
        if drill_reg and "LINHA_NUM" in df_r.columns:
            sel = st.selectbox(f"Drill-down {drill_reg} da linha:", ["—"]+[str(x) for x in df_r["LINHA_NUM"].tolist()])
            if sel != "—":
                ln = int(sel)
                df_d = df[(df["REG"]==drill_reg) & (df["LINHA_NUM"]>ln)].copy()
                prox = df_r[df_r["LINHA_NUM"]>ln]["LINHA_NUM"].min()
                if not pd.isna(prox):
                    df_d = df_d[df_d["LINHA_NUM"]<prox]
                st.dataframe(df_d, use_container_width=True, height=200)

    with t1:
        _show_tab("C100",["LINHA_NUM","COD_PART","COD_MOD","NUM_DOC","DT_DOC","VL_DOC",
                           "VL_ICMS","VL_PIS","VL_COFINS","COD_SIT"],"C170")
    with t2:
        _show_tab("C500",["LINHA_NUM","COD_PART","NUM_DOC","DT_DOC","VL_DOC",
                           "VL_ICMS","VL_PIS","VL_COFINS"],"C501")
    with t3:
        _show_tab("D100",["LINHA_NUM","COD_PART","COD_MOD","NUM_DOC","DT_DOC","VL_DOC",
                           "VL_SERV","VL_ICMS","VL_PIS","VL_COFINS"],"D101")


# =============================================================================
# ── PÁGINA: INCONSISTÊNCIAS ───────────────────────────────────────────────────
# =============================================================================
def pg_inconsistencias():
    df = st.session_state.df_editado
    if df is None:
        st.info("Carregue um arquivo SPED primeiro.")
        return
    st.markdown("""<div class="sped-header">
        <h1>⚠️ Inconsistências Fiscais</h1>
        <p>Motor de regras CST × CFOP: críticos e avisos com valores sugeridos.</p>
    </div>""", unsafe_allow_html=True)

    c1, _ = st.columns([1,4])
    with c1:
        if st.button("🔄 Re-validar", type="primary"):
            with st.spinner("Validando..."):
                st.session_state.df_incs = avaliar_dataframe(df, st.session_state.regras)
            st.rerun()

    df_incs = st.session_state.df_incs
    if df_incs.empty:
        st.success("✅ Nenhuma inconsistência ou validação não executada.")
        return

    c1,c2,c3 = st.columns(3)
    with c1: f_t = st.multiselect("Tipo",    df_incs["TIPO"].unique().tolist(),    default=df_incs["TIPO"].unique().tolist())
    with c2: f_r = st.multiselect("Tributo", df_incs["TRIBUTO"].unique().tolist(), default=df_incs["TRIBUTO"].unique().tolist())
    with c3: f_g = st.multiselect("Registro",df_incs["REGISTRO"].unique().tolist(),default=df_incs["REGISTRO"].unique().tolist())

    df_f = df_incs[df_incs["TIPO"].isin(f_t)&df_incs["TRIBUTO"].isin(f_r)&df_incs["REGISTRO"].isin(f_g)]

    n_c = len(df_f[df_f["TIPO"]=="CRITICO"])
    n_a = len(df_f[df_f["TIPO"]=="AVISO"])
    st.markdown('<div class="kpi-grid">'+_kpi(len(df_f),"Total")+_kpi(n_c,"Críticos","critico")+_kpi(n_a,"Avisos","aviso")+'</div>',unsafe_allow_html=True)

    def _cor(val):
        if val=="CRITICO": return "background:rgba(252,129,129,.15);color:#FC8181;font-weight:600"
        if val=="AVISO":   return "background:rgba(246,173,85,.15);color:#F6AD55;font-weight:600"
        return ""

    st.dataframe(
        df_f[["LINHA_NUM","REGISTRO","BLOCO","TIPO","TRIBUTO","CAMPO","DESCRICAO","VALOR_ATUAL","VALOR_SUGERIDO","REGRA"]]
            .style.applymap(_cor, subset=["TIPO"]),
        use_container_width=True, height=420
    )
    col_d1,col_d2 = st.columns(2)
    with col_d1:
        st.download_button("⬇️ CSV",  exportar_csv_incs(df_f), "inconsistencias.csv","text/csv")
    with col_d2:
        if st.button("➡️ Correções em Massa"):
            st.session_state.pagina="correcoes"; st.rerun()


# =============================================================================
# ── PÁGINA: CORREÇÕES EM MASSA ────────────────────────────────────────────────
# =============================================================================
def pg_correcoes():
    df = st.session_state.df_editado
    if df is None:
        st.info("Carregue um arquivo SPED primeiro.")
        return
    st.markdown("""<div class="sped-header">
        <h1>🔧 Correções em Massa</h1>
        <p>Filtro → prévia → confirmação com registro de auditoria completo.</p>
    </div>""", unsafe_allow_html=True)

    regras = st.session_state.regras
    c1,c2,c3,c4 = st.columns(4)
    with c1: f_reg = st.multiselect("Registros",sorted(df["REG"].dropna().unique().tolist()),
                                     default=["C170","C190","D101","D105","C501","C505","F100"])
    with c2: f_cst  = st.text_input("CST contém","")
    with c3: f_cfop = st.text_input("CFOP contém","")
    with c4: f_bl   = st.multiselect("Bloco",sorted(df["BLOCO"].dropna().unique().tolist()))

    df_sel = df[df["REG"].isin(f_reg)].copy()
    if f_bl:   df_sel = df_sel[df_sel["BLOCO"].isin(f_bl)]
    if f_cfop and "CFOP" in df_sel.columns:
        df_sel = df_sel[df_sel["CFOP"].str.contains(f_cfop,na=False)]

    st.markdown(f"**{len(df_sel):,} registros** selecionados")
    if df_sel.empty:
        st.info("Nenhum registro nos filtros selecionados.")
        return

    acao = st.radio("Ação", [
        "Calcular imposto automático (base × alíquota / 100)",
        "Preencher base com valor do item",
        "Preencher alíquota com valor fixo",
        "Zerar tributação onde CST não exige",
    ])
    aliq_fixa = 12.0
    if "alíquota fixa" in acao:
        aliq_fixa = st.number_input("Alíquota (%)", 0.0, 100.0, 12.0, 0.01)

    # Prévia
    previa = []
    for _, row in df_sel.head(200).iterrows():
        ln  = int(row.get("LINHA_NUM",0))
        reg = str(row.get("REG",""))
        if "Calcular imposto" in acao:
            for trib,bc_c,al_c,vl_c,cst_c in [
                ("ICMS","VL_BC_ICMS","ALIQ_ICMS","VL_ICMS","CST_ICMS"),
                ("PIS","VL_BC_PIS","ALIQ_PIS","VL_PIS","CST_PIS"),
                ("COFINS","VL_BC_COFINS","ALIQ_COFINS","VL_COFINS","CST_COFINS"),
            ]:
                if bc_c in row.index and al_c in row.index and vl_c in row.index:
                    cst = str(row.get(cst_c,"")).strip()
                    r_ = buscar_regra(regras,trib,cst)
                    if r_ and r_.exige_imposto:
                        bc = _to_decimal(row.get(bc_c)); al = _to_decimal(row.get(al_c))
                        if bc > 0 and al > 0:
                            novo = calcular_imposto(r_,str(bc),str(al))
                            atual= _to_decimal(row.get(vl_c))
                            if novo != atual:
                                previa.append({"LINHA_NUM":ln,"REG":reg,"CAMPO":vl_c,
                                               "ANTES":str(atual),"DEPOIS":str(novo),"TRIBUTO":trib})
        elif "base" in acao:
            for trib,bc_c,vi_c in [("ICMS","VL_BC_ICMS","VL_ITEM"),
                                     ("PIS","VL_BC_PIS","VL_ITEM"),
                                     ("COFINS","VL_BC_COFINS","VL_ITEM")]:
                if bc_c in row.index:
                    bc_at = _to_decimal(row.get(bc_c))
                    vi    = _to_decimal(row.get(vi_c,row.get("VL_OPER","0")))
                    if bc_at == 0 and vi > 0:
                        previa.append({"LINHA_NUM":ln,"REG":reg,"CAMPO":bc_c,
                                       "ANTES":str(bc_at),"DEPOIS":str(vi),"TRIBUTO":trib})
        elif "alíquota fixa" in acao or "Preencher alíquota" in acao:
            for al_c in ["ALIQ_ICMS","ALIQ_PIS","ALIQ_COFINS"]:
                if al_c in row.index and _to_decimal(row.get(al_c))==0:
                    previa.append({"LINHA_NUM":ln,"REG":reg,"CAMPO":al_c,
                                   "ANTES":"0","DEPOIS":str(aliq_fixa),"TRIBUTO":al_c.split("_")[1]})

    df_prev = pd.DataFrame(previa)
    if df_prev.empty:
        st.info("Nenhuma alteração prevista com os filtros e ação selecionados.")
        return

    st.markdown(f"**{len(df_prev)} alterações previstas:**")
    st.dataframe(df_prev, use_container_width=True, height=280)
    motivo = st.text_input("Motivo (auditoria)", "Correção em massa via motor de regras")

    if st.button("✅ Aplicar Correções", type="primary"):
        df_mut = st.session_state.df_editado.copy()
        for _, pr in df_prev.iterrows():
            mask = df_mut["LINHA_NUM"]==pr["LINHA_NUM"]
            if pr["CAMPO"] in df_mut.columns:
                df_mut.loc[mask, pr["CAMPO"]] = pr["DEPOIS"]
                _log(int(pr["LINHA_NUM"]),pr["REG"],pr["CAMPO"],pr["ANTES"],pr["DEPOIS"],acao[:50],motivo)
        st.session_state.df_editado = df_mut
        st.session_state.df_incs    = avaliar_dataframe(df_mut, regras)
        st.success(f"✅ {len(df_prev)} correções aplicadas!")
        st.rerun()


# =============================================================================
# ── PÁGINA: EDITOR MANUAL ─────────────────────────────────────────────────────
# =============================================================================
def pg_editor():
    df = st.session_state.df_editado
    if df is None:
        st.info("Carregue um arquivo SPED primeiro.")
        return
    st.markdown("""<div class="sped-header">
        <h1>✏️ Editor Manual</h1>
        <p>Edição campo a campo com comparação antes/depois e restauração ao original.</p>
    </div>""", unsafe_allow_html=True)

    c1,c2 = st.columns([1,2])
    with c1: reg_sel = st.selectbox("Registro", sorted(df["REG"].dropna().unique()))
    with c2: busca_ln = st.number_input("Número da linha",1,int(df["LINHA_NUM"].max()),1)

    df_reg = df[df["REG"]==reg_sel].copy()
    if df_reg.empty:
        st.info(f"Nenhum registro {reg_sel}.")
        return

    idx = (df_reg["LINHA_NUM"]-busca_ln).abs().idxmin()
    linha = df_reg.loc[idx]
    ln    = int(linha["LINHA_NUM"])
    st.markdown(f"**{reg_sel} — Linha {ln}**")

    campos_reg = CAMPOS_REGISTRO.get(reg_sel,[])
    campos_edit = [c for c in campos_reg if c in df.columns and c not in ("REG",)]
    if not campos_edit:
        campos_edit = [c for c in linha.index if c not in ("REG","BLOCO","LINHA_NUM")]

    df_orig = st.session_state.df_original
    nova_vals = {}
    col_ed, col_orig = st.columns(2)

    with col_ed:
        st.markdown("##### ✏️ Editar")
        for campo in campos_edit[:25]:
            val = str(linha.get(campo,"")) if pd.notna(linha.get(campo,"")) else ""
            nova_vals[campo] = st.text_input(campo, val, key=f"ed_{campo}_{ln}")

    with col_orig:
        st.markdown("##### 🔒 Original")
        if df_orig is not None:
            lo = df_orig[df_orig["LINHA_NUM"]==ln]
            if not lo.empty:
                for campo in campos_edit[:25]:
                    o = str(lo.iloc[0].get(campo,"")) if campo in lo.columns else ""
                    a = str(linha.get(campo,""))
                    cor = "#FC8181" if o != a else "#68D391"
                    st.markdown(
                        f'<div style="padding:3px 8px;margin:2px 0;background:#1A2535;border-radius:4px;color:{cor};font-size:.82rem;">'
                        f'<b>{campo}:</b> {o}</div>', unsafe_allow_html=True)

    motivo = st.text_input("Motivo","Correção manual")
    cb1, cb2 = st.columns(2)
    with cb1:
        if st.button("💾 Salvar", type="primary"):
            df_mut = st.session_state.df_editado.copy()
            mask   = df_mut["LINHA_NUM"]==ln
            for campo, nv in nova_vals.items():
                if campo in df_mut.columns:
                    antes = str(df_mut.loc[mask,campo].values[0]) if mask.any() else ""
                    if antes != nv:
                        df_mut.loc[mask,campo] = nv
                        _log(ln,reg_sel,campo,antes,nv,"editor_manual",motivo)
            st.session_state.df_editado = df_mut
            st.success("✅ Salvo!")
            st.rerun()
    with cb2:
        if st.button("↩️ Restaurar Original"):
            df_mut = st.session_state.df_editado.copy()
            mask   = df_mut["LINHA_NUM"]==ln
            if df_orig is not None:
                lo = df_orig[df_orig["LINHA_NUM"]==ln]
                if not lo.empty:
                    for campo in campos_edit:
                        if campo in df_mut.columns and campo in lo.columns:
                            vo = str(lo.iloc[0][campo])
                            va = str(df_mut.loc[mask,campo].values[0])
                            if va != vo:
                                df_mut.loc[mask,campo] = vo
                                _log(ln,reg_sel,campo,va,vo,"restaurar","Restaurado")
            st.session_state.df_editado = df_mut
            st.success("↩️ Restaurado!")
            st.rerun()


# =============================================================================
# ── PÁGINA: CT-e XML → BLOCO D ───────────────────────────────────────────────
# =============================================================================
def pg_cte():
    st.markdown("""<div class="sped-header">
        <h1>🚚 CT-e XML → Bloco D</h1>
        <p>Importe XMLs CT-e para gerar D100, D101 e D105 automaticamente.</p>
    </div>""", unsafe_allow_html=True)

    xmls = st.file_uploader("XMLs CT-e (múltiplos)", type=["xml"], accept_multiple_files=True)
    if not xmls:
        st.info("Faça upload de um ou mais arquivos XML CT-e.")
        return

    if st.button("🔄 Processar XMLs", type="primary"):
        rows = []; erros_g = []
        for x in xmls:
            res = parse_xml_cte(x.read())
            erros_g.extend([f"{x.name}: {e}" for e in res["erros"]])
            if res["d100"]:
                rows.append({**res["d100"],"REG":"D100","BLOCO":"D","LINHA_NUM":0,"_STATUS":"IMPORTADO"})
            if res["d101"]:
                rows.append({**res["d101"],"REG":"D101","BLOCO":"D","LINHA_NUM":0,"_STATUS":"IMPORTADO","_CHV":res["chave"]})
            if res["d105"]:
                rows.append({**res["d105"],"REG":"D105","BLOCO":"D","LINHA_NUM":0,"_STATUS":"IMPORTADO","_CHV":res["chave"]})
        for e in erros_g: st.warning(f"⚠️ {e}")
        if rows:
            st.session_state.df_ctes_import = pd.DataFrame(rows)
            st.success(f"✅ {len([r for r in rows if r.get('REG')=='D100'])} CT-e(s) processados.")
        else:
            st.error("Nenhum CT-e processado com sucesso.")
            return

    df_ctes = st.session_state.df_ctes_import
    if df_ctes.empty:
        return

    cols = [c for c in ["REG","NUM_DOC","DT_DOC","COD_PART","NOME_EMIT","VL_DOC",
                          "VL_SERV","VL_ICMS","VL_PIS","VL_COFINS","CHV_CTE"] if c in df_ctes.columns]
    st.dataframe(df_ctes[cols] if cols else df_ctes, use_container_width=True, height=280)

    st.markdown("#### ⚙️ Configurações antes de incorporar")
    c1,c2,c3 = st.columns(3)
    with c1: cod_part = st.text_input("COD_PART padrão (CNPJ transportador)","")
    with c2: nat_bc   = st.selectbox("NAT_BC_CRED padrão",["17","01","02","03","04","05","06","07"],
                                      help="17 = Frete (Lei 10.833/2003)")
    with c3: ind_nat  = st.selectbox("IND_NAT_FRT padrão",["07 - Outros","01 - Subcontratação",
                                                             "02 - Redespacho","03 - Redespacho intermediário"])

    if st.session_state.resultado is None:
        st.warning("Carregue um SPED principal antes de incorporar.")
    else:
        if st.button("✅ Incorporar ao Bloco D", type="primary"):
            df_m = st.session_state.df_editado.copy()
            max_ln = int(df_m["LINHA_NUM"].max()) if not df_m.empty else 0
            novos = []
            for i, (_, row) in enumerate(df_ctes.iterrows(), 1):
                n = row.to_dict()
                n["LINHA_NUM"] = max_ln + i
                if cod_part: n["COD_PART"] = cod_part
                n["NAT_BC_CRED"] = nat_bc
                n["IND_NAT_FRT"] = ind_nat[0:2].strip()
                novos.append(n)
            df_n = pd.DataFrame(novos)
            for col in df_m.columns:
                if col not in df_n.columns: df_n[col] = ""
            df_n = df_n.reindex(columns=df_m.columns, fill_value="")
            st.session_state.df_editado = pd.concat([df_m, df_n], ignore_index=True)
            _log(0,"D100/D101/D105","IMPORTACAO","",f"{len(df_n)} regs","import_cte",f"{len(xmls)} XMLs")
            st.success(f"✅ {len(df_n)} registros adicionados ao Bloco D!")
            st.rerun()

    csv_b = df_ctes.to_csv(index=False,sep=";",encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button("⬇️ CSV dos registros CT-e", csv_b, "cte_bloco_d.csv", "text/csv")


# =============================================================================
# ── PÁGINA: ENERGIA (C500/C501/C505) ─────────────────────────────────────────
# =============================================================================
def pg_energia():
    st.markdown("""<div class="sped-header">
        <h1>⚡ Energia Elétrica — C500 / C501 / C505</h1>
        <p>Notas fiscais de energia e crédito PIS/COFINS sobre consumo.</p>
    </div>""", unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["📄 C500 — Documento","🟦 C501 — PIS","🟨 C505 — COFINS"])

    def _lista_reg(reg):
        df = st.session_state.df_editado
        if df is None or "REG" not in df.columns:
            return pd.DataFrame()
        return df[df["REG"]==reg].copy()

    with tab1:
        df_r = _lista_reg("C500")
        if not df_r.empty:
            cols = [c for c in ["LINHA_NUM","COD_PART","NUM_DOC","DT_DOC","VL_DOC",
                                  "VL_BC_ICMS","VL_ICMS","VL_PIS","VL_COFINS"] if c in df_r.columns]
            st.dataframe(df_r[cols], use_container_width=True, height=180)
        st.markdown("#### ➕ Novo C500")
        with st.form("form_c500"):
            c1,c2,c3 = st.columns(3)
            with c1:
                cod_part = st.text_input("COD_PART (CNPJ distribuidora)","")
                num_doc  = st.text_input("NUM_DOC","")
                dt_doc   = st.text_input("DT_DOC (DDMMAAAA)", datetime.now().strftime("%d%m%Y"))
            with c2:
                vl_doc   = st.number_input("VL_DOC",0.0,step=.01,format="%.2f")
                vl_forn  = st.number_input("VL_FORN (energia)",0.0,step=.01,format="%.2f")
                vl_icms  = st.number_input("VL_ICMS",0.0,step=.01,format="%.2f")
            with c3:
                vl_pis   = st.number_input("VL_PIS",0.0,step=.0001,format="%.4f")
                vl_cof   = st.number_input("VL_COFINS",0.0,step=.0001,format="%.4f")
                ser      = st.text_input("SER","")
            if st.form_submit_button("💾 Adicionar C500", type="primary"):
                _adicionar_reg("C500","C",{
                    "IND_OPER":"0","IND_EMIT":"1","COD_PART":cod_part,"COD_MOD":"06",
                    "COD_SIT":"00","SER":ser,"SUB":"","NUM_DOC":num_doc,"DT_DOC":dt_doc,
                    "DT_E_S":dt_doc,"VL_DOC":str(round(vl_doc,2)),"VL_DESC":"0",
                    "VL_FORN":str(round(vl_forn,2)),"VL_SERV_NT":"0","VL_TERC":"0","VL_DA":"0",
                    "VL_BC_ICMS":"0","VL_ICMS":str(round(vl_icms,2)),
                    "VL_BC_ICMS_ST":"0","VL_ICMS_ST":"0","COD_INF":"",
                    "VL_PIS":str(round(vl_pis,4)),"VL_COFINS":str(round(vl_cof,4)),
                })
                st.success("✅ C500 adicionado!"); st.rerun()

    with tab2:
        df_r = _lista_reg("C501")
        if not df_r.empty:
            cols = [c for c in ["LINHA_NUM","CST_PIS","VL_ITEM","NAT_BC_CRED","VL_BC_PIS","ALIQ_PIS","VL_PIS"] if c in df_r.columns]
            st.dataframe(df_r[cols], use_container_width=True, height=180)
        st.markdown("#### ➕ Novo C501")
        with st.form("form_c501"):
            c1,c2,c3 = st.columns(3)
            with c1: cst_pis = st.selectbox("CST_PIS",["50","51","52","53","54","55","56","70","71","72","73","74","75","98","99"])
            with c2:
                vl_item  = st.number_input("VL_ITEM",0.0,step=.01,format="%.2f")
                nat_bc   = st.selectbox("NAT_BC_CRED",["06","01","02","03","04","05","07"])
            with c3:
                aliq_pis = st.number_input("ALIQ_PIS (%)",0.0,100.0,1.65,0.01)
                vl_bc_p  = st.number_input("VL_BC_PIS",0.0,step=.01,format="%.2f")
            vl_pis_calc = round(vl_bc_p * aliq_pis / 100, 4)
            st.info(f"PIS calculado: R$ {vl_pis_calc:.4f}")
            if st.form_submit_button("💾 Adicionar C501", type="primary"):
                _adicionar_reg("C501","C",{
                    "CST_PIS":cst_pis,"VL_ITEM":str(round(vl_item,2)),
                    "NAT_BC_CRED":nat_bc,"VL_BC_PIS":str(round(vl_bc_p,2)),
                    "ALIQ_PIS":str(aliq_pis),"VL_PIS":str(vl_pis_calc),
                })
                st.success("✅ C501 adicionado!"); st.rerun()

    with tab3:
        df_r = _lista_reg("C505")
        if not df_r.empty:
            cols = [c for c in ["LINHA_NUM","CST_COFINS","VL_ITEM","NAT_BC_CRED","VL_BC_COFINS","ALIQ_COFINS","VL_COFINS"] if c in df_r.columns]
            st.dataframe(df_r[cols], use_container_width=True, height=180)
        st.markdown("#### ➕ Novo C505")
        with st.form("form_c505"):
            c1,c2,c3 = st.columns(3)
            with c1: cst_cof = st.selectbox("CST_COFINS",["50","51","52","53","54","55","56","70","71","72","73","74","75","98","99"])
            with c2:
                vl_item2 = st.number_input("VL_ITEM",0.0,step=.01,format="%.2f",key="vi2")
                nat_bc2  = st.selectbox("NAT_BC_CRED",["06","01","02","03","04","05","07"],key="nb2")
            with c3:
                aliq_cof = st.number_input("ALIQ_COFINS (%)",0.0,100.0,7.6,0.01)
                vl_bc_c  = st.number_input("VL_BC_COFINS",0.0,step=.01,format="%.2f")
            vl_cof_calc = round(vl_bc_c * aliq_cof / 100, 4)
            st.info(f"COFINS calculado: R$ {vl_cof_calc:.4f}")
            if st.form_submit_button("💾 Adicionar C505", type="primary"):
                _adicionar_reg("C505","C",{
                    "CST_COFINS":cst_cof,"VL_ITEM":str(round(vl_item2,2)),
                    "NAT_BC_CRED":nat_bc2,"VL_BC_COFINS":str(round(vl_bc_c,2)),
                    "ALIQ_COFINS":str(aliq_cof),"VL_COFINS":str(vl_cof_calc),
                })
                st.success("✅ C505 adicionado!"); st.rerun()

    with st.expander("ℹ️ Regras para Crédito de Energia Elétrica"):
        st.markdown("""
**Lei 10.833/2003, art. 3º, II** — crédito sobre energia elétrica consumida nos estabelecimentos.

| Registro | Finalidade |
|----------|-----------|
| C500 | Nota fiscal de serviço de energia elétrica (mod. 06) |
| C501 | Detalhe PIS — discrimina CST, base, alíquota e valor |
| C505 | Detalhe COFINS — idem para COFINS |

**NAT_BC_CRED mais usados para energia:**
- `06` = Energia Elétrica e Térmica (inc. cogeração)
- `99` = Outras (quando não se enquadra acima)

**Alíquotas (não-cumulativo):** PIS = 1,65% · COFINS = 7,6%
""")


# =============================================================================
# ── PÁGINA: ALUGUEL (F100) ────────────────────────────────────────────────────
# =============================================================================
def pg_aluguel():
    st.markdown("""<div class="sped-header">
        <h1>🏢 Crédito de Aluguel — F100</h1>
        <p>EFD Contribuições: crédito PIS/COFINS sobre aluguéis e arrendamentos.</p>
    </div>""", unsafe_allow_html=True)

    df = st.session_state.df_editado
    if df is not None and "REG" in df.columns:
        df_f100 = df[df["REG"]=="F100"].copy()
        if not df_f100.empty:
            cols = [c for c in ["LINHA_NUM","COD_PART","DT_OPER","VL_OPER",
                                  "CST_PIS","ALIQ_PIS","VL_PIS",
                                  "CST_COFINS","ALIQ_COFINS","VL_COFINS",
                                  "NAT_BC_CRED","DESC_DOC_OPR"] if c in df_f100.columns]
            st.markdown(f"**{len(df_f100)} registros F100 existentes:**")
            st.dataframe(df_f100[cols], use_container_width=True, height=180)

    st.markdown("#### ➕ Adicionar Registro F100")
    with st.form("form_f100"):
        c1,c2,c3 = st.columns(3)
        with c1:
            ind_oper = st.selectbox("IND_OPER",["0 - Entrada","1 - Saída"])
            cod_part = st.text_input("COD_PART (CNPJ locador)","")
            cod_item = st.text_input("COD_ITEM","ALUGUEL")
            dt_oper  = st.text_input("DT_OPER (DDMMAAAA)",datetime.now().strftime("%d%m%Y"))
        with c2:
            vl_oper  = st.number_input("VL_OPER (valor aluguel)",0.0,step=.01,format="%.2f")
            cst_pis  = st.selectbox("CST_PIS",["50","51","52","53","54","55","56","70","71","72","73","74","75","98","99"])
            aliq_pis = st.number_input("ALIQ_PIS (%)",0.0,100.0,1.65,0.01)
            nat_bc   = st.selectbox("NAT_BC_CRED",["06","01","02","03","04","05","07"],
                                    help="06=Edificações, 05=Máquinas e Equipamentos")
        with c3:
            cst_cof  = st.selectbox("CST_COFINS",["50","51","52","53","54","55","56","70","71","72","73","74","75","98","99"])
            aliq_cof = st.number_input("ALIQ_COFINS (%)",0.0,100.0,7.6,0.01)
            cod_cta  = st.text_input("COD_CTA","")
            desc_doc = st.text_input("DESC_DOC_OPR","Crédito de aluguel de imóvel")

        vl_pis_c = round(vl_oper * aliq_pis / 100, 2)
        vl_cof_c = round(vl_oper * aliq_cof / 100, 2)
        st.markdown(f"""
        <div style="background:#1C4532;border-radius:6px;padding:.8rem;margin:.5rem 0;">
            <b>Prévia:</b> &nbsp;
            PIS = R$ {vl_pis_c:,.2f} ({aliq_pis}% × R$ {vl_oper:,.2f}) &nbsp;|&nbsp;
            COFINS = R$ {vl_cof_c:,.2f} ({aliq_cof}% × R$ {vl_oper:,.2f})
        </div>""", unsafe_allow_html=True)

        if st.form_submit_button("💾 Adicionar F100", type="primary"):
            _adicionar_reg("F100","F",{
                "IND_OPER":ind_oper[0],"COD_PART":cod_part,"COD_ITEM":cod_item,
                "DT_OPER":dt_oper,"VL_OPER":str(round(vl_oper,2)),
                "CST_PIS":cst_pis,"VL_BC_PIS":str(round(vl_oper,2)),
                "ALIQ_PIS":str(aliq_pis),"VL_PIS":str(vl_pis_c),
                "CST_COFINS":cst_cof,"VL_BC_COFINS":str(round(vl_oper,2)),
                "ALIQ_COFINS":str(aliq_cof),"VL_COFINS":str(vl_cof_c),
                "NAT_BC_CRED":nat_bc,"IND_ORIG_CRED":"0",
                "COD_CTA":cod_cta,"COD_CCUS":"","DESC_DOC_OPR":desc_doc,
            })
            st.success(f"✅ F100 adicionado! PIS: R$ {vl_pis_c:,.2f} | COFINS: R$ {vl_cof_c:,.2f}")
            st.rerun()

    with st.expander("ℹ️ Regras para Crédito de Aluguel"):
        st.markdown("""
**Fundamento:** Lei 10.637/2002 art. 3º, IV (PIS) e Lei 10.833/2003 art. 3º, IV (COFINS).

Aplica-se a imóveis utilizados nas atividades da empresa (regime não-cumulativo).

| CST | Uso típico |
|-----|-----------|
| 50  | Crédito vinculado a receitas tributadas MI |
| 51  | Crédito vinculado a receitas não-tributadas MI |
| 52  | Crédito vinculado a receitas de exportação |
| 53–56 | Créditos mistos |
| 70–75 | Sem direito a crédito |

**NAT_BC_CRED:**
- `06` = Aluguéis de imóveis utilizados nas atividades
- `05` = Máquinas, equipamentos e outros bens do ativo fixo

**Alíquotas não-cumulativo padrão:** PIS 1,65% · COFINS 7,6%
""")


# =============================================================================
# ── PÁGINA: MOTOR DE REGRAS ───────────────────────────────────────────────────
# =============================================================================
def pg_regras():
    st.markdown("""<div class="sped-header">
        <h1>⚙️ Motor de Regras Tributárias</h1>
        <p>Configure regras CST × CFOP para ICMS, PIS e COFINS. Editável, importável e exportável.</p>
    </div>""", unsafe_allow_html=True)

    regras = st.session_state.regras
    df_r = regras_para_df(regras)

    c1,c2 = st.columns(2)
    with c1: f_t = st.multiselect("Tributo",df_r["tributo"].unique().tolist(),default=df_r["tributo"].unique().tolist())
    with c2: f_b = st.text_input("Buscar CST ou descrição","")

    df_f = df_r[df_r["tributo"].isin(f_t)]
    if f_b: df_f = df_f[df_f["cst"].str.contains(f_b,na=False)|df_f["descricao"].str.contains(f_b,case=False,na=False)]

    df_ed = st.data_editor(
        df_f, use_container_width=True, height=380, num_rows="dynamic", key="ed_regras",
        column_config={
            "tributo":        st.column_config.SelectboxColumn("Tributo",options=["ICMS","PIS","COFINS","IPI"]),
            "exige_base":     st.column_config.CheckboxColumn("Exige Base"),
            "exige_aliquota": st.column_config.CheckboxColumn("Exige Alíquota"),
            "exige_imposto":  st.column_config.CheckboxColumn("Exige Imposto"),
        }
    )

    c1,c2,c3 = st.columns(3)
    with c1:
        if st.button("💾 Salvar Regras",type="primary"):
            df_nf = df_r[~df_r["tributo"].isin(f_t)]
            st.session_state.regras = df_para_regras(pd.concat([df_ed,df_nf],ignore_index=True))
            st.success("✅ Regras salvas!"); st.rerun()
    with c2:
        if st.button("🔄 Restaurar Padrão"):
            st.session_state.regras = list(TODAS_REGRAS_PADRAO)
            st.success("✅ Padrão restaurado."); st.rerun()
    with c3:
        csv_r = df_r.to_csv(index=False,sep=";").encode()
        st.download_button("⬇️ Exportar CSV", csv_r, "regras.csv","text/csv")

    with st.expander("📤 Importar Regras"):
        arq_r = st.file_uploader("CSV de regras",type=["csv"])
        if arq_r and st.button("Importar"):
            st.session_state.regras = df_para_regras(pd.read_csv(arq_r,sep=";"))
            st.success("✅ Importado!"); st.rerun()

    with st.expander("🧪 Testador de Cálculo"):
        tc1,tc2,tc3,tc4,tc5 = st.columns(5)
        t_tri = tc1.selectbox("Tributo",["ICMS","PIS","COFINS"])
        t_cst = tc2.text_input("CST","00")
        t_bc  = tc3.number_input("Base (R$)",0.0,step=.01)
        t_al  = tc4.number_input("Alíquota (%)",0.0,step=.01)
        if tc5.button("▶️ Calcular"):
            rg = buscar_regra(st.session_state.regras,t_tri,t_cst)
            if rg:
                res = calcular_imposto(rg,str(t_bc),str(t_al))
                st.success(f"**R$ {res}** — Fórmula: `{rg.formula}`")
            else:
                st.error(f"Regra não encontrada: {t_tri} / CST {t_cst}")


# =============================================================================
# ── PÁGINA: EXPORTAÇÃO ────────────────────────────────────────────────────────
# =============================================================================
def pg_exportacao():
    df = st.session_state.df_editado
    if df is None:
        st.info("Carregue um arquivo SPED primeiro.")
        return
    r = st.session_state.resultado
    st.markdown("""<div class="sped-header">
        <h1>📥 Exportação</h1>
        <p>SPED TXT corrigido, Excel de auditoria (5 abas) e CSV de inconsistências.</p>
    </div>""", unsafe_allow_html=True)

    df_incs = st.session_state.df_incs
    df_log  = st.session_state.df_log
    df_orig = st.session_state.df_original

    n_c = len(df_incs[df_incs["TIPO"]=="CRITICO"]) if not df_incs.empty else 0
    if n_c > 0:
        st.warning(f"⚠️ Ainda há **{n_c}** inconsistências críticas.")
    else:
        st.success("✅ Arquivo sem críticos — pronto para exportar.")

    try:
        mask_alt = (df.fillna("")!=df_orig.fillna("")).any(axis=1) if df_orig is not None else pd.Series([False]*len(df))
        df_alt = df[mask_alt]
    except Exception:
        df_alt = pd.DataFrame()

    st.markdown('<div class="kpi-grid">'+
        _kpi(n_c,"Críticos pendentes","critico" if n_c>0 else "ok")+
        _kpi(len(df_alt),"Registros alterados")+
        _kpi(len(df_log),"Entradas de log")+
    '</div>',unsafe_allow_html=True)

    col1,col2,col3 = st.columns(3)

    with col1:
        st.markdown("#### 📄 SPED TXT Corrigido")
        if r and r.linhas_raw:
            txt_b = reconstruir_txt(r.linhas_raw, df).encode("utf-8")
        else:
            linhas = []
            for _, row in df.sort_values("LINHA_NUM").iterrows():
                reg = str(row.get("REG",""))
                campos = CAMPOS_REGISTRO.get(reg,[])
                if campos:
                    vals = [str(row.get(c,"")) for c in campos if c in row.index]
                else:
                    vals = [str(v) for k,v in row.items() if k not in ("LINHA_NUM","BLOCO") and pd.notna(v)]
                linhas.append("|"+"|".join(vals)+"|")
            txt_b = "\n".join(linhas).encode("utf-8")
        st.download_button("⬇️ Baixar TXT", txt_b,
            f"sped_corrigido_{datetime.now().strftime('%Y%m%d_%H%M')}.txt","text/plain",type="primary")

    with col2:
        st.markdown("#### 📊 Excel de Auditoria")
        info = {
            "nome_empresa": r.nome_empresa if r else "",
            "cnpj":         r.cnpj if r else "",
            "ie":           r.ie if r else "",
            "uf":           r.uf if r else "",
            "dt_ini":       r.dt_ini if r else "",
            "dt_fin":       r.dt_fin if r else "",
            "tipo_arquivo": r.tipo_arquivo if r else "",
        }
        excel_b = exportar_excel_auditoria(df_incs, df_alt, df_log,
                                            regras_para_df(st.session_state.regras), info)
        st.download_button("⬇️ Baixar Excel",excel_b,
            f"auditoria_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")

    with col3:
        st.markdown("#### 📋 CSV Inconsistências")
        if not df_incs.empty:
            st.download_button("⬇️ Baixar CSV",exportar_csv_incs(df_incs),
                f"inconsistencias_{datetime.now().strftime('%Y%m%d_%H%M')}.csv","text/csv")
        else:
            st.info("Sem inconsistências.")

    if not df_log.empty:
        st.markdown("#### 📋 Log de Auditoria")
        st.dataframe(df_log, use_container_width=True, height=200)
        csv_l = df_log.to_csv(index=False,sep=";",encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button("⬇️ Log CSV", csv_l, "log_auditoria.csv","text/csv")


# =============================================================================
# ── PÁGINA: LOG DE AUDITORIA ─────────────────────────────────────────────────
# =============================================================================
def pg_log():
    st.markdown("""<div class="sped-header">
        <h1>🔍 Log de Auditoria</h1>
        <p>Rastreabilidade completa: data/hora, analista, campo, antes/depois, regra aplicada.</p>
    </div>""", unsafe_allow_html=True)

    df_log = st.session_state.df_log
    if df_log.empty:
        st.info("Nenhuma alteração registrada nesta sessão.")
        return

    c1,c2 = st.columns(2)
    with c1: f_r = st.multiselect("Registro",df_log["REGISTRO"].unique().tolist(),default=df_log["REGISTRO"].unique().tolist())
    with c2: f_c = st.multiselect("Campo",    df_log["CAMPO"].unique().tolist(),   default=df_log["CAMPO"].unique().tolist())

    df_f = df_log[df_log["REGISTRO"].isin(f_r)&df_log["CAMPO"].isin(f_c)]
    st.markdown(f"**{len(df_f)} entradas**")
    st.dataframe(df_f, use_container_width=True, height=400)

    csv_l = df_f.to_csv(index=False,sep=";",encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button("⬇️ Exportar Log","".join([csv_l.decode("utf-8-sig")]).encode("utf-8-sig"),
                        "log_auditoria.csv","text/csv")

    if st.button("🗑️ Limpar Log"):
        st.session_state.df_log = pd.DataFrame(columns=st.session_state.df_log.columns)
        st.rerun()


# =============================================================================
# ── ROTEADOR PRINCIPAL ────────────────────────────────────────────────────────
# =============================================================================
ROTAS = {
    "upload":          pg_upload,
    "dashboard":       pg_dashboard,
    "blocos":          pg_blocos,
    "notas":           pg_notas,
    "inconsistencias": pg_inconsistencias,
    "correcoes":       pg_correcoes,
    "editor":          pg_editor,
    "cte":             pg_cte,
    "energia":         pg_energia,
    "aluguel":         pg_aluguel,
    "regras":          pg_regras,
    "exportacao":      pg_exportacao,
    "log":             pg_log,
}


def main():
    _sidebar()
    fn = ROTAS.get(st.session_state.pagina, pg_upload)
    try:
        fn()
    except Exception as e:
        st.error(f"❌ Erro: {e}")
        logger.exception(f"Erro na página {st.session_state.pagina}")
        if st.button("🔄 Recarregar"):
            st.rerun()


if __name__ == "__main__":
    main()
