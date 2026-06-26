"""
╔═══════════════════════════════════════════════════════════════════════════════╗
║          SPED AUDITOR — FISCAL INTELLIGENCE PLATFORM  v3.1                  ║
║          EFD ICMS/IPI  +  EFD CONTRIBUIÇÕES (PIS/COFINS)                    ║
║          Arquivo único · Python + Streamlit                                  ║
║          Compatível com Streamlit 1.58+ / Python 3.14+                      ║
╚═══════════════════════════════════════════════════════════════════════════════╝

Execução:
    pip install streamlit pandas numpy openpyxl plotly
    streamlit run sped_auditor.py
"""

# ═══════════════════════════════════════════════════════════════════════════════
# IMPORTS
# ═══════════════════════════════════════════════════════════════════════════════

from __future__ import annotations

import io
import json
import os
import re
import sys
import uuid
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field, asdict
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Optional, List, Dict, Any, Tuple, Callable
from functools import lru_cache
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO
# ═══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="SPED Auditor — EFD ICMS/IPI + Contribuições",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={"About": "SPED Auditor v3.1 | EFD ICMS/IPI + EFD Contribuições + CT-e"},
)

# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  CONSTANTES E UTILITÁRIOS  █████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

TIPO_EFD_ICMS = "EFD_ICMS_IPI"
TIPO_EFD_CONTRIB = "EFD_CONTRIBUICOES"

CAMPOS_NUMERICOS = {
    "VL_ITEM", "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS",
    "VL_BC_ICMS_ST", "ALIQ_ST", "VL_ICMS_ST",
    "VL_IPI", "VL_BC_IPI", "ALIQ_IPI",
    "VL_BC_PIS", "ALIQ_PIS", "VL_PIS", "QUANT_BC_PIS",
    "VL_BC_COFINS", "ALIQ_COFINS", "VL_COFINS", "QUANT_BC_COFINS",
    "VL_DOC", "VL_MERC", "VL_DESC", "VL_OPER", "QTD",
    "VL_CRED", "VL_CONT_APUR", "VL_CONT_PER", "VL_REC_BRT", "VL_BC_CONT",
    "VL_ICMS_RECOLHER", "VL_TOT_DEBITOS", "VL_TOT_CREDITOS",
}

CORES = {
    "primaria": "#0F2540",
    "secundaria": "#1A6BAD",
    "critico": "#C0392B",
    "aviso": "#B7860D",
    "ok": "#1A7A35",
    "info": "#2980B9",
    "fundo": "#F7FAFF",
    "borda": "#E0E8F4",
    "texto_claro": "#FFFFFF",
    "texto_escuro": "#0F2540",
}

# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  UTILITÁRIOS  █████████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

class Utilitarios:
    """Classe com funções utilitárias reutilizáveis."""
    
    @staticmethod
    def decode_bytes(conteudo: bytes) -> str:
        """Decodifica bytes com múltiplos encodings."""
        for enc in ("utf-8", "latin-1", "cp1252"):
            try:
                return conteudo.decode(enc)
            except UnicodeDecodeError:
                continue
        raise ValueError("Encoding não reconhecido.")
    
    @staticmethod
    def to_float(valor: Any) -> Optional[float]:
        """Converte valor para float de forma segura."""
        if valor is None:
            return None
        if isinstance(valor, (int, float)):
            import math
            return None if (isinstance(valor, float) and math.isnan(valor)) else float(valor)
        try:
            s = str(valor).strip().replace(".", "").replace(",", ".")
            return None if not s else float(Decimal(s))
        except (ValueError, InvalidOperation):
            return None
    
    @staticmethod
    def to_sped_str(valor: float) -> str:
        """Converte float para string no formato SPED (vírgula decimal)."""
        if valor is None:
            return ""
        try:
            d = Decimal(str(valor)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            return str(d).replace(".", ",")
        except:
            return ""
    
    @staticmethod
    def fmt_brl(valor: Optional[float]) -> str:
        """Formata valor para exibição em reais."""
        if valor is None:
            return "—"
        try:
            d = Decimal(str(valor)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            i, dec = str(d).split(".")
            fmt = ""
            for j, c in enumerate(reversed(i)):
                if j and j % 3 == 0 and c != "-":
                    fmt = "." + fmt
                fmt = c + fmt
            return f"R$ {fmt},{dec}"
        except:
            return str(valor)
    
    @staticmethod
    def formatar_data_sped(data_iso: str) -> str:
        """Converte data ISO para formato DDMMAAAA do SPED."""
        if not data_iso:
            return ""
        data_iso = data_iso[:10]
        partes = data_iso.split("-")
        if len(partes) == 3:
            return partes[2] + partes[1] + partes[0]
        return ""
    
    @staticmethod
    def extrair_namespace(root: ET.Element) -> str:
        """Extrai namespace do elemento XML."""
        tag = root.tag
        if tag.startswith("{"):
            return tag[1:tag.index("}")]
        return ""


# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  MODELOS DE DADOS  █████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class LinhaRegistro:
    """Representa uma linha do arquivo SPED."""
    numero_linha: int
    bloco: str
    registro: str
    campos: List[str]
    linha_original: str
    
    def get(self, indice: int, padrao: str = "") -> str:
        try:
            return self.campos[indice] if indice < len(self.campos) else padrao
        except IndexError:
            return padrao
    
    def set(self, indice: int, valor: str) -> None:
        while len(self.campos) <= indice:
            self.campos.append("")
        self.campos[indice] = valor
    
    def to_sped(self) -> str:
        return "|" + "|".join(self.campos) + "|\n"


@dataclass
class MetadadosArquivo:
    tipo_escrituracao: str = ""
    periodo_apuracao: str = ""
    cnpj: str = ""
    nome_empresa: str = ""
    uf: str = ""
    ie: str = ""
    cod_municipio: str = ""
    ind_perfil: str = ""
    total_linhas: int = 0
    blocos_presentes: List[str] = field(default_factory=list)


@dataclass
class ResultadoParser:
    metadados: MetadadosArquivo
    linhas: List[LinhaRegistro]
    df: pd.DataFrame
    erros: List[str]
    tipo_arquivo: str


@dataclass
class RegraFiscal:
    """Regra tributária configurável."""
    id: str
    tributo: str
    cst: str
    cfop: str
    ind_oper: str
    descricao: str
    exige_base: bool
    exige_aliquota: bool
    exige_valor: bool
    base_campo_sugerido: str
    aliquota_padrao: Optional[float]
    formula: str
    permite_base_zero: bool
    permite_valor_zero: bool
    criticidade: str
    ativa: bool = True
    cod_mod: str = "*"
    
    def match_cst(self, valor: str) -> bool:
        return self.cst == "*" or valor == self.cst or valor.startswith(self.cst)
    
    def match_cfop(self, valor: str) -> bool:
        return self.cfop == "*" or valor == self.cfop or valor.startswith(self.cfop)
    
    def match_oper(self, valor: str) -> bool:
        return self.ind_oper in ("*", valor)
    
    def match_mod(self, valor: str) -> bool:
        if self.cod_mod == "*":
            return True
        if not valor:
            return self.cod_mod == "*"
        return valor == self.cod_mod


@dataclass
class EntradaAuditoria:
    id: str
    timestamp: str
    usuario: str
    numero_linha: int
    registro: str
    bloco: str
    campo: str
    valor_anterior: str
    valor_novo: str
    regra: str
    motivo: str
    tipo: str


@dataclass
class DadosCTe:
    chave: str = ""
    num_doc: str = ""
    serie: str = ""
    modelo: str = "57"
    dt_emissao: str = ""
    dt_acesso: str = ""
    cod_sit: str = "00"
    tipo_cte: str = "0"
    ind_oper: str = "0"
    ind_emit: str = "1"
    cnpj_emit: str = ""
    ie_emit: str = ""
    nome_emit: str = ""
    vl_doc: str = ""
    vl_serv: str = ""
    vl_desc: str = ""
    cst_icms: str = ""
    vl_bc_icms: str = ""
    aliq_icms: str = ""
    vl_icms: str = ""
    cst_pis: str = ""
    vl_bc_pis: str = ""
    aliq_pis: str = ""
    vl_pis: str = ""
    cst_cofins: str = ""
    vl_bc_cofins: str = ""
    aliq_cofins: str = ""
    vl_cofins: str = ""
    ind_nat_frt: str = "0"
    erros: List[str] = field(default_factory=list)


# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  MAPA DE CAMPOS  ███████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

class MapaCampos:
    """Gerencia o mapeamento de campos para cada registro SPED."""
    
    _MAPAS: Dict[str, Dict[str, int]] = {
        "0000": {"COD_VER": 1, "COD_FIN": 2, "DT_INI": 3, "DT_FIN": 4,
                 "NOME": 5, "CNPJ": 6, "CPF": 7, "UF": 8, "IE": 9,
                 "COD_MUN": 10, "IND_PERFIL": 13, "IND_OPORT": 14},
        "0150": {"COD_PART": 1, "NOME": 2, "COD_PAIS": 3, "CNPJ": 4, "CPF": 5,
                 "IE": 6, "COD_MUN": 7, "END": 9, "BAIRRO": 12},
        "0200": {"COD_ITEM": 1, "DESCR_ITEM": 2, "COD_BARRA": 3, "UNID_INV": 5,
                 "TIPO_ITEM": 6, "COD_NCM": 7, "ALIQ_ICMS": 11},
        "C100": {"IND_OPER": 1, "IND_EMIT": 2, "COD_PART": 3, "COD_MOD": 4,
                 "COD_SIT": 5, "SER": 6, "NUM_DOC": 7, "CHV_NFE": 8,
                 "DT_DOC": 9, "DT_E_S": 10, "VL_DOC": 11,
                 "VL_DESC": 13, "VL_MERC": 15,
                 "VL_BC_ICMS": 20, "VL_ICMS": 21,
                 "VL_BC_ICMS_ST": 22, "VL_ICMS_ST": 23,
                 "VL_IPI": 24, "VL_PIS": 25, "VL_COFINS": 26},
        "C170": {"NUM_ITEM": 1, "COD_ITEM": 2, "DESCR_COMPL": 3,
                 "QTD": 4, "UNID": 5, "VL_ITEM": 6, "VL_DESC": 7, "IND_MOV": 8,
                 "CST_ICMS": 9, "CFOP": 10,
                 "VL_BC_ICMS": 12, "ALIQ_ICMS": 13, "VL_ICMS": 14,
                 "VL_BC_ICMS_ST": 15, "ALIQ_ST": 16, "VL_ICMS_ST": 17,
                 "CST_IPI": 19, "VL_BC_IPI": 20, "ALIQ_IPI": 21, "VL_IPI": 22,
                 "CST_PIS": 23, "VL_BC_PIS": 24, "ALIQ_PIS": 25, "VL_PIS": 26,
                 "CST_COFINS": 27, "VL_BC_COFINS": 28, "ALIQ_COFINS": 29, "VL_COFINS": 30},
        "C190": {"CST_ICMS": 1, "CFOP": 2, "ALIQ_ICMS": 3,
                 "VL_OPR": 4, "VL_BC_ICMS": 5, "VL_ICMS": 6,
                 "VL_BC_ICMS_ST": 7, "VL_ICMS_ST": 8},
        "D100": {"IND_OPER": 1, "COD_PART": 3, "COD_MOD": 4, "COD_SIT": 5,
                 "NUM_DOC": 8, "DT_DOC": 10, "VL_DOC": 14,
                 "VL_BC_ICMS": 18, "ALIQ_ICMS": 19, "VL_ICMS": 20,
                 "CST_PIS": 21, "VL_BC_PIS": 22, "ALIQ_PIS": 23, "VL_PIS": 24,
                 "CST_COFINS": 25, "VL_BC_COFINS": 26, "ALIQ_COFINS": 27, "VL_COFINS": 28},
        "D190": {"CST_ICMS": 1, "CFOP": 2, "ALIQ_ICMS": 3,
                 "VL_OPR": 4, "VL_BC_ICMS": 5, "VL_ICMS": 6},
        "E110": {"VL_TOT_DEBITOS": 1, "VL_AJ_DEBITOS": 2,
                 "VL_TOT_CREDITOS": 5, "VL_AJ_CREDITOS": 6,
                 "VL_SLD_APURADO": 10, "VL_ICMS_RECOLHER": 12},
        "E116": {"COD_OR": 1, "VL_OR": 2, "DT_VCTO": 3, "COD_REC": 4},
        "A100": {"IND_OPER": 1, "IND_EMIT": 2, "COD_PART": 3, "COD_SIT": 4,
                 "SER": 5, "SUB": 6, "NUM_DOC": 7, "CHV_NFSE": 8,
                 "DT_DOC": 9, "DT_EXE_SERV": 10, "VL_DOC": 11,
                 "VL_DESC": 12, "VL_BC_PIS": 13, "ALIQ_PIS": 14, "VL_PIS": 15,
                 "VL_BC_COFINS": 16, "ALIQ_COFINS": 17, "VL_COFINS": 18},
        "A170": {"NUM_ITEM": 1, "COD_ITEM": 2, "DESCR_COMPL": 3,
                 "VL_ITEM": 4, "VL_DESC": 5, "NAT_BC_CRED": 6, "IND_ORIG_CRED": 7,
                 "CST_PIS": 8, "VL_BC_PIS": 9, "ALIQ_PIS": 10, "VL_PIS": 11,
                 "CST_COFINS": 12, "VL_BC_COFINS": 13, "ALIQ_COFINS": 14, "VL_COFINS": 15,
                 "COD_CTA": 16},
        "C010": {"IND_ESCRI": 1},
        "C180": {"COD_CRED": 1, "IND_ORIG_CRED": 2, "VL_BC_PIS": 3, "ALIQ_PIS": 4,
                 "QUANT_BC_PIS": 5, "VL_PIS": 6, "VL_BC_COFINS": 7, "ALIQ_COFINS": 8,
                 "QUANT_BC_COFINS": 9, "VL_COFINS": 10, "COD_CTA": 11},
        "C181": {"COD_CRED": 1, "IND_ORIG_CRED": 2, "CNPJ_CPF_PART": 3,
                 "COD_MOD": 4, "DT_OPER": 5, "CHV_NFE": 6, "NUM_DOC": 7,
                 "VL_OPER": 8, "CFOP": 9, "NAT_BC_CRED": 10,
                 "VL_BC_PIS": 12, "ALIQ_PIS": 13, "VL_PIS": 14,
                 "VL_BC_COFINS": 15, "ALIQ_COFINS": 16, "VL_COFINS": 17},
        "C185": {"NUM_ITEM": 1, "COD_ITEM": 2, "CST_PIS": 3, "COD_CRED": 4,
                 "VL_BC_PIS": 5, "ALIQ_PIS": 6, "VL_PIS": 7,
                 "CST_COFINS": 8, "VL_BC_COFINS": 9, "ALIQ_COFINS": 10, "VL_COFINS": 11},
        "C380": {"COD_MOD": 1, "DT_DOC_INI": 2, "DT_DOC_FIN": 3, "NUM_DOC_INI": 4,
                 "NUM_DOC_FIN": 5, "VL_DOC": 6, "VL_DOC_CANC": 7},
        "C481": {"CST_PIS": 1, "VL_ITEM": 2, "VL_BC_PIS": 3, "ALIQ_PIS": 4,
                 "QUANT_BC_PIS": 5, "VL_PIS": 6, "COD_CTA": 7},
        "C485": {"CST_COFINS": 1, "VL_ITEM": 2, "VL_BC_COFINS": 3, "ALIQ_COFINS": 4,
                 "QUANT_BC_COFINS": 5, "VL_COFINS": 6, "COD_CTA": 7},
        "D101": {"IND_NAT_FRT": 1, "VL_ITEM": 2, "CST_PIS": 3, "NAT_BC_CRED": 4,
                 "VL_BC_PIS": 5, "ALIQ_PIS": 6, "VL_PIS": 7, "COD_CTA": 8},
        "D105": {"IND_NAT_FRT": 1, "VL_ITEM": 2, "CST_COFINS": 3, "NAT_BC_CRED": 4,
                 "VL_BC_COFINS": 5, "ALIQ_COFINS": 6, "VL_COFINS": 7, "COD_CTA": 8},
        "F100": {"IND_OPER": 1, "COD_PART": 2, "DT_OPER": 3, "VL_OPER": 4,
                 "COD_CRED": 5, "IND_ORIG_CRED": 6, "VL_BC_PIS": 7, "ALIQ_PIS": 8,
                 "VL_PIS": 9, "VL_BC_COFINS": 10, "ALIQ_COFINS": 11, "VL_COFINS": 12},
        "F120": {"NAT_DESPESA": 1, "VL_AQUISICOES": 2, "VL_PARCELA": 3,
                 "VL_BC_PIS": 4, "ALIQ_PIS": 5, "VL_PIS": 6,
                 "VL_BC_COFINS": 7, "ALIQ_COFINS": 8, "VL_COFINS": 9},
        "F130": {"IND_ORIG_CRED": 1, "IND_UTIL_BENS": 2, "VL_OPER": 3,
                 "PARC_OPER": 4, "VL_BC_PIS": 5, "ALIQ_PIS": 6, "VL_PIS": 7,
                 "VL_BC_COFINS": 8, "ALIQ_COFINS": 9, "VL_COFINS": 10},
        "F150": {"NAT_BC_CRED": 1, "VL_TOT_EST": 2, "VL_BC_PIS": 3, "ALIQ_PIS": 4,
                 "VL_PIS": 5, "VL_BC_COFINS": 6, "ALIQ_COFINS": 7, "VL_COFINS": 8},
        "F200": {"IND_OPER": 1, "COD_PART": 2, "COD_ITEM": 3, "DT_OPER": 4,
                 "VL_OPER": 5, "CST_PIS": 6, "VL_BC_PIS": 7, "ALIQ_PIS": 8,
                 "VL_PIS": 9, "CST_COFINS": 10, "VL_BC_COFINS": 11, "ALIQ_COFINS": 12, "VL_COFINS": 13},
        "M100": {"COD_CRED": 1, "IND_CRED_ORI": 2, "VL_BC_PIS": 3, "ALIQ_PIS": 4,
                 "QUANT_BC_PIS": 5, "VL_CRED": 6, "VL_AJUS_ACRES": 7, "VL_AJUS_REDUC": 8,
                 "VL_CRED_DIF": 9, "VL_CRED_DISP": 10, "IND_DESC_CRED": 11, "VL_CRED_DESC": 12},
        "M110": {"IND_AJ": 1, "VL_AJ": 2, "COD_AJ": 3, "NUM_DOC": 4, "DESCR_AJ": 5, "DT_REF": 6},
        "M200": {"VL_TOT_CONT_NC_PER": 1, "VL_TOT_CRED_DESC": 2, "VL_TOT_CRED_DESC_ANT": 3,
                 "VL_TOT_CONT_NC_DEV": 4, "VL_RET_NC": 5, "VL_OUT_DED_NC": 6,
                 "VL_CONT_NC_REC": 7, "VL_TOT_CONT_CUM_PER": 8, "VL_RET_CUM": 9,
                 "VL_OUT_DED_CUM": 10, "VL_CONT_CUM_REC": 11, "VL_TOT_CONT_REC": 12},
        "M210": {"COD_CONT": 1, "VL_REC_BRT": 2, "VL_BC_CONT": 3, "ALIQ_PIS": 4,
                 "QUANT_BC_PIS": 5, "VL_CONT_APUR": 6, "VL_AJUS_ACRES": 7, "VL_AJUS_REDUC": 8,
                 "VL_CONT_DIFER": 9, "VL_CONT_DIFER_ANT": 10, "VL_CONT_PER": 11},
        "M500": {"COD_CRED": 1, "IND_CRED_ORI": 2, "VL_BC_COFINS": 3, "ALIQ_COFINS": 4,
                 "QUANT_BC_COFINS": 5, "VL_CRED": 6, "VL_AJUS_ACRES": 7, "VL_AJUS_REDUC": 8,
                 "VL_CRED_DIF": 9, "VL_CRED_DISP": 10, "IND_DESC_CRED": 11, "VL_CRED_DESC": 12},
        "M600": {"VL_TOT_CONT_NC_PER": 1, "VL_TOT_CRED_DESC": 2, "VL_TOT_CRED_DESC_ANT": 3,
                 "VL_TOT_CONT_NC_DEV": 4, "VL_RET_NC": 5, "VL_OUT_DED_NC": 6,
                 "VL_CONT_NC_REC": 7, "VL_TOT_CONT_CUM_PER": 8, "VL_RET_CUM": 9,
                 "VL_OUT_DED_CUM": 10, "VL_CONT_CUM_REC": 11, "VL_TOT_CONT_REC": 12},
        "M610": {"COD_CONT": 1, "VL_REC_BRT": 2, "VL_BC_CONT": 3, "ALIQ_COFINS": 4,
                 "QUANT_BC_COFINS": 5, "VL_CONT_APUR": 6, "VL_AJUS_ACRES": 7, "VL_AJUS_REDUC": 8,
                 "VL_CONT_DIFER": 9, "VL_CONT_DIFER_ANT": 10, "VL_CONT_PER": 11},
        "P001": {"IND_MOV": 1},
        "P010": {"CNPJ": 1},
        "P100": {"DT_INI": 1, "DT_FIN": 2, "VL_REC_TOT_EST": 3, "COD_ATIV_ECON": 4,
                 "VL_REC_ATIV_ESTAB": 5, "VL_EXC": 6, "VL_REC_BC": 7, "ALIQ_CONTRIB_APUR": 8,
                 "VL_CONTRIB_APUR": 9, "VL_CONT_RECOL": 10},
        "P110": {"NUM_CAMPO": 1, "COD_DET": 2, "DET_VALOR": 3},
        "9900": {"REG": 1, "QTD": 2},
        "9999": {"QTD": 1},
    }
    
    @classmethod
    def get_mapa(cls, registro: str) -> Dict[str, int]:
        return cls._MAPAS.get(registro, {})
    
    @classmethod
    def get_campos(cls, registro: str) -> List[str]:
        return list(cls.get_mapa(registro).keys())
    
    @classmethod
    def get_indice(cls, registro: str, campo: str) -> Optional[int]:
        mapa = cls.get_mapa(registro)
        return mapa.get(campo)


# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  REGRAS FISCAIS  ███████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

class GerenciadorRegras:
    """Gerencia o catálogo de regras fiscais."""
    
    ARQUIVO_REGRA = "sped_regras_fiscais.json"
    
    REGRAS_ICMS = [
        ("ICMS", "000", "*", "Tributado Integralmente", True, True, True, "VL_ITEM", 12.0, "critica"),
        ("ICMS", "010", "*", "Tributado + ST", True, True, True, "VL_ITEM", 12.0, "critica"),
        ("ICMS", "020", "*", "Tributado c/ Redução BC", True, True, True, "VL_ITEM", 12.0, "critica"),
        ("ICMS", "030", "*", "Isento c/ ST", False, False, False, "", 0.0, "aviso"),
        ("ICMS", "040", "*", "Isento", False, False, False, "", 0.0, "aviso"),
        ("ICMS", "041", "*", "Não Tributado", False, False, False, "", 0.0, "aviso"),
        ("ICMS", "050", "*", "Suspensão", False, False, False, "", 0.0, "aviso"),
        ("ICMS", "051", "*", "Diferimento", False, False, False, "", 0.0, "info"),
        ("ICMS", "060", "*", "ST Cobrado Anteriormente", False, False, False, "", 0.0, "info"),
        ("ICMS", "070", "*", "Redução BC + ST", True, True, True, "VL_ITEM", 12.0, "critica"),
        ("ICMS", "090", "*", "Outras Situações", False, False, False, "VL_ITEM", None, "aviso"),
    ]
    
    REGRAS_PIS = [
        ("PIS", "01", "*", "PIS Não Cumulativo Alíq. Básica", True, True, True, "VL_ITEM", 1.65, "critica"),
        ("PIS", "02", "*", "PIS Não Cumulativo Alíq. Diferenciada", True, True, True, "VL_ITEM", None, "critica"),
        ("PIS", "04", "*", "PIS Monofásico", False, False, False, "", 0.0, "aviso"),
        ("PIS", "05", "*", "PIS ST", False, False, False, "", 0.0, "aviso"),
        ("PIS", "06", "*", "PIS Alíq. Zero", False, False, False, "", 0.0, "aviso"),
        ("PIS", "07", "*", "PIS Isento", False, False, False, "", 0.0, "aviso"),
        ("PIS", "08", "*", "PIS Sem Incidência", False, False, False, "", 0.0, "aviso"),
        ("PIS", "09", "*", "PIS Suspensão", False, False, False, "", 0.0, "aviso"),
        ("PIS", "49", "*", "PIS Outras (saídas)", False, False, False, "VL_ITEM", None, "info"),
        ("PIS", "50", "*", "PIS Crédito Básico", True, True, True, "VL_ITEM", 1.65, "critica"),
        ("PIS", "51", "*", "PIS Crédito Presumido", True, True, True, "VL_ITEM", None, "critica"),
        ("PIS", "70", "*", "PIS Crédito Alíq. Básica", True, True, True, "VL_ITEM", 1.65, "critica"),
        ("PIS", "71", "*", "PIS Crédito Alíq. Diferenciada", True, True, True, "VL_ITEM", None, "critica"),
        ("PIS", "72", "*", "PIS Créd. Presumido Agroind.", True, True, True, "VL_ITEM", None, "critica"),
        ("PIS", "73", "*", "PIS Créd. Ass. Exportação", True, True, True, "VL_ITEM", None, "critica"),
        ("PIS", "74", "*", "PIS Créd. Pat. Amort.", True, True, True, "VL_ITEM", None, "critica"),
        ("PIS", "75", "*", "PIS Créd. Val. Estoques", True, True, True, "VL_ITEM", None, "critica"),
        ("PIS", "98", "*", "PIS Outras Entradas", False, False, False, "VL_ITEM", None, "info"),
        ("PIS", "99", "*", "PIS Outras Saídas", False, False, False, "VL_ITEM", None, "info"),
    ]
    
    REGRAS_COFINS = [
        ("COFINS", "01", "*", "COFINS Não Cumulativo Alíq. Básica", True, True, True, "VL_ITEM", 7.6, "critica"),
        ("COFINS", "02", "*", "COFINS Não Cumulativo Alíq. Diferenciada", True, True, True, "VL_ITEM", None, "critica"),
        ("COFINS", "04", "*", "COFINS Monofásico", False, False, False, "", 0.0, "aviso"),
        ("COFINS", "05", "*", "COFINS ST", False, False, False, "", 0.0, "aviso"),
        ("COFINS", "06", "*", "COFINS Alíq. Zero", False, False, False, "", 0.0, "aviso"),
        ("COFINS", "07", "*", "COFINS Isento", False, False, False, "", 0.0, "aviso"),
        ("COFINS", "08", "*", "COFINS Sem Incidência", False, False, False, "", 0.0, "aviso"),
        ("COFINS", "09", "*", "COFINS Suspensão", False, False, False, "", 0.0, "aviso"),
        ("COFINS", "49", "*", "COFINS Outras (saídas)", False, False, False, "VL_ITEM", None, "info"),
        ("COFINS", "50", "*", "COFINS Crédito Básico", True, True, True, "VL_ITEM", 7.6, "critica"),
        ("COFINS", "70", "*", "COFINS Créd. Alíq. Básica", True, True, True, "VL_ITEM", 7.6, "critica"),
        ("COFINS", "71", "*", "COFINS Créd. Alíq. Diferenciada", True, True, True, "VL_ITEM", None, "critica"),
        ("COFINS", "98", "*", "COFINS Outras Entradas", False, False, False, "VL_ITEM", None, "info"),
        ("COFINS", "99", "*", "COFINS Outras Saídas", False, False, False, "VL_ITEM", None, "info"),
    ]
    
    @classmethod
    def _build_regras_padrao(cls) -> List[RegraFiscal]:
        regras = []
        
        for trib, cst, cfop, desc, eb, ea, ev, bc, alp, crit in cls.REGRAS_ICMS:
            regras.append(RegraFiscal(
                id=f"{trib}_{cst}_{cfop}",
                tributo=trib,
                cst=cst,
                cfop=cfop,
                ind_oper="*",
                descricao=desc,
                exige_base=eb,
                exige_aliquota=ea,
                exige_valor=ev,
                base_campo_sugerido=bc,
                aliquota_padrao=alp,
                formula="base * aliq / 100" if ev else "zero",
                permite_base_zero=not eb,
                permite_valor_zero=not ev,
                criticidade=crit,
                cod_mod="*"
            ))
        
        for trib, cst, cfop, desc, eb, ea, ev, bc, alp, crit in cls.REGRAS_PIS + cls.REGRAS_COFINS:
            regras.append(RegraFiscal(
                id=f"{trib}_{cst}_{cfop}",
                tributo=trib,
                cst=cst,
                cfop=cfop,
                ind_oper="*",
                descricao=desc,
                exige_base=eb,
                exige_aliquota=ea,
                exige_valor=ev,
                base_campo_sugerido=bc,
                aliquota_padrao=alp,
                formula="base * aliq / 100" if ev else "zero",
                permite_base_zero=not eb,
                permite_valor_zero=not ev,
                criticidade=crit,
                cod_mod="*"
            ))
        
        return regras
    
    @classmethod
    def carregar(cls) -> List[RegraFiscal]:
        if os.path.exists(cls.ARQUIVO_REGRA):
            try:
                with open(cls.ARQUIVO_REGRA, encoding="utf-8") as f:
                    dados = json.load(f)
                    return [RegraFiscal(**d) for d in dados]
            except Exception:
                pass
        return cls._build_regras_padrao()
    
    @classmethod
    def salvar(cls, regras: List[RegraFiscal]) -> None:
        with open(cls.ARQUIVO_REGRA, "w", encoding="utf-8") as f:
            json.dump([asdict(r) for r in regras], f, ensure_ascii=False, indent=2)
    
    @classmethod
    def buscar(cls, regras: List[RegraFiscal], tributo: str, cst: str, cfop: str = "*", cod_mod: str = "*") -> Optional[RegraFiscal]:
        """Busca a regra mais específica que corresponde aos critérios."""
        if cod_mod is None or cod_mod == "":
            cod_mod = "*"
        if cfop is None or cfop == "":
            cfop = "*"
        
        candidatos = [
            r for r in regras 
            if r.ativa and 
            r.tributo == tributo and 
            r.match_cst(cst) and 
            r.match_cfop(cfop) and
            r.match_mod(cod_mod)
        ]
        
        if not candidatos:
            return None
        
        def pontuacao(r: RegraFiscal) -> int:
            score = 0
            if r.cst != "*":
                score += 2
            if r.cfop != "*":
                score += 2
            if r.cod_mod != "*":
                score += 1
            return score
        
        return max(candidatos, key=pontuacao)
    
    @classmethod
    def calcular(cls, regra: RegraFiscal, base: float, aliquota: float) -> float:
        if regra.formula == "zero":
            return 0.0
        r = Decimal(str(base)) * Decimal(str(aliquota)) / Decimal("100")
        return float(r.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  PARSER SPED  █████████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

class ParserSPED:
    """Parser para arquivos SPED."""
    
    @staticmethod
    def parse(conteudo: bytes) -> ResultadoParser:
        texto = Utilitarios.decode_bytes(conteudo)
        erros: List[str] = []
        linhas: List[LinhaRegistro] = []
        
        for i, raw in enumerate(texto.splitlines(), 1):
            raw = raw.strip()
            if not raw:
                continue
            
            if not (raw.startswith("|") and raw.endswith("|")):
                erros.append(f"L{i}: formato inválido — {raw[:60]}")
                continue
            
            campos = raw[1:-1].split("|")
            if not campos or not campos[0].strip():
                erros.append(f"L{i}: registro vazio")
                continue
            
            reg = campos[0].strip().upper()
            linhas.append(LinhaRegistro(
                numero_linha=i,
                bloco=reg[0] if reg else "?",
                registro=reg,
                campos=campos,
                linha_original=raw
            ))
        
        metadados = ParserSPED._extrair_metadados(linhas)
        df = ParserSPED._criar_dataframe(linhas)
        tipo_arquivo = ParserSPED._detectar_tipo(linhas, metadados)
        
        return ResultadoParser(metadados, linhas, df, erros, tipo_arquivo)
    
    @staticmethod
    def _extrair_metadados(linhas: List[LinhaRegistro]) -> MetadadosArquivo:
        meta = MetadadosArquivo()
        
        for l in linhas:
            if l.registro == "0000":
                meta.periodo_apuracao = f"{l.get(3)} a {l.get(4)}"
                meta.nome_empresa = l.get(5)
                meta.cnpj = l.get(6)
                meta.uf = l.get(8)
                meta.ie = l.get(9)
                meta.cod_municipio = l.get(10)
                meta.ind_perfil = l.get(13)
                break
        
        meta.total_linhas = len(linhas)
        meta.blocos_presentes = sorted(set(l.bloco for l in linhas if l.bloco != "?"))
        
        return meta
    
    @staticmethod
    def _criar_dataframe(linhas: List[LinhaRegistro]) -> pd.DataFrame:
        rows = []
        
        for l in linhas:
            row = {
                "numero_linha": l.numero_linha,
                "bloco": l.bloco,
                "registro": l.registro,
                "linha_original": l.linha_original,
            }
            
            row.update({f"campo_{j:02d}": l.get(j) for j in range(min(len(l.campos), 55))})
            
            mapa = MapaCampos.get_mapa(l.registro)
            if mapa:
                for nome, idx in mapa.items():
                    row[nome] = l.get(idx)
            
            rows.append(row)
        
        df = pd.DataFrame(rows)
        
        for col in CAMPOS_NUMERICOS:
            if col in df.columns:
                df[col] = df[col].apply(Utilitarios.to_float)
        
        return df
    
    @staticmethod
    def _detectar_tipo(linhas: List[LinhaRegistro], metadados: MetadadosArquivo) -> str:
        blocos = set(metadados.blocos_presentes)
        
        cod_ver = ""
        for l in linhas:
            if l.registro == "0000":
                cod_ver = l.get(1)
                break
        
        if blocos & {"M", "P"} or cod_ver.startswith("006") or cod_ver.startswith("007"):
            return TIPO_EFD_CONTRIB
        elif blocos & {"E", "G", "K", "H"}:
            return TIPO_EFD_ICMS
        
        return TIPO_EFD_ICMS
    
    @staticmethod
    def reconstruir(linhas: List[LinhaRegistro]) -> bytes:
        return "".join(l.to_sped() for l in sorted(linhas, key=lambda x: x.numero_linha)).encode("utf-8")
    
    @staticmethod
    def sync_df_linhas(df: pd.DataFrame, linhas: List[LinhaRegistro]) -> List[LinhaRegistro]:
        mapa = {l.numero_linha: l for l in linhas}
        
        for _, row in df.iterrows():
            nl = row.get("numero_linha")
            if nl not in mapa:
                continue
            
            l = mapa[nl]
            mapa_campos = MapaCampos.get_mapa(l.registro)
            if not mapa_campos:
                continue
            
            for nome, idx in mapa_campos.items():
                if nome not in row or pd.isna(row[nome]):
                    continue
                
                valor = row[nome]
                if isinstance(valor, float):
                    valor = Utilitarios.to_sped_str(valor)
                l.set(idx, str(valor))
        
        return list(mapa.values())


# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  VALIDADOR FISCAL  █████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

class ValidadorFiscal:
    """Validador de inconsistências fiscais."""
    
    @staticmethod
    def validar(df: pd.DataFrame, regras: List[RegraFiscal], tipo_arquivo: str) -> pd.DataFrame:
        inconsistencias: List[Dict[str, Any]] = []
        
        def adicionar(nl: int, reg: str, bloco: str, tipo: str, severidade: str,
                     campo: str, desc: str, va: str, vs: str, cst: str = "",
                     cfop: str = "", regra_id: str = "", num_doc: str = ""):
            inconsistencias.append({
                "numero_linha": nl,
                "bloco": bloco,
                "registro": reg,
                "tipo": tipo,
                "severidade": severidade,
                "campo_afetado": campo,
                "descricao": desc,
                "valor_atual": va,
                "valor_sugerido": vs,
                "cst": cst,
                "cfop": cfop,
                "num_doc": num_doc,
                "regra_aplicada": regra_id,
                "corrigido": False,
            })
        
        ValidadorFiscal._validar_tributos_c170(df, regras, adicionar)
        ValidadorFiscal._validar_tributos_a170(df, regras, adicionar)
        ValidadorFiscal._validar_tributos_c185(df, regras, adicionar)
        ValidadorFiscal._validar_tributos_d100(df, regras, adicionar)
        ValidadorFiscal._validar_f100(df, regras, adicionar)
        ValidadorFiscal._validar_totalizacao_c100(df, adicionar)
        ValidadorFiscal._validar_c190(df, regras, adicionar)
        ValidadorFiscal._validar_apuracao(df, adicionar)
        ValidadorFiscal._validar_blocos(df, adicionar)
        ValidadorFiscal._validar_campos_obrigatorios(df, adicionar)
        ValidadorFiscal._validar_negativos(df, adicionar)
        
        return pd.DataFrame(inconsistencias)
    
    @staticmethod
    def _validar_tributo(row: pd.Series, nl: int, reg: str, bloco: str,
                        tributo: str, campo_cst: str, campo_bc: str,
                        campo_aliq: str, campo_vl: str, campo_item: str,
                        regras: List[RegraFiscal], adicionar: Callable):
        cst = str(row.get(campo_cst, "") or "").strip()
        bc = Utilitarios.to_float(row.get(campo_bc))
        aliq = Utilitarios.to_float(row.get(campo_aliq))
        vl = Utilitarios.to_float(row.get(campo_vl))
        vi = Utilitarios.to_float(row.get(campo_item))
        cod_mod = str(row.get("COD_MOD", "*") or "*").strip()
        cfop = str(row.get("CFOP", "*") or "*").strip()
        
        if not cst:
            return
        
        regra = GerenciadorRegras.buscar(regras, tributo, cst, cfop, cod_mod)
        if not regra:
            return
        
        severidade = regra.criticidade.upper()
        rid = regra.id
        
        if regra.exige_base and (bc is None or bc == 0):
            adicionar(
                nl, reg, bloco, f"SEM_BASE_{tributo}", severidade,
                campo_bc,
                f"[{tributo}] CST {cst} exige base — ausente/zerada",
                str(row.get(campo_bc, "")),
                Utilitarios.to_sped_str(vi) if vi else "",
                cst, cfop, rid
            )
        
        if regra.exige_aliquota and (aliq is None or aliq == 0):
            adicionar(
                nl, reg, bloco, f"SEM_ALIQ_{tributo}", severidade,
                campo_aliq,
                f"[{tributo}] CST {cst} exige alíquota — ausente/zerada",
                str(row.get(campo_aliq, "")),
                Utilitarios.to_sped_str(regra.aliquota_padrao) if regra.aliquota_padrao else "",
                cst, cfop, rid
            )
        
        if regra.exige_valor and (vl is None or vl == 0):
            base = bc or vi or 0.0
            aliquota = aliq or (regra.aliquota_padrao or 0.0)
            calculado = GerenciadorRegras.calcular(regra, base, aliquota)
            adicionar(
                nl, reg, bloco, f"SEM_VALOR_{tributo}", severidade,
                campo_vl,
                f"[{tributo}] CST {cst} exige valor — ausente/zerado (sugerido: {Utilitarios.to_sped_str(calculado)})",
                str(row.get(campo_vl, "")),
                Utilitarios.to_sped_str(calculado) if calculado else "",
                cst, cfop, rid
            )
        
        if not regra.exige_valor and vl and vl > 0:
            adicionar(
                nl, reg, bloco, f"VALOR_INDEVIDO_{tributo}", "AVISO",
                campo_vl,
                f"[{tributo}] CST {cst} não deve ter valor tributário — preenchido indevidamente",
                Utilitarios.to_sped_str(vl),
                "0",
                cst, cfop, rid
            )
        
        if bc and aliq and vl is not None and regra.formula == "base * aliq / 100":
            esperado = GerenciadorRegras.calcular(regra, bc, aliq)
            if abs(esperado - vl) > 0.05:
                adicionar(
                    nl, reg, bloco, f"DIVERGENCIA_{tributo}", "AVISO",
                    campo_vl,
                    f"[{tributo}] Calculado {Utilitarios.fmt_brl(esperado)} ≠ Registrado {Utilitarios.fmt_brl(vl)}",
                    Utilitarios.to_sped_str(vl),
                    Utilitarios.to_sped_str(esperado),
                    cst, cfop, rid
                )
    
    @staticmethod
    def _validar_tributos_c170(df: pd.DataFrame, regras: List[RegraFiscal], adicionar: Callable):
        for _, row in df[df["registro"] == "C170"].iterrows():
            nl = int(row.get("numero_linha", 0))
            ValidadorFiscal._validar_tributo(
                row, nl, "C170", "C", "ICMS", "CST_ICMS",
                "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS", "VL_ITEM",
                regras, adicionar
            )
            ValidadorFiscal._validar_tributo(
                row, nl, "C170", "C", "PIS", "CST_PIS",
                "VL_BC_PIS", "ALIQ_PIS", "VL_PIS", "VL_ITEM",
                regras, adicionar
            )
            ValidadorFiscal._validar_tributo(
                row, nl, "C170", "C", "COFINS", "CST_COFINS",
                "VL_BC_COFINS", "ALIQ_COFINS", "VL_COFINS", "VL_ITEM",
                regras, adicionar
            )
    
    @staticmethod
    def _validar_tributos_a170(df: pd.DataFrame, regras: List[RegraFiscal], adicionar: Callable):
        for _, row in df[df["registro"] == "A170"].iterrows():
            nl = int(row.get("numero_linha", 0))
            ValidadorFiscal._validar_tributo(
                row, nl, "A170", "A", "PIS", "CST_PIS",
                "VL_BC_PIS", "ALIQ_PIS", "VL_PIS", "VL_ITEM",
                regras, adicionar
            )
            ValidadorFiscal._validar_tributo(
                row, nl, "A170", "A", "COFINS", "CST_COFINS",
                "VL_BC_COFINS", "ALIQ_COFINS", "VL_COFINS", "VL_ITEM",
                regras, adicionar
            )
    
    @staticmethod
    def _validar_tributos_c185(df: pd.DataFrame, regras: List[RegraFiscal], adicionar: Callable):
        for _, row in df[df["registro"] == "C185"].iterrows():
            nl = int(row.get("numero_linha", 0))
            ValidadorFiscal._validar_tributo(
                row, nl, "C185", "C", "PIS", "CST_PIS",
                "VL_BC_PIS", "ALIQ_PIS", "VL_PIS", "",
                regras, adicionar
            )
            ValidadorFiscal._validar_tributo(
                row, nl, "C185", "C", "COFINS", "CST_COFINS",
                "VL_BC_COFINS", "ALIQ_COFINS", "VL_COFINS", "",
                regras, adicionar
            )
    
    @staticmethod
    def _validar_tributos_d100(df: pd.DataFrame, regras: List[RegraFiscal], adicionar: Callable):
        for _, row in df[df["registro"] == "D100"].iterrows():
            nl = int(row.get("numero_linha", 0))
            ValidadorFiscal._validar_tributo(
                row, nl, "D100", "D", "PIS", "CST_PIS",
                "VL_BC_PIS", "ALIQ_PIS", "VL_PIS", "VL_DOC",
                regras, adicionar
            )
            ValidadorFiscal._validar_tributo(
                row, nl, "D100", "D", "COFINS", "CST_COFINS",
                "VL_BC_COFINS", "ALIQ_COFINS", "VL_COFINS", "VL_DOC",
                regras, adicionar
            )
    
    @staticmethod
    def _validar_f100(df: pd.DataFrame, regras: List[RegraFiscal], adicionar: Callable):
        regra_base = RegraFiscal(
            id="", tributo="PIS", cst="01", cfop="*", ind_oper="*",
            descricao="", exige_base=True, exige_aliquota=True, exige_valor=True,
            base_campo_sugerido="", aliquota_padrao=1.65,
            formula="base * aliq / 100", permite_base_zero=False,
            permite_valor_zero=False, criticidade="critica", ativa=True
        )
        
        for _, row in df[df["registro"] == "F100"].iterrows():
            nl = int(row.get("numero_linha", 0))
            
            bc_p = Utilitarios.to_float(row.get("VL_BC_PIS"))
            al_p = Utilitarios.to_float(row.get("ALIQ_PIS"))
            vl_p = Utilitarios.to_float(row.get("VL_PIS"))
            
            if bc_p and al_p and vl_p is not None:
                esperado = GerenciadorRegras.calcular(regra_base, bc_p, al_p)
                if abs(esperado - vl_p) > 0.05:
                    adicionar(
                        nl, "F100", "F", "DIVERGENCIA_PIS", "AVISO",
                        "VL_PIS",
                        f"[PIS] F100 calculado {Utilitarios.fmt_brl(esperado)} ≠ {Utilitarios.fmt_brl(vl_p)}",
                        Utilitarios.to_sped_str(vl_p),
                        Utilitarios.to_sped_str(esperado)
                    )
            
            bc_c = Utilitarios.to_float(row.get("VL_BC_COFINS"))
            al_c = Utilitarios.to_float(row.get("ALIQ_COFINS"))
            vl_c = Utilitarios.to_float(row.get("VL_COFINS"))
            
            if bc_c and al_c and vl_c is not None:
                regra_c = RegraFiscal(
                    id="", tributo="COFINS", cst="01", cfop="*", ind_oper="*",
                    descricao="", exige_base=True, exige_aliquota=True, exige_valor=True,
                    base_campo_sugerido="", aliquota_padrao=7.6,
                    formula="base * aliq / 100", permite_base_zero=False,
                    permite_valor_zero=False, criticidade="critica", ativa=True
                )
                esperado = GerenciadorRegras.calcular(regra_c, bc_c, al_c)
                if abs(esperado - vl_c) > 0.05:
                    adicionar(
                        nl, "F100", "F", "DIVERGENCIA_COFINS", "AVISO",
                        "VL_COFINS",
                        f"[COFINS] F100 calculado {Utilitarios.fmt_brl(esperado)} ≠ {Utilitarios.fmt_brl(vl_c)}",
                        Utilitarios.to_sped_str(vl_c),
                        Utilitarios.to_sped_str(esperado)
                    )
    
    @staticmethod
    def _validar_totalizacao_c100(df: pd.DataFrame, adicionar: Callable):
        df_docs = df[df["registro"].isin(["C100", "C170"])].sort_values("numero_linha")
        c100_atual = None
        c100_nl = 0
        soma_icms = 0.0
        
        for _, row in df_docs.iterrows():
            if row["registro"] == "C100":
                if c100_atual is not None:
                    vl_doc = Utilitarios.to_float(c100_atual.get("VL_ICMS")) or 0.0
                    if abs(soma_icms - vl_doc) > 0.10:
                        adicionar(
                            c100_nl, "C100", "C", "DIVERGENCIA_TOTAL_ICMS", "CRITICA",
                            "VL_ICMS",
                            f"Total C100 ICMS ({Utilitarios.fmt_brl(vl_doc)}) ≠ soma C170 ({Utilitarios.fmt_brl(soma_icms)})",
                            Utilitarios.to_sped_str(vl_doc),
                            Utilitarios.to_sped_str(soma_icms),
                            num_doc=str(c100_atual.get("NUM_DOC", ""))
                        )
                
                c100_atual = row
                c100_nl = int(row["numero_linha"])
                soma_icms = 0.0
            
            elif row["registro"] == "C170" and c100_atual is not None:
                soma_icms += Utilitarios.to_float(row.get("VL_ICMS")) or 0.0
    
    @staticmethod
    def _validar_c190(df: pd.DataFrame, regras: List[RegraFiscal], adicionar: Callable):
        for _, row in df[df["registro"] == "C190"].iterrows():
            cst = str(row.get("CST_ICMS", "") or "").strip()
            cfop = str(row.get("CFOP", "") or "").strip()
            nl = int(row.get("numero_linha", 0))
            
            regra = GerenciadorRegras.buscar(regras, "ICMS", cst)
            if regra and regra.exige_base:
                bc = Utilitarios.to_float(row.get("VL_BC_ICMS"))
                if not bc:
                    adicionar(
                        nl, "C190", "C", "C190_SEM_BASE", "AVISO",
                        "VL_BC_ICMS",
                        f"C190 CST {cst}/CFOP {cfop} sem base ICMS",
                        "",
                        str(row.get("VL_OPR", "")),
                        cst, cfop
                    )
    
    @staticmethod
    def _validar_apuracao(df: pd.DataFrame, adicionar: Callable):
        for reg_m, campo_rec in [("M200", "VL_TOT_CONT_REC"), ("M600", "VL_TOT_CONT_REC")]:
            for _, row in df[df["registro"] == reg_m].iterrows():
                nl = int(row.get("numero_linha", 0))
                vr = Utilitarios.to_float(row.get(campo_rec))
                nc = Utilitarios.to_float(row.get("VL_CONT_NC_REC"))
                cu = Utilitarios.to_float(row.get("VL_CONT_CUM_REC"))
                trib = "PIS" if reg_m == "M200" else "COFINS"
                
                if vr is not None and nc is not None and cu is not None:
                    esperado = (nc or 0) + (cu or 0)
                    if abs(esperado - (vr or 0)) > 0.05:
                        adicionar(
                            nl, reg_m, "M", f"DIVERGENCIA_APURACAO_{trib}", "AVISO",
                            campo_rec,
                            f"[{trib}] {reg_m}: NC ({Utilitarios.fmt_brl(nc)}) + CUM ({Utilitarios.fmt_brl(cu)}) ≠ Total ({Utilitarios.fmt_brl(vr)})",
                            Utilitarios.to_sped_str(vr or 0),
                            Utilitarios.to_sped_str(esperado)
                        )
    
    @staticmethod
    def _validar_blocos(df: pd.DataFrame, adicionar: Callable):
        for bloco in df["bloco"].unique():
            if bloco == "?":
                continue
            
            for suf, tipo in [("001", "BLOCO_SEM_ABERTURA"), ("990", "BLOCO_SEM_FECHAMENTO")]:
                reg_bloco = f"{bloco}{suf}"
                if df[df["registro"] == reg_bloco].empty:
                    adicionar(
                        0, reg_bloco, bloco, tipo, "CRITICA",
                        "registro",
                        f"Bloco {bloco}: {reg_bloco} ausente",
                        "ausente",
                        ""
                    )
    
    @staticmethod
    def _validar_campos_obrigatorios(df: pd.DataFrame, adicionar: Callable):
        df0 = df[df["registro"] == "0000"]
        if df0.empty:
            adicionar(
                0, "0000", "0", "CAMPO_AUSENTE", "CRITICA",
                "0000",
                "Registro 0000 não encontrado",
                "ausente",
                ""
            )
        else:
            row0 = df0.iloc[0]
            for campo in ("CNPJ", "NOME"):
                if not str(row0.get(campo, "") or "").strip():
                    adicionar(
                        int(row0.get("numero_linha", 1)),
                        "0000", "0", "CAMPO_AUSENTE", "CRITICA",
                        campo,
                        f"Campo {campo} obrigatório ausente no 0000",
                        "",
                        ""
                    )
    
    @staticmethod
    def _validar_negativos(df: pd.DataFrame, adicionar: Callable):
        df_c170 = df[df["registro"] == "C170"]
        campos_negativos = ["VL_ITEM", "VL_BC_ICMS", "VL_ICMS", "ALIQ_ICMS",
                           "VL_BC_PIS", "VL_PIS", "VL_BC_COFINS", "VL_COFINS"]
        
        for campo in campos_negativos:
            if campo not in df_c170.columns:
                continue
            
            negativos = df_c170[df_c170[campo].apply(
                lambda x: x is not None and isinstance(x, float) and x < 0
            )]
            
            for _, row in negativos.iterrows():
                adicionar(
                    int(row.get("numero_linha", 0)),
                    "C170", "C", "VALOR_NEGATIVO", "AVISO",
                    campo,
                    f"Valor negativo em {campo}: {row[campo]:.2f}",
                    Utilitarios.to_sped_str(row[campo]),
                    Utilitarios.to_sped_str(abs(row[campo])),
                    str(row.get("CST_ICMS", "")),
                    str(row.get("CFOP", ""))
                )


# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  EDITOR E AUDITORIA  ███████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

class TrilhaAuditoria:
    """Gerencia o log de auditoria."""
    
    def __init__(self):
        self._log: List[EntradaAuditoria] = []
    
    def registrar(self, nl: int, reg: str, bloco: str, campo: str,
                  ant: str, novo: str, regra: str = "", motivo: str = "",
                  tipo: str = "MANUAL", usuario: str = "analista") -> None:
        self._log.append(EntradaAuditoria(
            str(uuid.uuid4())[:8],
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            usuario,
            nl,
            reg,
            bloco,
            campo,
            ant,
            novo,
            regra,
            motivo,
            tipo
        ))
    
    def to_df(self) -> pd.DataFrame:
        if not self._log:
            return pd.DataFrame()
        
        return pd.DataFrame([{
            "ID": e.id,
            "Data/Hora": e.timestamp,
            "Usuário": e.usuario,
            "Linha": e.numero_linha,
            "Registro": e.registro,
            "Bloco": e.bloco,
            "Campo": e.campo,
            "Anterior": e.valor_anterior,
            "Novo": e.valor_novo,
            "Regra": e.regra,
            "Motivo": e.motivo,
            "Tipo": e.tipo,
        } for e in self._log])
    
    def total(self) -> int:
        return len(self._log)


class EditorSped:
    """Editor de dados SPED com controle de versão."""
    
    def __init__(self, df: pd.DataFrame):
        self._original = df.copy()
        self._atual = df.copy()
        self._historico: List[pd.DataFrame] = []
        self.auditoria = TrilhaAuditoria()
    
    @property
    def df(self) -> pd.DataFrame:
        return self._atual
    
    @property
    def df_original(self) -> pd.DataFrame:
        return self._original
    
    def _set_valor(self, idx: int, campo: str, valor: Any) -> None:
        if campo in CAMPOS_NUMERICOS and campo in self._atual.columns:
            vn = Utilitarios.to_float(valor)
            if vn is not None:
                self._atual.at[idx, campo] = vn
            else:
                if self._atual[campo].dtype != object:
                    self._atual[campo] = self._atual[campo].astype(object)
                self._atual.at[idx, campo] = valor
        else:
            if campo in self._atual.columns and self._atual[campo].dtype != object:
                self._atual[campo] = self._atual[campo].astype(object)
            self._atual.at[idx, campo] = valor
    
    def editar(self, nl: int, campo: str, valor: Any, motivo: str = "",
               regra: str = "", usuario: str = "analista") -> bool:
        idx = self._atual[self._atual["numero_linha"] == nl].index
        if idx.empty:
            return False
        
        self._historico.append(self._atual.copy())
        i = idx[0]
        row = self._atual.loc[i]
        ant = str(row.get(campo, ""))
        
        self._set_valor(i, campo, valor)
        
        self.auditoria.registrar(
            nl,
            str(row.get("registro", "")),
            str(row.get("bloco", "")),
            campo,
            ant,
            str(valor),
            regra,
            motivo,
            "MANUAL",
            usuario
        )
        
        return True
    
    def massa(self, nls: List[int], campo: str, valor: Any,
              motivo: str = "Correção em massa",
              regra: str = "", usuario: str = "analista") -> int:
        self._historico.append(self._atual.copy())
        count = 0
        
        for nl in nls:
            idx = self._atual[self._atual["numero_linha"] == nl].index
            if idx.empty:
                continue
            
            i = idx[0]
            row = self._atual.loc[i]
            ant = str(row.get(campo, ""))
            
            if ant == str(valor):
                continue
            
            self._set_valor(i, campo, valor)
            
            self.auditoria.registrar(
                nl,
                str(row.get("registro", "")),
                str(row.get("bloco", "")),
                campo,
                ant,
                str(valor),
                regra,
                motivo,
                "MASSA",
                usuario
            )
            count += 1
        
        return count
    
    def recalcular_massa(self, nls: List[int], regras: List[RegraFiscal],
                         tributo: str, campo_cst: str, campo_bc: str,
                         campo_aliq: str, campo_vl: str,
                         motivo: str = "Recálculo automático",
                         usuario: str = "analista") -> int:
        self._historico.append(self._atual.copy())
        count = 0
        
        for nl in nls:
            idx = self._atual[self._atual["numero_linha"] == nl].index
            if idx.empty:
                continue
            
            i = idx[0]
            row = self._atual.loc[i]
            
            cst = str(row.get(campo_cst, "") or "").strip()
            bc = Utilitarios.to_float(row.get(campo_bc))
            aliq = Utilitarios.to_float(row.get(campo_aliq))
            
            if bc is None or aliq is None:
                continue
            
            regra = GerenciadorRegras.buscar(regras, tributo, cst)
            if not regra:
                continue
            
            novo_valor = GerenciadorRegras.calcular(regra, bc, aliq)
            ant = str(row.get(campo_vl, ""))
            
            self._set_valor(i, campo_vl, novo_valor)
            
            self.auditoria.registrar(
                nl,
                str(row.get("registro", "")),
                str(row.get("bloco", "")),
                campo_vl,
                ant,
                Utilitarios.to_sped_str(novo_valor),
                regra.id,
                motivo,
                "AUTO",
                usuario
            )
            count += 1
        
        return count
    
    def desfazer(self) -> bool:
        if not self._historico:
            return False
        self._atual = self._historico.pop()
        return True
    
    def restaurar(self, nls: Optional[List[int]] = None) -> None:
        self._historico.append(self._atual.copy())
        
        if nls:
            for nl in nls:
                idx_atual = self._atual[self._atual["numero_linha"] == nl].index
                idx_orig = self._original[self._original["numero_linha"] == nl].index
                if not idx_atual.empty and not idx_orig.empty:
                    self._atual.loc[idx_atual[0]] = self._original.loc[idx_orig[0]]
        else:
            self._atual = self._original.copy()
    
    def get_alterados(self) -> pd.DataFrame:
        try:
            return self._atual[self._atual.ne(self._original).any(axis=1)]
        except:
            return pd.DataFrame()
    
    def preview(self, nls: List[int]) -> pd.DataFrame:
        rows = []
        
        for nl in nls:
            atual = self._atual[self._atual["numero_linha"] == nl]
            original = self._original[self._original["numero_linha"] == nl]
            
            if atual.empty or original.empty:
                continue
            
            ra = atual.iloc[0]
            ro = original.iloc[0]
            
            for campo in CAMPOS_NUMERICOS:
                if campo in ra.index and ra.get(campo) != ro.get(campo):
                    rows.append({
                        "Linha": nl,
                        "Registro": ra.get("registro", ""),
                        "Campo": campo,
                        "Original": ro.get(campo),
                        "Atual": ra.get(campo),
                    })
        
        return pd.DataFrame(rows)


# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  EXPORTADOR  ███████████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

class Exportador:
    """Gerencia exportação de dados."""
    
    _CORES = {
        "cabecalho": "FF1A3A5C",
        "critico": "FFC0392B",
        "aviso": "FFB7860D",
        "ok": "FF1A7A35",
        "info": "FF2C5F8A",
        "fundo_par": "FFEBF0F7",
        "fundo_impar": "FFFFFFFF",
    }
    
    @classmethod
    def _estilo_borda(cls) -> Border:
        return Border(bottom=Side(style="thin", color="CCCCCC"),
                      right=Side(style="thin", color="CCCCCC"))
    
    @classmethod
    def _criar_aba(cls, ws, df: pd.DataFrame, titulo: str, col_sev: str = "") -> None:
        ws.sheet_view.showGridLines = False
        
        if df is None or df.empty:
            ws["A1"] = f"{titulo} — sem dados"
            ws["A1"].font = Font(bold=True, italic=True, color="888888")
            return
        
        n_cols = len(df.columns)
        ultima_col = get_column_letter(n_cols)
        
        ws.merge_cells(f"A1:{ultima_col}1")
        c = ws["A1"]
        c.value = titulo
        c.font = Font(bold=True, size=13, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=cls._CORES["cabecalho"])
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 26
        
        borda = cls._estilo_borda()
        for ci, col in enumerate(df.columns, 1):
            h = ws.cell(2, ci, str(col).upper())
            h.font = Font(bold=True, size=10, color="FFFFFF")
            h.fill = PatternFill("solid", fgColor=cls._CORES["info"])
            h.alignment = Alignment(horizontal="center")
            h.border = borda
        ws.row_dimensions[2].height = 22
        
        mapa_cores = {
            "CRITICA": "FFF0F0",
            "AVISO": "FFF8DC",
            "INFO": "EAF4FF",
        }
        
        for ri, (_, row) in enumerate(df.iterrows(), 3):
            cor_fundo = cls._CORES["fundo_par"] if ri % 2 == 0 else cls._CORES["fundo_impar"]
            sev = str(row.get(col_sev, "")).upper() if col_sev else ""
            cor = mapa_cores.get(sev, cor_fundo)
            
            for ci, col in enumerate(df.columns, 1):
                cell = ws.cell(ri, ci, row[col])
                cell.font = Font(size=10)
                cell.fill = PatternFill("solid", fgColor=cor)
                cell.border = borda
                cell.alignment = Alignment(vertical="center")
        
        for ci, col in enumerate(df.columns, 1):
            largura = max(len(str(col)), df[col].astype(str).str.len().max() if len(df) > 0 else 10)
            ws.column_dimensions[get_column_letter(ci)].width = min(largura + 4, 42)
        
        ws.freeze_panes = "A3"
    
    @classmethod
    def _criar_resumo(cls, ws, df_inc: pd.DataFrame, df_alt: pd.DataFrame, meta: Dict[str, str]) -> None:
        ws.sheet_view.showGridLines = False
        
        ws.merge_cells("A1:F1")
        c = ws["A1"]
        c.value = "SPED AUDITOR — RELATÓRIO GERENCIAL"
        c.font = Font(name="Calibri", bold=True, size=16, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=cls._CORES["cabecalho"])
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 36
        
        dados = [
            ("Empresa:", meta.get("nome_empresa", "—")),
            ("CNPJ:", meta.get("cnpj", "—")),
            ("Período:", meta.get("periodo_apuracao", "—")),
            ("Tipo:", meta.get("tipo_arquivo", "—")),
            ("UF:", meta.get("uf", "—")),
            ("Gerado em:", datetime.now().strftime("%d/%m/%Y %H:%M")),
        ]
        
        for i, (label, valor) in enumerate(dados, 3):
            ws.cell(i, 1, label).font = Font(bold=True, size=11)
            ws.cell(i, 2, valor).font = Font(size=11)
        
        ws.cell(11, 1, "INDICADORES").font = Font(bold=True, size=12, color=cls._CORES["cabecalho"][2:])
        
        def contar(tipo: str) -> int:
            if df_inc.empty or "tipo" not in df_inc.columns:
                return 0
            return int((df_inc["tipo"] == tipo).sum())
        
        def contar_sev(sev: str) -> int:
            if df_inc.empty or "severidade" not in df_inc.columns:
                return 0
            return int((df_inc["severidade"] == sev).sum())
        
        indicadores = [
            ("Total Inconsistências", len(df_inc) if not df_inc.empty else 0, cls._CORES["cabecalho"]),
            ("Críticas", contar_sev("CRITICA"), cls._CORES["critico"]),
            ("Avisos", contar_sev("AVISO"), cls._CORES["aviso"]),
            ("Sem Base ICMS", contar("SEM_BASE_ICMS"), cls._CORES["critico"]),
            ("Sem Valor ICMS", contar("SEM_VALOR_ICMS"), cls._CORES["critico"]),
            ("Sem Base PIS", contar("SEM_BASE_PIS"), cls._CORES["critico"]),
            ("Sem Valor PIS", contar("SEM_VALOR_PIS"), cls._CORES["critico"]),
            ("Sem Base COFINS", contar("SEM_BASE_COFINS"), cls._CORES["critico"]),
            ("Sem Valor COFINS", contar("SEM_VALOR_COFINS"), cls._CORES["critico"]),
            ("Registros Alterados", len(df_alt) if not df_alt.empty else 0, cls._CORES["ok"]),
        ]
        
        for j, (label, valor, cor) in enumerate(indicadores, 13):
            ws.cell(j, 1, label).font = Font(size=11, bold=True)
            c2 = ws.cell(j, 2, valor)
            c2.font = Font(size=11, bold=True, color="FFFFFF")
            c2.fill = PatternFill("solid", fgColor=cor)
            c2.alignment = Alignment(horizontal="center")
        
        ws.column_dimensions["A"].width = 34
        ws.column_dimensions["B"].width = 20
    
    @classmethod
    def to_excel(cls, df_inc: pd.DataFrame, df_alt: pd.DataFrame,
                 df_aud: pd.DataFrame, df_reg: pd.DataFrame,
                 meta: Dict[str, str]) -> bytes:
        buf = io.BytesIO()
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        
        cls._criar_resumo(wb.create_sheet("Resumo Gerencial"), df_inc, df_alt, meta)
        cls._criar_aba(wb.create_sheet("Inconsistências"), df_inc, "Inconsistências Fiscais", "severidade")
        cls._criar_aba(wb.create_sheet("Registros Alterados"), df_alt, "Registros Modificados")
        cls._criar_aba(wb.create_sheet("Log de Auditoria"), df_aud, "Trilha de Auditoria")
        cls._criar_aba(wb.create_sheet("Regras Tributárias"), df_reg, "Catálogo de Regras")
        
        wb.save(buf)
        buf.seek(0)
        return buf.read()
    
    @classmethod
    def to_csv(cls, df: pd.DataFrame) -> bytes:
        buf = io.StringIO()
        df.to_csv(buf, index=False, sep=";", encoding="utf-8-sig")
        return buf.getvalue().encode("utf-8-sig")


# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  IMPORTAÇÃO CT-e  █████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

class ImportadorCTe:
    """Importador de XML CT-e para SPED."""
    
    @staticmethod
    def _txt(elem: ET.Element, *caminho: str, ns: str = "") -> str:
        atual = elem
        for tag in caminho:
            tag_com_ns = f"{{{ns}}}{tag}" if ns else tag
            atual = atual.find(tag_com_ns)
            if atual is None:
                return ""
        return (atual.text or "").strip()
    
    @staticmethod
    def _fmt_valor(valor: str) -> str:
        if not valor:
            return ""
        try:
            d = Decimal(valor.replace(",", "."))
            return str(d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)).replace(".", ",")
        except:
            return valor
    
    @classmethod
    def parse_xml(cls, conteudo_bytes: bytes) -> DadosCTe:
        dados = DadosCTe()
        ns = ""
        
        try:
            root = ET.fromstring(conteudo_bytes)
            ns = Utilitarios.extrair_namespace(root)
        except ET.ParseError as e:
            dados.erros.append(f"XML inválido: {e}")
            return dados
        
        def find_any(elem, tag):
            r = elem.find(f"{{{ns}}}{tag}" if ns else tag)
            if r is None and ns:
                r = elem.find(tag)
            return r
        
        inf_cte = find_any(root, "infCte")
        if inf_cte is None:
            cte_node = find_any(root, "CTe")
            if cte_node is not None:
                inf_cte = find_any(cte_node, "infCte")
        
        if inf_cte is None:
            dados.erros.append("Elemento infCte não encontrado no XML.")
            return dados
        
        ide = find_any(inf_cte, "ide")
        if ide is not None:
            dados.num_doc = cls._txt(ide, "nCT", ns=ns)
            dados.serie = cls._txt(ide, "serie", ns=ns)
            dados.modelo = cls._txt(ide, "mod", ns=ns)
            dados.tipo_cte = cls._txt(ide, "tpCTe", ns=ns)
            dados.cod_sit = "00"
            dt_emi = cls._txt(ide, "dhEmi", ns=ns) or cls._txt(ide, "dEmi", ns=ns)
            dados.dt_emissao = Utilitarios.formatar_data_sped(dt_emi)
        
        chave_raw = inf_cte.get("Id", "")
        if chave_raw.startswith("CTe"):
            dados.chave = chave_raw[3:]
        else:
            dados.chave = chave_raw
        
        emit = find_any(inf_cte, "emit")
        if emit is not None:
            dados.cnpj_emit = cls._txt(emit, "CNPJ", ns=ns)
            dados.ie_emit = cls._txt(emit, "IE", ns=ns)
            dados.nome_emit = cls._txt(emit, "xNome", ns=ns)
        
        v_prest = find_any(inf_cte, "vPrest")
        if v_prest is not None:
            dados.vl_doc = cls._fmt_valor(cls._txt(v_prest, "vTPrest", ns=ns))
            dados.vl_serv = cls._fmt_valor(cls._txt(v_prest, "vRec", ns=ns))
        
        imp = find_any(inf_cte, "imp")
        if imp is not None:
            icms_node = find_any(imp, "ICMS")
            if icms_node is not None:
                for tag in ["ICMS00", "ICMS20", "ICMS45", "ICMS60", "ICMS90", "ICMSOutUF", "ICMSSn"]:
                    filho = find_any(icms_node, tag)
                    if filho is not None:
                        dados.cst_icms = cls._txt(filho, "CST", ns=ns)
                        dados.vl_bc_icms = cls._fmt_valor(cls._txt(filho, "vBC", ns=ns))
                        dados.aliq_icms = cls._fmt_valor(cls._txt(filho, "pICMS", ns=ns))
                        dados.vl_icms = cls._fmt_valor(cls._txt(filho, "vICMS", ns=ns))
                        break
            
            pis_node = find_any(imp, "PIS")
            if pis_node is not None:
                for tag in ["PISAliq", "PISQtde", "PISNT", "PISOutr"]:
                    filho = find_any(pis_node, tag)
                    if filho is not None:
                        dados.cst_pis = cls._txt(filho, "CST", ns=ns)
                        dados.vl_bc_pis = cls._fmt_valor(cls._txt(filho, "vBC", ns=ns))
                        dados.aliq_pis = cls._fmt_valor(cls._txt(filho, "pPIS", ns=ns))
                        dados.vl_pis = cls._fmt_valor(cls._txt(filho, "vPIS", ns=ns))
                        break
            
            cof_node = find_any(imp, "COFINS")
            if cof_node is not None:
                for tag in ["COFINSAliq", "COFINSQtde", "COFINSNT", "COFINSOutr"]:
                    filho = find_any(cof_node, tag)
                    if filho is not None:
                        dados.cst_cofins = cls._txt(filho, "CST", ns=ns)
                        dados.vl_bc_cofins = cls._fmt_valor(cls._txt(filho, "vBC", ns=ns))
                        dados.aliq_cofins = cls._fmt_valor(cls._txt(filho, "pCOFINS", ns=ns))
                        dados.vl_cofins = cls._fmt_valor(cls._txt(filho, "vCOFINS", ns=ns))
                        break
        
        return dados
    
    @classmethod
    def para_linhas_d(cls, dados: DadosCTe, cod_part: str, nl_inicio: int,
                      aliq_pis_padrao: str = "1,65",
                      aliq_cofins_padrao: str = "7,60",
                      cst_pis_padrao: str = "50",
                      cst_cofins_padrao: str = "50",
                      ind_nat_frt: str = "0") -> List[LinhaRegistro]:
        linhas: List[LinhaRegistro] = []
        nl = nl_inicio
        
        cst_p = dados.cst_pis or cst_pis_padrao
        bc_p = dados.vl_bc_pis or dados.vl_serv or dados.vl_doc
        aliq_p = dados.aliq_pis or aliq_pis_padrao
        vl_p = dados.vl_pis
        
        if not vl_p and bc_p and aliq_p:
            try:
                vl_p = cls._fmt_valor(str(
                    Decimal(bc_p.replace(",", ".")) *
                    Decimal(aliq_p.replace(",", ".")) / Decimal("100")
                ))
            except:
                vl_p = ""
        
        cst_c = dados.cst_cofins or cst_cofins_padrao
        bc_c = dados.vl_bc_cofins or dados.vl_serv or dados.vl_doc
        aliq_c = dados.aliq_cofins or aliq_cofins_padrao
        vl_c = dados.vl_cofins
        
        if not vl_c and bc_c and aliq_c:
            try:
                vl_c = cls._fmt_valor(str(
                    Decimal(bc_c.replace(",", ".")) *
                    Decimal(aliq_c.replace(",", ".")) / Decimal("100")
                ))
            except:
                vl_c = ""
        
        campos_d100 = [
            "D100",
            dados.ind_oper,
            dados.ind_emit,
            cod_part,
            dados.modelo,
            dados.cod_sit,
            dados.serie,
            "",
            dados.num_doc,
            dados.chave,
            dados.dt_emissao,
            dados.dt_acesso or dados.dt_emissao,
            dados.tipo_cte,
            "",
            dados.vl_doc,
            dados.vl_desc or "",
            ind_nat_frt,
            dados.vl_serv,
            dados.vl_bc_icms,
            dados.aliq_icms,
            dados.vl_icms,
            "",
            "",
            "",
            cst_p,
            bc_p,
            aliq_p,
            vl_p,
            cst_c,
            bc_c,
            aliq_c,
            vl_c,
        ]
        
        linha_raw = "|" + "|".join(campos_d100) + "|"
        linhas.append(LinhaRegistro(nl, "D", "D100", campos_d100, linha_raw))
        nl += 1
        
        vl_item_frt = dados.vl_serv or dados.vl_doc
        campos_d101 = [
            "D101",
            ind_nat_frt,
            vl_item_frt,
            cst_p,
            "05",
            bc_p,
            aliq_p,
            vl_p,
            "",
        ]
        linha_raw = "|" + "|".join(campos_d101) + "|"
        linhas.append(LinhaRegistro(nl, "D", "D101", campos_d101, linha_raw))
        nl += 1
        
        campos_d105 = [
            "D105",
            ind_nat_frt,
            vl_item_frt,
            cst_c,
            "05",
            bc_c,
            aliq_c,
            vl_c,
            "",
        ]
        linha_raw = "|" + "|".join(campos_d105) + "|"
        linhas.append(LinhaRegistro(nl, "D", "D105", campos_d105, linha_raw))
        
        return linhas


# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  DEMOS  ████████████████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

class DemosSPED:
    """Dados de demonstração para teste."""
    
    ICMS = """\
|0000|015|0|01012024|31012024|MOVEIS PREMIUM LTDA|12345678000199||SP|123456789|3550308|||A|1|
|0001|0|
|0150|001|FORNECEDOR MADEIRAS SA|105|11222333000155||IE999999|3550308|||RUA DAS MADEIRAS|100||CENTRO|
|0150|002|CLIENTE VAREJO LTDA|105|44555666000177||IE888888|3550308|||AV COMERCIO|200||VILA|
|0190|UN|UNIDADE|
|0200|MOV001|PAINEL MDF 18MM|||M2|00|9403990010|||12,00|
|0200|MOV002|PORTA MDF 210X90|||UN|00|9403990010|||12,00|
|0990|8|
|C001|0|
|C100|0|0|001|55|00|001|000001|43240112345678000199550010000001011000000001|01012024|03012024|10000,00|0|0,00|0,00|10000,00|0|0,00|0,00|0,00|0,00|1200,00|0,00|0,00|0,00|0,00|
|C170|1|MOV001|PAINEL MDF 18MM|50,000|M2|5000,00|0,00|0|000|5102||5000,00|12,00|600,00|0,00|0,00|0,00||||||01|5000,00|1,65|82,50|01|5000,00|7,60|380,00|
|C170|2|MOV002|PORTA MDF 210X90|10,000|UN|5000,00|0,00|0|000|5102||5000,00|12,00|600,00|0,00|0,00|0,00||||||01|5000,00|1,65|82,50|01|5000,00|7,60|380,00|
|C190|000|5102|12,00|10000,00|10000,00|1200,00|0,00|0,00|0,00||
|C100|0|0|001|55|00|001|000003|43240112345678000199550010000003031000000003|10012024|12012024|6000,00|0|0,00|0,00|6000,00|0|0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|
|C170|1|MOV001|PAINEL MDF 18MM|30,000|M2|3000,00|0,00|0|000|5102||||0,00|0,00|0,00||||||01||1,65||01||7,60||
|C170|2|MOV002|PORTA MDF 210X90|6,000|UN|3000,00|0,00|0|000|5102||||0,00|0,00|0,00||||||01||1,65||01||7,60||
|C190|000|5102|12,00|6000,00|0,00|0,00|0,00|0,00|0,00||
|C990|10|
|E001|0|
|E110|1200,00|0,00|1200,00|0,00|0,00|0,00|0,00|0,00|0,00|1200,00|0,00|1200,00|0,00|0,00|
|E990|3|
|9001|0|
|9900|0000|1|
|9900|C100|2|
|9900|C170|4|
|9900|C190|2|
|9990|5|
|9999|30|
"""
    
    CONTRIB = """\
|0000|006|0|01012024|31012024|MOVEIS PREMIUM LTDA|12345678000199||SP|123456789|3550308|||A|1|
|0001|0|
|0150|001|FORNECEDOR MADEIRAS SA|105|11222333000155||IE999999|3550308|||RUA DAS MADEIRAS|100||CENTRO|
|0150|002|CLIENTE VAREJO LTDA|105|44555666000177||IE888888|3550308|||AV COMERCIO|200||VILA|
|0190|UN|UNIDADE|
|0200|MOV001|PAINEL MDF 18MM|||M2|00|9403990010|||12,00|
|0200|MOV002|PORTA MDF 210X90|||UN|00|9403990010|||12,00|
|0990|8|
|A001|0|
|A010|12345678000199|
|A100|1|1|002|00|NF|001|000010||01012024|01012024|5000,00|0,00|5000,00|1,65|82,50|5000,00|7,60|380,00|
|A170|1|SERVICO MANUTENCAO||5000,00|0,00|05|1|01|5000,00|1,65|82,50|01|5000,00|7,60|380,00||
|A990|4|
|C001|0|
|C010|0|
|C100|0|0|001|55|00|001|000001|43240112345678000199550010000001011000000001|01012024|03012024|10000,00|0|0,00|0,00|10000,00|0|0,00|0,00|0,00|0,00|1200,00|0,00|165,00|760,00|
|C170|1|MOV001|PAINEL MDF 18MM|50,000|M2|5000,00|0,00|0|000|5102||5000,00|12,00|600,00|0,00|0,00|0,00||||||01|5000,00|1,65|82,50|01|5000,00|7,60|380,00|
|C170|2|MOV002|PORTA MDF 210X90|10,000|UN|5000,00|0,00|0|000|5102||5000,00|12,00|600,00|0,00|0,00|0,00||||||01|5000,00|1,65|82,50|01|5000,00|7,60|380,00|
|C185|1|MOV001|01|01|5000,00|1,65|82,50|01|5000,00|7,60|380,00|
|C185|2|MOV002|01|01|5000,00|1,65|82,50|01|5000,00|7,60|380,00|
|C100|0|0|001|55|00|001|000002|43240112345678000199550010000002021000000002|10012024|12012024|6000,00|0|0,00|0,00|6000,00|0|0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|
|C170|1|MOV001|PAINEL MDF 18MM|30,000|M2|3000,00|0,00|0|000|5102||||0,00|0,00|0,00||||||01||1,65||01||7,60||
|C170|2|MOV002|PORTA MDF 210X90|6,000|UN|3000,00|0,00|0|000|5102||||0,00|0,00|0,00||||||01||1,65||01||7,60||
|C990|13|
|D001|0|
|D100|0|001|57|00|001|CT000001||01012024|01012024|2000,00||||||||||01|2000,00|1,65|33,00|01|2000,00|7,60|152,00|
|D101|0|2000,00|50|05|2000,00|1,65|33,00||
|D105|0|2000,00|50|05|2000,00|7,60|152,00||
|D990|5|
|F001|0|
|F100|1|001|05012024|3000,00|01|1|3000,00|1,65|49,50|3000,00|7,60|228,00|
|F990|3|
|M001|0|
|M100|01|0|20000,00|1,65|0|330,00|0,00|0,00|0,00|330,00|1|330,00|
|M200|0,00|330,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|
|M210|01|21000,00|20000,00|1,65|0|330,00|0,00|0,00|0,00|0,00|330,00|
|M500|01|0|20000,00|7,60|0|1520,00|0,00|0,00|0,00|1520,00|1|1520,00|
|M600|0,00|1520,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|
|M610|04|21000,00|20000,00|7,60|0|1520,00|0,00|0,00|0,00|0,00|1520,00|
|M990|8|
|9001|0|
|9900|0000|1|
|9900|A100|1|
|9900|A170|1|
|9900|C100|2|
|9900|C170|4|
|9900|C185|2|
|9900|D100|1|
|9900|F100|1|
|9900|M100|1|
|9900|M200|1|
|9900|M500|1|
|9900|M600|1|
|9990|14|
|9999|55|
"""


# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  INTERFACE DE USUÁRIO  ████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

class UI:
    """Componentes da interface de usuário."""
    
    PAGINAS = [
        "📊 Dashboard",
        "📂 Upload",
        "🗂️ Blocos",
        "📋 Registros",
        "🧾 Notas Fiscais",
        "📦 Itens (C170/A170)",
        "💰 PIS/COFINS",
        "⚠️ Inconsistências",
        "🔧 Correções em Massa",
        "✏️ Editor Manual",
        "📤 Exportação",
        "📜 Log de Auditoria",
        "⚙️ Motor de Regras",
        "🚚 Importar CT-e",
    ]
    
    @staticmethod
    def css():
        st.markdown(f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        
        html, body, [class*="css"] {{
            font-family: 'Inter', sans-serif;
        }}
        
        section[data-testid="stSidebar"] {{
            background: {CORES["primaria"]} !important;
            color: #fff;
        }}
        
        section[data-testid="stSidebar"] .stRadio label {{
            color: #BDD6F5 !important;
            font-size: 0.88rem;
        }}
        
        section[data-testid="stSidebar"] .stRadio label:hover {{
            color: #fff !important;
        }}
        
        section[data-testid="stSidebar"] .stRadio [data-testid="stWidgetLabel"] {{
            display: none;
        }}
        
        .sec {{
            font-size: 1.05rem;
            font-weight: 700;
            color: {CORES["primaria"]};
            border-left: 4px solid {CORES["secundaria"]};
            padding-left: 10px;
            margin: 16px 0 10px 0;
        }}
        
        .stButton > button {{
            border-radius: 6px;
            font-weight: 600;
            font-size: 0.87rem;
        }}
        
        hr {{
            border-color: {CORES["borda"]};
        }}
        
        div[data-testid="metric-container"] {{
            background: {CORES["fundo"]};
            border: 1px solid {CORES["borda"]};
            border-radius: 8px;
            padding: 10px 14px;
        }}
        </style>
        """, unsafe_allow_html=True)
    
    @staticmethod
    def card(col, label: str, valor: Any, classe: str = ""):
        cores = {
            "critico": CORES["critico"],
            "aviso": CORES["aviso"],
            "ok": CORES["ok"],
            "": CORES["primaria"],
        }
        cor = cores.get(classe, CORES["primaria"])
        
        col.markdown(f"""
        <div style="background:#fff;border:1px solid {CORES["borda"]};border-radius:8px;
                    padding:12px 14px;text-align:center;box-shadow:0 2px 6px rgba(26,58,92,.07);">
            <div style="font-size:.68rem;color:#5A7398;font-weight:700;text-transform:uppercase;
                        letter-spacing:.04em;">{label}</div>
            <div style="font-size:1.55rem;font-weight:800;color:{cor};margin-top:3px;">{valor}</div>
        </div>
        """, unsafe_allow_html=True)
    
    @staticmethod
    def sidebar() -> str:
        with st.sidebar:
            st.markdown(f"""
            <div style="padding:10px 0 18px;text-align:center;">
                <span style="font-size:1.4rem;font-weight:800;color:#fff;">🔍 SPED Auditor</span><br>
                <span style="font-size:.68rem;color:#7BAFD4;">EFD ICMS/IPI + Contribuições  v3.1</span>
            </div>
            """, unsafe_allow_html=True)
            
            res = st.session_state.get("resultado")
            if res:
                meta = res.metadados
                tipo_badge = "🟦 ICMS/IPI" if res.tipo_arquivo == TIPO_EFD_ICMS else "🟩 CONTRIBUIÇÕES"
                st.markdown(f"""
                <div style="background:#1A3A5C;border-radius:6px;padding:9px 11px;margin-bottom:14px;">
                    <div style="font-size:.63rem;color:#7BAFD4;font-weight:700;">{tipo_badge}</div>
                    <div style="font-size:.8rem;color:#fff;margin-top:3px;">{(meta.nome_empresa or '—')[:26]}</div>
                    <div style="font-size:.67rem;color:#BDD6F5;">CNPJ: {meta.cnpj or '—'}</div>
                    <div style="font-size:.67rem;color:#BDD6F5;">{meta.periodo_apuracao or '—'}</div>
                    <div style="font-size:.67rem;color:#BDD6F5;">Blocos: {', '.join(meta.blocos_presentes)}</div>
                </div>
                """, unsafe_allow_html=True)
            
            st.markdown("**Navegação**")
            pagina = st.radio("Navegação", UI.PAGINAS, label_visibility="collapsed")
            
            ed = st.session_state.get("editor")
            if ed and ed.auditoria.total():
                st.markdown(f"""
                <div style="color:#7BAFD4;font-size:.71rem;margin-top:8px;">
                    ✏️ {ed.auditoria.total()} alteração(ões)
                </div>
                """, unsafe_allow_html=True)
        
        return pagina
    
    @staticmethod
    def _cols_rel(df: pd.DataFrame) -> List[str]:
        base = ["numero_linha", "bloco", "registro"]
        extras = [
            c for c in df.columns
            if not c.startswith("campo_") and c not in base + ["linha_original"]
            and df[c].astype(str).str.strip().ne("").any()
        ]
        return base + extras[:30]
    
    @staticmethod
    def _metricas(df: pd.DataFrame, tipo: str) -> Dict[str, int]:
        metrics = {
            "total_linhas": len(df),
            "total_c100": int((df["registro"] == "C100").sum()) if "registro" in df.columns else 0,
            "total_c170": int((df["registro"] == "C170").sum()) if "registro" in df.columns else 0,
            "total_a170": int((df["registro"] == "A170").sum()) if "registro" in df.columns else 0,
            "total_d100": int((df["registro"] == "D100").sum()) if "registro" in df.columns else 0,
        }
        
        df17 = df[df["registro"] == "C170"] if "registro" in df.columns else pd.DataFrame()
        
        def sem(col):
            if col in df17.columns:
                return int(df17[col].apply(
                    lambda x: x is None or (isinstance(x, float) and (x == 0 or x != x))
                ).sum())
            return 0
        
        metrics["sem_base_icms"] = sem("VL_BC_ICMS")
        metrics["sem_icms"] = sem("VL_ICMS")
        metrics["sem_base_pis"] = sem("VL_BC_PIS")
        metrics["sem_pis"] = sem("VL_PIS")
        metrics["sem_base_cof"] = sem("VL_BC_COFINS")
        metrics["sem_cof"] = sem("VL_COFINS")
        
        return metrics


# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  PÁGINAS  ██████████████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

class Paginas:
    """Implementação das páginas da aplicação."""
    
    @staticmethod
    def upload():
        st.markdown('<div class="sec">Upload do Arquivo SPED</div>', unsafe_allow_html=True)
        
        c1, c2 = st.columns([3, 2])
        
        with c1:
            arquivo = st.file_uploader("Selecione o arquivo SPED (.txt)", type=["txt"])
            tipo_demo = st.radio(
                "Demo:",
                ["EFD ICMS/IPI", "EFD Contribuições (PIS/COFINS)"],
                horizontal=True,
                key="td"
            )
            usar_demo = st.checkbox("📁 Usar arquivo de demonstração", value=not bool(arquivo))
            usuario = st.text_input("Usuário da sessão:", "analista")
            
            if st.button("▶ Processar Arquivo", type="primary"):
                with st.spinner("Analisando arquivo SPED…"):
                    try:
                        if usar_demo or not arquivo:
                            conteudo = (DemosSPED.ICMS if "ICMS" in tipo_demo else DemosSPED.CONTRIB).encode("utf-8")
                            st.info(f"Demo carregado: {tipo_demo}")
                        else:
                            conteudo = arquivo.read()
                        
                        resultado = ParserSPED.parse(conteudo)
                        st.session_state.update({
                            "resultado": resultado,
                            "linhas_orig": list(resultado.linhas),
                            "editor": EditorSped(resultado.df),
                            "_usuario": usuario,
                        })
                        
                        regras = st.session_state.get("regras") or GerenciadorRegras.carregar()
                        df_inc = ValidadorFiscal.validar(resultado.df, regras, resultado.tipo_arquivo)
                        st.session_state["df_inc"] = df_inc
                        
                        st.success(f"✅ {resultado.metadados.total_linhas:,} linhas | {resultado.tipo_arquivo.replace('_', ' ')} | {len(df_inc)} inconsistência(s)")
                        
                        if resultado.erros:
                            with st.expander(f"⚠️ {len(resultado.erros)} aviso(s) do parser"):
                                for e in resultado.erros[:20]:
                                    st.text(e)
                    
                    except Exception as e:
                        st.error(f"Erro: {e}")
        
        with c2:
            st.markdown("**Tipos suportados:**")
            st.markdown("- ✅ **EFD ICMS/IPI** — Blocos A B C D E G H K")
            st.markdown("- ✅ **EFD Contribuições** — Blocos A C D F M P")
            st.markdown("- 🔄 ECD/ECF *(arquitetura pronta)*")
            st.markdown("**Encoding:** UTF-8 · Latin-1 · CP1252")
            
            res = st.session_state.get("resultado")
            if res:
                meta = res.metadados
                tipo_label = "🟦 EFD ICMS/IPI" if res.tipo_arquivo == TIPO_EFD_ICMS else "🟩 EFD Contribuições"
                st.success(f"""
                **{tipo_label}**
                
                **{meta.nome_empresa}**
                
                CNPJ: `{meta.cnpj}`
                
                Período: {meta.periodo_apuracao}
                
                Blocos: `{', '.join(meta.blocos_presentes)}`
                
                Linhas: `{meta.total_linhas:,}`
                """)
    
    @staticmethod
    def dashboard():
        st.markdown('<div class="sec">Dashboard — Visão Geral</div>', unsafe_allow_html=True)
        
        res = st.session_state.get("resultado")
        ed = st.session_state.get("editor")
        df_inc = st.session_state.get("df_inc", pd.DataFrame())
        
        if not res:
            st.info("Carregue um arquivo em **Upload**.")
            return
        
        df = ed.df if ed else res.df
        meta = res.metadados
        metricas = UI._metricas(df, res.tipo_arquivo)
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Empresa", (meta.nome_empresa or "—")[:22])
        c2.metric("CNPJ", meta.cnpj or "—")
        c3.metric("Período", meta.periodo_apuracao or "—")
        tipo_label = "🟦 EFD ICMS/IPI" if res.tipo_arquivo == TIPO_EFD_ICMS else "🟩 EFD Contribuições"
        c4.metric("Tipo", tipo_label)
        
        st.markdown("---")
        
        cols = st.columns(8)
        UI.card(cols[0], "Total Linhas", f"{metricas['total_linhas']:,}")
        UI.card(cols[1], "NF-e (C100)", f"{metricas['total_c100']:,}")
        UI.card(cols[2], "Itens (C170)", f"{metricas['total_c170']:,}")
        UI.card(cols[3], "Serv (A170)", f"{metricas.get('total_a170', 0):,}")
        UI.card(cols[4], "Transp (D100)", f"{metricas.get('total_d100', 0):,}")
        
        ni = len(df_inc) if not df_inc.empty else 0
        nc = int((df_inc["severidade"] == "CRITICA").sum()) if not df_inc.empty and "severidade" in df_inc.columns else 0
        UI.card(cols[5], "Inconsistências", f"{ni:,}", "critico" if ni else "ok")
        UI.card(cols[6], "Críticas", f"{nc:,}", "critico" if nc else "ok")
        
        alt = len(ed.get_alterados()) if ed else 0
        UI.card(cols[7], "Alterados", f"{alt:,}", "aviso" if alt else "ok")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        cols = st.columns(6)
        UI.card(cols[0], "Sem Base ICMS", f"{metricas['sem_base_icms']:,}", "critico" if metricas['sem_base_icms'] else "ok")
        UI.card(cols[1], "Sem VL ICMS", f"{metricas['sem_icms']:,}", "critico" if metricas['sem_icms'] else "ok")
        UI.card(cols[2], "Sem Base PIS", f"{metricas['sem_base_pis']:,}", "critico" if metricas['sem_base_pis'] else "ok")
        UI.card(cols[3], "Sem VL PIS", f"{metricas['sem_pis']:,}", "critico" if metricas['sem_pis'] else "ok")
        UI.card(cols[4], "Sem Base COFINS", f"{metricas['sem_base_cof']:,}", "critico" if metricas['sem_base_cof'] else "ok")
        UI.card(cols[5], "Sem VL COFINS", f"{metricas['sem_cof']:,}", "critico" if metricas['sem_cof'] else "ok")
        
        st.markdown("---")
        
        g1, g2 = st.columns(2)
        
        with g1:
            st.markdown('<div class="sec">Registros por Bloco</div>', unsafe_allow_html=True)
            df_bloco = df["bloco"].value_counts().reset_index()
            df_bloco.columns = ["Bloco", "Qtd"]
            fig = px.bar(
                df_bloco,
                x="Bloco",
                y="Qtd",
                color="Bloco",
                color_discrete_sequence=px.colors.qualitative.Set2,
                template="plotly_white"
            )
            fig.update_layout(showlegend=False, height=240, margin=dict(l=5, r=5, t=5, b=5))
            st.plotly_chart(fig, width="stretch")
        
        with g2:
            st.markdown('<div class="sec">Inconsistências por Tipo</div>', unsafe_allow_html=True)
            if not df_inc.empty and "tipo" in df_inc.columns:
                df_tipo = df_inc["tipo"].value_counts().head(12).reset_index()
                df_tipo.columns = ["Tipo", "Qtd"]
                fig2 = px.bar(
                    df_tipo,
                    x="Qtd",
                    y="Tipo",
                    orientation="h",
                    template="plotly_white",
                    color="Qtd",
                    color_continuous_scale=["#FFF0F0", CORES["critico"]]
                )
                fig2.update_layout(
                    height=240,
                    margin=dict(l=5, r=5, t=5, b=5),
                    coloraxis_showscale=False
                )
                st.plotly_chart(fig2, width="stretch")
        
        g3, g4 = st.columns(2)
        
        with g3:
            st.markdown('<div class="sec">CST ICMS — C170</div>', unsafe_allow_html=True)
            df17 = df[df["registro"] == "C170"]
            if not df17.empty and "CST_ICMS" in df17.columns:
                df_cst = df17["CST_ICMS"].value_counts().reset_index()
                df_cst.columns = ["CST", "Qtd"]
                fig3 = px.pie(
                    df_cst,
                    names="CST",
                    values="Qtd",
                    hole=0.4,
                    template="plotly_white",
                    color_discrete_sequence=px.colors.qualitative.Pastel
                )
                fig3.update_layout(height=240, margin=dict(l=5, r=5, t=5, b=5))
                st.plotly_chart(fig3, width="stretch")
        
        with g4:
            st.markdown('<div class="sec">CST PIS — C170</div>', unsafe_allow_html=True)
            if not df17.empty and "CST_PIS" in df17.columns:
                df_pis = df17["CST_PIS"].value_counts().reset_index()
                df_pis.columns = ["CST", "Qtd"]
                fig4 = px.bar(
                    df_pis,
                    x="CST",
                    y="Qtd",
                    template="plotly_white",
                    color_discrete_sequence=[CORES["secundaria"]]
                )
                fig4.update_layout(height=240, margin=dict(l=5, r=5, t=5, b=5))
                st.plotly_chart(fig4, width="stretch")
    
    @staticmethod
    def blocos():
        st.markdown('<div class="sec">Visão por Blocos</div>', unsafe_allow_html=True)
        
        ed = st.session_state.get("editor")
        if not ed:
            st.info("Carregue um arquivo.")
            return
        
        df = ed.df
        blocos = sorted(df["bloco"].unique())
        
        cols = st.columns(min(len(blocos), 7))
        for i, bloco in enumerate(blocos):
            cols[i % 7].metric(f"Bloco {bloco}", f"{int((df['bloco'] == bloco).sum()):,}")
        
        st.markdown("---")
        
        bloco_sel = st.selectbox("Explorar bloco:", blocos)
        df_bloco = df[df["bloco"] == bloco_sel].copy()
        registros = sorted(df_bloco["registro"].unique())
        reg_sel = st.multiselect("Filtrar registros:", registros, default=registros[:6])
        
        if reg_sel:
            df_bloco = df_bloco[df_bloco["registro"].isin(reg_sel)]
        
        busca = st.text_input("Buscar em qualquer campo:")
        if busca:
            mask = df_bloco.astype(str).apply(
                lambda c: c.str.contains(busca, case=False)
            ).any(axis=1)
            df_bloco = df_bloco[mask]
        
        cols_exib = UI._cols_rel(df_bloco)
        st.dataframe(df_bloco[cols_exib].fillna("").head(500), width="stretch", height=420)
        st.caption(f"{len(df_bloco):,} registros")
    
    @staticmethod
    def registros():
        st.markdown('<div class="sec">Visão por Registros</div>', unsafe_allow_html=True)
        
        ed = st.session_state.get("editor")
        if not ed:
            st.info("Carregue um arquivo.")
            return
        
        df = ed.df
        registros = sorted(df["registro"].unique())
        
        c1, c2 = st.columns([2, 3])
        with c1:
            reg = st.selectbox("Registro:", registros)
        with c2:
            busca = st.text_input("Buscar:", placeholder="CST, CFOP, valor…")
        
        df_reg = df[df["registro"] == reg].copy()
        if busca:
            mask = df_reg.astype(str).apply(
                lambda c: c.str.contains(busca, case=False)
            ).any(axis=1)
            df_reg = df_reg[mask]
        
        cols_exib = UI._cols_rel(df_reg)
        st.dataframe(df_reg[cols_exib].fillna(""), width="stretch", height=440)
        st.caption(f"{len(df_reg):,} registros do tipo {reg}")
    
    @staticmethod
    def notas():
        st.markdown('<div class="sec">Notas Fiscais / Documentos (C100 · A100 · D100)</div>', unsafe_allow_html=True)
        
        ed = st.session_state.get("editor")
        if not ed:
            st.info("Carregue um arquivo.")
            return
        
        df = ed.df
        
        abas = st.tabs(["NF-e (C100)", "Serviços (A100)", "Transportes (D100)"])
        
        configs = [
            ("C100", ["numero_linha", "NUM_DOC", "DT_DOC", "COD_PART", "VL_DOC",
                      "VL_BC_ICMS", "VL_ICMS", "VL_PIS", "VL_COFINS", "COD_SIT"]),
            ("A100", ["numero_linha", "NUM_DOC", "DT_DOC", "COD_PART", "VL_DOC",
                      "VL_BC_PIS", "ALIQ_PIS", "VL_PIS", "VL_BC_COFINS", "ALIQ_COFINS", "VL_COFINS"]),
            ("D100", ["numero_linha", "NUM_DOC", "DT_DOC", "COD_PART", "VL_DOC",
                      "VL_BC_ICMS", "VL_ICMS", "VL_BC_PIS", "VL_PIS", "VL_COFINS"]),
        ]
        
        for tab, (reg, campos_vis) in zip(abas, configs):
            with tab:
                df_doc = df[df["registro"] == reg].copy()
                if df_doc.empty:
                    st.info(f"Nenhum {reg} encontrado.")
                    continue
                
                cols = [c for c in campos_vis if c in df_doc.columns]
                
                c1, c2 = st.columns(2)
                with c1:
                    parceiros = ["Todos"] + sorted(df_doc["COD_PART"].dropna().unique().tolist()) if "COD_PART" in df_doc.columns else ["Todos"]
                    parc_sel = st.selectbox("Parceiro:", parceiros, key=f"pf_{reg}")
                    if parc_sel != "Todos" and "COD_PART" in df_doc.columns:
                        df_doc = df_doc[df_doc["COD_PART"] == parc_sel]
                
                st.dataframe(df_doc[cols].fillna(""), width="stretch", height=380)
                st.caption(f"{len(df_doc):,} documentos")
                
                nums = [c for c in ["VL_DOC", "VL_ICMS", "VL_PIS", "VL_COFINS"] if c in df_doc.columns]
                if nums:
                    cols_metric = st.columns(len(nums))
                    for i, c in enumerate(nums):
                        cols_metric[i].metric(c, Utilitarios.fmt_brl(pd.to_numeric(df_doc[c], errors="coerce").sum()))
    
    @staticmethod
    def itens():
        st.markdown('<div class="sec">Itens de NF-e (C170) e Serviços (A170)</div>', unsafe_allow_html=True)
        
        ed = st.session_state.get("editor")
        if not ed:
            st.info("Carregue um arquivo.")
            return
        
        df = ed.df
        df_inc = st.session_state.get("df_inc", pd.DataFrame())
        
        abas = st.tabs(["C170 — Itens NF-e", "A170 — Itens Serviço", "C185 — PIS/COFINS Itens", "F100 — Outras Op."])
        
        configs = [
            ("C170", ["numero_linha", "COD_ITEM", "DESCR_COMPL", "QTD", "VL_ITEM",
                      "CST_ICMS", "CFOP", "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS",
                      "CST_PIS", "VL_BC_PIS", "ALIQ_PIS", "VL_PIS",
                      "CST_COFINS", "VL_BC_COFINS", "ALIQ_COFINS", "VL_COFINS"], "CST_ICMS"),
            ("A170", ["numero_linha", "COD_ITEM", "DESCR_COMPL", "VL_ITEM",
                      "CST_PIS", "VL_BC_PIS", "ALIQ_PIS", "VL_PIS",
                      "CST_COFINS", "VL_BC_COFINS", "ALIQ_COFINS", "VL_COFINS"], "CST_PIS"),
            ("C185", ["numero_linha", "COD_ITEM", "CST_PIS", "VL_BC_PIS", "ALIQ_PIS", "VL_PIS",
                      "CST_COFINS", "VL_BC_COFINS", "ALIQ_COFINS", "VL_COFINS"], "CST_PIS"),
            ("F100", ["numero_linha", "COD_PART", "DT_OPER", "VL_OPER",
                      "VL_BC_PIS", "ALIQ_PIS", "VL_PIS",
                      "VL_BC_COFINS", "ALIQ_COFINS", "VL_COFINS"], ""),
        ]
        
        for tab, (reg, campos_vis, campo_cst) in zip(abas, configs):
            with tab:
                df_reg = df[df["registro"] == reg].copy()
                if df_reg.empty:
                    st.info(f"Nenhum {reg} encontrado.")
                    continue
                
                c1, c2, c3 = st.columns(3)
                
                with c1:
                    if campo_cst and campo_cst in df_reg.columns:
                        csts = ["Todos"] + sorted(df_reg[campo_cst].dropna().unique().tolist())
                        cst_sel = st.selectbox("CST:", csts, key=f"cst_{reg}")
                        if cst_sel != "Todos":
                            df_reg = df_reg[df_reg[campo_cst] == cst_sel]
                
                with c2:
                    if "CFOP" in df_reg.columns:
                        cfops = ["Todos"] + sorted(df_reg["CFOP"].dropna().unique().tolist())
                        cfop_sel = st.selectbox("CFOP:", cfops, key=f"cfop_{reg}")
                        if cfop_sel != "Todos":
                            df_reg = df_reg[df_reg["CFOP"] == cfop_sel]
                
                with c3:
                    apenas_inc = st.checkbox("Apenas com inconsistências", key=f"ap_{reg}")
                    if apenas_inc and not df_inc.empty:
                        linhas_inc = set(df_inc["numero_linha"].tolist())
                        df_reg = df_reg[df_reg["numero_linha"].isin(linhas_inc)]
                
                cols = [c for c in campos_vis if c in df_reg.columns]
                st.dataframe(df_reg[cols].fillna(""), width="stretch", height=400)
                st.caption(f"{len(df_reg):,} registros")
    
    @staticmethod
    def pis_cofins():
        st.markdown('<div class="sec">PIS / COFINS — Apuração (Bloco M)</div>', unsafe_allow_html=True)
        
        ed = st.session_state.get("editor")
        if not ed:
            st.info("Carregue um arquivo.")
            return
        
        df = ed.df
        res = st.session_state.get("resultado")
        
        if res and res.tipo_arquivo != TIPO_EFD_CONTRIB:
            st.warning("⚠️ Arquivo identificado como EFD ICMS/IPI. Bloco M pode estar ausente.")
        
        abas = st.tabs([
            "M100 — Créditos PIS", "M200 — Apuração PIS", "M210 — Detalhamento PIS",
            "M500 — Créditos COFINS", "M600 — Apuração COFINS", "M610 — Detalhamento COFINS",
            "P100 — Contrib. Previdenciária"
        ])
        
        for tab, reg in zip(abas, ["M100", "M200", "M210", "M500", "M600", "M610", "P100"]):
            with tab:
                df_reg = df[df["registro"] == reg].copy()
                if df_reg.empty:
                    st.info(f"Nenhum registro {reg} encontrado.")
                    continue
                
                cols_exib = UI._cols_rel(df_reg)
                st.dataframe(df_reg[cols_exib].fillna(""), width="stretch", height=350)
                st.caption(f"{len(df_reg):,} registros {reg}")
                
                if reg in ("M200", "M600"):
                    cols_tot = [c for c in ["VL_TOT_CONT_NC_PER", "VL_TOT_CRED_DESC", "VL_TOT_CONT_REC",
                                            "VL_CONT_NC_REC", "VL_CONT_CUM_REC"] if c in df_reg.columns]
                    if cols_tot:
                        cols_metric = st.columns(len(cols_tot))
                        for i, c in enumerate(cols_tot):
                            cols_metric[i].metric(c, Utilitarios.fmt_brl(pd.to_numeric(df_reg[c], errors="coerce").sum()))
    
    @staticmethod
    def inconsistencias():
        st.markdown('<div class="sec">Inconsistências Fiscais</div>', unsafe_allow_html=True)
        
        df_inc = st.session_state.get("df_inc", pd.DataFrame())
        res = st.session_state.get("resultado")
        
        if res and (df_inc is None or df_inc.empty):
            st.success("✅ Nenhuma inconsistência detectada.")
            return
        
        if not res:
            st.info("Carregue um arquivo.")
            return
        
        c1, c2, c3, c4 = st.columns(4)
        nc = int((df_inc["severidade"] == "CRITICA").sum()) if "severidade" in df_inc.columns else 0
        na = int((df_inc["severidade"] == "AVISO").sum()) if "severidade" in df_inc.columns else 0
        UI.card(c1, "Total", len(df_inc))
        UI.card(c2, "Críticas", nc, "critico")
        UI.card(c3, "Avisos", na, "aviso")
        UI.card(c4, "Corrigidos", int(df_inc["corrigido"].sum()) if "corrigido" in df_inc.columns else 0, "ok")
        
        st.markdown("---")
        
        c1, c2, c3, c4, c5 = st.columns(5)
        
        sev_opts = ["Todos"] + sorted(df_inc["severidade"].unique().tolist()) if "severidade" in df_inc.columns else ["Todos"]
        tipo_opts = ["Todos"] + sorted(df_inc["tipo"].unique().tolist()) if "tipo" in df_inc.columns else ["Todos"]
        trib_opts = ["Todos", "ICMS", "PIS", "COFINS", "IPI", "GERAL"]
        cst_opts = ["Todos"] + sorted(df_inc["cst"].dropna().unique().tolist()) if "cst" in df_inc.columns else ["Todos"]
        
        with c1:
            sev_filtro = st.selectbox("Severidade:", sev_opts)
        with c2:
            tipo_filtro = st.selectbox("Tipo:", tipo_opts)
        with c3:
            trib_filtro = st.selectbox("Tributo:", trib_opts)
        with c4:
            cst_filtro = st.selectbox("CST:", cst_opts)
        with c5:
            busca = st.text_input("Buscar descrição:")
        
        df_filtrado = df_inc.copy()
        if sev_filtro != "Todos":
            df_filtrado = df_filtrado[df_filtrado["severidade"] == sev_filtro]
        if tipo_filtro != "Todos":
            df_filtrado = df_filtrado[df_filtrado["tipo"] == tipo_filtro]
        if trib_filtro != "Todos":
            df_filtrado = df_filtrado[df_filtrado["tipo"].str.contains(trib_filtro, na=False)]
        if cst_filtro != "Todos":
            df_filtrado = df_filtrado[df_filtrado["cst"] == cst_filtro]
        if busca:
            df_filtrado = df_filtrado[df_filtrado["descricao"].astype(str).str.contains(busca, case=False)]
        
        def cor_linha(row):
            if row.get("severidade") == "CRITICA":
                return ["background-color:#FFF0F0"] * len(row)
            if row.get("severidade") == "AVISO":
                return ["background-color:#FFFDE8"] * len(row)
            return [""] * len(row)
        
        cols_show = [c for c in df_filtrado.columns if c != "linha_original"]
        st.dataframe(
            df_filtrado[cols_show].style.apply(cor_linha, axis=1),
            width="stretch",
            height=460
        )
        st.caption(f"{len(df_filtrado):,} inconsistências")
        
        if st.button("🔄 Revalidar Arquivo"):
            ed = st.session_state.get("editor")
            regras = st.session_state.get("regras")
            if ed and regras:
                st.session_state["df_inc"] = ValidadorFiscal.validar(ed.df, regras, res.tipo_arquivo)
                st.rerun()
    
    @staticmethod
    def massa():
        st.markdown('<div class="sec">Correções em Massa</div>', unsafe_allow_html=True)
        
        ed = st.session_state.get("editor")
        regras = st.session_state.get("regras", [])
        df_inc = st.session_state.get("df_inc", pd.DataFrame())
        usuario = st.session_state.get("_usuario", "analista")
        
        if not ed:
            st.info("Carregue um arquivo.")
            return
        
        c1, c2 = st.columns(2)
        with c1:
            tributo = st.selectbox("Tributo a corrigir:", ["ICMS", "PIS", "COFINS"], key="m_trib")
        with c2:
            reg_alvo = st.selectbox("Registro alvo:", ["C170", "A170", "C185", "D100", "F100"], key="m_reg")
        
        mapa_campos = {
            "ICMS": {"cst": "CST_ICMS", "bc": "VL_BC_ICMS", "aliq": "ALIQ_ICMS", "vl": "VL_ICMS"},
            "PIS": {"cst": "CST_PIS", "bc": "VL_BC_PIS", "aliq": "ALIQ_PIS", "vl": "VL_PIS"},
            "COFINS": {"cst": "CST_COFINS", "bc": "VL_BC_COFINS", "aliq": "ALIQ_COFINS", "vl": "VL_COFINS"},
        }
        campos = mapa_campos[tributo]
        
        df_reg = ed.df[ed.df["registro"] == reg_alvo].copy()
        if df_reg.empty:
            st.warning(f"Nenhum {reg_alvo} encontrado.")
            return
        
        st.markdown("**Filtros de seleção**")
        c1, c2, c3, c4 = st.columns(4)
        
        with c1:
            cst_opts = ["Todos"] + sorted(df_reg[campos["cst"]].dropna().unique().tolist()) if campos["cst"] in df_reg.columns else ["Todos"]
            cst_filtro = st.selectbox(f"CST {tributo}:", cst_opts, key="mc_cst")
        with c2:
            cfop_opts = ["Todos"] + sorted(df_reg["CFOP"].dropna().unique().tolist()) if "CFOP" in df_reg.columns else ["Todos"]
            cfop_filtro = st.selectbox("CFOP:", cfop_opts, key="mc_cfop")
        with c3:
            tipo_opts = ["Todos"] + (sorted(df_inc["tipo"].unique().tolist()) if not df_inc.empty and "tipo" in df_inc.columns else [])
            tipo_filtro = st.selectbox("Tipo inconsistência:", tipo_opts, key="mc_tipo")
        with c4:
            apenas_inc = st.checkbox("Apenas com inconsistência", value=True, key="mc_ap")
        
        df_alvo = df_reg.copy()
        if cst_filtro != "Todos" and campos["cst"] in df_alvo.columns:
            df_alvo = df_alvo[df_alvo[campos["cst"]] == cst_filtro]
        if cfop_filtro != "Todos" and "CFOP" in df_alvo.columns:
            df_alvo = df_alvo[df_alvo["CFOP"] == cfop_filtro]
        if apenas_inc and not df_inc.empty:
            linhas_inc = set(df_inc[df_inc["tipo"] == tipo_filtro]["numero_linha"].tolist() if tipo_filtro != "Todos" else df_inc["numero_linha"].tolist())
            df_alvo = df_alvo[df_alvo["numero_linha"].isin(linhas_inc)]
        
        st.info(f"**{len(df_alvo)} item(ns) selecionado(s) em {reg_alvo} — tributo {tributo}**")
        
        cols_vis = [c for c in ["numero_linha", campos["cst"], "CFOP", "VL_ITEM", campos["bc"], campos["aliq"], campos["vl"]] if c in df_alvo.columns]
        st.dataframe(df_alvo[cols_vis].fillna("").head(200), width="stretch", height=180)
        
        if df_alvo.empty:
            return
        
        nls = df_alvo["numero_linha"].tolist()
        
        st.markdown("---")
        tab1, tab2, tab3, tab4 = st.tabs(["📐 Preencher Base", "📊 Preencher Alíquota", "🧮 Recalcular Valor", "👁️ Prévia"])
        
        with tab1:
            fonte = st.radio("Origem:", ["VL_ITEM do item", "VL_DOC do documento", "Valor manual"], key="fb")
            valor_manual = 0.0
            if "manual" in fonte:
                valor_manual = st.number_input("Valor (R$):", min_value=0.0, step=0.01, key="bman")
            motivo = st.text_input("Motivo:", f"Base {tributo} preenchida via {fonte}", key="mb")
            
            if st.button(f"▶ Aplicar Base {tributo}", type="primary"):
                count = 0
                for _, row in df_alvo.iterrows():
                    nl = int(row["numero_linha"])
                    if "manual" in fonte:
                        valor = Utilitarios.to_sped_str(valor_manual)
                    elif "VL_ITEM" in fonte:
                        vi = Utilitarios.to_float(row.get("VL_ITEM"))
                        valor = Utilitarios.to_sped_str(vi) if vi else ""
                    else:
                        vd = Utilitarios.to_float(row.get("VL_DOC"))
                        valor = Utilitarios.to_sped_str(vd) if vd else ""
                    
                    if valor and ed.editar(nl, campos["bc"], valor, motivo, "", usuario):
                        count += 1
                
                st.success(f"✅ {count} base(s) de {tributo} preenchidas.")
        
        with tab2:
            aliq_padrao = {"ICMS": 12.0, "PIS": 1.65, "COFINS": 7.6}[tributo]
            aliq_val = st.number_input("Alíquota (%):", 0.0, 100.0, aliq_padrao, 0.01, key="maliq")
            motivo2 = st.text_input("Motivo:", f"Alíquota {tributo} padrão aplicada", key="ma2")
            
            if st.button(f"▶ Aplicar Alíquota {tributo}", type="primary"):
                count = ed.massa(nls, campos["aliq"], Utilitarios.to_sped_str(aliq_val), motivo2, "", usuario)
                st.success(f"✅ {count} alíquota(s) de {tributo} preenchida(s).")
        
        with tab3:
            motivo3 = st.text_input("Motivo:", f"Recálculo {tributo} = Base × Alíq ÷ 100", key="mc3")
            
            if st.button(f"▶ Recalcular {tributo}", type="primary"):
                count = ed.recalcular_massa(
                    nls, regras, tributo,
                    campos["cst"], campos["bc"], campos["aliq"], campos["vl"],
                    motivo3, usuario
                )
                st.success(f"✅ {count} valor(es) de {tributo} recalculado(s).")
        
        with tab4:
            preview = ed.preview(nls)
            if not preview.empty:
                st.dataframe(preview, width="stretch")
            else:
                st.info("Nenhuma alteração pendente.")
        
        st.markdown("---")
        if st.button("↩️ Desfazer última operação"):
            if ed.desfazer():
                st.success("✅ Desfeito.")
                st.rerun()
            else:
                st.warning("Nada para desfazer.")
    
    @staticmethod
    def editor():
        st.markdown('<div class="sec">Editor Manual de Registros</div>', unsafe_allow_html=True)
        
        ed = st.session_state.get("editor")
        usuario = st.session_state.get("_usuario", "analista")
        
        if not ed:
            st.info("Carregue um arquivo.")
            return
        
        df = ed.df
        registros = sorted(df["registro"].unique())
        
        c1, c2, c3 = st.columns([2, 2, 1])
        with c1:
            reg_sel = st.selectbox("Registro:", registros)
        
        df_reg = df[df["registro"] == reg_sel]
        with c2:
            nl_min = int(df_reg["numero_linha"].min()) if not df_reg.empty else 1
            nl_max = int(df_reg["numero_linha"].max()) if not df_reg.empty else 1
            nl = st.number_input("Linha:", nl_min, nl_max, nl_min)
        with c3:
            motivo = st.text_input("Motivo:", key="em")
        
        df_linha = df[df["numero_linha"] == nl]
        if df_linha.empty:
            st.warning("Linha não encontrada.")
            return
        
        row_atual = df_linha.iloc[0]
        row_orig = ed.df_original[ed.df_original["numero_linha"] == nl]
        row_orig = row_orig.iloc[0] if not row_orig.empty else row_atual
        
        mapa = MapaCampos.get_mapa(reg_sel)
        campos = list(mapa.keys()) if mapa else [
            c for c in df.columns
            if not c.startswith("campo_") and c not in ["numero_linha", "bloco", "registro", "linha_original"]
        ]
        
        alterados = [c for c in campos if str(row_atual.get(c, "")) != str(row_orig.get(c, ""))]
        if alterados:
            st.warning(f"⚠️ Campo(s) alterado(s): **{', '.join(alterados)}**")
        
        st.markdown(f"**Editando: {reg_sel} — Linha {nl}**")
        
        pendentes = {}
        cols = st.columns(3)
        
        for i, campo in enumerate(campos):
            valor_atual = str(row_atual.get(campo, "") or "")
            valor_orig = str(row_orig.get(campo, "") or "")
            destaque = "🔴 " if valor_atual != valor_orig else ""
            novo = cols[i % 3].text_input(f"{destaque}{campo}", value=valor_atual, key=f"ed_{campo}_{nl}")
            if novo != valor_atual:
                pendentes[campo] = novo
        
        col_a, col_b, col_c = st.columns(3)
        
        with col_a:
            if st.button("💾 Salvar", type="primary", disabled=not pendentes):
                for campo, valor in pendentes.items():
                    ed.editar(nl, campo, valor, motivo, "", usuario)
                st.success(f"✅ {len(pendentes)} campo(s) salvo(s).")
                st.rerun()
        
        with col_b:
            if st.button("↩️ Desfazer"):
                if ed.desfazer():
                    st.success("✅ Desfeito.")
                    st.rerun()
                else:
                    st.warning("Nada para desfazer.")
        
        with col_c:
            if st.button("🔄 Restaurar linha"):
                ed.restaurar([nl])
                st.success("✅ Restaurado.")
                st.rerun()
        
        st.markdown("---")
        
        diferencas = [{
            "Campo": c,
            "Original": str(row_orig.get(c, "")),
            "Atual": str(row_atual.get(c, ""))
        } for c in campos if str(row_atual.get(c, "")) != str(row_orig.get(c, ""))]
        
        if diferencas:
            st.dataframe(pd.DataFrame(diferencas), width="stretch")
        else:
            st.success("Nenhuma diferença nesta linha.")
    
    @staticmethod
    def exportacao():
        st.markdown('<div class="sec">Exportação</div>', unsafe_allow_html=True)
        
        ed = st.session_state.get("editor")
        res = st.session_state.get("resultado")
        df_inc = st.session_state.get("df_inc", pd.DataFrame())
        regras = st.session_state.get("regras", [])
        
        if not ed or not res:
            st.info("Carregue um arquivo.")
            return
        
        meta = {
            "nome_empresa": res.metadados.nome_empresa,
            "cnpj": res.metadados.cnpj,
            "periodo_apuracao": res.metadados.periodo_apuracao,
            "uf": res.metadados.uf,
            "tipo_arquivo": res.tipo_arquivo,
        }
        
        c1, c2 = st.columns(2)
        
        with c1:
            st.markdown("#### 📄 SPED TXT Corrigido")
            if st.button("Gerar SPED TXT", type="primary"):
                linhas_orig = st.session_state.get("linhas_orig", res.linhas)
                linhas_corrigidas = ParserSPED.sync_df_linhas(ed.df, linhas_orig)
                bytes_txt = ParserSPED.reconstruir(linhas_corrigidas)
                nome = f"SPED_CORRIGIDO_{meta['cnpj']}_{res.tipo_arquivo}.txt"
                st.download_button("⬇️ Baixar TXT", bytes_txt, nome, "text/plain")
        
        with c2:
            st.markdown("#### 📊 Excel de Auditoria (5 abas)")
            if st.button("Gerar Excel", type="primary"):
                df_alt = ed.get_alterados()
                df_aud = ed.auditoria.to_df()
                df_reg = pd.DataFrame([asdict(r) for r in regras])
                bytes_excel = Exportador.to_excel(
                    df_inc or pd.DataFrame(),
                    df_alt,
                    df_aud,
                    df_reg,
                    meta
                )
                nome = f"AUDITORIA_{meta['cnpj']}.xlsx"
                st.download_button(
                    "⬇️ Baixar Excel",
                    bytes_excel,
                    nome,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        st.markdown("---")
        
        c3, c4 = st.columns(2)
        
        with c3:
            st.markdown("#### 📋 CSV Inconsistências")
            if df_inc is not None and not df_inc.empty:
                st.download_button(
                    "⬇️ Baixar CSV",
                    Exportador.to_csv(df_inc),
                    "inconsistencias.csv",
                    "text/csv"
                )
            else:
                st.info("Sem inconsistências.")
        
        with c4:
            st.markdown("#### 📋 CSV Auditoria")
            df_aud = ed.auditoria.to_df()
            if not df_aud.empty:
                st.download_button(
                    "⬇️ Baixar CSV",
                    Exportador.to_csv(df_aud),
                    "log_auditoria.csv",
                    "text/csv"
                )
            else:
                st.info("Sem alterações.")
    
    @staticmethod
    def auditoria():
        st.markdown('<div class="sec">Log de Auditoria</div>', unsafe_allow_html=True)
        
        ed = st.session_state.get("editor")
        if not ed:
            st.info("Carregue um arquivo.")
            return
        
        df_aud = ed.auditoria.to_df()
        if df_aud.empty:
            st.info("Nenhuma alteração registrada.")
            return
        
        tipos = df_aud["Tipo"].value_counts().to_dict() if "Tipo" in df_aud.columns else {}
        
        c1, c2, c3 = st.columns(3)
        UI.card(c1, "Total Alterações", ed.auditoria.total())
        UI.card(c2, "Manuais", tipos.get("MANUAL", 0))
        UI.card(c3, "Massa/Auto", tipos.get("MASSA", 0) + tipos.get("AUTO", 0))
        
        st.markdown("---")
        
        c1, c2, c3 = st.columns(3)
        with c1:
            tipo_filtro = st.selectbox("Tipo:", ["Todos"] + sorted(df_aud["Tipo"].unique().tolist()) if "Tipo" in df_aud.columns else ["Todos"])
        with c2:
            campo_filtro = st.selectbox("Campo:", ["Todos"] + sorted(df_aud["Campo"].unique().tolist()) if "Campo" in df_aud.columns else ["Todos"])
        with c3:
            busca = st.text_input("Buscar motivo/regra:")
        
        df_filtrado = df_aud.copy()
        if tipo_filtro != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Tipo"] == tipo_filtro]
        if campo_filtro != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Campo"] == campo_filtro]
        if busca:
            df_filtrado = df_filtrado[
                df_filtrado["Motivo"].astype(str).str.contains(busca, case=False) |
                df_filtrado["Regra"].astype(str).str.contains(busca, case=False)
            ]
        
        st.dataframe(df_filtrado, width="stretch", height=490)
    
    @staticmethod
    def regras():
        st.markdown('<div class="sec">Motor de Regras Tributárias</div>', unsafe_allow_html=True)
        
        regras = st.session_state.get("regras") or GerenciadorRegras.carregar()
        st.session_state["regras"] = regras
        
        tab1, tab2 = st.tabs(["📋 Regras Ativas", "➕ Cadastrar / Editar"])
        
        with tab1:
            trib_filtro = st.selectbox("Filtrar por tributo:", ["Todos", "ICMS", "PIS", "COFINS", "IPI"], key="rf_t")
            df_regras = pd.DataFrame([asdict(r) for r in regras])
            
            if trib_filtro != "Todos" and "tributo" in df_regras.columns:
                df_regras = df_regras[df_regras["tributo"] == trib_filtro]
            
            st.dataframe(df_regras, width="stretch", height=420)
            
            c1, c2 = st.columns(2)
            with c1:
                if st.button("🔄 Resetar para padrão"):
                    st.session_state["regras"] = GerenciadorRegras._build_regras_padrao()
                    GerenciadorRegras.salvar(st.session_state["regras"])
                    st.success("✅ Regras padrão restauradas.")
                    st.rerun()
            
            with c2:
                id_remover = st.text_input("ID para remover:", key="rdel")
                if st.button("🗑️ Remover") and id_remover:
                    st.session_state["regras"] = [r for r in regras if r.id != id_remover]
                    GerenciadorRegras.salvar(st.session_state["regras"])
                    st.success(f"✅ Removido '{id_remover}'.")
                    st.rerun()
        
        with tab2:
            c1, c2, c3 = st.columns(3)
            
            with c1:
                novo_id = st.text_input("ID:", key="ni")
                novo_trib = st.selectbox("Tributo:", ["ICMS", "PIS", "COFINS", "IPI"], key="nt")
                novo_cst = st.text_input("CST (* = todos):", key="nc")
                novo_cfop = st.text_input("CFOP (* = todos):", value="*", key="nf")
                novo_oper = st.selectbox("Operação:", ["*", "E", "S"], key="no")
                novo_mod = st.text_input("COD_MOD (* = todos):", value="*", key="nm")
            
            with c2:
                novo_desc = st.text_area("Descrição:", key="nd")
                novo_exige_base = st.checkbox("Exige base", value=True, key="neb")
                novo_exige_aliq = st.checkbox("Exige alíquota", value=True, key="nea")
                novo_exige_valor = st.checkbox("Exige valor", value=True, key="nev")
            
            with c3:
                novo_aliq = st.number_input("Alíquota padrão (%):", 0.0, 100.0, 12.0, key="nal")
                novo_bc = st.selectbox("Campo base sugerida:", ["VL_ITEM", "VL_DOC", "VL_OPER", ""], key="nbc")
                novo_formula = st.selectbox("Fórmula:", ["base * aliq / 100", "zero"], key="nfo")
                novo_crit = st.selectbox("Criticidade:", ["critica", "aviso", "info"], key="ncr")
            
            if st.button("💾 Salvar Regra", type="primary"):
                if not novo_id or not novo_cst:
                    st.error("ID e CST obrigatórios.")
                else:
                    nova = RegraFiscal(
                        id=novo_id,
                        tributo=novo_trib,
                        cst=novo_cst,
                        cfop=novo_cfop,
                        ind_oper=novo_oper,
                        descricao=novo_desc,
                        exige_base=novo_exige_base,
                        exige_aliquota=novo_exige_aliq,
                        exige_valor=novo_exige_valor,
                        base_campo_sugerido=novo_bc,
                        aliquota_padrao=novo_aliq,
                        formula=novo_formula,
                        permite_base_zero=not novo_exige_base,
                        permite_valor_zero=not novo_exige_valor,
                        criticidade=novo_crit,
                        cod_mod=novo_mod,
                    )
                    st.session_state["regras"] = [r for r in regras if r.id != novo_id] + [nova]
                    GerenciadorRegras.salvar(st.session_state["regras"])
                    st.success(f"✅ Regra '{novo_id}' salva.")
                    st.rerun()
    
    @staticmethod
    def importar_cte():
        st.markdown('<div class="sec">Importar XML CT-e → Bloco D (D100 · D101 · D105)</div>', unsafe_allow_html=True)
        
        ed = st.session_state.get("editor")
        res = st.session_state.get("resultado")
        
        if not ed or not res:
            st.warning("⚠️ Carregue um arquivo SPED primeiro na tela **Upload**.")
            return
        
        if res.tipo_arquivo != TIPO_EFD_CONTRIB:
            st.info("ℹ️ Esta função é projetada para **EFD Contribuições**. "
                    "Arquivo atual identificado como EFD ICMS/IPI — os registros D100/D101/D105 "
                    "serão inseridos mesmo assim, mas verifique a compatibilidade com o layout.")
        
        st.markdown("""
        **Fluxo:**
        1. Faça upload dos XMLs de CT-e (um ou vários)
        2. Confira os dados extraídos na prévia
        3. Configure parâmetros de PIS/COFINS se necessário
        4. Clique em **Inserir no SPED** para adicionar os registros D100/D101/D105
        """)
        
        st.markdown("---")
        st.markdown("**1. Upload dos XMLs de CT-e**")
        
        xmls = st.file_uploader(
            "Selecione um ou mais arquivos XML de CT-e",
            type=["xml"],
            accept_multiple_files=True,
            key="cte_upload",
        )
        
        if not xmls:
            st.info("Nenhum XML selecionado.")
            return
        
        st.markdown("---")
        st.markdown("**2. Parâmetros de PIS/COFINS (aplicados quando ausentes no XML)**")
        
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        with c1:
            cst_pis_pad = st.text_input("CST PIS:", value="50", key="cte_cst_pis")
        with c2:
            aliq_pis_pad = st.text_input("Alíq. PIS (%):", value="1,65", key="cte_aliq_pis")
        with c3:
            cst_cof_pad = st.text_input("CST COFINS:", value="50", key="cte_cst_cof")
        with c4:
            aliq_cof_pad = st.text_input("Alíq. COFINS (%):", value="7,60", key="cte_aliq_cof")
        with c5:
            ind_nat_frt = st.selectbox(
                "Nat. Frete (D101/D105):",
                options=["0", "1", "2", "3", "4", "9"],
                format_func=lambda x: {
                    "0": "0-Própria do tomador",
                    "1": "1-Redespacho",
                    "2": "2-Redespacho Intermediário",
                    "3": "3-Serv. Vinculado",
                    "4": "4-Própria do emitente",
                    "9": "9-Sem frete"
                }.get(x, x),
                key="cte_nat_frt",
            )
        with c6:
            cod_part_pad = st.text_input(
                "COD_PART padrão:",
                value="",
                help="Código do participante (transportador) no SPED. Deixe vazio para preencher por CT-e.",
                key="cte_cod_part",
            )
        
        st.markdown("---")
        st.markdown("**3. Prévia dos dados extraídos**")
        
        dados_lista: List[Tuple[str, DadosCTe]] = []
        for arquivo in xmls:
            conteudo = arquivo.read()
            dados = ImportadorCTe.parse_xml(conteudo)
            dados_lista.append((arquivo.name, dados))
        
        rows_prev = []
        for nome, d in dados_lista:
            rows_prev.append({
                "Arquivo": nome,
                "Chave": d.chave[:22] + "…" if len(d.chave) > 22 else d.chave,
                "Num. CT-e": d.num_doc,
                "Série": d.serie,
                "Modelo": d.modelo,
                "DT Emissão": d.dt_emissao,
                "Emitente": d.nome_emit[:20] if d.nome_emit else d.cnpj_emit,
                "VL Doc": d.vl_doc,
                "VL Serviço": d.vl_serv,
                "CST ICMS": d.cst_icms,
                "VL ICMS": d.vl_icms,
                "CST PIS": d.cst_pis or cst_pis_pad,
                "VL PIS": d.vl_pis,
                "CST COFINS": d.cst_cofins or cst_cof_pad,
                "VL COFINS": d.vl_cofins,
                "Erros": "; ".join(d.erros) if d.erros else "✅",
            })
        
        df_prev = pd.DataFrame(rows_prev)
        
        def cor_erro(row):
            if row.get("Erros") != "✅":
                return ["background-color:#FFF0F0"] * len(row)
            return [""] * len(row)
        
        st.dataframe(
            df_prev.style.apply(cor_erro, axis=1),
            width="stretch",
            height=280,
        )
        
        erros_totais = sum(len(d.erros) for _, d in dados_lista)
        validos = [(n, d) for n, d in dados_lista if not d.erros]
        st.caption(f"{len(dados_lista)} XML(s) lido(s) · {len(validos)} válido(s) · {erros_totais} erro(s)")
        
        if not validos:
            st.error("Nenhum CT-e válido para inserir.")
            return
        
        st.markdown("---")
        st.markdown("**4. Inserir no SPED**")
        
        col_a, col_b = st.columns(2)
        with col_a:
            posicao = st.radio(
                "Posição de inserção:",
                ["Após o último D100 existente", "No final do Bloco D", "Número de linha específico"],
                key="cte_posicao",
            )
        with col_b:
            nl_especifica = st.number_input(
                "Número da linha (se específico):",
                min_value=1,
                value=1,
                step=1,
                key="cte_nl_esp",
            )
        
        usuario = st.session_state.get("_usuario", "analista")
        
        if st.button("▶ Inserir D100/D101/D105 no SPED", type="primary"):
            df_atual = ed.df
            
            if "específico" in posicao:
                nl_ini = int(nl_especifica)
            elif "último D100" in posicao:
                d100s = df_atual[df_atual["registro"] == "D100"]
                nl_ini = int(d100s["numero_linha"].max()) + 10 if not d100s.empty else int(df_atual["numero_linha"].max()) + 1
            else:
                bloco_d = df_atual[df_atual["bloco"] == "D"]
                nl_ini = int(bloco_d["numero_linha"].max()) + 1 if not bloco_d.empty else int(df_atual["numero_linha"].max()) + 1
            
            nls_existentes = set(df_atual["numero_linha"].tolist())
            while nl_ini in nls_existentes or (nl_ini + 1) in nls_existentes or (nl_ini + 2) in nls_existentes:
                nl_ini += 10
            
            novas_linhas: List[LinhaRegistro] = []
            novas_rows: List[Dict[str, Any]] = []
            nl_cursor = nl_ini
            
            for nome_arq, dados in validos:
                cod_part = cod_part_pad or dados.cnpj_emit or ""
                lns = ImportadorCTe.para_linhas_d(
                    dados, cod_part, nl_cursor,
                    aliq_pis_padrao=aliq_pis_pad,
                    aliq_cofins_padrao=aliq_cof_pad,
                    cst_pis_padrao=cst_pis_pad,
                    cst_cofins_padrao=cst_cof_pad,
                    ind_nat_frt=ind_nat_frt,
                )
                novas_linhas.extend(lns)
                
                for ln in lns:
                    r: Dict[str, Any] = {
                        "numero_linha": ln.numero_linha,
                        "bloco": ln.bloco,
                        "registro": ln.registro,
                        "linha_original": ln.linha_original,
                    }
                    r.update({f"campo_{j:02d}": ln.get(j) for j in range(min(len(ln.campos), 55))})
                    mapa = MapaCampos.get_mapa(ln.registro)
                    if mapa:
                        for nome_c, idx in mapa.items():
                            r[nome_c] = ln.get(idx)
                    novas_rows.append(r)
                
                nl_cursor += 10
            
            df_novas = pd.DataFrame(novas_rows)
            for col in CAMPOS_NUMERICOS:
                if col in df_novas.columns:
                    df_novas[col] = df_novas[col].apply(Utilitarios.to_float)
            
            df_novo = pd.concat([ed.df, df_novas], ignore_index=True)
            df_novo = df_novo.sort_values("numero_linha").reset_index(drop=True)
            
            ed._atual = df_novo
            
            for ln in novas_linhas:
                ed.auditoria.registrar(
                    nl=ln.numero_linha,
                    reg=ln.registro,
                    bloco=ln.bloco,
                    campo="[importação CT-e]",
                    ant="",
                    novo=ln.linha_original[:80],
                    regra="importador_cte",
                    motivo=f"CT-e importado via XML",
                    tipo="AUTO",
                    usuario=usuario,
                )
            
            linhas_orig = st.session_state.get("linhas_orig", [])
            linhas_orig.extend(novas_linhas)
            st.session_state["linhas_orig"] = linhas_orig
            
            regras = st.session_state.get("regras", [])
            if regras:
                st.session_state["df_inc"] = ValidadorFiscal.validar(ed.df, regras, res.tipo_arquivo)
            
            n_d100 = sum(1 for l in novas_linhas if l.registro == "D100")
            n_d101 = sum(1 for l in novas_linhas if l.registro == "D101")
            n_d105 = sum(1 for l in novas_linhas if l.registro == "D105")
            
            st.success(
                f"✅ {n_d100} D100 · {n_d101} D101 · {n_d105} D105 inseridos "
                f"a partir da linha {nl_ini}. "
                f"Acesse **Exportação** para gerar o SPED corrigido."
            )
            st.rerun()


# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  INICIALIZAÇÃO E ROTEAMENTO  ███████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

def inicializar():
    """Inicializa o estado da sessão."""
    estado_inicial = {
        "resultado": None,
        "linhas_orig": None,
        "editor": None,
        "regras": None,
        "df_inc": None,
        "_usuario": "analista",
    }
    
    for key, valor in estado_inicial.items():
        if key not in st.session_state:
            st.session_state[key] = valor
    
    if st.session_state["regras"] is None:
        st.session_state["regras"] = GerenciadorRegras.carregar()


ROTEADOR = {
    "📊 Dashboard": Paginas.dashboard,
    "📂 Upload": Paginas.upload,
    "🗂️ Blocos": Paginas.blocos,
    "📋 Registros": Paginas.registros,
    "🧾 Notas Fiscais": Paginas.notas,
    "📦 Itens (C170/A170)": Paginas.itens,
    "💰 PIS/COFINS": Paginas.pis_cofins,
    "⚠️ Inconsistências": Paginas.inconsistencias,
    "🔧 Correções em Massa": Paginas.massa,
    "✏️ Editor Manual": Paginas.editor,
    "📤 Exportação": Paginas.exportacao,
    "📜 Log de Auditoria": Paginas.auditoria,
    "⚙️ Motor de Regras": Paginas.regras,
    "🚚 Importar CT-e": Paginas.importar_cte,
}


def main():
    """Função principal da aplicação."""
    inicializar()
    UI.css()
    pagina = UI.sidebar()
    
    funcao = ROTEADOR.get(pagina)
    if funcao:
        funcao()
    else:
        st.error(f"Página não encontrada: {pagina}")


if __name__ == "__main__":
    main()