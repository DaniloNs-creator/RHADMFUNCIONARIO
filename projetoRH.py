"""
╔═══════════════════════════════════════════════════════════════════════════════╗
║          SPED AUDITOR — FISCAL INTELLIGENCE PLATFORM  v2.1                  ║
║          EFD ICMS/IPI  +  EFD CONTRIBUIÇÕES (PIS/COFINS)                    ║
║          Arquivo único · Python + Streamlit                                  ║
║          Compatível com Streamlit 1.58+ / Python 3.14+                      ║
╚═══════════════════════════════════════════════════════════════════════════════╝

Correções v2.1 (baseadas nos logs do Streamlit Cloud):
  - st.radio("", ...) -> st.radio("Navegação", ...) — label vazio proibido em 3.14+
  - use_container_width=True  -> width='stretch'  — deprecado após 2025-12-31
  - use_container_width=False -> width='content'  — deprecado após 2025-12-31

Execução:
    pip install streamlit pandas numpy openpyxl plotly
    streamlit run sped_auditor.py
"""

# ═══════════════════════════════════════════════════════════════════════════════
# IMPORTS
# ═══════════════════════════════════════════════════════════════════════════════

from __future__ import annotations

import io, json, os, re, sys, uuid
from dataclasses import dataclass, field, asdict
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Optional

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
    menu_items={"About": "SPED Auditor v2.1 | EFD ICMS/IPI + EFD Contribuições"},
)

# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  MODELOS DE DADOS  ████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class LinhaRegistro:
    numero_linha: int
    bloco: str
    registro: str
    campos: list[str]
    linha_original: str

    def get(self, i: int, default: str = "") -> str:
        try: return self.campos[i] if i < len(self.campos) else default
        except IndexError: return default

    def set(self, i: int, v: str) -> None:
        while len(self.campos) <= i: self.campos.append("")
        self.campos[i] = v

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
    blocos_presentes: list[str] = field(default_factory=list)


@dataclass
class ResultadoParser:
    metadados: MetadadosArquivo
    linhas: list[LinhaRegistro]
    df: pd.DataFrame
    erros: list[str]
    tipo_arquivo: str   # 'EFD_ICMS_IPI' | 'EFD_CONTRIBUICOES'


@dataclass
class RegraFiscal:
    """Regra tributária configurável — unifica ICMS e PIS/COFINS."""
    id: str
    tributo: str           # "ICMS" | "PIS" | "COFINS" | "IPI"
    cst: str               # valor exato ou "*"
    cfop: str              # valor exato, prefixo ou "*"
    ind_oper: str          # "E" | "S" | "*"
    descricao: str
    exige_base: bool
    exige_aliquota: bool
    exige_valor: bool
    base_campo_sugerido: str
    aliquota_padrao: Optional[float]
    formula: str           # "base * aliq / 100" | "zero"
    permite_base_zero: bool
    permite_valor_zero: bool
    criticidade: str       # "critica" | "aviso" | "info"
    ativa: bool = True

    def match_cst(self, c: str) -> bool:
        return self.cst == "*" or c == self.cst or c.startswith(self.cst)
    def match_cfop(self, c: str) -> bool:
        return self.cfop == "*" or c == self.cfop or c.startswith(self.cfop)
    def match_oper(self, o: str) -> bool:
        return self.ind_oper in ("*", o)


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


# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  MAPA DE CAMPOS  ███████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

MAPA_CAMPOS: dict[str, dict[str, int]] = {

    # ── Bloco 0 — Abertura (comum) ────────────────────────────────────────────
    "0000": {"COD_VER":1,"COD_FIN":2,"DT_INI":3,"DT_FIN":4,
             "NOME":5,"CNPJ":6,"CPF":7,"UF":8,"IE":9,
             "COD_MUN":10,"IND_PERFIL":13,"IND_OPORT":14},
    "0150": {"COD_PART":1,"NOME":2,"COD_PAIS":3,"CNPJ":4,"CPF":5,
             "IE":6,"COD_MUN":7,"END":9,"BAIRRO":12},
    "0200": {"COD_ITEM":1,"DESCR_ITEM":2,"COD_BARRA":3,"UNID_INV":5,
             "TIPO_ITEM":6,"COD_NCM":7,"ALIQ_ICMS":11},

    # ── EFD ICMS/IPI — Bloco C ───────────────────────────────────────────────
    "C100": {"IND_OPER":1,"IND_EMIT":2,"COD_PART":3,"COD_MOD":4,
             "COD_SIT":5,"SER":6,"NUM_DOC":7,"CHV_NFE":8,
             "DT_DOC":9,"DT_E_S":10,"VL_DOC":11,
             "VL_DESC":13,"VL_MERC":15,
             "VL_BC_ICMS":20,"VL_ICMS":21,
             "VL_BC_ICMS_ST":22,"VL_ICMS_ST":23,
             "VL_IPI":24,"VL_PIS":25,"VL_COFINS":26},
    "C170": {"NUM_ITEM":1,"COD_ITEM":2,"DESCR_COMPL":3,
             "QTD":4,"UNID":5,"VL_ITEM":6,"VL_DESC":7,"IND_MOV":8,
             "CST_ICMS":9,"CFOP":10,
             "VL_BC_ICMS":12,"ALIQ_ICMS":13,"VL_ICMS":14,
             "VL_BC_ICMS_ST":15,"ALIQ_ST":16,"VL_ICMS_ST":17,
             "CST_IPI":19,"VL_BC_IPI":20,"ALIQ_IPI":21,"VL_IPI":22,
             "CST_PIS":23,"VL_BC_PIS":24,"ALIQ_PIS":25,"VL_PIS":26,
             "CST_COFINS":27,"VL_BC_COFINS":28,"ALIQ_COFINS":29,"VL_COFINS":30},
    "C190": {"CST_ICMS":1,"CFOP":2,"ALIQ_ICMS":3,
             "VL_OPR":4,"VL_BC_ICMS":5,"VL_ICMS":6,
             "VL_BC_ICMS_ST":7,"VL_ICMS_ST":8},

    # ── EFD ICMS/IPI — Bloco D (transportes) ─────────────────────────────────
    "D100": {"IND_OPER":1,"COD_PART":3,"COD_MOD":4,"COD_SIT":5,
             "NUM_DOC":8,"DT_DOC":10,"VL_DOC":14,
             "VL_BC_ICMS":18,"ALIQ_ICMS":19,"VL_ICMS":20,
             "CST_PIS":21,"VL_BC_PIS":22,"ALIQ_PIS":23,"VL_PIS":24,
             "CST_COFINS":25,"VL_BC_COFINS":26,"ALIQ_COFINS":27,"VL_COFINS":28},
    "D190": {"CST_ICMS":1,"CFOP":2,"ALIQ_ICMS":3,
             "VL_OPR":4,"VL_BC_ICMS":5,"VL_ICMS":6},

    # ── EFD ICMS/IPI — Bloco E (apuração ICMS) ───────────────────────────────
    "E110": {"VL_TOT_DEBITOS":1,"VL_AJ_DEBITOS":2,
             "VL_TOT_CREDITOS":5,"VL_AJ_CREDITOS":6,
             "VL_SLD_APURADO":10,"VL_ICMS_RECOLHER":12},
    "E116": {"COD_OR":1,"VL_OR":2,"DT_VCTO":3,"COD_REC":4},

    # ── EFD Contribuições — Bloco A (serviços) ───────────────────────────────
    "A100": {"IND_OPER":1,"IND_EMIT":2,"COD_PART":3,"COD_SIT":4,
             "SER":5,"SUB":6,"NUM_DOC":7,"CHV_NFSE":8,
             "DT_DOC":9,"DT_EXE_SERV":10,"VL_DOC":11,
             "VL_DESC":12,"VL_BC_PIS":13,"ALIQ_PIS":14,"VL_PIS":15,
             "VL_BC_COFINS":16,"ALIQ_COFINS":17,"VL_COFINS":18},
    "A170": {"NUM_ITEM":1,"COD_ITEM":2,"DESCR_COMPL":3,
             "VL_ITEM":4,"VL_DESC":5,"NAT_BC_CRED":6,"IND_ORIG_CRED":7,
             "CST_PIS":8,"VL_BC_PIS":9,"ALIQ_PIS":10,"VL_PIS":11,
             "CST_COFINS":12,"VL_BC_COFINS":13,"ALIQ_COFINS":14,"VL_COFINS":15,
             "COD_CTA":16},

    # ── EFD Contribuições — Bloco C ───────────────────────────────────────────
    "C010": {"IND_ESCRI":1},
    "C180": {"COD_CRED":1,"IND_ORIG_CRED":2,"VL_BC_PIS":3,"ALIQ_PIS":4,
             "QUANT_BC_PIS":5,"VL_PIS":6,"VL_BC_COFINS":7,"ALIQ_COFINS":8,
             "QUANT_BC_COFINS":9,"VL_COFINS":10,"COD_CTA":11},
    "C181": {"COD_CRED":1,"IND_ORIG_CRED":2,"CNPJ_CPF_PART":3,
             "COD_MOD":4,"DT_OPER":5,"CHV_NFE":6,"NUM_DOC":7,
             "VL_OPER":8,"CFOP":9,"NAT_BC_CRED":10,
             "VL_BC_PIS":12,"ALIQ_PIS":13,"VL_PIS":14,
             "VL_BC_COFINS":15,"ALIQ_COFINS":16,"VL_COFINS":17},
    "C185": {"NUM_ITEM":1,"COD_ITEM":2,"CST_PIS":3,"COD_CRED":4,
             "VL_BC_PIS":5,"ALIQ_PIS":6,"VL_PIS":7,
             "CST_COFINS":8,"VL_BC_COFINS":9,"ALIQ_COFINS":10,"VL_COFINS":11},
    "C380": {"COD_MOD":1,"DT_DOC_INI":2,"DT_DOC_FIN":3,"NUM_DOC_INI":4,
             "NUM_DOC_FIN":5,"VL_DOC":6,"VL_DOC_CANC":7},
    "C481": {"CST_PIS":1,"VL_ITEM":2,"VL_BC_PIS":3,"ALIQ_PIS":4,
             "QUANT_BC_PIS":5,"VL_PIS":6,"COD_CTA":7},
    "C485": {"CST_COFINS":1,"VL_ITEM":2,"VL_BC_COFINS":3,"ALIQ_COFINS":4,
             "QUANT_BC_COFINS":5,"VL_COFINS":6,"COD_CTA":7},

    # ── EFD Contribuições — Bloco D ───────────────────────────────────────────
    "D101": {"IND_NAT_FRT":1,"VL_ITEM":2,"CST_PIS":3,"NAT_BC_CRED":4,
             "VL_BC_PIS":5,"ALIQ_PIS":6,"VL_PIS":7,"COD_CTA":8},
    "D105": {"IND_NAT_FRT":1,"VL_ITEM":2,"CST_COFINS":3,"NAT_BC_CRED":4,
             "VL_BC_COFINS":5,"ALIQ_COFINS":6,"VL_COFINS":7,"COD_CTA":8},

    # ── EFD Contribuições — Bloco F (demais documentos) ──────────────────────
    "F100": {"IND_OPER":1,"COD_PART":2,"DT_OPER":3,"VL_OPER":4,
             "COD_CRED":5,"IND_ORIG_CRED":6,"VL_BC_PIS":7,"ALIQ_PIS":8,
             "VL_PIS":9,"VL_BC_COFINS":10,"ALIQ_COFINS":11,"VL_COFINS":12},
    "F120": {"NAT_DESPESA":1,"VL_AQUISICOES":2,"VL_PARCELA":3,
             "VL_BC_PIS":4,"ALIQ_PIS":5,"VL_PIS":6,
             "VL_BC_COFINS":7,"ALIQ_COFINS":8,"VL_COFINS":9},
    "F130": {"IND_ORIG_CRED":1,"IND_UTIL_BENS":2,"VL_OPER":3,
             "PARC_OPER":4,"VL_BC_PIS":5,"ALIQ_PIS":6,"VL_PIS":7,
             "VL_BC_COFINS":8,"ALIQ_COFINS":9,"VL_COFINS":10},
    "F150": {"NAT_BC_CRED":1,"VL_TOT_EST":2,"VL_BC_PIS":3,"ALIQ_PIS":4,
             "VL_PIS":5,"VL_BC_COFINS":6,"ALIQ_COFINS":7,"VL_COFINS":8},
    "F200": {"IND_OPER":1,"COD_PART":2,"COD_ITEM":3,"DT_OPER":4,
             "VL_OPER":5,"CST_PIS":6,"VL_BC_PIS":7,"ALIQ_PIS":8,
             "VL_PIS":9,"CST_COFINS":10,"VL_BC_COFINS":11,"ALIQ_COFINS":12,"VL_COFINS":13},

    # ── EFD Contribuições — Bloco M (apuração PIS/COFINS) ─────────────────────
    "M100": {"COD_CRED":1,"IND_CRED_ORI":2,"VL_BC_PIS":3,"ALIQ_PIS":4,
             "QUANT_BC_PIS":5,"VL_CRED":6,"VL_AJUS_ACRES":7,"VL_AJUS_REDUC":8,
             "VL_CRED_DIF":9,"VL_CRED_DISP":10,"IND_DESC_CRED":11,"VL_CRED_DESC":12},
    "M110": {"IND_AJ":1,"VL_AJ":2,"COD_AJ":3,"NUM_DOC":4,"DESCR_AJ":5,"DT_REF":6},
    "M200": {"VL_TOT_CONT_NC_PER":1,"VL_TOT_CRED_DESC":2,"VL_TOT_CRED_DESC_ANT":3,
             "VL_TOT_CONT_NC_DEV":4,"VL_RET_NC":5,"VL_OUT_DED_NC":6,
             "VL_CONT_NC_REC":7,"VL_TOT_CONT_CUM_PER":8,"VL_RET_CUM":9,
             "VL_OUT_DED_CUM":10,"VL_CONT_CUM_REC":11,"VL_TOT_CONT_REC":12},
    "M210": {"COD_CONT":1,"VL_REC_BRT":2,"VL_BC_CONT":3,"ALIQ_PIS":4,
             "QUANT_BC_PIS":5,"VL_CONT_APUR":6,"VL_AJUS_ACRES":7,"VL_AJUS_REDUC":8,
             "VL_CONT_DIFER":9,"VL_CONT_DIFER_ANT":10,"VL_CONT_PER":11},
    "M500": {"COD_CRED":1,"IND_CRED_ORI":2,"VL_BC_COFINS":3,"ALIQ_COFINS":4,
             "QUANT_BC_COFINS":5,"VL_CRED":6,"VL_AJUS_ACRES":7,"VL_AJUS_REDUC":8,
             "VL_CRED_DIF":9,"VL_CRED_DISP":10,"IND_DESC_CRED":11,"VL_CRED_DESC":12},
    "M600": {"VL_TOT_CONT_NC_PER":1,"VL_TOT_CRED_DESC":2,"VL_TOT_CRED_DESC_ANT":3,
             "VL_TOT_CONT_NC_DEV":4,"VL_RET_NC":5,"VL_OUT_DED_NC":6,
             "VL_CONT_NC_REC":7,"VL_TOT_CONT_CUM_PER":8,"VL_RET_CUM":9,
             "VL_OUT_DED_CUM":10,"VL_CONT_CUM_REC":11,"VL_TOT_CONT_REC":12},
    "M610": {"COD_CONT":1,"VL_REC_BRT":2,"VL_BC_CONT":3,"ALIQ_COFINS":4,
             "QUANT_BC_COFINS":5,"VL_CONT_APUR":6,"VL_AJUS_ACRES":7,"VL_AJUS_REDUC":8,
             "VL_CONT_DIFER":9,"VL_CONT_DIFER_ANT":10,"VL_CONT_PER":11},

    # ── EFD Contribuições — Bloco P (contribuição previdenciária) ──────────────
    "P001": {"IND_MOV":1},
    "P010": {"CNPJ":1},
    "P100": {"DT_INI":1,"DT_FIN":2,"VL_REC_TOT_EST":3,"COD_ATIV_ECON":4,
             "VL_REC_ATIV_ESTAB":5,"VL_EXC":6,"VL_REC_BC":7,"ALIQ_CONTRIB_APUR":8,
             "VL_CONTRIB_APUR":9,"VL_CONT_RECOL":10},
    "P110": {"NUM_CAMPO":1,"COD_DET":2,"DET_VALOR":3},

    # ── Bloco 9 — Encerramento (comum) ───────────────────────────────────────
    "9900": {"REG":1,"QTD":2},
    "9999": {"QTD":1},
}

# Campos numéricos — todos os tributos
CAMPOS_NUMERICOS = {
    "VL_ITEM","VL_BC_ICMS","ALIQ_ICMS","VL_ICMS",
    "VL_BC_ICMS_ST","ALIQ_ST","VL_ICMS_ST",
    "VL_IPI","VL_BC_IPI","ALIQ_IPI",
    "VL_BC_PIS","ALIQ_PIS","VL_PIS","QUANT_BC_PIS",
    "VL_BC_COFINS","ALIQ_COFINS","VL_COFINS","QUANT_BC_COFINS",
    "VL_DOC","VL_MERC","VL_DESC","VL_OPER","QTD",
    "VL_CRED","VL_CONT_APUR","VL_CONT_PER","VL_REC_BRT","VL_BC_CONT",
    "VL_ICMS_RECOLHER","VL_TOT_DEBITOS","VL_TOT_CREDITOS",
}

# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  REGRAS FISCAIS PADRÃO  ████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

_REGRAS_ICMS = [
    ("ICMS","000","*","Tributado Integralmente",True,True,True,"VL_ITEM",12.0,"critica"),
    ("ICMS","010","*","Tributado + ST",         True,True,True,"VL_ITEM",12.0,"critica"),
    ("ICMS","020","*","Tributado c/ Redução BC",True,True,True,"VL_ITEM",12.0,"critica"),
    ("ICMS","030","*","Isento c/ ST",           False,False,False,"",0.0,"aviso"),
    ("ICMS","040","*","Isento",                 False,False,False,"",0.0,"aviso"),
    ("ICMS","041","*","Não Tributado",          False,False,False,"",0.0,"aviso"),
    ("ICMS","050","*","Suspensão",              False,False,False,"",0.0,"aviso"),
    ("ICMS","051","*","Diferimento",            False,False,False,"",0.0,"info"),
    ("ICMS","060","*","ST Cobrado Anteriorm.",  False,False,False,"",0.0,"info"),
    ("ICMS","070","*","Redução BC + ST",        True,True,True,"VL_ITEM",12.0,"critica"),
    ("ICMS","090","*","Outras Situações",       False,False,False,"VL_ITEM",None,"aviso"),
]

_REGRAS_PIS = [
    ("PIS","01","*","PIS Não Cumulativo Alíq. Básica",True,True,True,"VL_ITEM",1.65,"critica"),
    ("PIS","02","*","PIS Não Cumulativo Alíq. Difenc.",True,True,True,"VL_ITEM",None,"critica"),
    ("PIS","04","*","PIS Monofásico",            False,False,False,"",0.0,"aviso"),
    ("PIS","05","*","PIS ST",                    False,False,False,"",0.0,"aviso"),
    ("PIS","06","*","PIS Alíq. Zero",            False,False,False,"",0.0,"aviso"),
    ("PIS","07","*","PIS Isento",                False,False,False,"",0.0,"aviso"),
    ("PIS","08","*","PIS Sem Incidência",        False,False,False,"",0.0,"aviso"),
    ("PIS","09","*","PIS Suspensão",             False,False,False,"",0.0,"aviso"),
    ("PIS","49","*","PIS Outras (saídas)",       False,False,False,"VL_ITEM",None,"info"),
    ("PIS","50","*","PIS Crédito Básico",        True,True,True,"VL_ITEM",1.65,"critica"),
    ("PIS","51","*","PIS Crédito Presumido",     True,True,True,"VL_ITEM",None,"critica"),
    ("PIS","70","*","PIS Crédito Alíq. Básica", True,True,True,"VL_ITEM",1.65,"critica"),
    ("PIS","71","*","PIS Crédito Alíq. Difenc.",True,True,True,"VL_ITEM",None,"critica"),
    ("PIS","72","*","PIS Créd. Pres. Agroind.", True,True,True,"VL_ITEM",None,"critica"),
    ("PIS","73","*","PIS Créd. Ass. Exportação",True,True,True,"VL_ITEM",None,"critica"),
    ("PIS","74","*","PIS Créd. Pat. Amort.",    True,True,True,"VL_ITEM",None,"critica"),
    ("PIS","75","*","PIS Créd. Val. Estoques",  True,True,True,"VL_ITEM",None,"critica"),
    ("PIS","98","*","PIS Outras Entradas",       False,False,False,"VL_ITEM",None,"info"),
    ("PIS","99","*","PIS Outras Saídas",         False,False,False,"VL_ITEM",None,"info"),
]

_REGRAS_COFINS = [
    ("COFINS","01","*","COFINS Não Cumulativo Alíq. Básica",True,True,True,"VL_ITEM",7.6,"critica"),
    ("COFINS","02","*","COFINS Não Cumulativo Alíq. Difenc.",True,True,True,"VL_ITEM",None,"critica"),
    ("COFINS","04","*","COFINS Monofásico",       False,False,False,"",0.0,"aviso"),
    ("COFINS","05","*","COFINS ST",               False,False,False,"",0.0,"aviso"),
    ("COFINS","06","*","COFINS Alíq. Zero",       False,False,False,"",0.0,"aviso"),
    ("COFINS","07","*","COFINS Isento",           False,False,False,"",0.0,"aviso"),
    ("COFINS","08","*","COFINS Sem Incidência",   False,False,False,"",0.0,"aviso"),
    ("COFINS","09","*","COFINS Suspensão",        False,False,False,"",0.0,"aviso"),
    ("COFINS","49","*","COFINS Outras (saídas)",  False,False,False,"VL_ITEM",None,"info"),
    ("COFINS","50","*","COFINS Crédito Básico",   True,True,True,"VL_ITEM",7.6,"critica"),
    ("COFINS","70","*","COFINS Créd. Alíq. Básica",True,True,True,"VL_ITEM",7.6,"critica"),
    ("COFINS","71","*","COFINS Créd. Alíq. Difenc.",True,True,True,"VL_ITEM",None,"critica"),
    ("COFINS","98","*","COFINS Outras Entradas",  False,False,False,"VL_ITEM",None,"info"),
    ("COFINS","99","*","COFINS Outras Saídas",    False,False,False,"VL_ITEM",None,"info"),
]

def _build_regras() -> list[RegraFiscal]:
    out = []
    for trib, cst, cfop, desc, eb, ea, ev, bc, alp, crit in _REGRAS_ICMS:
        out.append(RegraFiscal(
            id=f"{trib}_{cst}_{cfop}", tributo=trib, cst=cst, cfop=cfop, ind_oper="*",
            descricao=desc, exige_base=eb, exige_aliquota=ea, exige_valor=ev,
            base_campo_sugerido=bc, aliquota_padrao=alp,
            formula="base * aliq / 100" if ev else "zero",
            permite_base_zero=not eb, permite_valor_zero=not ev, criticidade=crit,
        ))
    for trib, cst, cfop, desc, eb, ea, ev, bc, alp, crit in _REGRAS_PIS + _REGRAS_COFINS:
        out.append(RegraFiscal(
            id=f"{trib}_{cst}_{cfop}", tributo=trib, cst=cst, cfop=cfop, ind_oper="*",
            descricao=desc, exige_base=eb, exige_aliquota=ea, exige_valor=ev,
            base_campo_sugerido=bc, aliquota_padrao=alp,
            formula="base * aliq / 100" if ev else "zero",
            permite_base_zero=not eb, permite_valor_zero=not ev, criticidade=crit,
        ))
    return out

REGRAS_PADRAO: list[RegraFiscal] = _build_regras()

ARQUIVO_REGRAS = "sped_regras_fiscais.json"

def carregar_regras() -> list[RegraFiscal]:
    if os.path.exists(ARQUIVO_REGRAS):
        try:
            with open(ARQUIVO_REGRAS, encoding="utf-8") as f:
                return [RegraFiscal(**d) for d in json.load(f)]
        except Exception:
            pass
    return list(REGRAS_PADRAO)

def salvar_regras(regras: list[RegraFiscal]) -> None:
    with open(ARQUIVO_REGRAS, "w", encoding="utf-8") as f:
        json.dump([asdict(r) for r in regras], f, ensure_ascii=False, indent=2)

def buscar_regra(regras: list[RegraFiscal], tributo: str, cst: str, cfop: str = "*") -> Optional[RegraFiscal]:
    cands = [r for r in regras if r.ativa and r.tributo == tributo and r.match_cst(cst) and r.match_cfop(cfop)]
    if not cands: return None
    def sc(r):
        s = 0
        if r.cst != "*": s += 2
        if r.cfop != "*": s += 2
        return s
    return max(cands, key=sc)

def calcular_tributo(regra: RegraFiscal, base: float, aliquota: float) -> float:
    if regra.formula == "zero": return 0.0
    r = Decimal(str(base)) * Decimal(str(aliquota)) / Decimal("100")
    return float(r.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  DEMOS SPED  ███████████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

SPED_DEMO_ICMS = """\
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

SPED_DEMO_CONTRIB = """\
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
# ████████████  PARSER  ████████████████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

def _decode(b: bytes) -> str:
    for enc in ("utf-8","latin-1","cp1252"):
        try: return b.decode(enc)
        except UnicodeDecodeError: pass
    raise ValueError("Encoding não reconhecido.")

def _to_float(v) -> Optional[float]:
    if v is None: return None
    if isinstance(v,(int,float)):
        import math; return None if (isinstance(v,float) and math.isnan(v)) else float(v)
    s = str(v).strip().replace(".","").replace(",",".")
    if not s: return None
    try: return float(Decimal(s))
    except: return None

def _to_sped_str(v: float) -> str:
    d = Decimal(str(v)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return str(d).replace(".",",")

def _fmt_brl(v: Optional[float]) -> str:
    if v is None: return "—"
    try:
        d = Decimal(str(v)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        i,dec = str(d).split(".")
        fmt=""
        for j,c in enumerate(reversed(i)):
            if j and j%3==0 and c!="-": fmt="."+fmt
            fmt=c+fmt
        return f"R$ {fmt},{dec}"
    except: return str(v)

def parse_sped(conteudo: bytes) -> ResultadoParser:
    texto = _decode(conteudo)
    erros: list[str] = []
    linhas: list[LinhaRegistro] = []

    for i, raw in enumerate(texto.splitlines(), 1):
        raw = raw.strip()
        if not raw: continue
        if not (raw.startswith("|") and raw.endswith("|")):
            erros.append(f"L{i}: formato inválido — {raw[:60]}"); continue
        campos = raw[1:-1].split("|")
        if not campos or not campos[0].strip():
            erros.append(f"L{i}: registro vazio"); continue
        reg = campos[0].strip().upper()
        linhas.append(LinhaRegistro(i, reg[0] if reg else "?", reg, campos, raw))

    meta = MetadadosArquivo()
    for l in linhas:
        if l.registro == "0000":
            meta.periodo_apuracao = f"{l.get(3)} a {l.get(4)}"
            meta.nome_empresa = l.get(5); meta.cnpj = l.get(6)
            meta.uf = l.get(8); meta.ie = l.get(9)
            meta.cod_municipio = l.get(10); meta.ind_perfil = l.get(13)
            break

    meta.total_linhas = len(linhas)
    meta.blocos_presentes = sorted(set(l.bloco for l in linhas))

    blocos = set(meta.blocos_presentes)
    cod_ver = ""
    for l in linhas:
        if l.registro == "0000": cod_ver = l.get(1); break
    if blocos & {"M","P"} or cod_ver.startswith("006") or cod_ver.startswith("007"):
        tipo = "EFD_CONTRIBUICOES"
    elif blocos & {"E","G","K","H"}:
        tipo = "EFD_ICMS_IPI"
    else:
        tipo = "EFD_ICMS_IPI"

    rows = []
    for l in linhas:
        r = {"numero_linha":l.numero_linha,"bloco":l.bloco,"registro":l.registro,"linha_original":l.linha_original}
        r.update({f"campo_{j:02d}":l.get(j) for j in range(min(len(l.campos),55))})
        mapa = MAPA_CAMPOS.get(l.registro)
        if mapa:
            for nome, idx in mapa.items():
                r[nome] = l.get(idx)
        rows.append(r)

    df = pd.DataFrame(rows)
    for col in CAMPOS_NUMERICOS:
        if col in df.columns:
            df[col] = df[col].apply(_to_float)

    return ResultadoParser(meta, linhas, df, erros, tipo)


def reconstruir_sped(linhas: list[LinhaRegistro]) -> bytes:
    return "".join(l.to_sped() for l in sorted(linhas, key=lambda x: x.numero_linha)).encode("utf-8")


def sync_df_linhas(df: pd.DataFrame, linhas: list[LinhaRegistro]) -> list[LinhaRegistro]:
    mapa = {l.numero_linha: l for l in linhas}
    for _, row in df.iterrows():
        nl = row.get("numero_linha")
        if nl not in mapa: continue
        l = mapa[nl]
        for nome, idx in (MAPA_CAMPOS.get(l.registro) or {}).items():
            if nome not in row or pd.isna(row[nome]): continue
            v = row[nome]
            if isinstance(v, float): v = _to_sped_str(v)
            l.set(idx, str(v))
    return list(mapa.values())

# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  VALIDADOR FISCAL  ████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

def validar(df: pd.DataFrame, regras: list[RegraFiscal], tipo_arquivo: str) -> pd.DataFrame:
    inc: list[dict] = []

    def add(nl,reg,bloco,tipo,sev,campo,desc,va,vs,cst="",cfop="",regra_id="",num_doc=""):
        inc.append({"numero_linha":nl,"bloco":bloco,"registro":reg,
                    "tipo":tipo,"severidade":sev,"campo_afetado":campo,
                    "descricao":desc,"valor_atual":va,"valor_sugerido":vs,
                    "cst":cst,"cfop":cfop,"num_doc":num_doc,
                    "regra_aplicada":regra_id,"corrigido":False})

    def val_tributo(row, nl, reg, bloco,
                    tributo, campo_cst, campo_bc, campo_aliq, campo_vl, campo_item="VL_ITEM"):
        cst  = str(row.get(campo_cst,"") or "").strip()
        bc   = _to_float(row.get(campo_bc))
        aliq = _to_float(row.get(campo_aliq))
        vl   = _to_float(row.get(campo_vl))
        vi   = _to_float(row.get(campo_item))
        if not cst: return
        regra = buscar_regra(regras, tributo, cst)
        if not regra: return
        sev = regra.criticidade.upper()
        cfop = str(row.get("CFOP","") or "").strip()
        rid  = regra.id
        if regra.exige_base and (bc is None or bc==0):
            add(nl,reg,bloco,"SEM_BASE_"+tributo,sev,campo_bc,
                f"[{tributo}] CST {cst} exige base — ausente/zerada",
                str(row.get(campo_bc,"")),_to_sped_str(vi) if vi else "",cst,cfop,rid)
        if regra.exige_aliquota and (aliq is None or aliq==0):
            add(nl,reg,bloco,"SEM_ALIQ_"+tributo,sev,campo_aliq,
                f"[{tributo}] CST {cst} exige alíquota — ausente/zerada",
                str(row.get(campo_aliq,"")),
                _to_sped_str(regra.aliquota_padrao) if regra.aliquota_padrao else "",
                cst,cfop,rid)
        if regra.exige_valor and (vl is None or vl==0):
            b_s = bc or vi or 0.0; a_s = aliq or (regra.aliquota_padrao or 0.0)
            calc = calcular_tributo(regra, b_s, a_s)
            add(nl,reg,bloco,"SEM_VALOR_"+tributo,sev,campo_vl,
                f"[{tributo}] CST {cst} exige valor — ausente/zerado (sugerido: {_to_sped_str(calc)})",
                str(row.get(campo_vl,"")),_to_sped_str(calc) if calc else "",cst,cfop,rid)
        if not regra.exige_valor and vl and vl>0:
            add(nl,reg,bloco,"VALOR_INDEVIDO_"+tributo,"AVISO",campo_vl,
                f"[{tributo}] CST {cst} não deve ter valor tributário — preenchido indevidamente",
                _to_sped_str(vl),"0",cst,cfop,rid)
        if bc and aliq and vl is not None and regra.formula=="base * aliq / 100":
            esp = calcular_tributo(regra, bc, aliq)
            if abs(esp-vl)>0.05:
                add(nl,reg,bloco,"DIVERGENCIA_"+tributo,"AVISO",campo_vl,
                    f"[{tributo}] Calculado {_fmt_brl(esp)} ≠ Registrado {_fmt_brl(vl)}",
                    _to_sped_str(vl),_to_sped_str(esp),cst,cfop,rid)

    df_c170 = df[df["registro"]=="C170"].copy()
    for _, row in df_c170.iterrows():
        nl = int(row.get("numero_linha",0))
        val_tributo(row,nl,"C170","C","ICMS","CST_ICMS","VL_BC_ICMS","ALIQ_ICMS","VL_ICMS")
        val_tributo(row,nl,"C170","C","PIS","CST_PIS","VL_BC_PIS","ALIQ_PIS","VL_PIS")
        val_tributo(row,nl,"C170","C","COFINS","CST_COFINS","VL_BC_COFINS","ALIQ_COFINS","VL_COFINS")

    df_a170 = df[df["registro"]=="A170"].copy()
    for _, row in df_a170.iterrows():
        nl = int(row.get("numero_linha",0))
        val_tributo(row,nl,"A170","A","PIS","CST_PIS","VL_BC_PIS","ALIQ_PIS","VL_PIS","VL_ITEM")
        val_tributo(row,nl,"A170","A","COFINS","CST_COFINS","VL_BC_COFINS","ALIQ_COFINS","VL_COFINS","VL_ITEM")

    df_c185 = df[df["registro"]=="C185"].copy()
    for _, row in df_c185.iterrows():
        nl = int(row.get("numero_linha",0))
        val_tributo(row,nl,"C185","C","PIS","CST_PIS","VL_BC_PIS","ALIQ_PIS","VL_PIS")
        val_tributo(row,nl,"C185","C","COFINS","CST_COFINS","VL_BC_COFINS","ALIQ_COFINS","VL_COFINS")

    df_d100 = df[df["registro"]=="D100"].copy()
    for _, row in df_d100.iterrows():
        nl = int(row.get("numero_linha",0))
        val_tributo(row,nl,"D100","D","PIS","CST_PIS","VL_BC_PIS","ALIQ_PIS","VL_PIS","VL_DOC")
        val_tributo(row,nl,"D100","D","COFINS","CST_COFINS","VL_BC_COFINS","ALIQ_COFINS","VL_COFINS","VL_DOC")

    df_f100 = df[df["registro"]=="F100"].copy()
    for _, row in df_f100.iterrows():
        nl = int(row.get("numero_linha",0))
        bc_p = _to_float(row.get("VL_BC_PIS")); al_p = _to_float(row.get("ALIQ_PIS")); vl_p = _to_float(row.get("VL_PIS"))
        bc_c = _to_float(row.get("VL_BC_COFINS")); al_c = _to_float(row.get("ALIQ_COFINS")); vl_c = _to_float(row.get("VL_COFINS"))
        if bc_p and al_p and vl_p is not None:
            esp = calcular_tributo(RegraFiscal("","PIS","01","*","*","",True,True,True,"",1.65,"base * aliq / 100",False,False,"critica"),bc_p,al_p)
            if abs(esp-vl_p)>0.05:
                add(nl,"F100","F","DIVERGENCIA_PIS","AVISO","VL_PIS",f"[PIS] F100 calculado {_fmt_brl(esp)} ≠ {_fmt_brl(vl_p)}",_to_sped_str(vl_p),_to_sped_str(esp))
        if bc_c and al_c and vl_c is not None:
            esp = calcular_tributo(RegraFiscal("","COFINS","01","*","*","",True,True,True,"",7.6,"base * aliq / 100",False,False,"critica"),bc_c,al_c)
            if abs(esp-vl_c)>0.05:
                add(nl,"F100","F","DIVERGENCIA_COFINS","AVISO","VL_COFINS",f"[COFINS] F100 calculado {_fmt_brl(esp)} ≠ {_fmt_brl(vl_c)}",_to_sped_str(vl_c),_to_sped_str(esp))

    df_docs = df[df["registro"].isin(["C100","C170"])].sort_values("numero_linha")
    c100a=None; c100nl=0; aicms=0.0
    for _, row in df_docs.iterrows():
        if row["registro"]=="C100":
            if c100a is not None:
                vd = _to_float(c100a.get("VL_ICMS")) or 0.0
                if abs(aicms-vd)>0.10:
                    add(c100nl,"C100","C","DIVERGENCIA_TOTAL_ICMS","CRITICA","VL_ICMS",
                        f"Total C100 ICMS ({_fmt_brl(vd)}) ≠ soma C170 ({_fmt_brl(aicms)})",
                        _to_sped_str(vd),_to_sped_str(aicms),num_doc=str(c100a.get("NUM_DOC","")))
            c100a=row; c100nl=int(row["numero_linha"]); aicms=0.0
        elif row["registro"]=="C170" and c100a is not None:
            aicms += _to_float(row.get("VL_ICMS")) or 0.0

    for _, row in df[df["registro"]=="C190"].iterrows():
        cst=str(row.get("CST_ICMS","") or "").strip()
        cfop=str(row.get("CFOP","") or "").strip()
        nl=int(row.get("numero_linha",0))
        r=buscar_regra(regras,"ICMS",cst)
        if r and r.exige_base:
            bc=_to_float(row.get("VL_BC_ICMS"))
            if not bc:
                add(nl,"C190","C","C190_SEM_BASE","AVISO","VL_BC_ICMS",
                    f"C190 CST {cst}/CFOP {cfop} sem base ICMS","",str(row.get("VL_OPR","")),cst,cfop)

    for reg_m, campo_rec in [("M200","VL_TOT_CONT_REC"),("M600","VL_TOT_CONT_REC")]:
        for _, row in df[df["registro"]==reg_m].iterrows():
            nl=int(row.get("numero_linha",0))
            vr = _to_float(row.get(campo_rec))
            nc = _to_float(row.get("VL_CONT_NC_REC"))
            cu = _to_float(row.get("VL_CONT_CUM_REC"))
            trib = "PIS" if reg_m=="M200" else "COFINS"
            if vr is not None and nc is not None and cu is not None:
                esp = (nc or 0)+(cu or 0)
                if abs(esp-(vr or 0))>0.05:
                    add(nl,reg_m,"M",f"DIVERGENCIA_APURACAO_{trib}","AVISO",campo_rec,
                        f"[{trib}] {reg_m}: NC ({_fmt_brl(nc)}) + CUM ({_fmt_brl(cu)}) ≠ Total ({_fmt_brl(vr)})",
                        _to_sped_str(vr or 0),_to_sped_str(esp))

    for bloco in df["bloco"].unique():
        if bloco=="?": continue
        for suf,te in [("001","BLOCO_SEM_ABERTURA"),("990","BLOCO_SEM_FECHAMENTO")]:
            re_ = f"{bloco}{suf}"
            if df[df["registro"]==re_].empty:
                add(0,re_,bloco,te,"CRITICA","registro",f"Bloco {bloco}: {re_} ausente","ausente","")

    df0=df[df["registro"]=="0000"]
    if df0.empty:
        add(0,"0000","0","CAMPO_AUSENTE","CRITICA","0000","Registro 0000 não encontrado","ausente","")
    else:
        row0=df0.iloc[0]
        for c in ("CNPJ","NOME"):
            if not str(row0.get(c,"") or "").strip():
                add(int(row0.get("numero_linha",1)),"0000","0","CAMPO_AUSENTE","CRITICA",c,
                    f"Campo {c} obrigatório ausente no 0000","","")

    for col in ["VL_ITEM","VL_BC_ICMS","VL_ICMS","ALIQ_ICMS","VL_BC_PIS","VL_PIS","VL_BC_COFINS","VL_COFINS"]:
        if col not in df_c170.columns: continue
        neg=df_c170[df_c170[col].apply(lambda x: x is not None and isinstance(x,float) and x<0)]
        for _,row in neg.iterrows():
            add(int(row.get("numero_linha",0)),"C170","C","VALOR_NEGATIVO","AVISO",col,
                f"Valor negativo em {col}: {row[col]:.2f}",_to_sped_str(row[col]),_to_sped_str(abs(row[col])),
                str(row.get("CST_ICMS","")),str(row.get("CFOP","")))

    return pd.DataFrame(inc) if inc else pd.DataFrame()

# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  EDITOR + AUDITORIA  ██████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

class TrilhaAuditoria:
    def __init__(self): self._log: list[EntradaAuditoria] = []
    def registrar(self,nl,reg,bloco,campo,ant,novo,regra="",motivo="",tipo="MANUAL",usuario="analista"):
        self._log.append(EntradaAuditoria(
            str(uuid.uuid4())[:8],datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            usuario,nl,reg,bloco,campo,ant,novo,regra,motivo,tipo))
    def to_df(self):
        if not self._log: return pd.DataFrame()
        return pd.DataFrame([{"ID":e.id,"Data/Hora":e.timestamp,"Usuário":e.usuario,
            "Linha":e.numero_linha,"Registro":e.registro,"Bloco":e.bloco,
            "Campo":e.campo,"Anterior":e.valor_anterior,"Novo":e.valor_novo,
            "Regra":e.regra,"Motivo":e.motivo,"Tipo":e.tipo} for e in self._log])
    def total(self): return len(self._log)


class EditorSped:
    def __init__(self, df):
        self._orig = df.copy(); self._atual = df.copy()
        self._hist: list[pd.DataFrame] = []
        self.auditoria = TrilhaAuditoria()

    @property
    def df(self): return self._atual
    @property
    def df_original(self): return self._orig

    def _set(self, i, campo, valor):
        if campo in CAMPOS_NUMERICOS and campo in self._atual.columns and self._atual[campo].dtype==float:
            vn = _to_float(valor)
            self._atual[campo] = self._atual[campo].astype(object)
            self._atual.at[i,campo] = vn if vn is not None else valor
        else:
            if campo in self._atual.columns and self._atual[campo].dtype!=object:
                self._atual[campo] = self._atual[campo].astype(object)
            self._atual.at[i,campo] = valor

    def editar(self,nl,campo,valor,motivo="",regra="",usuario="analista"):
        idx=self._atual[self._atual["numero_linha"]==nl].index
        if idx.empty: return False
        self._hist.append(self._atual.copy()); i=idx[0]
        row=self._atual.loc[i]; ant=str(row.get(campo,""))
        self._set(i,campo,valor)
        self.auditoria.registrar(nl,str(row.get("registro","")),str(row.get("bloco","")),
                                 campo,ant,str(valor),regra,motivo,"MANUAL",usuario)
        return True

    def massa(self,nls,campo,valor,motivo="Correção em massa",regra="",usuario="analista"):
        self._hist.append(self._atual.copy()); n=0
        for nl in nls:
            idx=self._atual[self._atual["numero_linha"]==nl].index
            if idx.empty: continue
            i=idx[0]; row=self._atual.loc[i]; ant=str(row.get(campo,""))
            if ant==str(valor): continue
            self._set(i,campo,valor)
            self.auditoria.registrar(nl,str(row.get("registro","")),str(row.get("bloco","")),
                                     campo,ant,str(valor),regra,motivo,"MASSA",usuario); n+=1
        return n

    def recalcular_massa(self,nls,regras,tributo="ICMS",
                         campo_cst="CST_ICMS",campo_bc="VL_BC_ICMS",campo_aliq="ALIQ_ICMS",campo_vl="VL_ICMS",
                         motivo="Recálculo automático",usuario="analista"):
        self._hist.append(self._atual.copy()); n=0
        for nl in nls:
            idx=self._atual[self._atual["numero_linha"]==nl].index
            if idx.empty: continue
            i=idx[0]; row=self._atual.loc[i]
            cst=str(row.get(campo_cst,"") or "").strip()
            bc=_to_float(row.get(campo_bc)); aliq=_to_float(row.get(campo_aliq))
            if bc is None or aliq is None: continue
            regra=buscar_regra(regras,tributo,cst)
            if not regra: continue
            novo=calcular_tributo(regra,bc,aliq)
            ant=str(row.get(campo_vl,""))
            self._set(i,campo_vl,novo)
            self.auditoria.registrar(nl,str(row.get("registro","")),str(row.get("bloco","")),
                                     campo_vl,ant,_to_sped_str(novo),regra.id,motivo,"AUTO",usuario); n+=1
        return n

    def desfazer(self):
        if not self._hist: return False
        self._atual=self._hist.pop(); return True

    def restaurar(self,nls=None):
        self._hist.append(self._atual.copy())
        if nls:
            for nl in nls:
                ia=self._atual[self._atual["numero_linha"]==nl].index
                io=self._orig[self._orig["numero_linha"]==nl].index
                if not ia.empty and not io.empty:
                    self._atual.loc[ia[0]]=self._orig.loc[io[0]]
        else: self._atual=self._orig.copy()

    def get_alterados(self):
        try: return self._atual[self._atual.ne(self._orig).any(axis=1)]
        except: return pd.DataFrame()

    def preview(self,nls):
        rows=[]
        for nl in nls:
            a=self._atual[self._atual["numero_linha"]==nl]
            o=self._orig[self._orig["numero_linha"]==nl]
            if a.empty or o.empty: continue
            ra,ro=a.iloc[0],o.iloc[0]
            for c in list(CAMPOS_NUMERICOS):
                if c in ra.index and ra.get(c)!=ro.get(c):
                    rows.append({"Linha":nl,"Registro":ra.get("registro",""),"Campo":c,
                                 "Original":ro.get(c),"Atual":ra.get(c)})
        return pd.DataFrame(rows)

# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  EXPORTAÇÕES  ████████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

_CH="FF1A3A5C"; _CC="FFC0392B"; _CA="FFB7860D"; _CO="FF1A7A35"
_CZ="FFEBF0F7"; _CW="FFFFFFFF"; _CS="FF2C5F8A"

def _bd(): s=Side(style="thin",color="CCCCCC"); return Border(bottom=s,right=s)

def _aba(ws,df,titulo,col_sev=""):
    ws.sheet_view.showGridLines=False
    if df is None or df.empty:
        ws["A1"]=f"{titulo} — sem dados"; ws["A1"].font=Font(bold=True,italic=True,color="888888"); return
    n=len(df.columns); ul=get_column_letter(n)
    ws.merge_cells(f"A1:{ul}1"); c=ws["A1"]
    c.value=titulo; c.font=Font(bold=True,size=13,color=_CW)
    c.fill=PatternFill("solid",fgColor=_CH); c.alignment=Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[1].height=26; bd=_bd()
    for ci,col in enumerate(df.columns,1):
        h=ws.cell(2,ci,str(col).upper())
        h.font=Font(bold=True,size=10,color=_CW); h.fill=PatternFill("solid",fgColor=_CS)
        h.alignment=Alignment(horizontal="center"); h.border=bd
    ws.row_dimensions[2].height=22
    mc={"CRITICA":"FCE4E4","AVISO":"FFF8DC","INFO":"EAF4FF"}
    for ri,(_,row) in enumerate(df.iterrows(),3):
        zb=_CZ if ri%2==0 else _CW
        sev=str(row.get(col_sev,"")).upper() if col_sev else ""
        cor=mc.get(sev,zb)
        for ci,col in enumerate(df.columns,1):
            cell=ws.cell(ri,ci,row[col])
            cell.font=Font(size=10); cell.fill=PatternFill("solid",fgColor=cor)
            cell.border=bd; cell.alignment=Alignment(vertical="center")
    for ci,col in enumerate(df.columns,1):
        mw=max(len(str(col)),df[col].astype(str).str.len().max() if len(df)>0 else 10)
        ws.column_dimensions[get_column_letter(ci)].width=min(mw+4,42)
    ws.freeze_panes="A3"

def _aba_resumo(ws,df_inc,df_alt,meta):
    ws.sheet_view.showGridLines=False
    ws.merge_cells("A1:F1"); c=ws["A1"]
    c.value="SPED AUDITOR — RELATÓRIO GERENCIAL"
    c.font=Font(name="Calibri",bold=True,size=16,color=_CW)
    c.fill=PatternFill("solid",fgColor=_CH); c.alignment=Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[1].height=36
    for i,(l,v) in enumerate([("Empresa:",meta.get("nome_empresa","—")),("CNPJ:",meta.get("cnpj","—")),
        ("Período:",meta.get("periodo_apuracao","—")),("Tipo:",meta.get("tipo_arquivo","—")),
        ("UF:",meta.get("uf","—")),("Gerado em:",datetime.now().strftime("%d/%m/%Y %H:%M"))],3):
        ws.cell(i,1,l).font=Font(bold=True,size=11); ws.cell(i,2,v).font=Font(size=11)
    ws.cell(11,1,"INDICADORES").font=Font(bold=True,size=12,color=_CH)
    def n(t): return int((df_inc["tipo"]==t).sum()) if not df_inc.empty and "tipo" in df_inc.columns else 0
    def s(s_): return int((df_inc["severidade"]==s_).sum()) if not df_inc.empty and "severidade" in df_inc.columns else 0
    inds=[("Total Inconsistências",len(df_inc) if not df_inc.empty else 0,_CH),
          ("Críticas",s("CRITICA"),_CC),("Avisos",s("AVISO"),_CA),
          ("Sem Base ICMS",n("SEM_BASE_ICMS"),_CC),("Sem Valor ICMS",n("SEM_VALOR_ICMS"),_CC),
          ("Sem Base PIS",n("SEM_BASE_PIS"),_CC),("Sem Valor PIS",n("SEM_VALOR_PIS"),_CC),
          ("Sem Base COFINS",n("SEM_BASE_COFINS"),_CC),("Sem Valor COFINS",n("SEM_VALOR_COFINS"),_CC),
          ("Registros Alterados",len(df_alt) if not df_alt.empty else 0,_CO)]
    for j,(l,v,cor) in enumerate(inds,13):
        ws.cell(j,1,l).font=Font(size=11,bold=True)
        c2=ws.cell(j,2,v); c2.font=Font(size=11,bold=True,color=_CW)
        c2.fill=PatternFill("solid",fgColor=cor); c2.alignment=Alignment(horizontal="center")
    ws.column_dimensions["A"].width=34; ws.column_dimensions["B"].width=20

def exportar_excel(df_inc,df_alt,df_aud,df_reg,meta):
    buf=io.BytesIO(); wb=openpyxl.Workbook(); wb.remove(wb.active)
    _aba_resumo(wb.create_sheet("Resumo Gerencial"),df_inc,df_alt,meta)
    _aba(wb.create_sheet("Inconsistências"),df_inc,"Inconsistências Fiscais",col_sev="severidade")
    _aba(wb.create_sheet("Registros Alterados"),df_alt,"Registros Modificados")
    _aba(wb.create_sheet("Log de Auditoria"),df_aud,"Trilha de Auditoria")
    _aba(wb.create_sheet("Regras Tributárias"),df_reg,"Catálogo de Regras")
    wb.save(buf); buf.seek(0); return buf.read()

def exportar_csv(df):
    buf=io.StringIO(); df.to_csv(buf,index=False,sep=";",encoding="utf-8-sig")
    return buf.getvalue().encode("utf-8-sig")

# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  UTILITÁRIOS UI  ████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

CSS="""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif;}
section[data-testid="stSidebar"]{background:#0F2540 !important;color:#fff;}
section[data-testid="stSidebar"] .stRadio label{color:#BDD6F5 !important;font-size:.88rem;}
section[data-testid="stSidebar"] .stRadio label:hover{color:#fff !important;}
section[data-testid="stSidebar"] .stRadio [data-testid="stWidgetLabel"]{display:none;}
.sec{font-size:1.05rem;font-weight:700;color:#0F2540;border-left:4px solid #1A6BAD;
     padding-left:10px;margin:16px 0 10px 0;}
.stButton>button{border-radius:6px;font-weight:600;font-size:.87rem;}
hr{border-color:#E0E8F4;}
div[data-testid="metric-container"]{background:#F7FAFF;border:1px solid #E0E8F4;
  border-radius:8px;padding:10px 14px;}
</style>
"""

def _card(col,label,valor,classe=""):
    cores={"critico":"#C0392B","aviso":"#B7860D","ok":"#1A7A35","":"#0F2540"}
    cor=cores.get(classe,"#0F2540")
    col.markdown(f"""
    <div style="background:#fff;border:1px solid #E0E8F4;border-radius:8px;
                padding:12px 14px;text-align:center;box-shadow:0 2px 6px rgba(26,58,92,.07);">
      <div style="font-size:.68rem;color:#5A7398;font-weight:700;text-transform:uppercase;
                  letter-spacing:.04em;">{label}</div>
      <div style="font-size:1.55rem;font-weight:800;color:{cor};margin-top:3px;">{valor}</div>
    </div>""",unsafe_allow_html=True)

def _cols_rel(df):
    base=["numero_linha","bloco","registro"]
    extras=[c for c in df.columns if not c.startswith("campo_") and c not in base+["linha_original"]
            and df[c].astype(str).str.strip().ne("").any()]
    return base+extras[:30]

def _metricas(df,tipo):
    r={"total_linhas":len(df),
       "total_c100":int((df["registro"]=="C100").sum()) if "registro" in df.columns else 0,
       "total_c170":int((df["registro"]=="C170").sum()) if "registro" in df.columns else 0,
       "total_a170":int((df["registro"]=="A170").sum()) if "registro" in df.columns else 0,
       "total_d100":int((df["registro"]=="D100").sum()) if "registro" in df.columns else 0}
    df17=df[df["registro"]=="C170"] if "registro" in df.columns else pd.DataFrame()
    def sem(col): return int(df17[col].apply(lambda x:x is None or(isinstance(x,float) and(x==0 or x!=x))).sum()) if col in df17.columns else 0
    r["sem_base_icms"]=sem("VL_BC_ICMS"); r["sem_icms"]=sem("VL_ICMS")
    r["sem_base_pis"]=sem("VL_BC_PIS");  r["sem_pis"]=sem("VL_PIS")
    r["sem_base_cof"]=sem("VL_BC_COFINS"); r["sem_cof"]=sem("VL_COFINS")
    return r

# ─────────────────────────────────────────────────────────────────────────────
# Helper: width compatível com Streamlit ≥1.45 e <1.45
# ─────────────────────────────────────────────────────────────────────────────
def _w(stretch=True):
    """Retorna kwargs de largura compatíveis com Streamlit 1.45+ (width=) e anterior (use_container_width=)."""
    import streamlit as _st
    ver = tuple(int(x) for x in _st.__version__.split(".")[:2])
    if ver >= (1, 45):
        return {"width": "stretch" if stretch else "content"}
    return {"use_container_width": stretch}

# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  INICIALIZAÇÃO  █████████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

def _init():
    for k,v in {"resultado":None,"linhas_orig":None,"editor":None,
                "regras":None,"df_inc":None,"_usuario":"analista"}.items():
        if k not in st.session_state: st.session_state[k]=v
    if st.session_state["regras"] is None:
        st.session_state["regras"]=carregar_regras()

# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  TELAS  ██████████████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

PAGINAS=["📊 Dashboard","📂 Upload","🗂️ Blocos","📋 Registros",
         "🧾 Notas Fiscais","📦 Itens (C170/A170)","💰 PIS/COFINS",
         "⚠️ Inconsistências","🔧 Correções em Massa",
         "✏️ Editor Manual","📤 Exportação","📜 Log de Auditoria","⚙️ Motor de Regras"]

# ── Sidebar ────────────────────────────────────────────────────────────────────
def sidebar():
    with st.sidebar:
        st.markdown("""
        <div style="padding:10px 0 18px;text-align:center;">
          <span style="font-size:1.4rem;font-weight:800;color:#fff;">🔍 SPED Auditor</span><br>
          <span style="font-size:.68rem;color:#7BAFD4;">EFD ICMS/IPI + Contribuições  v2.1</span>
        </div>""",unsafe_allow_html=True)
        res=st.session_state.get("resultado")
        if res:
            m=res.metadados
            tipo_badge="🟦 ICMS/IPI" if res.tipo_arquivo=="EFD_ICMS_IPI" else "🟩 CONTRIBUIÇÕES"
            st.markdown(f"""
            <div style="background:#1A3A5C;border-radius:6px;padding:9px 11px;margin-bottom:14px;">
              <div style="font-size:.63rem;color:#7BAFD4;font-weight:700;">{tipo_badge}</div>
              <div style="font-size:.8rem;color:#fff;margin-top:3px;">{(m.nome_empresa or '—')[:26]}</div>
              <div style="font-size:.67rem;color:#BDD6F5;">CNPJ: {m.cnpj or '—'}</div>
              <div style="font-size:.67rem;color:#BDD6F5;">{m.periodo_apuracao or '—'}</div>
              <div style="font-size:.67rem;color:#BDD6F5;">Blocos: {', '.join(m.blocos_presentes)}</div>
            </div>""",unsafe_allow_html=True)
        st.markdown("**Navegação**")
        # ── CORREÇÃO 1: label não pode ser vazio no Streamlit 1.58+ / Python 3.14+
        # label_visibility="collapsed" oculta o label visualmente sem gerar warning
        pag=st.radio("Navegação",PAGINAS,label_visibility="collapsed")
        ed=st.session_state.get("editor")
        if ed and ed.auditoria.total():
            st.markdown(f'<div style="color:#7BAFD4;font-size:.71rem;margin-top:8px;">✏️ {ed.auditoria.total()} alteração(ões)</div>',unsafe_allow_html=True)
    return pag

# ── Upload ─────────────────────────────────────────────────────────────────────
def pg_upload():
    st.markdown('<div class="sec">Upload do Arquivo SPED</div>',unsafe_allow_html=True)
    c1,c2=st.columns([3,2])
    with c1:
        arq=st.file_uploader("Selecione o arquivo SPED (.txt)",type=["txt"])
        tipo_demo=st.radio("Demo:",["EFD ICMS/IPI","EFD Contribuições (PIS/COFINS)"],horizontal=True,key="td")
        usar_demo=st.checkbox("📁 Usar arquivo de demonstração",value=not bool(arq))
        usuario=st.text_input("Usuário da sessão:","analista")
        if st.button("▶ Processar Arquivo",type="primary",**_w()):
            _processar(arq,usar_demo,tipo_demo,usuario)
    with c2:
        st.markdown("**Tipos suportados:**")
        st.markdown("- ✅ **EFD ICMS/IPI** — Blocos A B C D E G H K\n- ✅ **EFD Contribuições** — Blocos A C D F M P\n- 🔄 ECD/ECF *(arquitetura pronta)*")
        st.markdown("**Encoding:** UTF-8 · Latin-1 · CP1252")
        res=st.session_state.get("resultado")
        if res:
            m=res.metadados
            tipo_label="🟦 EFD ICMS/IPI" if res.tipo_arquivo=="EFD_ICMS_IPI" else "🟩 EFD Contribuições"
            st.success(f"**{tipo_label}**\n\n**{m.nome_empresa}**\n\nCNPJ: `{m.cnpj}`\n\nPeríodo: {m.periodo_apuracao}\n\nBlocos: `{', '.join(m.blocos_presentes)}`\n\nLinhas: `{m.total_linhas:,}`")

def _processar(arq,usar_demo,tipo_demo,usuario):
    with st.spinner("Analisando arquivo SPED…"):
        try:
            if usar_demo or not arq:
                conteudo=(SPED_DEMO_ICMS if "ICMS" in tipo_demo else SPED_DEMO_CONTRIB).encode("utf-8")
                st.info(f"Demo carregado: {tipo_demo}")
            else:
                conteudo=arq.read()
            res=parse_sped(conteudo)
            st.session_state.update({"resultado":res,"linhas_orig":list(res.linhas),
                                     "editor":EditorSped(res.df),"_usuario":usuario})
            regras=st.session_state["regras"] or carregar_regras()
            df_inc=validar(res.df,regras,res.tipo_arquivo)
            st.session_state["df_inc"]=df_inc
            st.success(f"✅ {res.metadados.total_linhas:,} linhas | {res.tipo_arquivo.replace('_',' ')} | {len(df_inc)} inconsistência(s)")
            if res.erros:
                with st.expander(f"⚠️ {len(res.erros)} aviso(s) do parser"): [st.text(e) for e in res.erros[:20]]
        except Exception as e:
            st.error(f"Erro: {e}")

# ── Dashboard ──────────────────────────────────────────────────────────────────
def pg_dashboard():
    st.markdown('<div class="sec">Dashboard — Visão Geral</div>',unsafe_allow_html=True)
    res=st.session_state.get("resultado"); ed=st.session_state.get("editor")
    df_inc=st.session_state.get("df_inc",pd.DataFrame())
    if not res: st.info("Carregue um arquivo em **Upload**."); return
    df=ed.df if ed else res.df; m=res.metadados
    mt=_metricas(df,res.tipo_arquivo)
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Empresa",(m.nome_empresa or "—")[:22]); c2.metric("CNPJ",m.cnpj or "—")
    c3.metric("Período",m.periodo_apuracao or "—")
    tipo_label="🟦 EFD ICMS/IPI" if res.tipo_arquivo=="EFD_ICMS_IPI" else "🟩 EFD Contribuições"
    c4.metric("Tipo",tipo_label)
    st.markdown("---")
    cc=st.columns(8)
    _card(cc[0],"Total Linhas",f"{mt['total_linhas']:,}")
    _card(cc[1],"NF-e (C100)",f"{mt['total_c100']:,}")
    _card(cc[2],"Itens (C170)",f"{mt['total_c170']:,}")
    _card(cc[3],"Serv (A170)",f"{mt.get('total_a170',0):,}")
    _card(cc[4],"Transp (D100)",f"{mt.get('total_d100',0):,}")
    ni=len(df_inc) if not df_inc.empty else 0
    nc=int((df_inc["severidade"]=="CRITICA").sum()) if not df_inc.empty and "severidade" in df_inc.columns else 0
    _card(cc[5],"Inconsistências",f"{ni:,}","critico" if ni else "ok")
    _card(cc[6],"Críticas",f"{nc:,}","critico" if nc else "ok")
    alt=len(ed.get_alterados()) if ed else 0
    _card(cc[7],"Alterados",f"{alt:,}","aviso" if alt else "ok")
    st.markdown("<br>",unsafe_allow_html=True)
    ct=st.columns(6)
    _card(ct[0],"Sem Base ICMS",f"{mt['sem_base_icms']:,}","critico" if mt['sem_base_icms'] else "ok")
    _card(ct[1],"Sem VL ICMS",f"{mt['sem_icms']:,}","critico" if mt['sem_icms'] else "ok")
    _card(ct[2],"Sem Base PIS",f"{mt['sem_base_pis']:,}","critico" if mt['sem_base_pis'] else "ok")
    _card(ct[3],"Sem VL PIS",f"{mt['sem_pis']:,}","critico" if mt['sem_pis'] else "ok")
    _card(ct[4],"Sem Base COFINS",f"{mt['sem_base_cof']:,}","critico" if mt['sem_base_cof'] else "ok")
    _card(ct[5],"Sem VL COFINS",f"{mt['sem_cof']:,}","critico" if mt['sem_cof'] else "ok")
    st.markdown("---")
    g1,g2=st.columns(2)
    with g1:
        st.markdown('<div class="sec">Registros por Bloco</div>',unsafe_allow_html=True)
        dfb=df["bloco"].value_counts().reset_index(); dfb.columns=["Bloco","Qtd"]
        fig=px.bar(dfb,x="Bloco",y="Qtd",color="Bloco",color_discrete_sequence=px.colors.qualitative.Set2,template="plotly_white")
        fig.update_layout(showlegend=False,height=240,margin=dict(l=5,r=5,t=5,b=5))
        # ── CORREÇÃO 2: use_container_width -> width='stretch'
        st.plotly_chart(fig,width="stretch")
    with g2:
        st.markdown('<div class="sec">Inconsistências por Tipo</div>',unsafe_allow_html=True)
        if not df_inc.empty and "tipo" in df_inc.columns:
            dft=df_inc["tipo"].value_counts().head(12).reset_index(); dft.columns=["Tipo","Qtd"]
            fig2=px.bar(dft,x="Qtd",y="Tipo",orientation="h",template="plotly_white",
                       color="Qtd",color_continuous_scale=["#FFF0F0","#C0392B"])
            fig2.update_layout(height=240,margin=dict(l=5,r=5,t=5,b=5),coloraxis_showscale=False)
            st.plotly_chart(fig2,width="stretch")
    g3,g4=st.columns(2)
    with g3:
        st.markdown('<div class="sec">CST ICMS — C170</div>',unsafe_allow_html=True)
        df17=df[df["registro"]=="C170"]
        if not df17.empty and "CST_ICMS" in df17.columns:
            dfc=df17["CST_ICMS"].value_counts().reset_index(); dfc.columns=["CST","Qtd"]
            fig3=px.pie(dfc,names="CST",values="Qtd",hole=0.4,template="plotly_white",
                       color_discrete_sequence=px.colors.qualitative.Pastel)
            fig3.update_layout(height=240,margin=dict(l=5,r=5,t=5,b=5))
            st.plotly_chart(fig3,width="stretch")
    with g4:
        st.markdown('<div class="sec">CST PIS — C170</div>',unsafe_allow_html=True)
        if not df17.empty and "CST_PIS" in df17.columns:
            dfp=df17["CST_PIS"].value_counts().reset_index(); dfp.columns=["CST","Qtd"]
            fig4=px.bar(dfp,x="CST",y="Qtd",template="plotly_white",color_discrete_sequence=["#2980B9"])
            fig4.update_layout(height=240,margin=dict(l=5,r=5,t=5,b=5))
            st.plotly_chart(fig4,width="stretch")

# ── Blocos ─────────────────────────────────────────────────────────────────────
def pg_blocos():
    st.markdown('<div class="sec">Visão por Blocos</div>',unsafe_allow_html=True)
    ed=st.session_state.get("editor")
    if not ed: st.info("Carregue um arquivo."); return
    df=ed.df; blocos=sorted(df["bloco"].unique())
    cols=st.columns(min(len(blocos),7))
    for i,b in enumerate(blocos): cols[i%7].metric(f"Bloco {b}",f"{int((df['bloco']==b).sum()):,}")
    st.markdown("---")
    b_sel=st.selectbox("Explorar bloco:",blocos)
    df_b=df[df["bloco"]==b_sel].copy()
    regs=sorted(df_b["registro"].unique())
    r_sel=st.multiselect("Filtrar registros:",regs,default=regs[:6])
    if r_sel: df_b=df_b[df_b["registro"].isin(r_sel)]
    busca=st.text_input("Buscar em qualquer campo:")
    if busca: mask=df_b.astype(str).apply(lambda c:c.str.contains(busca,case=False)).any(axis=1); df_b=df_b[mask]
    st.dataframe(df_b[_cols_rel(df_b)].fillna("").head(500),width="stretch",height=420)
    st.caption(f"{len(df_b):,} registros")

# ── Registros ──────────────────────────────────────────────────────────────────
def pg_registros():
    st.markdown('<div class="sec">Visão por Registros</div>',unsafe_allow_html=True)
    ed=st.session_state.get("editor")
    if not ed: st.info("Carregue um arquivo."); return
    df=ed.df; regs=sorted(df["registro"].unique())
    c1,c2=st.columns([2,3])
    with c1: reg=st.selectbox("Registro:",regs)
    with c2: busca=st.text_input("Buscar:",placeholder="CST, CFOP, valor…")
    df_r=df[df["registro"]==reg].copy()
    if busca: mask=df_r.astype(str).apply(lambda c:c.str.contains(busca,case=False)).any(axis=1); df_r=df_r[mask]
    st.dataframe(df_r[_cols_rel(df_r)].fillna(""),width="stretch",height=440)
    st.caption(f"{len(df_r):,} registros do tipo {reg}")

# ── Notas Fiscais ─────────────────────────────────────────────────────────────
def pg_notas():
    st.markdown('<div class="sec">Notas Fiscais / Documentos (C100 · A100 · D100)</div>',unsafe_allow_html=True)
    ed=st.session_state.get("editor")
    if not ed: st.info("Carregue um arquivo."); return
    df=ed.df
    aba=st.tabs(["NF-e (C100)","Serviços (A100)","Transportes (D100)"])
    for tab,reg,campos_vis in zip(aba,
        ["C100","A100","D100"],
        [["numero_linha","NUM_DOC","DT_DOC","COD_PART","VL_DOC","VL_BC_ICMS","VL_ICMS","VL_PIS","VL_COFINS","COD_SIT"],
         ["numero_linha","NUM_DOC","DT_DOC","COD_PART","VL_DOC","VL_BC_PIS","ALIQ_PIS","VL_PIS","VL_BC_COFINS","ALIQ_COFINS","VL_COFINS"],
         ["numero_linha","NUM_DOC","DT_DOC","COD_PART","VL_DOC","VL_BC_ICMS","VL_ICMS","VL_BC_PIS","VL_PIS","VL_COFINS"]]):
        with tab:
            dfd=df[df["registro"]==reg].copy()
            if dfd.empty: st.info(f"Nenhum {reg} encontrado."); continue
            cols=[c for c in campos_vis if c in dfd.columns]
            c1,c2=st.columns(2)
            with c1:
                ps=["Todos"]+sorted(dfd["COD_PART"].dropna().unique().tolist()) if "COD_PART" in dfd.columns else ["Todos"]
                pf=st.selectbox("Parceiro:",ps,key=f"pf_{reg}")
            if pf!="Todos" and "COD_PART" in dfd.columns: dfd=dfd[dfd["COD_PART"]==pf]
            st.dataframe(dfd[cols].fillna(""),width="stretch",height=380)
            st.caption(f"{len(dfd):,} documentos")
            ns=[c for c in ["VL_DOC","VL_ICMS","VL_PIS","VL_COFINS"] if c in dfd.columns]
            if ns:
                ct=st.columns(len(ns))
                for i,c in enumerate(ns): ct[i].metric(c,_fmt_brl(pd.to_numeric(dfd[c],errors="coerce").sum()))

# ── Itens C170/A170 ───────────────────────────────────────────────────────────
def pg_itens():
    st.markdown('<div class="sec">Itens de NF-e (C170) e Serviços (A170)</div>',unsafe_allow_html=True)
    ed=st.session_state.get("editor")
    if not ed: st.info("Carregue um arquivo."); return
    df=ed.df; df_inc=st.session_state.get("df_inc",pd.DataFrame())
    aba=st.tabs(["C170 — Itens NF-e","A170 — Itens Serviço","C185 — PIS/COFINS Itens","F100 — Outras Op."])
    configs=[
        ("C170",["numero_linha","COD_ITEM","DESCR_COMPL","QTD","VL_ITEM","CST_ICMS","CFOP",
                 "VL_BC_ICMS","ALIQ_ICMS","VL_ICMS","CST_PIS","VL_BC_PIS","ALIQ_PIS","VL_PIS",
                 "CST_COFINS","VL_BC_COFINS","ALIQ_COFINS","VL_COFINS"],"CST_ICMS"),
        ("A170",["numero_linha","COD_ITEM","DESCR_COMPL","VL_ITEM","CST_PIS","VL_BC_PIS",
                 "ALIQ_PIS","VL_PIS","CST_COFINS","VL_BC_COFINS","ALIQ_COFINS","VL_COFINS"],"CST_PIS"),
        ("C185",["numero_linha","COD_ITEM","CST_PIS","VL_BC_PIS","ALIQ_PIS","VL_PIS",
                 "CST_COFINS","VL_BC_COFINS","ALIQ_COFINS","VL_COFINS"],"CST_PIS"),
        ("F100",["numero_linha","COD_PART","DT_OPER","VL_OPER","VL_BC_PIS","ALIQ_PIS","VL_PIS",
                 "VL_BC_COFINS","ALIQ_COFINS","VL_COFINS"],""),
    ]
    for tab,(reg,cv,campo_cst) in zip(aba,configs):
        with tab:
            dfd=df[df["registro"]==reg].copy()
            if dfd.empty: st.info(f"Nenhum {reg} encontrado."); continue
            c1,c2,c3=st.columns(3)
            with c1:
                if campo_cst and campo_cst in dfd.columns:
                    csts=["Todos"]+sorted(dfd[campo_cst].dropna().unique().tolist())
                    cf=st.selectbox("CST:",csts,key=f"cst_{reg}")
                    if cf!="Todos": dfd=dfd[dfd[campo_cst]==cf]
            with c2:
                if "CFOP" in dfd.columns:
                    cfops=["Todos"]+sorted(dfd["CFOP"].dropna().unique().tolist())
                    cff=st.selectbox("CFOP:",cfops,key=f"cfop_{reg}")
                    if cff!="Todos": dfd=dfd[dfd["CFOP"]==cff]
            with c3:
                ap=st.checkbox("Apenas com inconsistências",key=f"ap_{reg}")
                if ap and not df_inc.empty:
                    li=set(df_inc["numero_linha"].tolist()); dfd=dfd[dfd["numero_linha"].isin(li)]
            cols=[c for c in cv if c in dfd.columns]
            st.dataframe(dfd[cols].fillna(""),width="stretch",height=400)
            st.caption(f"{len(dfd):,} registros")

# ── PIS/COFINS — apuração ─────────────────────────────────────────────────────
def pg_pis_cofins():
    st.markdown('<div class="sec">PIS / COFINS — Apuração (Bloco M)</div>',unsafe_allow_html=True)
    ed=st.session_state.get("editor")
    if not ed: st.info("Carregue um arquivo."); return
    df=ed.df
    res=st.session_state.get("resultado")
    if res and res.tipo_arquivo!="EFD_CONTRIBUICOES":
        st.warning("⚠️ Arquivo identificado como EFD ICMS/IPI. Bloco M pode estar ausente.")
    tabs=st.tabs(["M100 — Créditos PIS","M200 — Apuração PIS","M210 — Detalhamento PIS",
                  "M500 — Créditos COFINS","M600 — Apuração COFINS","M610 — Detalhamento COFINS",
                  "P100 — Contrib. Previdenciária"])
    for tab,reg in zip(tabs,["M100","M200","M210","M500","M600","M610","P100"]):
        with tab:
            dfd=df[df["registro"]==reg].copy()
            if dfd.empty: st.info(f"Nenhum registro {reg} encontrado."); continue
            st.dataframe(dfd[_cols_rel(dfd)].fillna(""),width="stretch",height=350)
            st.caption(f"{len(dfd):,} registros {reg}")
            if reg in ("M200","M600"):
                cols_tot=[c for c in ["VL_TOT_CONT_NC_PER","VL_TOT_CRED_DESC","VL_TOT_CONT_REC",
                                      "VL_CONT_NC_REC","VL_CONT_CUM_REC"] if c in dfd.columns]
                if cols_tot:
                    ct=st.columns(len(cols_tot))
                    for i,c in enumerate(cols_tot):
                        ct[i].metric(c,_fmt_brl(pd.to_numeric(dfd[c],errors="coerce").sum()))

# ── Inconsistências ───────────────────────────────────────────────────────────
def pg_inconsistencias():
    st.markdown('<div class="sec">Inconsistências Fiscais</div>',unsafe_allow_html=True)
    df_inc=st.session_state.get("df_inc",pd.DataFrame())
    res=st.session_state.get("resultado")
    if res and (df_inc is None or df_inc.empty):
        st.success("✅ Nenhuma inconsistência detectada."); return
    if not res: st.info("Carregue um arquivo."); return
    c1,c2,c3,c4=st.columns(4)
    nc=int((df_inc["severidade"]=="CRITICA").sum()) if "severidade" in df_inc.columns else 0
    na=int((df_inc["severidade"]=="AVISO").sum()) if "severidade" in df_inc.columns else 0
    _card(c1,"Total",len(df_inc)); _card(c2,"Críticas",nc,"critico")
    _card(c3,"Avisos",na,"aviso")
    _card(c4,"Corrigidos",int(df_inc["corrigido"].sum()) if "corrigido" in df_inc.columns else 0,"ok")
    st.markdown("---")
    c1,c2,c3,c4,c5=st.columns(5)
    sev_ops=["Todos"]+sorted(df_inc["severidade"].unique().tolist()) if "severidade" in df_inc.columns else ["Todos"]
    tipo_ops=["Todos"]+sorted(df_inc["tipo"].unique().tolist()) if "tipo" in df_inc.columns else ["Todos"]
    trib_ops=["Todos","ICMS","PIS","COFINS","IPI","GERAL"]
    cst_ops=["Todos"]+sorted(df_inc["cst"].dropna().unique().tolist()) if "cst" in df_inc.columns else ["Todos"]
    with c1: sf=st.selectbox("Severidade:",sev_ops)
    with c2: tf=st.selectbox("Tipo:",tipo_ops)
    with c3: trif=st.selectbox("Tributo:",trib_ops)
    with c4: cf=st.selectbox("CST:",cst_ops)
    with c5: bf=st.text_input("Buscar descrição:")
    df_f=df_inc.copy()
    if sf!="Todos": df_f=df_f[df_f["severidade"]==sf]
    if tf!="Todos": df_f=df_f[df_f["tipo"]==tf]
    if trif!="Todos": df_f=df_f[df_f["tipo"].str.contains(trif,na=False)]
    if cf!="Todos": df_f=df_f[df_f["cst"]==cf]
    if bf: df_f=df_f[df_f["descricao"].astype(str).str.contains(bf,case=False)]
    def cor_linha(row):
        if row.get("severidade")=="CRITICA": return ["background-color:#FFF0F0"]*len(row)
        if row.get("severidade")=="AVISO":   return ["background-color:#FFFDE8"]*len(row)
        return [""]*len(row)
    cols_show=[c for c in df_f.columns if c!="linha_original"]
    st.dataframe(df_f[cols_show].style.apply(cor_linha,axis=1),width="stretch",height=460)
    st.caption(f"{len(df_f):,} inconsistências")
    if st.button("🔄 Revalidar Arquivo"):
        ed=st.session_state.get("editor"); regras=st.session_state.get("regras")
        if ed and regras:
            st.session_state["df_inc"]=validar(ed.df,regras,res.tipo_arquivo); st.rerun()

# ── Correções em Massa ────────────────────────────────────────────────────────
def pg_massa():
    st.markdown('<div class="sec">Correções em Massa</div>',unsafe_allow_html=True)
    ed=st.session_state.get("editor"); regras=st.session_state.get("regras",[])
    df_inc=st.session_state.get("df_inc",pd.DataFrame())
    usuario=st.session_state.get("_usuario","analista")
    if not ed: st.info("Carregue um arquivo."); return

    c1,c2=st.columns(2)
    with c1: tributo=st.selectbox("Tributo a corrigir:",["ICMS","PIS","COFINS"],key="m_trib")
    with c2:
        reg_opts=["C170","A170","C185","D100","F100"]
        reg_alvo=st.selectbox("Registro alvo:",reg_opts,key="m_reg")

    mapa_campos_trib={
        "ICMS":{"cst":"CST_ICMS","bc":"VL_BC_ICMS","aliq":"ALIQ_ICMS","vl":"VL_ICMS"},
        "PIS": {"cst":"CST_PIS", "bc":"VL_BC_PIS", "aliq":"ALIQ_PIS", "vl":"VL_PIS"},
        "COFINS":{"cst":"CST_COFINS","bc":"VL_BC_COFINS","aliq":"ALIQ_COFINS","vl":"VL_COFINS"},
    }
    fc=mapa_campos_trib[tributo]

    df_reg=ed.df[ed.df["registro"]==reg_alvo].copy()
    if df_reg.empty: st.warning(f"Nenhum {reg_alvo} encontrado."); return

    st.markdown("**Filtros de seleção**")
    c1,c2,c3,c4=st.columns(4)
    with c1:
        cs_opts=["Todos"]+sorted(df_reg[fc["cst"]].dropna().unique().tolist()) if fc["cst"] in df_reg.columns else ["Todos"]
        csf=st.selectbox(f"CST {tributo}:",cs_opts,key="mc_cst")
    with c2:
        cf_opts=["Todos"]+sorted(df_reg["CFOP"].dropna().unique().tolist()) if "CFOP" in df_reg.columns else ["Todos"]
        cff=st.selectbox("CFOP:",cf_opts,key="mc_cfop")
    with c3:
        ti_ops=["Todos"]+(sorted(df_inc["tipo"].unique().tolist()) if not df_inc.empty and "tipo" in df_inc.columns else [])
        tif=st.selectbox("Tipo inconsistência:",ti_ops,key="mc_tipo")
    with c4:
        ap=st.checkbox("Apenas com inconsistência",value=True,key="mc_ap")

    df_alvo=df_reg.copy()
    if csf!="Todos" and fc["cst"] in df_alvo.columns: df_alvo=df_alvo[df_alvo[fc["cst"]]==csf]
    if cff!="Todos" and "CFOP" in df_alvo.columns: df_alvo=df_alvo[df_alvo["CFOP"]==cff]
    if ap and not df_inc.empty:
        li=set(df_inc[df_inc["tipo"]==tif]["numero_linha"].tolist() if tif!="Todos" else df_inc["numero_linha"].tolist())
        df_alvo=df_alvo[df_alvo["numero_linha"].isin(li)]

    st.info(f"**{len(df_alvo)} item(ns) selecionado(s) em {reg_alvo} — tributo {tributo}**")
    cols_vis=[c for c in ["numero_linha",fc["cst"],"CFOP","VL_ITEM",fc["bc"],fc["aliq"],fc["vl"]] if c in df_alvo.columns]
    st.dataframe(df_alvo[cols_vis].fillna("").head(200),width="stretch",height=180)
    if df_alvo.empty: return

    nls=df_alvo["numero_linha"].tolist()
    st.markdown("---")
    tab1,tab2,tab3,tab4=st.tabs(["📐 Preencher Base","📊 Preencher Alíquota","🧮 Recalcular Valor","👁️ Prévia"])

    with tab1:
        fonte=st.radio("Origem:",["VL_ITEM do item","VL_DOC do documento","Valor manual"],key="fb")
        vman=0.0
        if "manual" in fonte: vman=st.number_input("Valor (R$):",min_value=0.0,step=0.01,key="bman")
        mot=st.text_input("Motivo:",f"Base {tributo} preenchida via {fonte}",key="mb")
        if st.button(f"▶ Aplicar Base {tributo}",type="primary"):
            n=0
            for _,row in df_alvo.iterrows():
                nl=int(row["numero_linha"])
                if "manual" in fonte: v=_to_sped_str(vman)
                elif "VL_ITEM" in fonte: vi=_to_float(row.get("VL_ITEM")); v=_to_sped_str(vi) if vi else ""
                else: vd=_to_float(row.get("VL_DOC")); v=_to_sped_str(vd) if vd else ""
                if v and ed.editar(nl,fc["bc"],v,mot,"",usuario): n+=1
            st.success(f"✅ {n} base(s) de {tributo} preenchidas.")

    with tab2:
        aliq_v=st.number_input("Alíquota (%):",0.0,100.0,{"ICMS":12.0,"PIS":1.65,"COFINS":7.6}[tributo],0.01,key="maliq")
        mot2=st.text_input("Motivo:",f"Alíquota {tributo} padrão aplicada",key="ma2")
        if st.button(f"▶ Aplicar Alíquota {tributo}",type="primary"):
            n=ed.massa(nls,fc["aliq"],_to_sped_str(aliq_v),mot2,"",usuario)
            st.success(f"✅ {n} alíquota(s) de {tributo} preenchida(s).")

    with tab3:
        mot3=st.text_input("Motivo:",f"Recálculo {tributo} = Base × Alíq ÷ 100",key="mc3")
        if st.button(f"▶ Recalcular {tributo}",type="primary"):
            n=ed.recalcular_massa(nls,regras,tributo,fc["cst"],fc["bc"],fc["aliq"],fc["vl"],mot3,usuario)
            st.success(f"✅ {n} valor(es) de {tributo} recalculado(s).")

    with tab4:
        pv=ed.preview(nls)
        if not pv.empty: st.dataframe(pv,width="stretch")
        else: st.info("Nenhuma alteração pendente.")

    st.markdown("---")
    if st.button("↩️ Desfazer última operação"):
        if ed.desfazer(): st.success("✅ Desfeito."); st.rerun()
        else: st.warning("Nada para desfazer.")

# ── Editor Manual ─────────────────────────────────────────────────────────────
def pg_editor():
    st.markdown('<div class="sec">Editor Manual de Registros</div>',unsafe_allow_html=True)
    ed=st.session_state.get("editor"); usuario=st.session_state.get("_usuario","analista")
    if not ed: st.info("Carregue um arquivo."); return
    df=ed.df; regs=sorted(df["registro"].unique())
    c1,c2,c3=st.columns([2,2,1])
    with c1: reg_sel=st.selectbox("Registro:",regs)
    df_reg=df[df["registro"]==reg_sel]
    with c2:
        nl_min=int(df_reg["numero_linha"].min()) if not df_reg.empty else 1
        nl_max=int(df_reg["numero_linha"].max()) if not df_reg.empty else 1
        nl=st.number_input("Linha:",nl_min,nl_max,nl_min)
    with c3: motivo=st.text_input("Motivo:",key="em")
    df_l=df[df["numero_linha"]==nl]
    if df_l.empty: st.warning("Linha não encontrada."); return
    row_a=df_l.iloc[0]
    row_o_df=ed.df_original[ed.df_original["numero_linha"]==nl]
    row_o=row_o_df.iloc[0] if not row_o_df.empty else row_a
    mapa=MAPA_CAMPOS.get(reg_sel,{})
    campos=list(mapa.keys()) if mapa else [c for c in df.columns if not c.startswith("campo_") and c not in ["numero_linha","bloco","registro","linha_original"]]
    alts=[c for c in campos if str(row_a.get(c,""))!=str(row_o.get(c,""))]
    if alts: st.warning(f"⚠️ Campo(s) alterado(s): **{', '.join(alts)}**")
    st.markdown(f"**Editando: {reg_sel} — Linha {nl}**")
    pend={}; cc=st.columns(3)
    for i,campo in enumerate(campos):
        va=str(row_a.get(campo,"") or ""); vo=str(row_o.get(campo,"") or "")
        dest="🔴 " if va!=vo else ""
        novo=cc[i%3].text_input(f"{dest}{campo}",value=va,key=f"ed_{campo}_{nl}")
        if novo!=va: pend[campo]=novo
    ca,cb,cc2=st.columns(3)
    with ca:
        if st.button("💾 Salvar",type="primary",disabled=not pend):
            for c,v in pend.items(): ed.editar(nl,c,v,motivo,"",usuario)
            st.success(f"✅ {len(pend)} campo(s) salvo(s)."); st.rerun()
    with cb:
        if st.button("↩️ Desfazer"):
            if ed.desfazer(): st.success("✅ Desfeito."); st.rerun()
            else: st.warning("Nada para desfazer.")
    with cc2:
        if st.button("🔄 Restaurar linha"):
            ed.restaurar([nl]); st.success("✅ Restaurado."); st.rerun()
    st.markdown("---")
    diffs=[{"Campo":c,"Original":str(row_o.get(c,"")),"Atual":str(row_a.get(c,""))} for c in campos if str(row_a.get(c,""))!=str(row_o.get(c,""))]
    if diffs: st.dataframe(pd.DataFrame(diffs),width="stretch")
    else: st.success("Nenhuma diferença nesta linha.")

# ── Exportação ────────────────────────────────────────────────────────────────
def pg_exportacao():
    st.markdown('<div class="sec">Exportação</div>',unsafe_allow_html=True)
    ed=st.session_state.get("editor"); res=st.session_state.get("resultado")
    df_inc=st.session_state.get("df_inc",pd.DataFrame()); regras=st.session_state.get("regras",[])
    if not ed or not res: st.info("Carregue um arquivo."); return
    m=res.metadados
    meta={"nome_empresa":m.nome_empresa,"cnpj":m.cnpj,"periodo_apuracao":m.periodo_apuracao,"uf":m.uf,"tipo_arquivo":res.tipo_arquivo}
    c1,c2=st.columns(2)
    with c1:
        st.markdown("#### 📄 SPED TXT Corrigido")
        if st.button("Gerar SPED TXT",type="primary"):
            lo=st.session_state.get("linhas_orig",res.linhas)
            lc=sync_df_linhas(ed.df,lo)
            bts=reconstruir_sped(lc)
            nome=f"SPED_CORRIGIDO_{m.cnpj}_{res.tipo_arquivo}.txt"
            st.download_button("⬇️ Baixar TXT",bts,nome,"text/plain")
    with c2:
        st.markdown("#### 📊 Excel de Auditoria (5 abas)")
        if st.button("Gerar Excel",type="primary"):
            df_alt=ed.get_alterados(); df_aud=ed.auditoria.to_df()
            df_reg=pd.DataFrame([asdict(r) for r in regras])
            bts=exportar_excel(df_inc or pd.DataFrame(),df_alt,df_aud,df_reg,meta)
            st.download_button("⬇️ Baixar Excel",bts,f"AUDITORIA_{m.cnpj}.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.markdown("---"); c3,c4=st.columns(2)
    with c3:
        st.markdown("#### 📋 CSV Inconsistências")
        if df_inc is not None and not df_inc.empty:
            st.download_button("⬇️ Baixar CSV",exportar_csv(df_inc),"inconsistencias.csv","text/csv")
        else: st.info("Sem inconsistências.")
    with c4:
        st.markdown("#### 📋 CSV Auditoria")
        df_aud=ed.auditoria.to_df()
        if not df_aud.empty:
            st.download_button("⬇️ Baixar CSV",exportar_csv(df_aud),"log_auditoria.csv","text/csv")
        else: st.info("Sem alterações.")

# ── Log de Auditoria ──────────────────────────────────────────────────────────
def pg_auditoria():
    st.markdown('<div class="sec">Log de Auditoria</div>',unsafe_allow_html=True)
    ed=st.session_state.get("editor")
    if not ed: st.info("Carregue um arquivo."); return
    df_aud=ed.auditoria.to_df()
    if df_aud.empty: st.info("Nenhuma alteração registrada."); return
    tipos=df_aud["Tipo"].value_counts().to_dict() if "Tipo" in df_aud.columns else {}
    c1,c2,c3=st.columns(3)
    _card(c1,"Total Alterações",ed.auditoria.total())
    _card(c2,"Manuais",tipos.get("MANUAL",0))
    _card(c3,"Massa/Auto",tipos.get("MASSA",0)+tipos.get("AUTO",0))
    st.markdown("---")
    c1,c2,c3=st.columns(3)
    with c1: tf=st.selectbox("Tipo:",["Todos"]+sorted(df_aud["Tipo"].unique().tolist()) if "Tipo" in df_aud.columns else ["Todos"])
    with c2: cf=st.selectbox("Campo:",["Todos"]+sorted(df_aud["Campo"].unique().tolist()) if "Campo" in df_aud.columns else ["Todos"])
    with c3: bf=st.text_input("Buscar motivo/regra:")
    df_f=df_aud.copy()
    if tf!="Todos": df_f=df_f[df_f["Tipo"]==tf]
    if cf!="Todos": df_f=df_f[df_f["Campo"]==cf]
    if bf: df_f=df_f[df_f["Motivo"].astype(str).str.contains(bf,case=False)|df_f["Regra"].astype(str).str.contains(bf,case=False)]
    st.dataframe(df_f,width="stretch",height=490)

# ── Motor de Regras ───────────────────────────────────────────────────────────
def pg_regras():
    st.markdown('<div class="sec">Motor de Regras Tributárias</div>',unsafe_allow_html=True)
    regras=st.session_state.get("regras") or carregar_regras()
    st.session_state["regras"]=regras
    tab1,tab2=st.tabs(["📋 Regras Ativas","➕ Cadastrar / Editar"])
    with tab1:
        trib_f=st.selectbox("Filtrar por tributo:",["Todos","ICMS","PIS","COFINS","IPI"],key="rf_t")
        df_r=pd.DataFrame([asdict(r) for r in regras])
        if trib_f!="Todos" and "tributo" in df_r.columns: df_r=df_r[df_r["tributo"]==trib_f]
        st.dataframe(df_r,width="stretch",height=420)
        c1,c2=st.columns(2)
        with c1:
            if st.button("🔄 Resetar para padrão"):
                st.session_state["regras"]=list(REGRAS_PADRAO); salvar_regras(REGRAS_PADRAO)
                st.success("✅ Regras padrão restauradas."); st.rerun()
        with c2:
            id_del=st.text_input("ID para remover:",key="rdel")
            if st.button("🗑️ Remover") and id_del:
                st.session_state["regras"]=[r for r in regras if r.id!=id_del]
                salvar_regras(st.session_state["regras"]); st.success(f"✅ Removido '{id_del}'."); st.rerun()
    with tab2:
        c1,c2,c3=st.columns(3)
        with c1:
            ni=st.text_input("ID:",key="ni"); nt=st.selectbox("Tributo:",["ICMS","PIS","COFINS","IPI"],key="nt")
            nc=st.text_input("CST (* = todos):",key="nc"); nf=st.text_input("CFOP (* = todos):",value="*",key="nf")
            no=st.selectbox("Operação:",["*","E","S"],key="no")
        with c2:
            nd=st.text_area("Descrição:",key="nd")
            neb=st.checkbox("Exige base",value=True,key="neb"); nea=st.checkbox("Exige alíquota",value=True,key="nea")
            nev=st.checkbox("Exige valor",value=True,key="nev")
        with c3:
            nal=st.number_input("Alíquota padrão (%):",0.0,100.0,12.0,key="nal")
            nbc=st.selectbox("Campo base sugerida:",["VL_ITEM","VL_DOC","VL_OPER",""],key="nbc")
            nfo=st.selectbox("Fórmula:",["base * aliq / 100","zero"],key="nfo")
            ncr=st.selectbox("Criticidade:",["critica","aviso","info"],key="ncr")
        if st.button("💾 Salvar Regra",type="primary"):
            if not ni or not nc: st.error("ID e CST obrigatórios.")
            else:
                nova=RegraFiscal(id=ni,tributo=nt,cst=nc,cfop=nf,ind_oper=no,descricao=nd,
                                 exige_base=neb,exige_aliquota=nea,exige_valor=nev,
                                 base_campo_sugerido=nbc,aliquota_padrao=nal,formula=nfo,
                                 permite_base_zero=not neb,permite_valor_zero=not nev,criticidade=ncr)
                st.session_state["regras"]=[r for r in regras if r.id!=ni]+[nova]
                salvar_regras(st.session_state["regras"]); st.success(f"✅ Regra '{ni}' salva."); st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# ████████████  MAIN  ████████████████████████████████████████████████████████
# ═══════════════════════════════════════════════════════════════════════════════

ROTEADOR={
    "📊 Dashboard":          pg_dashboard,
    "📂 Upload":             pg_upload,
    "🗂️ Blocos":             pg_blocos,
    "📋 Registros":          pg_registros,
    "🧾 Notas Fiscais":      pg_notas,
    "📦 Itens (C170/A170)":  pg_itens,
    "💰 PIS/COFINS":         pg_pis_cofins,
    "⚠️ Inconsistências":    pg_inconsistencias,
    "🔧 Correções em Massa": pg_massa,
    "✏️ Editor Manual":      pg_editor,
    "📤 Exportação":         pg_exportacao,
    "📜 Log de Auditoria":   pg_auditoria,
    "⚙️ Motor de Regras":    pg_regras,
}

def main():
    _init()
    st.markdown(CSS,unsafe_allow_html=True)
    pagina=sidebar()
    fn=ROTEADOR.get(pagina)
    if fn: fn()
    else: st.error(f"Página não encontrada: {pagina}")

if __name__=="__main__":
    main()
