"""
╔══════════════════════════════════════════════════════════════════════════════════╗
║              SPED AUDITOR — FISCAL INTELLIGENCE PLATFORM                        ║
║              Versão 1.0 | Python + Streamlit | Arquivo único                   ║
║──────────────────────────────────────────────────────────────────────────────────║
║  Módulos embutidos:                                                              ║
║    • Parser SPED manual (EFD ICMS/IPI e EFD Contribuições)                      ║
║    • Motor de regras tributárias configurável (CST × CFOP)                      ║
║    • Validador fiscal (11 tipos de inconsistência)                               ║
║    • Editor com trilha de auditoria imutável                                     ║
║    • Correções em massa com prévia e desfazer                                    ║
║    • Exportação: SPED TXT corrigido, Excel 5-abas, CSV                          ║
╚══════════════════════════════════════════════════════════════════════════════════╝

Execução:
    pip install streamlit pandas numpy openpyxl plotly
    streamlit run sped_auditor.py
"""

# ════════════════════════════════════════════════════════════════════════════════
# IMPORTS
# ════════════════════════════════════════════════════════════════════════════════

from __future__ import annotations

import io
import json
import os
import re
import sys
import uuid
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

# ════════════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO DA PÁGINA — deve preceder qualquer chamada st.*
# ════════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="SPED Auditor — Fiscal Intelligence",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={"About": "SPED Auditor v1.0 | Auditoria Fiscal de Arquivos SPED"},
)

# ════════════════════════════════════════════════════════════════════════════════
# ████████████████████████  CAMADA DE DADOS  ████████████████████████████████████
# ════════════════════════════════════════════════════════════════════════════════

# ─────────────────────────────────────────────────────────────────────────────
# 1. MODELOS DE DADOS
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class LinhaRegistro:
    """Uma linha do arquivo SPED após tokenização por '|'."""
    numero_linha: int
    bloco: str
    registro: str
    campos: list[str]
    linha_original: str

    def get(self, indice: int, default: str = "") -> str:
        try:
            return self.campos[indice] if indice < len(self.campos) else default
        except IndexError:
            return default

    def set(self, indice: int, valor: str) -> None:
        while len(self.campos) <= indice:
            self.campos.append("")
        self.campos[indice] = valor

    def to_sped(self) -> str:
        return "|" + "|".join(self.campos) + "|\n"


@dataclass
class MetadadosArquivo:
    """Metadados extraídos do registro 0000."""
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
    """Saída completa do parser."""
    metadados: MetadadosArquivo
    linhas: list[LinhaRegistro]
    df: pd.DataFrame
    erros: list[str]
    tipo_arquivo: str  # 'EFD_ICMS_IPI' | 'EFD_CONTRIBUICOES' | 'DESCONHECIDO'


@dataclass
class RegraFiscal:
    """Define o comportamento esperado para CST + CFOP."""
    id: str
    cst_icms: str          # valor exato ou "*"
    cfop: str              # valor exato, prefixo ou "*"
    ind_oper: str          # "E" | "S" | "*"
    descricao: str
    exige_base: bool
    exige_aliquota: bool
    exige_valor_icms: bool
    base_campo_sugerido: str   # "VL_ITEM" | "VL_DOC" | ""
    aliquota_padrao: Optional[float]
    formula: str           # "base * aliq / 100" | "zero"
    permite_base_zero: bool
    permite_icms_zero: bool
    criticidade: str       # "critica" | "aviso" | "info"
    ativa: bool = True

    def match_cst(self, cst: str) -> bool:
        return self.cst_icms == "*" or cst == self.cst_icms or cst.startswith(self.cst_icms)

    def match_cfop(self, cfop: str) -> bool:
        return self.cfop == "*" or cfop == self.cfop or cfop.startswith(self.cfop)

    def match_oper(self, op: str) -> bool:
        return self.ind_oper in ("*", op)


@dataclass
class EntradaAuditoria:
    """Registro imutável de uma alteração realizada."""
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
    tipo: str  # "MANUAL" | "MASSA" | "AUTO"


# ─────────────────────────────────────────────────────────────────────────────
# 2. MAPA DE CAMPOS POR REGISTRO (índice posicional → nome semântico)
# ─────────────────────────────────────────────────────────────────────────────

MAPA_CAMPOS: dict[str, dict[str, int]] = {
    "0000": {
        "COD_VER": 1, "COD_FIN": 2, "DT_INI": 3, "DT_FIN": 4,
        "NOME": 5, "CNPJ": 6, "CPF": 7, "UF": 8, "IE": 9,
        "COD_MUN": 10, "IND_PERFIL": 13,
    },
    "0150": {
        "COD_PART": 1, "NOME": 2, "COD_PAIS": 3, "CNPJ": 4,
        "CPF": 5, "IE": 6, "COD_MUN": 7, "END": 9, "BAIRRO": 12,
    },
    "0200": {
        "COD_ITEM": 1, "DESCR_ITEM": 2, "COD_BARRA": 3,
        "UNID_INV": 5, "TIPO_ITEM": 6, "COD_NCM": 7, "ALIQ_ICMS": 11,
    },
    "C100": {
        "IND_OPER": 1, "IND_EMIT": 2, "COD_PART": 3, "COD_MOD": 4,
        "COD_SIT": 5, "SER": 6, "NUM_DOC": 7, "CHV_NFE": 8,
        "DT_DOC": 9, "DT_E_S": 10, "VL_DOC": 11,
        "VL_DESC": 13, "VL_MERC": 15,
        "VL_BC_ICMS": 20, "VL_ICMS": 21,
        "VL_BC_ICMS_ST": 22, "VL_ICMS_ST": 23,
        "VL_IPI": 24, "VL_PIS": 25, "VL_COFINS": 26,
    },
    "C170": {
        "NUM_ITEM": 1, "COD_ITEM": 2, "DESCR_COMPL": 3,
        "QTD": 4, "UNID": 5, "VL_ITEM": 6, "VL_DESC": 7, "IND_MOV": 8,
        "CST_ICMS": 9, "CFOP": 10,
        "VL_BC_ICMS": 12, "ALIQ_ICMS": 13, "VL_ICMS": 14,
        "VL_BC_ICMS_ST": 15, "ALIQ_ST": 16, "VL_ICMS_ST": 17,
        "CST_IPI": 19, "VL_BC_IPI": 20, "ALIQ_IPI": 21, "VL_IPI": 22,
        "CST_PIS": 23, "VL_BC_PIS": 24, "ALIQ_PIS": 25, "VL_PIS": 26,
        "CST_COFINS": 27, "VL_BC_COFINS": 28, "ALIQ_COFINS": 29, "VL_COFINS": 30,
    },
    "C190": {
        "CST_ICMS": 1, "CFOP": 2, "ALIQ_ICMS": 3,
        "VL_OPR": 4, "VL_BC_ICMS": 5, "VL_ICMS": 6,
        "VL_BC_ICMS_ST": 7, "VL_ICMS_ST": 8,
    },
    "D100": {
        "IND_OPER": 1, "COD_PART": 3, "COD_MOD": 4, "COD_SIT": 5,
        "NUM_DOC": 8, "DT_DOC": 10, "VL_DOC": 14,
        "VL_BC_ICMS": 18, "ALIQ_ICMS": 19, "VL_ICMS": 20,
    },
    "E110": {
        "VL_TOT_DEBITOS": 1, "VL_TOT_CREDITOS": 5,
        "VL_SLD_APURADO": 10, "VL_ICMS_RECOLHER": 12,
    },
}

CAMPOS_NUMERICOS = {
    "VL_ITEM", "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS",
    "VL_BC_ICMS_ST", "ALIQ_ST", "VL_ICMS_ST",
    "VL_IPI", "VL_BC_IPI", "ALIQ_IPI",
    "VL_PIS", "VL_COFINS", "VL_BC_PIS", "VL_BC_COFINS",
    "ALIQ_PIS", "ALIQ_COFINS", "VL_DOC", "VL_MERC", "QTD",
}

# ─────────────────────────────────────────────────────────────────────────────
# 3. REGRAS FISCAIS PADRÃO
# ─────────────────────────────────────────────────────────────────────────────

REGRAS_PADRAO: list[RegraFiscal] = [
    RegraFiscal(
        id="000_*", cst_icms="000", cfop="*", ind_oper="*",
        descricao="Tributado Integralmente — base, alíquota e imposto obrigatórios",
        exige_base=True, exige_aliquota=True, exige_valor_icms=True,
        base_campo_sugerido="VL_ITEM", aliquota_padrao=12.0,
        formula="base * aliq / 100",
        permite_base_zero=False, permite_icms_zero=False, criticidade="critica",
    ),
    RegraFiscal(
        id="010_*", cst_icms="010", cfop="*", ind_oper="*",
        descricao="Tributado + ST — base, alíquota e imposto obrigatórios",
        exige_base=True, exige_aliquota=True, exige_valor_icms=True,
        base_campo_sugerido="VL_ITEM", aliquota_padrao=12.0,
        formula="base * aliq / 100",
        permite_base_zero=False, permite_icms_zero=False, criticidade="critica",
    ),
    RegraFiscal(
        id="020_*", cst_icms="020", cfop="*", ind_oper="*",
        descricao="Tributado com Redução de BC — base reduzida obrigatória",
        exige_base=True, exige_aliquota=True, exige_valor_icms=True,
        base_campo_sugerido="VL_ITEM", aliquota_padrao=12.0,
        formula="base * aliq / 100",
        permite_base_zero=False, permite_icms_zero=False, criticidade="critica",
    ),
    RegraFiscal(
        id="040_*", cst_icms="040", cfop="*", ind_oper="*",
        descricao="Isento — não deve ter valores tributários",
        exige_base=False, exige_aliquota=False, exige_valor_icms=False,
        base_campo_sugerido="", aliquota_padrao=0.0,
        formula="zero",
        permite_base_zero=True, permite_icms_zero=True, criticidade="aviso",
    ),
    RegraFiscal(
        id="041_*", cst_icms="041", cfop="*", ind_oper="*",
        descricao="Não Tributado — sem valores tributários",
        exige_base=False, exige_aliquota=False, exige_valor_icms=False,
        base_campo_sugerido="", aliquota_padrao=0.0,
        formula="zero",
        permite_base_zero=True, permite_icms_zero=True, criticidade="aviso",
    ),
    RegraFiscal(
        id="050_*", cst_icms="050", cfop="*", ind_oper="*",
        descricao="Suspensão — sem valores tributários",
        exige_base=False, exige_aliquota=False, exige_valor_icms=False,
        base_campo_sugerido="", aliquota_padrao=0.0,
        formula="zero",
        permite_base_zero=True, permite_icms_zero=True, criticidade="aviso",
    ),
    RegraFiscal(
        id="060_*", cst_icms="060", cfop="*", ind_oper="*",
        descricao="ST Cobrado Anteriormente — verificar campos ST",
        exige_base=False, exige_aliquota=False, exige_valor_icms=False,
        base_campo_sugerido="", aliquota_padrao=0.0,
        formula="zero",
        permite_base_zero=True, permite_icms_zero=True, criticidade="info",
    ),
    RegraFiscal(
        id="070_*", cst_icms="070", cfop="*", ind_oper="*",
        descricao="Redução BC + ST — exige próprio e ST",
        exige_base=True, exige_aliquota=True, exige_valor_icms=True,
        base_campo_sugerido="VL_ITEM", aliquota_padrao=12.0,
        formula="base * aliq / 100",
        permite_base_zero=False, permite_icms_zero=False, criticidade="critica",
    ),
    RegraFiscal(
        id="090_*", cst_icms="090", cfop="*", ind_oper="*",
        descricao="Outras Situações — verificar manualmente",
        exige_base=False, exige_aliquota=False, exige_valor_icms=False,
        base_campo_sugerido="VL_ITEM", aliquota_padrao=None,
        formula="base * aliq / 100",
        permite_base_zero=True, permite_icms_zero=True, criticidade="aviso",
    ),
]

# ─────────────────────────────────────────────────────────────────────────────
# 4. ARQUIVO SPED DE DEMONSTRAÇÃO (com inconsistências intencionais)
# ─────────────────────────────────────────────────────────────────────────────

SPED_DEMO = """\
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
|C170|1|MOV001|PAINEL MDF 18MM|50,000|M2|5000,00|0,00|0|000|5102||5000,00|12,00|600,00|0,00|0,00|0,00|||||||||||||
|C170|2|MOV002|PORTA MDF 210X90|10,000|UN|5000,00|0,00|0|000|5102||5000,00|12,00|600,00|0,00|0,00|0,00|||||||||||||
|C190|000|5102|12,00|10000,00|10000,00|1200,00|0,00|0,00|0,00||
|C100|1|0|002|55|00|001|000002|43240112345678000199550010000002021000000002|05012024|07012024|8000,00|0|0,00|0,00|8000,00|0|0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|
|C170|1|MOV001|PAINEL MDF 18MM|40,000|M2|4000,00|0,00|0|040|5102||||0,00|0,00|0,00|||||||||||||
|C170|2|MOV002|PORTA MDF 210X90|8,000|UN|4000,00|0,00|0|040|5102||||0,00|0,00|0,00|||||||||||||
|C190|040|5102|0,00|8000,00|0,00|0,00|0,00|0,00|0,00||
|C100|0|0|001|55|00|001|000003|43240112345678000199550010000003031000000003|10012024|12012024|6000,00|0|0,00|0,00|6000,00|0|0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|
|C170|1|MOV001|PAINEL MDF 18MM|30,000|M2|3000,00|0,00|0|000|5102||||0,00|0,00|0,00|||||||||||||
|C170|2|MOV002|PORTA MDF 210X90|6,000|UN|3000,00|0,00|0|000|5102||||0,00|0,00|0,00|||||||||||||
|C190|000|5102|12,00|6000,00|0,00|0,00|0,00|0,00|0,00||
|C990|14|
|E001|0|
|E110|1200,00|0,00|1200,00|0,00|0,00|0,00|0,00|0,00|0,00|1200,00|0,00|1200,00|0,00|0,00|
|E990|3|
|9001|0|
|9900|0000|1|
|9900|C100|3|
|9900|C170|6|
|9900|C190|3|
|9990|5|
|9999|36|
"""


# ════════════════════════════════════════════════════════════════════════════════
# ████████████████████████  PARSER SPED  ████████████████████████████████████████
# ════════════════════════════════════════════════════════════════════════════════

def _decode_bytes(b: bytes) -> str:
    for enc in ("utf-8", "latin-1", "cp1252"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            continue
    raise ValueError("Encoding do arquivo SPED não reconhecido.")


def _parse_decimal_br(valor) -> Optional[float]:
    """Converte string BRL ('1.234,56') ou float para float Python. Retorna None se inválido."""
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        import math
        return None if (isinstance(valor, float) and math.isnan(valor)) else float(valor)
    s = str(valor).strip().replace(".", "").replace(",", ".")
    if not s:
        return None
    try:
        return float(Decimal(s))
    except (InvalidOperation, ValueError):
        return None


def _float_para_sped(valor: float) -> str:
    """float → string SPED com vírgula decimal e 2 casas."""
    d = Decimal(str(valor)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return str(d).replace(".", ",")


def _formatar_brl(valor: Optional[float]) -> str:
    """float → string formatada BRL para exibição."""
    if valor is None:
        return "—"
    try:
        d = Decimal(str(valor)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        inteiro, dec = str(d).split(".")
        fmt = ""
        for i, c in enumerate(reversed(inteiro)):
            if i and i % 3 == 0 and c != "-":
                fmt = "." + fmt
            fmt = c + fmt
        return f"R$ {fmt},{dec}"
    except Exception:
        return str(valor)


def _enriquecer(row: dict, linha: LinhaRegistro) -> None:
    """Adiciona colunas com nomes semânticos para registros mapeados."""
    mapa = MAPA_CAMPOS.get(linha.registro)
    if not mapa:
        return
    for nome, idx in mapa.items():
        row[nome] = linha.get(idx)


def parse_sped(conteudo_bytes: bytes) -> ResultadoParser:
    """
    Parser principal.
    Tokeniza cada linha por '|', identifica bloco/registro,
    monta DataFrame com colunas semânticas e detecta tipo de escrituração.
    """
    texto = _decode_bytes(conteudo_bytes)
    linhas_raw = texto.splitlines()
    erros: list[str] = []
    linhas: list[LinhaRegistro] = []

    for i, raw in enumerate(linhas_raw, 1):
        raw = raw.strip()
        if not raw:
            continue
        if not (raw.startswith("|") and raw.endswith("|")):
            erros.append(f"L{i}: formato inválido — {raw[:60]}")
            continue
        campos = raw[1:-1].split("|")
        if not campos or not campos[0].strip():
            erros.append(f"L{i}: tipo de registro vazio")
            continue
        registro = campos[0].strip().upper()
        bloco = registro[0] if registro else "?"
        linhas.append(LinhaRegistro(i, bloco, registro, campos, raw))

    # Metadados do 0000
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
    meta.blocos_presentes = sorted(set(l.bloco for l in linhas))

    # Tipo de escrituração
    blocos = set(meta.blocos_presentes)
    if blocos & {"E", "G", "K", "H"}:
        tipo = "EFD_ICMS_IPI"
    elif blocos & {"M", "P", "F"}:
        tipo = "EFD_CONTRIBUICOES"
    else:
        tipo = "EFD_ICMS_IPI"

    # DataFrame
    registros = []
    for l in linhas:
        r: dict = {
            "numero_linha": l.numero_linha,
            "bloco": l.bloco,
            "registro": l.registro,
            "linha_original": l.linha_original,
        }
        r.update({f"campo_{i:02d}": l.get(i) for i in range(min(len(l.campos), 50))})
        _enriquecer(r, l)
        registros.append(r)

    df = pd.DataFrame(registros)

    # Converter campos numéricos
    for col in CAMPOS_NUMERICOS:
        if col in df.columns:
            df[col] = df[col].apply(_parse_decimal_br)

    return ResultadoParser(meta, linhas, df, erros, tipo)


def reconstruir_sped(linhas: list[LinhaRegistro]) -> bytes:
    """Reconstrói arquivo SPED a partir das LinhaRegistro (possivelmente editadas)."""
    ordenadas = sorted(linhas, key=lambda l: l.numero_linha)
    return "".join(l.to_sped() for l in ordenadas).encode("utf-8")


def sincronizar_df_para_linhas(df: pd.DataFrame, linhas: list[LinhaRegistro]) -> list[LinhaRegistro]:
    """Propaga alterações do DataFrame de volta para LinhaRegistro."""
    mapa = {l.numero_linha: l for l in linhas}
    for _, row in df.iterrows():
        num = row.get("numero_linha")
        if num not in mapa:
            continue
        linha = mapa[num]
        campos_reg = MAPA_CAMPOS.get(linha.registro, {})
        for nome_campo, idx in campos_reg.items():
            if nome_campo not in row or pd.isna(row[nome_campo]):
                continue
            val = row[nome_campo]
            if isinstance(val, float):
                val = _float_para_sped(val)
            linha.set(idx, str(val))
    return list(mapa.values())


# ════════════════════════════════════════════════════════════════════════════════
# ████████████████████████  MOTOR DE REGRAS  ████████████████████████████════════
# ════════════════════════════════════════════════════════════════════════════════

ARQUIVO_REGRAS = "sped_regras_fiscais.json"


def carregar_regras() -> list[RegraFiscal]:
    if os.path.exists(ARQUIVO_REGRAS):
        try:
            with open(ARQUIVO_REGRAS, encoding="utf-8") as f:
                dados = json.load(f)
            return [RegraFiscal(**d) for d in dados]
        except Exception:
            pass
    return list(REGRAS_PADRAO)


def salvar_regras(regras: list[RegraFiscal]) -> None:
    with open(ARQUIVO_REGRAS, "w", encoding="utf-8") as f:
        json.dump([asdict(r) for r in regras], f, ensure_ascii=False, indent=2)


def buscar_regra(
    regras: list[RegraFiscal], cst: str, cfop: str, ind_oper: str = "*"
) -> Optional[RegraFiscal]:
    """
    Matching com prioridade: mais específico vence.
    Pontuação: CST exato +2, CFOP exato +2, operação exata +1.
    """
    candidatas = [
        r for r in regras
        if r.ativa and r.match_cst(cst) and r.match_cfop(cfop) and r.match_oper(ind_oper)
    ]
    if not candidatas:
        return None

    def score(r: RegraFiscal) -> int:
        s = 0
        if r.cst_icms != "*":
            s += 2
        if r.cfop != "*":
            s += 2
        if r.ind_oper != "*":
            s += 1
        return s

    return max(candidatas, key=score)


def calcular_imposto(regra: RegraFiscal, base: float, aliquota: float) -> float:
    """Aplica fórmula da regra com precisão Decimal."""
    if regra.formula == "zero":
        return 0.0
    resultado = Decimal(str(base)) * Decimal(str(aliquota)) / Decimal("100")
    return float(resultado.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


# ════════════════════════════════════════════════════════════════════════════════
# ████████████████████████  VALIDADOR FISCAL  ███████████████════════════════════
# ════════════════════════════════════════════════════════════════════════════════

def validar(df: pd.DataFrame, regras: list[RegraFiscal]) -> pd.DataFrame:
    """
    Executa todas as validações fiscais.
    Retorna DataFrame de inconsistências com colunas padronizadas.
    """
    inc: list[dict] = []

    # ------------------------------------------------------------------
    # A) Validação C170 — itens de NF-e
    # ------------------------------------------------------------------
    df_c170 = df[df["registro"] == "C170"].copy()
    for _, row in df_c170.iterrows():
        cst  = str(row.get("CST_ICMS", "") or "").strip()
        cfop = str(row.get("CFOP", "") or "").strip()
        nl   = int(row.get("numero_linha", 0))

        base    = _parse_decimal_br(row.get("VL_BC_ICMS"))
        aliq    = _parse_decimal_br(row.get("ALIQ_ICMS"))
        vl_icms = _parse_decimal_br(row.get("VL_ICMS"))
        vl_item = _parse_decimal_br(row.get("VL_ITEM"))

        if not cst:
            inc.append(_inc(nl, "C170", "C", "CAMPO_AUSENTE", "CRITICA",
                           "CST_ICMS", "CST_ICMS ausente no item", "", "", cst, cfop))
            continue

        regra = buscar_regra(regras, cst, cfop)
        sev   = regra.criticidade.upper() if regra else "AVISO"

        # --- base ausente em CST tributado ---
        if regra and regra.exige_base and (base is None or base == 0.0):
            sug = _float_para_sped(vl_item) if vl_item else ""
            inc.append(_inc(nl, "C170", "C", "CST_SEM_BASE", sev,
                           "VL_BC_ICMS",
                           f"CST {cst} exige base de cálculo — ausente/zerada",
                           str(row.get("VL_BC_ICMS", "")), sug, cst, cfop,
                           regra_id=regra.id if regra else ""))

        # --- alíquota ausente ---
        if regra and regra.exige_aliquota and (aliq is None or aliq == 0.0):
            sug = _float_para_sped(regra.aliquota_padrao) if regra.aliquota_padrao else ""
            inc.append(_inc(nl, "C170", "C", "CST_SEM_ALIQUOTA", sev,
                           "ALIQ_ICMS",
                           f"CST {cst} exige alíquota — ausente/zerada",
                           str(row.get("ALIQ_ICMS", "")), sug, cst, cfop,
                           regra_id=regra.id if regra else ""))

        # --- imposto ausente ---
        if regra and regra.exige_valor_icms and (vl_icms is None or vl_icms == 0.0):
            b_sug = base or vl_item or 0.0
            a_sug = aliq or (regra.aliquota_padrao or 0.0)
            imp_calc = calcular_imposto(regra, b_sug, a_sug) if regra else 0.0
            sug = _float_para_sped(imp_calc) if imp_calc else ""
            inc.append(_inc(nl, "C170", "C", "CST_SEM_IMPOSTO", sev,
                           "VL_ICMS",
                           f"CST {cst} exige VL_ICMS — ausente/zerado (sugerido: {sug})",
                           str(row.get("VL_ICMS", "")), sug, cst, cfop,
                           regra_id=regra.id if regra else ""))

        # --- imposto indevido em isento ---
        if regra and not regra.exige_valor_icms and vl_icms and vl_icms > 0:
            inc.append(_inc(nl, "C170", "C", "IMPOSTO_INDEVIDO", "AVISO",
                           "VL_ICMS",
                           f"CST {cst} não deve ter VL_ICMS — campo preenchido indevidamente",
                           _float_para_sped(vl_icms), "0", cst, cfop,
                           regra_id=regra.id if regra else ""))

        # --- divergência de cálculo ---
        if base and aliq and vl_icms is not None and regra:
            if regra.formula == "base * aliq / 100":
                esperado = calcular_imposto(regra, base, aliq)
                if abs(esperado - vl_icms) > 0.05:
                    inc.append(_inc(nl, "C170", "C", "DIVERGENCIA_CALCULO", "AVISO",
                                   "VL_ICMS",
                                   f"ICMS calculado ({_formatar_brl(esperado)}) ≠ registrado ({_formatar_brl(vl_icms)})",
                                   _float_para_sped(vl_icms), _float_para_sped(esperado),
                                   cst, cfop, regra_id=regra.id if regra else ""))

    # ------------------------------------------------------------------
    # B) Divergência de totais C100 vs soma C170
    # ------------------------------------------------------------------
    df_docs = df[df["registro"].isin(["C100", "C170"])].sort_values("numero_linha")
    c100_atual = None
    c100_nl    = 0
    acum_icms  = 0.0

    for _, row in df_docs.iterrows():
        if row["registro"] == "C100":
            if c100_atual is not None:
                vl_doc = _parse_decimal_br(c100_atual.get("VL_ICMS")) or 0.0
                if abs(acum_icms - vl_doc) > 0.10:
                    inc.append(_inc(c100_nl, "C100", "C", "DIVERGENCIA_TOTAL", "CRITICA",
                                   "VL_ICMS",
                                   f"Total C100 ({_formatar_brl(vl_doc)}) ≠ soma C170 ({_formatar_brl(acum_icms)})",
                                   _float_para_sped(vl_doc), _float_para_sped(acum_icms),
                                   num_doc=str(c100_atual.get("NUM_DOC", ""))))
            c100_atual = row
            c100_nl    = int(row["numero_linha"])
            acum_icms  = 0.0
        elif row["registro"] == "C170" and c100_atual is not None:
            acum_icms += _parse_decimal_br(row.get("VL_ICMS")) or 0.0

    # ------------------------------------------------------------------
    # C) C190 sem base para CST tributado
    # ------------------------------------------------------------------
    df_c190 = df[df["registro"] == "C190"]
    for _, row in df_c190.iterrows():
        cst  = str(row.get("CST_ICMS", "") or "").strip()
        cfop = str(row.get("CFOP", "") or "").strip()
        nl   = int(row.get("numero_linha", 0))
        regra = buscar_regra(regras, cst, cfop)
        if regra and regra.exige_base:
            bc = _parse_decimal_br(row.get("VL_BC_ICMS"))
            if bc is None or bc == 0:
                inc.append(_inc(nl, "C190", "C", "C190_SEM_BASE", "AVISO",
                               "VL_BC_ICMS",
                               f"C190 CST {cst}/CFOP {cfop} sem base de cálculo",
                               "", str(row.get("VL_OPR", "")), cst, cfop))

    # ------------------------------------------------------------------
    # D) Integridade de blocos (abertura e fechamento)
    # ------------------------------------------------------------------
    blocos_presentes = df["bloco"].unique()
    for bloco in blocos_presentes:
        if bloco in ("?",):
            continue
        for sufixo, tipo_err in [("001", "BLOCO_SEM_ABERTURA"), ("990", "BLOCO_SEM_FECHAMENTO")]:
            reg_esperado = f"{bloco}{sufixo}"
            if df[df["registro"] == reg_esperado].empty:
                inc.append(_inc(0, reg_esperado, bloco, tipo_err, "CRITICA",
                               "registro",
                               f"Bloco {bloco}: registro {reg_esperado} ausente",
                               "ausente", ""))

    # ------------------------------------------------------------------
    # E) Campos obrigatórios no 0000
    # ------------------------------------------------------------------
    df0 = df[df["registro"] == "0000"]
    if df0.empty:
        inc.append(_inc(0, "0000", "0", "CAMPO_AUSENTE", "CRITICA",
                       "0000", "Registro de abertura 0000 não encontrado", "ausente", ""))
    else:
        row0 = df0.iloc[0]
        for campo in ("CNPJ", "NOME", "DT_INI", "DT_FIN"):
            alias = {"DT_INI": "campo_03", "DT_FIN": "campo_04", "NOME": "NOME", "CNPJ": "CNPJ"}
            val = str(row0.get(campo, row0.get(alias.get(campo, ""), "")) or "").strip()
            if not val:
                inc.append(_inc(int(row0.get("numero_linha", 1)), "0000", "0",
                               "CAMPO_AUSENTE", "CRITICA", campo,
                               f"Campo {campo} obrigatório ausente no registro 0000", "", ""))

    # ------------------------------------------------------------------
    # F) Valores negativos inesperados
    # ------------------------------------------------------------------
    campos_nao_neg = ["VL_ITEM", "VL_BC_ICMS", "VL_ICMS", "ALIQ_ICMS"]
    for col in campos_nao_neg:
        if col not in df_c170.columns:
            continue
        neg = df_c170[df_c170[col].apply(lambda x: x is not None and isinstance(x, float) and x < 0)]
        for _, row in neg.iterrows():
            inc.append(_inc(int(row.get("numero_linha", 0)), "C170", "C",
                           "VALOR_NEGATIVO", "AVISO", col,
                           f"Valor negativo inesperado em {col}: {row[col]:.2f}",
                           _float_para_sped(row[col]), _float_para_sped(abs(row[col])),
                           str(row.get("CST_ICMS", "")), str(row.get("CFOP", ""))))

    return pd.DataFrame(inc) if inc else pd.DataFrame()


def _inc(
    numero_linha: int, registro: str, bloco: str,
    tipo: str, severidade: str, campo: str, descricao: str,
    valor_atual: str, valor_sugerido: str,
    cst: str = "", cfop: str = "", num_doc: str = "", regra_id: str = "",
) -> dict:
    return {
        "numero_linha": numero_linha, "bloco": bloco, "registro": registro,
        "tipo": tipo, "severidade": severidade, "campo_afetado": campo,
        "descricao": descricao, "valor_atual": valor_atual,
        "valor_sugerido": valor_sugerido, "cst": cst, "cfop": cfop,
        "num_doc": num_doc, "regra_aplicada": regra_id, "corrigido": False,
    }


# ════════════════════════════════════════════════════════════════════════════════
# ████████████████████████  EDITOR + AUDITORIA  █████════════════════════════════
# ════════════════════════════════════════════════════════════════════════════════

class TrilhaAuditoria:
    """Log imutável de todas as alterações da sessão."""

    def __init__(self):
        self._log: list[EntradaAuditoria] = []

    def registrar(
        self, numero_linha: int, registro: str, bloco: str,
        campo: str, valor_anterior: str, valor_novo: str,
        regra: str = "", motivo: str = "", tipo: str = "MANUAL",
        usuario: str = "analista",
    ) -> None:
        self._log.append(EntradaAuditoria(
            id=str(uuid.uuid4())[:8],
            timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            usuario=usuario, numero_linha=numero_linha,
            registro=registro, bloco=bloco, campo=campo,
            valor_anterior=valor_anterior, valor_novo=valor_novo,
            regra=regra, motivo=motivo, tipo=tipo,
        ))

    def to_df(self) -> pd.DataFrame:
        if not self._log:
            return pd.DataFrame()
        return pd.DataFrame([
            {"ID": e.id, "Data/Hora": e.timestamp, "Usuário": e.usuario,
             "Linha": e.numero_linha, "Registro": e.registro, "Bloco": e.bloco,
             "Campo": e.campo, "Valor Anterior": e.valor_anterior,
             "Valor Novo": e.valor_novo, "Regra": e.regra,
             "Motivo": e.motivo, "Tipo": e.tipo}
            for e in self._log
        ])

    def total(self) -> int:
        return len(self._log)


class EditorSped:
    """Gerencia o estado do DataFrame com suporte a desfazer e auditoria."""

    def __init__(self, df_original: pd.DataFrame):
        self._orig = df_original.copy()
        self._atual = df_original.copy()
        self._hist: list[pd.DataFrame] = []
        self.auditoria = TrilhaAuditoria()

    @property
    def df(self) -> pd.DataFrame:
        return self._atual

    @property
    def df_original(self) -> pd.DataFrame:
        return self._orig

    def get_alterados(self) -> pd.DataFrame:
        """Linhas que diferem do original."""
        try:
            mask = self._atual.ne(self._orig).any(axis=1)
            return self._atual[mask]
        except Exception:
            return pd.DataFrame()

    def editar(
        self, numero_linha: int, campo: str, valor_novo: str,
        motivo: str = "", regra: str = "", usuario: str = "analista",
    ) -> bool:
        idx = self._atual[self._atual["numero_linha"] == numero_linha].index
        if idx.empty:
            return False
        self._hist.append(self._atual.copy())
        i = idx[0]
        row = self._atual.loc[i]
        val_ant = str(row.get(campo, ""))
        # Converter para float se coluna numérica, senão manter como object
        if campo in CAMPOS_NUMERICOS and self._atual[campo].dtype == float:
            val_num = _parse_decimal_br(valor_novo)
            self._atual[campo] = self._atual[campo].astype(object)
            self._atual.at[i, campo] = val_num if val_num is not None else valor_novo
        else:
            if campo in self._atual.columns and self._atual[campo].dtype != object:
                self._atual[campo] = self._atual[campo].astype(object)
            self._atual.at[i, campo] = valor_novo
        self.auditoria.registrar(
            numero_linha, str(row.get("registro", "")), str(row.get("bloco", "")),
            campo, val_ant, valor_novo, regra, motivo, "MANUAL", usuario,
        )
        return True

    def massa(
        self, numeros_linha: list[int], campo: str, valor: str,
        motivo: str = "Correção em massa", regra: str = "", usuario: str = "analista",
    ) -> int:
        self._hist.append(self._atual.copy())
        n = 0
        for nl in numeros_linha:
            idx = self._atual[self._atual["numero_linha"] == nl].index
            if idx.empty:
                continue
            i = idx[0]
            row = self._atual.loc[i]
            ant = str(row.get(campo, ""))
            if ant == valor:
                continue
            if campo in CAMPOS_NUMERICOS and campo in self._atual.columns and self._atual[campo].dtype == float:
                val_num = _parse_decimal_br(valor)
                self._atual[campo] = self._atual[campo].astype(object)
                self._atual.at[i, campo] = val_num if val_num is not None else valor
            else:
                if campo in self._atual.columns and self._atual[campo].dtype != object:
                    self._atual[campo] = self._atual[campo].astype(object)
                self._atual.at[i, campo] = valor
            self.auditoria.registrar(
                nl, str(row.get("registro", "")), str(row.get("bloco", "")),
                campo, ant, valor, regra, motivo, "MASSA", usuario,
            )
            n += 1
        return n

    def recalcular_massa(
        self, numeros_linha: list[int], regras: list[RegraFiscal],
        motivo: str = "Recálculo automático", usuario: str = "analista",
    ) -> int:
        self._hist.append(self._atual.copy())
        n = 0
        for nl in numeros_linha:
            idx = self._atual[self._atual["numero_linha"] == nl].index
            if idx.empty:
                continue
            i = idx[0]
            row = self._atual.loc[i]
            cst  = str(row.get("CST_ICMS", "") or "").strip()
            cfop = str(row.get("CFOP", "") or "").strip()
            base = _parse_decimal_br(row.get("VL_BC_ICMS"))
            aliq = _parse_decimal_br(row.get("ALIQ_ICMS"))
            if base is None or aliq is None:
                continue
            regra = buscar_regra(regras, cst, cfop)
            if not regra:
                continue
            novo = calcular_imposto(regra, base, aliq)
            ant  = str(row.get("VL_ICMS", ""))
            self._atual.at[i, "VL_ICMS"] = novo
            self.auditoria.registrar(
                nl, str(row.get("registro", "")), str(row.get("bloco", "")),
                "VL_ICMS", ant, _float_para_sped(novo), regra.id, motivo, "AUTO", usuario,
            )
            n += 1
        return n

    def desfazer(self) -> bool:
        if not self._hist:
            return False
        self._atual = self._hist.pop()
        return True

    def restaurar(self, numeros_linha: Optional[list[int]] = None) -> None:
        self._hist.append(self._atual.copy())
        if numeros_linha:
            for nl in numeros_linha:
                ia = self._atual[self._atual["numero_linha"] == nl].index
                io = self._orig[self._orig["numero_linha"] == nl].index
                if not ia.empty and not io.empty:
                    self._atual.loc[ia[0]] = self._orig.loc[io[0]]
        else:
            self._atual = self._orig.copy()

    def preview(self, numeros_linha: list[int]) -> pd.DataFrame:
        rows = []
        for nl in numeros_linha:
            a = self._atual[self._atual["numero_linha"] == nl]
            o = self._orig[self._orig["numero_linha"] == nl]
            if a.empty or o.empty:
                continue
            ra, ro = a.iloc[0], o.iloc[0]
            for campo in ["VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS", "VL_ITEM",
                          "VL_BC_ICMS_ST", "ALIQ_ST", "VL_ICMS_ST"]:
                if ra.get(campo) != ro.get(campo):
                    rows.append({
                        "Linha": nl, "Registro": ra.get("registro", ""),
                        "Campo": campo, "Original": ro.get(campo),
                        "Atual": ra.get(campo),
                    })
        return pd.DataFrame(rows)


# ════════════════════════════════════════════════════════════════════════════════
# ████████████████████████  EXPORTAÇÕES  ████════════════════════════════════════
# ════════════════════════════════════════════════════════════════════════════════

# Paleta corporativa
_COR_HEADER  = "1A3A5C"
_COR_CRITICA = "C0392B"
_COR_AVISO   = "B7860D"
_COR_OK      = "1A7A35"
_COR_ZEBRA   = "EBF0F7"
_COR_BRANCO  = "FFFFFF"
_COR_SUB     = "2C5F8A"


def exportar_excel(
    df_inc: pd.DataFrame, df_alt: pd.DataFrame,
    df_aud: pd.DataFrame, df_reg: pd.DataFrame,
    meta: dict,
) -> bytes:
    buf = io.BytesIO()
    wb  = openpyxl.Workbook()
    wb.remove(wb.active)

    _aba_resumo(wb.create_sheet("Resumo Gerencial"), df_inc, df_alt, meta)
    _aba_dados(wb.create_sheet("Inconsistências"), df_inc,
               "Inconsistências Fiscais Detectadas", col_sev="severidade")
    _aba_dados(wb.create_sheet("Registros Alterados"), df_alt,
               "Registros Modificados na Sessão")
    _aba_dados(wb.create_sheet("Log de Auditoria"), df_aud,
               "Trilha de Auditoria — Alterações Realizadas")
    _aba_dados(wb.create_sheet("Regras Tributárias"), df_reg,
               "Catálogo de Regras Configuradas")

    wb.save(buf)
    buf.seek(0)
    return buf.read()


def _borda():
    s = Side(style="thin", color="CCCCCC")
    return Border(bottom=s, right=s)


def _aba_resumo(ws, df_inc, df_alt, meta):
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:F1")
    c = ws["A1"]
    c.value = "SPED AUDITOR — RELATÓRIO GERENCIAL"
    c.font  = Font(name="Calibri", bold=True, size=16, color=_COR_BRANCO)
    c.fill  = PatternFill("solid", fgColor=_COR_HEADER)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    campos_meta = [
        ("Empresa:", meta.get("nome_empresa", "—")),
        ("CNPJ:", meta.get("cnpj", "—")),
        ("Período:", meta.get("periodo_apuracao", "—")),
        ("UF:", meta.get("uf", "—")),
        ("Tipo:", meta.get("tipo_arquivo", "—")),
        ("Gerado em:", datetime.now().strftime("%d/%m/%Y %H:%M")),
    ]
    for i, (lbl, val) in enumerate(campos_meta, 3):
        ws.cell(i, 1, lbl).font = Font(bold=True, size=11)
        ws.cell(i, 2, val).font = Font(size=11)

    ws.cell(11, 1, "INDICADORES").font = Font(bold=True, size=12, color=_COR_HEADER)

    n_total  = len(df_inc) if not df_inc.empty else 0
    n_crit   = int((df_inc["severidade"] == "CRITICA").sum()) if not df_inc.empty and "severidade" in df_inc.columns else 0
    n_aviso  = int((df_inc["severidade"] == "AVISO").sum()) if not df_inc.empty and "severidade" in df_inc.columns else 0
    n_sem_b  = int((df_inc["tipo"] == "CST_SEM_BASE").sum()) if not df_inc.empty and "tipo" in df_inc.columns else 0
    n_sem_al = int((df_inc["tipo"] == "CST_SEM_ALIQUOTA").sum()) if not df_inc.empty and "tipo" in df_inc.columns else 0
    n_sem_im = int((df_inc["tipo"] == "CST_SEM_IMPOSTO").sum()) if not df_inc.empty and "tipo" in df_inc.columns else 0
    n_alt    = len(df_alt) if not df_alt.empty else 0

    indicadores = [
        ("Total de Inconsistências", n_total, _COR_HEADER),
        ("Críticas", n_crit, _COR_CRITICA),
        ("Avisos", n_aviso, _COR_AVISO),
        ("Itens sem Base de Cálculo", n_sem_b, _COR_CRITICA),
        ("Itens sem Alíquota", n_sem_al, _COR_CRITICA),
        ("Itens sem Valor ICMS", n_sem_im, _COR_CRITICA),
        ("Registros Alterados", n_alt, _COR_OK),
    ]
    for j, (lbl, val, cor) in enumerate(indicadores, 13):
        ws.cell(j, 1, lbl).font = Font(size=11, bold=True)
        c2 = ws.cell(j, 2, val)
        c2.font  = Font(size=11, bold=True, color=_COR_BRANCO)
        c2.fill  = PatternFill("solid", fgColor=cor)
        c2.alignment = Alignment(horizontal="center")

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 20


def _aba_dados(ws, df, titulo: str, col_sev: str = ""):
    ws.sheet_view.showGridLines = False
    if df is None or df.empty:
        ws["A1"] = f"{titulo} — sem dados"
        ws["A1"].font = Font(bold=True, italic=True, color="888888")
        return

    n = len(df.columns)
    ul = get_column_letter(n)
    ws.merge_cells(f"A1:{ul}1")
    c = ws["A1"]
    c.value = titulo
    c.font  = Font(bold=True, size=13, color=_COR_BRANCO)
    c.fill  = PatternFill("solid", fgColor=_COR_HEADER)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    bd = _borda()
    for ci, col in enumerate(df.columns, 1):
        h = ws.cell(2, ci, str(col).upper())
        h.font  = Font(bold=True, size=10, color=_COR_BRANCO)
        h.fill  = PatternFill("solid", fgColor=_COR_SUB)
        h.alignment = Alignment(horizontal="center", vertical="center")
        h.border = bd
    ws.row_dimensions[2].height = 22

    mapa_cor = {"CRITICA": "FCE4E4", "AVISO": "FFF8DC", "INFO": "EAF4FF"}
    for ri, (_, row) in enumerate(df.iterrows(), 3):
        zebra = _COR_ZEBRA if ri % 2 == 0 else _COR_BRANCO
        sev   = str(row.get(col_sev, "")).upper() if col_sev else ""
        cor   = mapa_cor.get(sev, zebra)
        for ci, col in enumerate(df.columns, 1):
            cell = ws.cell(ri, ci, row[col])
            cell.font   = Font(size=10)
            cell.fill   = PatternFill("solid", fgColor=cor)
            cell.border = bd
            cell.alignment = Alignment(vertical="center")

    for ci, col in enumerate(df.columns, 1):
        max_w = max(len(str(col)), df[col].astype(str).str.len().max() if len(df) > 0 else 10)
        ws.column_dimensions[get_column_letter(ci)].width = min(max_w + 4, 42)

    ws.freeze_panes = "A3"


def exportar_csv(df: pd.DataFrame) -> bytes:
    buf = io.StringIO()
    df.to_csv(buf, index=False, sep=";", encoding="utf-8-sig")
    return buf.getvalue().encode("utf-8-sig")


# ════════════════════════════════════════════════════════════════════════════════
# ████████████████████████  UTILITÁRIOS DE UI  ██════════════════════════════════
# ════════════════════════════════════════════════════════════════════════════════

def _cols_relevantes(df: pd.DataFrame) -> list[str]:
    base = ["numero_linha", "bloco", "registro"]
    extras = [
        c for c in df.columns
        if not c.startswith("campo_") and c not in base + ["linha_original"]
        and df[c].astype(str).str.strip().ne("").any()
    ]
    return base + extras[:28]


def _card(col, label: str, valor, classe: str = "") -> None:
    cores = {"critico": "#C0392B", "aviso": "#B7860D", "ok": "#1A7A35", "": "#0F2540"}
    cor = cores.get(classe, "#0F2540")
    col.markdown(
        f"""
        <div style="background:#fff;border:1px solid #E0E8F4;border-radius:8px;
                    padding:14px 16px;text-align:center;
                    box-shadow:0 2px 6px rgba(26,58,92,.07);">
          <div style="font-size:.7rem;color:#5A7398;font-weight:700;
                      text-transform:uppercase;letter-spacing:.04em;">{label}</div>
          <div style="font-size:1.7rem;font-weight:800;color:{cor};margin-top:4px;">{valor}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _metricas(df: pd.DataFrame) -> dict:
    df_c170 = df[df["registro"] == "C170"] if "registro" in df.columns else pd.DataFrame()

    def sem(col):
        if col not in df_c170.columns:
            return 0
        return int(df_c170[col].apply(
            lambda x: x is None or (isinstance(x, float) and (x == 0.0 or x != x))
        ).sum())

    return {
        "total_linhas": len(df),
        "total_c100": int((df["registro"] == "C100").sum()) if "registro" in df.columns else 0,
        "total_c170": int((df["registro"] == "C170").sum()) if "registro" in df.columns else 0,
        "c170_sem_base": sem("VL_BC_ICMS"),
        "c170_sem_aliq": sem("ALIQ_ICMS"),
        "c170_sem_icms": sem("VL_ICMS"),
    }


# ════════════════════════════════════════════════════════════════════════════════
# ████████████████████████  CSS GLOBAL  █████════════════════════════════════════
# ════════════════════════════════════════════════════════════════════════════════

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family:'Inter',sans-serif; }

/* ── Sidebar ── */
section[data-testid="stSidebar"] {
  background:#0F2540 !important; color:#fff;
}
section[data-testid="stSidebar"] .stRadio label {
  color:#BDD6F5 !important; font-size:.88rem; padding:3px 0;
}
section[data-testid="stSidebar"] .stRadio label:hover { color:#fff !important; }
section[data-testid="stSidebar"] h3 { color:#7BAFD4; }

/* ── Título de seção ── */
.sec {
  font-size:1.05rem; font-weight:700; color:#0F2540;
  border-left:4px solid #1A6BAD; padding-left:10px;
  margin:18px 0 12px 0;
}

/* ── Badges ── */
.badge-critica { background:#FCE4E4;color:#C0392B;padding:2px 9px;
                 border-radius:12px;font-size:.73rem;font-weight:700; }
.badge-aviso   { background:#FFF8DC;color:#B7860D;padding:2px 9px;
                 border-radius:12px;font-size:.73rem;font-weight:700; }
.badge-info    { background:#EAF4FF;color:#1A5C9E;padding:2px 9px;
                 border-radius:12px;font-size:.73rem;font-weight:700; }

/* ── Botões ── */
.stButton > button { border-radius:6px; font-weight:600; font-size:.87rem; }
hr { border-color:#E0E8F4; }
.stAlert { border-radius:6px; }
div[data-testid="metric-container"] { background:#F7FAFF;
  border:1px solid #E0E8F4; border-radius:8px; padding:10px 14px; }
</style>
"""

# ════════════════════════════════════════════════════════════════════════════════
# ████████████████████████  INICIALIZAÇÃO  ██════════════════════════════════════
# ════════════════════════════════════════════════════════════════════════════════

def _init():
    for k, v in {
        "resultado": None,
        "linhas_orig": None,
        "editor": None,
        "regras": None,
        "df_inc": None,
    }.items():
        if k not in st.session_state:
            st.session_state[k] = v

    if st.session_state["regras"] is None:
        st.session_state["regras"] = carregar_regras()


# ════════════════════════════════════════════════════════════════════════════════
# ████████████████████████  TELAS DA APLICAÇÃO  █════════════════════════════════
# ════════════════════════════════════════════════════════════════════════════════

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────

PAGINAS = [
    "📊 Dashboard",
    "📂 Upload",
    "🗂️ Blocos",
    "📋 Registros",
    "🧾 Notas Fiscais",
    "📦 Itens (C170)",
    "⚠️ Inconsistências",
    "🔧 Correções em Massa",
    "✏️ Editor Manual",
    "📤 Exportação",
    "📜 Log de Auditoria",
    "⚙️ Motor de Regras",
]


def sidebar() -> str:
    with st.sidebar:
        st.markdown(
            """
            <div style="padding:10px 0 20px;text-align:center;">
              <span style="font-size:1.5rem;font-weight:800;color:#fff;letter-spacing:.02em;">
                🔍 SPED Auditor
              </span><br>
              <span style="font-size:.7rem;color:#7BAFD4;">Fiscal Intelligence Platform v1.0</span>
            </div>
            """, unsafe_allow_html=True,
        )

        res = st.session_state.get("resultado")
        if res:
            m = res.metadados
            st.markdown(
                f"""
                <div style="background:#1A3A5C;border-radius:6px;padding:9px 11px;margin-bottom:14px;">
                  <div style="font-size:.65rem;color:#7BAFD4;font-weight:700;">ARQUIVO CARREGADO</div>
                  <div style="font-size:.8rem;color:#fff;margin-top:3px;">{(m.nome_empresa or '—')[:26]}</div>
                  <div style="font-size:.68rem;color:#BDD6F5;">CNPJ: {m.cnpj or '—'}</div>
                  <div style="font-size:.68rem;color:#BDD6F5;">{m.periodo_apuracao or '—'}</div>
                  <div style="font-size:.68rem;color:#BDD6F5;">Blocos: {', '.join(m.blocos_presentes)}</div>
                </div>
                """, unsafe_allow_html=True,
            )

        st.markdown("**Navegação**")
        pag = st.radio("", PAGINAS, label_visibility="collapsed")

        ed = st.session_state.get("editor")
        if ed and ed.auditoria.total():
            st.markdown(
                f'<div style="color:#7BAFD4;font-size:.72rem;margin-top:10px;">'
                f'✏️ {ed.auditoria.total()} alteração(ões) na sessão</div>',
                unsafe_allow_html=True,
            )
    return pag


# ─────────────────────────────────────────────────────────────────────────────
# UPLOAD
# ─────────────────────────────────────────────────────────────────────────────

def pg_upload():
    st.markdown('<div class="sec">Upload do Arquivo SPED</div>', unsafe_allow_html=True)

    col1, col2 = st.columns([3, 2])
    with col1:
        arq = st.file_uploader("Selecione o arquivo SPED (.txt)", type=["txt"])
        demo = st.checkbox("📁 Usar arquivo de demonstração (inclui inconsistências intencionais)",
                           value=not bool(arq))
        usuario = st.text_input("Usuário da sessão:", value="analista")

        if st.button("▶ Processar Arquivo", type="primary", use_container_width=True):
            _processar(arq, demo, usuario)

    with col2:
        st.markdown("**Suporte:**")
        st.markdown("- ✅ EFD ICMS/IPI  \n- ✅ EFD Contribuições  \n- 🔄 ECD/ECF *(arquitetura pronta)*")
        st.markdown("**Encoding:** UTF-8, Latin-1, CP1252")

        res = st.session_state.get("resultado")
        if res:
            m = res.metadados
            st.success(
                f"✅ **{m.nome_empresa}**\n\n"
                f"CNPJ: `{m.cnpj}` | UF: {m.uf}\n\n"
                f"Período: {m.periodo_apuracao}\n\n"
                f"Tipo: {res.tipo_arquivo.replace('_', ' ')}\n\n"
                f"Blocos: `{', '.join(m.blocos_presentes)}`\n\n"
                f"Linhas: `{m.total_linhas:,}`"
            )


def _processar(arq, demo: bool, usuario: str):
    with st.spinner("Analisando arquivo SPED…"):
        try:
            conteudo = SPED_DEMO.encode("utf-8") if (demo or not arq) else arq.read()
            if demo or not arq:
                st.info("Arquivo de demonstração carregado.")

            res = parse_sped(conteudo)
            st.session_state["resultado"]  = res
            st.session_state["linhas_orig"] = list(res.linhas)
            st.session_state["editor"]     = EditorSped(res.df)
            st.session_state["_usuario"]   = usuario

            regras = st.session_state["regras"] or carregar_regras()
            df_inc = validar(res.df, regras)
            st.session_state["df_inc"] = df_inc

            st.success(
                f"✅ {res.metadados.total_linhas:,} linhas processadas. "
                f"{len(df_inc)} inconsistência(s) detectada(s)."
            )
            if res.erros:
                with st.expander(f"⚠️ {len(res.erros)} aviso(s) do parser"):
                    for e in res.erros[:20]:
                        st.text(e)
        except Exception as e:
            st.error(f"Erro ao processar: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# DASHBOARD
# ─────────────────────────────────────────────────────────────────────────────

def pg_dashboard():
    st.markdown('<div class="sec">Dashboard — Visão Geral</div>', unsafe_allow_html=True)

    res = st.session_state.get("resultado")
    ed  = st.session_state.get("editor")
    df_inc = st.session_state.get("df_inc", pd.DataFrame())

    if not res:
        st.info("Nenhum arquivo carregado. Acesse **Upload** para começar."); return

    df  = ed.df if ed else res.df
    m   = res.metadados
    met = _metricas(df)

    # Metadados
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Empresa", (m.nome_empresa or "—")[:20])
    c2.metric("CNPJ", m.cnpj or "—")
    c3.metric("Período", m.periodo_apuracao or "—")
    c4.metric("Tipo", res.tipo_arquivo.replace("_", " "))

    st.markdown("---")

    # Cards
    cols = st.columns(7)
    _card(cols[0], "Total Linhas",   f"{met['total_linhas']:,}")
    _card(cols[1], "NF-e (C100)",    f"{met['total_c100']:,}")
    _card(cols[2], "Itens (C170)",   f"{met['total_c170']:,}")
    n_inc  = len(df_inc) if not df_inc.empty else 0
    n_crit = int((df_inc["severidade"] == "CRITICA").sum()) if not df_inc.empty and "severidade" in df_inc.columns else 0
    _card(cols[3], "Inconsistências", f"{n_inc:,}", "critico" if n_inc else "ok")
    _card(cols[4], "Críticas",        f"{n_crit:,}", "critico" if n_crit else "ok")
    _card(cols[5], "Sem Base",        f"{met['c170_sem_base']:,}", "critico" if met['c170_sem_base'] else "ok")
    _card(cols[6], "Sem ICMS",        f"{met['c170_sem_icms']:,}", "critico" if met['c170_sem_icms'] else "ok")

    st.markdown("<br>", unsafe_allow_html=True)

    # Gráficos
    g1, g2 = st.columns(2)

    with g1:
        st.markdown('<div class="sec">Registros por Bloco</div>', unsafe_allow_html=True)
        df_b = df["bloco"].value_counts().reset_index()
        df_b.columns = ["Bloco", "Qtd"]
        fig = px.bar(df_b, x="Bloco", y="Qtd", color="Bloco",
                     color_discrete_sequence=px.colors.qualitative.Set2,
                     template="plotly_white")
        fig.update_layout(showlegend=False, height=260,
                          margin=dict(l=5, r=5, t=5, b=5))
        st.plotly_chart(fig, use_container_width=True)

    with g2:
        st.markdown('<div class="sec">Itens por CST ICMS (C170)</div>', unsafe_allow_html=True)
        df_c170 = df[df["registro"] == "C170"]
        if not df_c170.empty and "CST_ICMS" in df_c170.columns:
            df_cst = df_c170["CST_ICMS"].value_counts().reset_index()
            df_cst.columns = ["CST", "Qtd"]
            fig2 = px.pie(df_cst, names="CST", values="Qtd", hole=0.4,
                          color_discrete_sequence=px.colors.qualitative.Pastel,
                          template="plotly_white")
            fig2.update_layout(height=260, margin=dict(l=5, r=5, t=5, b=5))
            st.plotly_chart(fig2, use_container_width=True)

    g3, g4 = st.columns(2)

    with g3:
        st.markdown('<div class="sec">Top 10 CFOPs (C170)</div>', unsafe_allow_html=True)
        if not df_c170.empty and "CFOP" in df_c170.columns:
            df_cf = df_c170["CFOP"].value_counts().head(10).reset_index()
            df_cf.columns = ["CFOP", "Qtd"]
            fig3 = px.bar(df_cf, x="CFOP", y="Qtd", template="plotly_white",
                          color_discrete_sequence=["#1A6BAD"])
            fig3.update_layout(height=250, margin=dict(l=5, r=5, t=5, b=5))
            st.plotly_chart(fig3, use_container_width=True)

    with g4:
        st.markdown('<div class="sec">Inconsistências por Tipo</div>', unsafe_allow_html=True)
        if not df_inc.empty and "tipo" in df_inc.columns:
            df_tp = df_inc["tipo"].value_counts().reset_index()
            df_tp.columns = ["Tipo", "Qtd"]
            fig4 = px.bar(df_tp, x="Qtd", y="Tipo", orientation="h",
                          template="plotly_white",
                          color="Qtd",
                          color_continuous_scale=["#FFF0F0", "#C0392B"])
            fig4.update_layout(height=250, margin=dict(l=5, r=5, t=5, b=5),
                               coloraxis_showscale=False)
            st.plotly_chart(fig4, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# BLOCOS
# ─────────────────────────────────────────────────────────────────────────────

def pg_blocos():
    st.markdown('<div class="sec">Visão por Blocos</div>', unsafe_allow_html=True)
    ed = st.session_state.get("editor")
    if not ed:
        _no_file(); return

    df     = ed.df
    blocos = sorted(df["bloco"].unique())

    cols = st.columns(min(len(blocos), 6))
    for i, b in enumerate(blocos):
        qtd = int((df["bloco"] == b).sum())
        cols[i % 6].metric(f"Bloco {b}", f"{qtd:,}")

    st.markdown("---")
    b_sel = st.selectbox("Explorar bloco:", blocos)
    df_b  = df[df["bloco"] == b_sel].copy()
    regs  = sorted(df_b["registro"].unique())
    r_sel = st.multiselect("Filtrar registros:", regs, default=regs[:6])
    if r_sel:
        df_b = df_b[df_b["registro"].isin(r_sel)]

    st.dataframe(df_b[_cols_relevantes(df_b)].fillna("").head(500),
                 use_container_width=True, height=420)
    st.caption(f"{len(df_b):,} registros")


# ─────────────────────────────────────────────────────────────────────────────
# REGISTROS
# ─────────────────────────────────────────────────────────────────────────────

def pg_registros():
    st.markdown('<div class="sec">Visão por Registros</div>', unsafe_allow_html=True)
    ed = st.session_state.get("editor")
    if not ed:
        _no_file(); return

    df   = ed.df
    regs = sorted(df["registro"].unique())
    c1, c2 = st.columns([2, 3])
    with c1:
        reg = st.selectbox("Registro:", regs)
    with c2:
        busca = st.text_input("Buscar:", placeholder="CST, CFOP, CNPJ…")

    df_r = df[df["registro"] == reg].copy()
    if busca:
        mask = df_r.astype(str).apply(lambda c: c.str.contains(busca, case=False)).any(axis=1)
        df_r = df_r[mask]

    st.dataframe(df_r[_cols_relevantes(df_r)].fillna(""), use_container_width=True, height=430)
    st.caption(f"{len(df_r):,} registros do tipo {reg}")


# ─────────────────────────────────────────────────────────────────────────────
# NOTAS FISCAIS
# ─────────────────────────────────────────────────────────────────────────────

def pg_notas():
    st.markdown('<div class="sec">Notas Fiscais (C100)</div>', unsafe_allow_html=True)
    ed = st.session_state.get("editor")
    if not ed:
        _no_file(); return

    df_c100 = ed.df[ed.df["registro"] == "C100"].copy()
    if df_c100.empty:
        st.warning("Nenhum C100 encontrado."); return

    cols_nf = [c for c in ["numero_linha", "NUM_DOC", "DT_DOC", "COD_PART",
                            "VL_DOC", "VL_BC_ICMS", "VL_ICMS", "COD_SIT"] if c in df_c100.columns]

    c1, c2 = st.columns(2)
    with c1:
        parceiros = ["Todos"] + sorted(df_c100["COD_PART"].dropna().unique().tolist()) \
                    if "COD_PART" in df_c100.columns else ["Todos"]
        parc = st.selectbox("Parceiro:", parceiros)
    with c2:
        sit = ["Todos"] + sorted(df_c100["COD_SIT"].dropna().unique().tolist()) \
              if "COD_SIT" in df_c100.columns else ["Todos"]
        cod_sit = st.selectbox("Situação:", sit)

    df_f = df_c100.copy()
    if parc != "Todos" and "COD_PART" in df_f.columns:
        df_f = df_f[df_f["COD_PART"] == parc]
    if cod_sit != "Todos" and "COD_SIT" in df_f.columns:
        df_f = df_f[df_f["COD_SIT"] == cod_sit]

    st.dataframe(df_f[cols_nf].fillna(""), use_container_width=True, height=420)
    st.caption(f"{len(df_f):,} notas")

    # Totalizadores
    cns = [c for c in ["VL_DOC", "VL_BC_ICMS", "VL_ICMS"] if c in df_f.columns]
    if cns:
        ct = st.columns(len(cns))
        for i, c in enumerate(cns):
            tot = pd.to_numeric(df_f[c], errors="coerce").sum()
            ct[i].metric(c, _formatar_brl(tot))


# ─────────────────────────────────────────────────────────────────────────────
# ITENS C170
# ─────────────────────────────────────────────────────────────────────────────

def pg_itens():
    st.markdown('<div class="sec">Itens de NF-e (C170)</div>', unsafe_allow_html=True)
    ed = st.session_state.get("editor")
    if not ed:
        _no_file(); return

    df_c170 = ed.df[ed.df["registro"] == "C170"].copy()
    if df_c170.empty:
        st.warning("Nenhum C170 encontrado."); return

    c1, c2, c3 = st.columns(3)
    with c1:
        csts = ["Todos"] + sorted(df_c170["CST_ICMS"].dropna().unique().tolist()) \
               if "CST_ICMS" in df_c170.columns else ["Todos"]
        cst_f = st.selectbox("CST ICMS:", csts)
    with c2:
        cfops = ["Todos"] + sorted(df_c170["CFOP"].dropna().unique().tolist()) \
                if "CFOP" in df_c170.columns else ["Todos"]
        cfop_f = st.selectbox("CFOP:", cfops)
    with c3:
        ap = st.checkbox("Apenas com inconsistências", value=False)

    df_f = df_c170.copy()
    if cst_f != "Todos":
        df_f = df_f[df_f["CST_ICMS"] == cst_f]
    if cfop_f != "Todos":
        df_f = df_f[df_f["CFOP"] == cfop_f]
    if ap:
        df_inc = st.session_state.get("df_inc", pd.DataFrame())
        if not df_inc.empty:
            linhas_inc = set(df_inc["numero_linha"].tolist())
            df_f = df_f[df_f["numero_linha"].isin(linhas_inc)]

    cols_item = [c for c in ["numero_linha", "COD_ITEM", "DESCR_COMPL", "QTD", "UNID",
                              "VL_ITEM", "CST_ICMS", "CFOP", "VL_BC_ICMS",
                              "ALIQ_ICMS", "VL_ICMS", "CST_PIS", "VL_PIS",
                              "CST_COFINS", "VL_COFINS"] if c in df_f.columns]
    st.dataframe(df_f[cols_item].fillna(""), use_container_width=True, height=420)
    st.caption(f"{len(df_f):,} itens")


# ─────────────────────────────────────────────────────────────────────────────
# INCONSISTÊNCIAS
# ─────────────────────────────────────────────────────────────────────────────

def pg_inconsistencias():
    st.markdown('<div class="sec">Inconsistências Fiscais</div>', unsafe_allow_html=True)

    df_inc = st.session_state.get("df_inc", pd.DataFrame())
    res    = st.session_state.get("resultado")

    if res and (df_inc is None or df_inc.empty):
        st.success("✅ Nenhuma inconsistência detectada."); return
    if not res:
        _no_file(); return

    # Métricas
    c1, c2, c3, c4 = st.columns(4)
    n_crit = int((df_inc["severidade"] == "CRITICA").sum()) if "severidade" in df_inc.columns else 0
    n_avi  = int((df_inc["severidade"] == "AVISO").sum()) if "severidade" in df_inc.columns else 0
    n_corr = int(df_inc["corrigido"].sum()) if "corrigido" in df_inc.columns else 0
    _card(c1, "Total", len(df_inc))
    _card(c2, "Críticas", n_crit, "critico")
    _card(c3, "Avisos", n_avi, "aviso")
    _card(c4, "Corrigidos", n_corr, "ok")

    st.markdown("---")

    # Filtros
    c1, c2, c3, c4 = st.columns(4)
    sev_ops  = ["Todos"] + sorted(df_inc["severidade"].unique().tolist()) if "severidade" in df_inc.columns else ["Todos"]
    tipo_ops = ["Todos"] + sorted(df_inc["tipo"].unique().tolist()) if "tipo" in df_inc.columns else ["Todos"]
    cst_ops  = ["Todos"] + sorted(df_inc["cst"].dropna().unique().tolist()) if "cst" in df_inc.columns else ["Todos"]

    with c1: sev_f  = st.selectbox("Severidade:", sev_ops)
    with c2: tipo_f = st.selectbox("Tipo:", tipo_ops)
    with c3: cst_f  = st.selectbox("CST:", cst_ops)
    with c4: busca  = st.text_input("Buscar na descrição:")

    df_f = df_inc.copy()
    if sev_f  != "Todos": df_f = df_f[df_f["severidade"] == sev_f]
    if tipo_f != "Todos": df_f = df_f[df_f["tipo"] == tipo_f]
    if cst_f  != "Todos": df_f = df_f[df_f["cst"] == cst_f]
    if busca: df_f = df_f[df_f["descricao"].astype(str).str.contains(busca, case=False)]

    def colorir(row):
        if row.get("severidade") == "CRITICA":
            return ["background-color:#FFF0F0"] * len(row)
        if row.get("severidade") == "AVISO":
            return ["background-color:#FFFDE8"] * len(row)
        return [""] * len(row)

    cols_show = [c for c in df_f.columns if c not in ["linha_original"]]
    st.dataframe(df_f[cols_show].style.apply(colorir, axis=1),
                 use_container_width=True, height=450)
    st.caption(f"{len(df_f):,} inconsistências")

    if st.button("🔄 Revalidar Arquivo"):
        ed = st.session_state.get("editor")
        regras = st.session_state.get("regras")
        if ed and regras:
            st.session_state["df_inc"] = validar(ed.df, regras)
            st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# CORREÇÕES EM MASSA
# ─────────────────────────────────────────────────────────────────────────────

def pg_massa():
    st.markdown('<div class="sec">Correções em Massa</div>', unsafe_allow_html=True)
    ed    = st.session_state.get("editor")
    regras = st.session_state.get("regras", [])
    df_inc = st.session_state.get("df_inc", pd.DataFrame())
    usuario = st.session_state.get("_usuario", "analista")

    if not ed:
        _no_file(); return

    df_c170 = ed.df[ed.df["registro"] == "C170"].copy()
    if df_c170.empty:
        st.warning("Nenhum C170 disponível."); return

    # ── Filtros
    st.markdown("**1. Filtros de seleção de itens**")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        csts  = ["Todos"] + sorted(df_c170["CST_ICMS"].dropna().unique().tolist()) if "CST_ICMS" in df_c170.columns else ["Todos"]
        cst_f = st.selectbox("CST ICMS:", csts, key="m_cst")
    with c2:
        cfops  = ["Todos"] + sorted(df_c170["CFOP"].dropna().unique().tolist()) if "CFOP" in df_c170.columns else ["Todos"]
        cfop_f = st.selectbox("CFOP:", cfops, key="m_cfop")
    with c3:
        ti_ops = ["Todos"]
        if not df_inc.empty and "tipo" in df_inc.columns:
            ti_ops += sorted(df_inc["tipo"].unique().tolist())
        tipo_f = st.selectbox("Tipo inconsistência:", ti_ops, key="m_tipo")
    with c4:
        ap_inc = st.checkbox("Apenas com inconsistência", value=True, key="m_inc")

    df_alvo = df_c170.copy()
    if cst_f  != "Todos": df_alvo = df_alvo[df_alvo["CST_ICMS"] == cst_f]
    if cfop_f != "Todos": df_alvo = df_alvo[df_alvo["CFOP"] == cfop_f]
    if ap_inc and not df_inc.empty:
        linhas_t = set(
            df_inc[df_inc["tipo"] == tipo_f]["numero_linha"].tolist()
            if tipo_f != "Todos" else df_inc["numero_linha"].tolist()
        )
        df_alvo = df_alvo[df_alvo["numero_linha"].isin(linhas_t)]

    n_alvo = len(df_alvo)
    st.info(f"**{n_alvo} item(ns) selecionado(s)**")

    cols_vis = [c for c in ["numero_linha", "CST_ICMS", "CFOP", "VL_ITEM",
                             "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS"] if c in df_alvo.columns]
    st.dataframe(df_alvo[cols_vis].fillna("").head(200),
                 use_container_width=True, height=200)

    if df_alvo.empty:
        return

    nls = df_alvo["numero_linha"].tolist()

    st.markdown("---")
    st.markdown("**2. Ações em lote**")

    tab1, tab2, tab3, tab4 = st.tabs(
        ["📐 Preencher Base", "📊 Preencher Alíquota", "🧮 Recalcular Imposto", "👁️ Prévia"]
    )

    with tab1:
        fonte = st.radio("Origem da base:", ["VL_ITEM do próprio item", "Valor manual"], key="fonte_b")
        if "manual" in fonte:
            val_b = st.number_input("Valor (R$):", min_value=0.0, step=0.01, key="b_manual")
        motivo_b = st.text_input("Motivo:", "Base preenchida via VL_ITEM", key="mot_b")
        if st.button("▶ Aplicar Base", type="primary"):
            alt = 0
            for _, row in df_alvo.iterrows():
                nl = int(row["numero_linha"])
                if "manual" in fonte:
                    v = _float_para_sped(val_b)
                else:
                    vi = _parse_decimal_br(row.get("VL_ITEM"))
                    v  = _float_para_sped(vi) if vi else ""
                if v:
                    ok = ed.editar(nl, "VL_BC_ICMS", v, motivo_b, "", usuario)
                    if ok: alt += 1
            st.success(f"✅ {alt} base(s) preenchida(s).")

    with tab2:
        aliq_v  = st.number_input("Alíquota (%):", 0.0, 100.0, 12.0, 0.5, key="aliq_massa")
        motivo_a = st.text_input("Motivo:", "Alíquota padrão aplicada", key="mot_a")
        if st.button("▶ Aplicar Alíquota", type="primary"):
            alt = ed.massa(nls, "ALIQ_ICMS", _float_para_sped(aliq_v), motivo_a, "", usuario)
            st.success(f"✅ {alt} alíquota(s) preenchida(s).")

    with tab3:
        motivo_c = st.text_input("Motivo:", "Recálculo automático pelo motor de regras", key="mot_c")
        st.caption("Fórmula: VL_ICMS = VL_BC_ICMS × ALIQ_ICMS ÷ 100 (conforme regra CST+CFOP)")
        if st.button("▶ Recalcular Imposto", type="primary"):
            alt = ed.recalcular_massa(nls, regras, motivo_c, usuario)
            st.success(f"✅ {alt} imposto(s) recalculado(s).")

    with tab4:
        df_prev = ed.preview(nls)
        if not df_prev.empty:
            st.dataframe(df_prev, use_container_width=True)
        else:
            st.info("Nenhuma alteração pendente.")

    st.markdown("---")
    if st.button("↩️ Desfazer última operação"):
        if ed.desfazer():
            st.success("✅ Desfeito."); st.rerun()
        else:
            st.warning("Nada para desfazer.")


# ─────────────────────────────────────────────────────────────────────────────
# EDITOR MANUAL
# ─────────────────────────────────────────────────────────────────────────────

def pg_editor():
    st.markdown('<div class="sec">Editor Manual de Registros</div>', unsafe_allow_html=True)
    ed = st.session_state.get("editor")
    usuario = st.session_state.get("_usuario", "analista")

    if not ed:
        _no_file(); return

    df   = ed.df
    regs = sorted(df["registro"].unique())

    c1, c2, c3 = st.columns([2, 2, 1])
    with c1:
        reg_sel = st.selectbox("Tipo de registro:", regs)
    df_reg = df[df["registro"] == reg_sel]
    with c2:
        nl_min = int(df_reg["numero_linha"].min()) if not df_reg.empty else 1
        nl_max = int(df_reg["numero_linha"].max()) if not df_reg.empty else 1
        nl_sel = st.number_input("Número da linha:", nl_min, nl_max, nl_min)
    with c3:
        motivo = st.text_input("Motivo:", key="ed_motivo")

    df_l = df[df["numero_linha"] == nl_sel]
    if df_l.empty:
        st.warning("Linha não encontrada."); return

    row_a = df_l.iloc[0]
    row_o_df = ed.df_original[ed.df_original["numero_linha"] == nl_sel]
    row_o = row_o_df.iloc[0] if not row_o_df.empty else row_a

    # Campos editáveis
    mapa = MAPA_CAMPOS.get(reg_sel, {})
    campos = list(mapa.keys()) if mapa else [
        c for c in df.columns
        if not c.startswith("campo_") and c not in ["numero_linha", "bloco", "registro", "linha_original"]
    ]

    # Campos alterados
    alt_campos = [c for c in campos if str(row_a.get(c, "")) != str(row_o.get(c, ""))]
    if alt_campos:
        st.warning(f"⚠️ Campo(s) alterado(s) nesta linha: **{', '.join(alt_campos)}**")

    st.markdown(f"**Editando: {reg_sel} — Linha {nl_sel}**")

    pendentes = {}
    cols3 = st.columns(3)
    for i, campo in enumerate(campos):
        val_a = str(row_a.get(campo, "") or "")
        val_o = str(row_o.get(campo, "") or "")
        dest  = "🔴 " if val_a != val_o else ""
        novo  = cols3[i % 3].text_input(
            f"{dest}{campo}", value=val_a, key=f"ed_{campo}_{nl_sel}"
        )
        if novo != val_a:
            pendentes[campo] = novo

    # Botões de ação
    ca, cb, cc = st.columns(3)
    with ca:
        if st.button("💾 Salvar", type="primary", disabled=not pendentes):
            for c, v in pendentes.items():
                ed.editar(nl_sel, c, v, motivo, "", usuario)
            st.success(f"✅ {len(pendentes)} campo(s) salvo(s)."); st.rerun()
    with cb:
        if st.button("↩️ Desfazer"):
            if ed.desfazer():
                st.success("✅ Desfeito."); st.rerun()
            else:
                st.warning("Nada para desfazer.")
    with cc:
        if st.button("🔄 Restaurar linha"):
            ed.restaurar([nl_sel])
            st.success("✅ Restaurado."); st.rerun()

    # Comparação
    st.markdown("---")
    st.markdown("**Comparação Original vs. Atual**")
    diffs = []
    for c in campos:
        va = str(row_a.get(c, "") or "")
        vo = str(row_o.get(c, "") or "")
        if va != vo:
            diffs.append({"Campo": c, "Original": vo, "Atual": va})
    if diffs:
        st.dataframe(pd.DataFrame(diffs), use_container_width=True)
    else:
        st.success("Nenhuma diferença nesta linha.")


# ─────────────────────────────────────────────────────────────────────────────
# EXPORTAÇÃO
# ─────────────────────────────────────────────────────────────────────────────

def pg_exportacao():
    st.markdown('<div class="sec">Exportação</div>', unsafe_allow_html=True)

    ed     = st.session_state.get("editor")
    res    = st.session_state.get("resultado")
    df_inc = st.session_state.get("df_inc", pd.DataFrame())
    regras = st.session_state.get("regras", [])

    if not ed or not res:
        _no_file(); return

    m = res.metadados
    meta_dict = {
        "nome_empresa": m.nome_empresa, "cnpj": m.cnpj,
        "periodo_apuracao": m.periodo_apuracao, "uf": m.uf,
        "tipo_arquivo": res.tipo_arquivo,
    }

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("#### 📄 SPED TXT Corrigido")
        st.markdown("Arquivo reconstruído com todas as correções da sessão.")
        if st.button("Gerar SPED TXT", type="primary"):
            linhas_orig = st.session_state.get("linhas_orig", res.linhas)
            linhas_corr = sincronizar_df_para_linhas(ed.df, linhas_orig)
            bts = reconstruir_sped(linhas_corr)
            nome = f"SPED_CORRIGIDO_{m.cnpj}_{m.periodo_apuracao.replace(' ', '').replace('/', '')[:12]}.txt"
            st.download_button("⬇️ Baixar TXT", bts, nome, "text/plain")

    with c2:
        st.markdown("#### 📊 Relatório Excel (5 abas)")
        st.markdown("Resumo · Inconsistências · Alterados · Auditoria · Regras")
        if st.button("Gerar Excel", type="primary"):
            df_alt = ed.get_alterados()
            df_aud = ed.auditoria.to_df()
            df_reg = pd.DataFrame([asdict(r) for r in regras])
            bts = exportar_excel(df_inc or pd.DataFrame(), df_alt, df_aud, df_reg, meta_dict)
            nome = f"AUDITORIA_SPED_{m.cnpj}.xlsx"
            st.download_button("⬇️ Baixar Excel", bts, nome,
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.markdown("---")
    c3, c4 = st.columns(2)

    with c3:
        st.markdown("#### 📋 CSV Inconsistências")
        if df_inc is not None and not df_inc.empty:
            csv = exportar_csv(df_inc)
            st.download_button("⬇️ Baixar CSV", csv, "inconsistencias_sped.csv", "text/csv")
        else:
            st.info("Sem inconsistências para exportar.")

    with c4:
        st.markdown("#### 📋 CSV Log de Auditoria")
        df_aud = ed.auditoria.to_df()
        if not df_aud.empty:
            csv_a = exportar_csv(df_aud)
            st.download_button("⬇️ Baixar CSV", csv_a, "log_auditoria.csv", "text/csv")
        else:
            st.info("Sem alterações registradas.")


# ─────────────────────────────────────────────────────────────────────────────
# LOG DE AUDITORIA
# ─────────────────────────────────────────────────────────────────────────────

def pg_auditoria():
    st.markdown('<div class="sec">Log de Auditoria — Trilha de Alterações</div>', unsafe_allow_html=True)
    ed = st.session_state.get("editor")
    if not ed:
        _no_file(); return

    df_aud = ed.auditoria.to_df()
    if df_aud.empty:
        st.info("Nenhuma alteração registrada nesta sessão."); return

    c1, c2, c3 = st.columns(3)
    _card(c1, "Total Alterações", ed.auditoria.total())
    tipos = df_aud["Tipo"].value_counts().to_dict() if "Tipo" in df_aud.columns else {}
    _card(c2, "Manuais", tipos.get("MANUAL", 0))
    _card(c3, "Em Massa / Auto", tipos.get("MASSA", 0) + tipos.get("AUTO", 0))

    st.markdown("---")

    # Filtros
    cc1, cc2 = st.columns(2)
    with cc1:
        t_ops = ["Todos"] + sorted(df_aud["Tipo"].unique().tolist()) if "Tipo" in df_aud.columns else ["Todos"]
        t_f   = st.selectbox("Tipo de operação:", t_ops)
    with cc2:
        c_ops = ["Todos"] + sorted(df_aud["Campo"].unique().tolist()) if "Campo" in df_aud.columns else ["Todos"]
        c_f   = st.selectbox("Campo alterado:", c_ops)

    df_f = df_aud.copy()
    if t_f != "Todos": df_f = df_f[df_f["Tipo"] == t_f]
    if c_f != "Todos": df_f = df_f[df_f["Campo"] == c_f]

    st.dataframe(df_f, use_container_width=True, height=480)


# ─────────────────────────────────────────────────────────────────────────────
# MOTOR DE REGRAS
# ─────────────────────────────────────────────────────────────────────────────

def pg_regras():
    st.markdown('<div class="sec">Motor de Regras Tributárias</div>', unsafe_allow_html=True)

    regras: list[RegraFiscal] = st.session_state.get("regras") or carregar_regras()
    st.session_state["regras"] = regras

    tab1, tab2 = st.tabs(["📋 Regras Ativas", "➕ Cadastrar / Editar"])

    with tab1:
        df_r = pd.DataFrame([asdict(r) for r in regras])
        st.dataframe(df_r, use_container_width=True, height=400)

        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 Resetar para regras padrão"):
                st.session_state["regras"] = list(REGRAS_PADRAO)
                salvar_regras(REGRAS_PADRAO)
                st.success("✅ Regras padrão restauradas."); st.rerun()
        with c2:
            id_del = st.text_input("ID da regra para remover:", key="id_del")
            if st.button("🗑️ Remover regra") and id_del:
                st.session_state["regras"] = [r for r in regras if r.id != id_del]
                salvar_regras(st.session_state["regras"])
                st.success(f"✅ Regra '{id_del}' removida."); st.rerun()

    with tab2:
        st.markdown("**Preencha os campos da nova regra:**")
        c1, c2, c3 = st.columns(3)
        with c1:
            n_id      = st.text_input("ID (ex: 000_5102):", key="n_id")
            n_cst     = st.text_input("CST ICMS (* = todos):", key="n_cst")
            n_cfop    = st.text_input("CFOP (* = todos):", key="n_cfop")
            n_oper    = st.selectbox("Operação:", ["*", "E", "S"], key="n_oper")
        with c2:
            n_desc    = st.text_area("Descrição:", key="n_desc")
            n_exb     = st.checkbox("Exige base", value=True, key="n_exb")
            n_exal    = st.checkbox("Exige alíquota", value=True, key="n_exal")
            n_eximp   = st.checkbox("Exige valor ICMS", value=True, key="n_eximp")
        with c3:
            n_alpad   = st.number_input("Alíquota padrão (%):", 0.0, 100.0, 12.0, key="n_alpad")
            n_bcampo  = st.selectbox("Campo p/ sugerir base:", ["VL_ITEM", "VL_DOC", ""], key="n_bcampo")
            n_formula = st.selectbox("Fórmula:", ["base * aliq / 100", "zero"], key="n_form")
            n_crit    = st.selectbox("Criticidade:", ["critica", "aviso", "info"], key="n_crit")

        if st.button("💾 Salvar Regra", type="primary"):
            if not n_id or not n_cst:
                st.error("ID e CST são obrigatórios.")
            else:
                nova = RegraFiscal(
                    id=n_id, cst_icms=n_cst, cfop=n_cfop, ind_oper=n_oper,
                    descricao=n_desc, exige_base=n_exb, exige_aliquota=n_exal,
                    exige_valor_icms=n_eximp, base_campo_sugerido=n_bcampo,
                    aliquota_padrao=n_alpad, formula=n_formula,
                    permite_base_zero=not n_exb, permite_icms_zero=not n_eximp,
                    criticidade=n_crit,
                )
                st.session_state["regras"] = [r for r in regras if r.id != n_id] + [nova]
                salvar_regras(st.session_state["regras"])
                st.success(f"✅ Regra '{n_id}' salva."); st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _no_file():
    st.info("Nenhum arquivo SPED carregado. Acesse **Upload** para iniciar.")


# ════════════════════════════════════════════════════════════════════════════════
# ████████████████████████  MAIN  ███════════════════════════════════════════════
# ════════════════════════════════════════════════════════════════════════════════

ROTEADOR = {
    "📊 Dashboard":          pg_dashboard,
    "📂 Upload":             pg_upload,
    "🗂️ Blocos":             pg_blocos,
    "📋 Registros":          pg_registros,
    "🧾 Notas Fiscais":      pg_notas,
    "📦 Itens (C170)":       pg_itens,
    "⚠️ Inconsistências":    pg_inconsistencias,
    "🔧 Correções em Massa": pg_massa,
    "✏️ Editor Manual":      pg_editor,
    "📤 Exportação":         pg_exportacao,
    "📜 Log de Auditoria":   pg_auditoria,
    "⚙️ Motor de Regras":    pg_regras,
}


def main():
    _init()
    st.markdown(CSS, unsafe_allow_html=True)
    pagina = sidebar()
    fn = ROTEADOR.get(pagina)
    if fn:
        fn()
    else:
        st.error(f"Página não encontrada: {pagina}")


if __name__ == "__main__":
    main()
