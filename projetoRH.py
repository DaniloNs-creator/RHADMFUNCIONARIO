"""
SPED FISCAL ANALYTICS - SISTEMA COMPLETO
Versão: 1.0.0
Arquitetura: Streamlit + Pandas + Regras Fiscais

Este sistema permite:
- Upload e leitura de arquivos SPED (EFD ICMS/IPI, EFD Contribuições)
- Análise fiscal com detecção de inconsistências
- Correção em massa de itens com problemas (base, alíquota, imposto)
- Suporte a blocos D100, C500, F100
- Exportação do arquivo corrigido e relatórios
- Auditoria completa de todas as alterações

Autor: Arquiteto de Software Sênior
Data: 2024
"""

import streamlit as st
import pandas as pd
import numpy as np
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime, date
from typing import Dict, List, Any, Optional, Tuple
import re
import io
import json
import copy
from dataclasses import dataclass, field, asdict
import plotly.express as px
import plotly.graph_objects as go

# ============================================================================
# CONFIGURAÇÕES E CONSTANTES
# ============================================================================

# Constantes fiscais
CST_ICMS = {
    "00": "Tributada integralmente",
    "10": "Tributada e com cobrança do ICMS por substituição tributária",
    "20": "Com redução de base de cálculo",
    "30": "Isenta / não tributada e com cobrança do ICMS por substituição tributária",
    "40": "Isenta",
    "41": "Não tributada",
    "50": "Suspensão",
    "51": "Diferimento",
    "60": "ICMS cobrado anteriormente por substituição tributária",
    "70": "Com redução de base de cálculo e cobrança do ICMS por substituição tributária",
    "90": "Outras"
}

CFOP_DESCRIPTIONS = {
    "5101": "Venda de produção do estabelecimento",
    "5102": "Venda de mercadoria adquirida ou recebida de terceiros",
    "5103": "Venda de produção do estabelecimento - Remetente",
    "5105": "Venda de mercadoria adquirida ou recebida de terceiros - Remetente",
    "5106": "Venda de produção do estabelecimento - Destinatário",
    "5109": "Venda de produção do estabelecimento - Exterior",
    "5110": "Venda de mercadoria adquirida ou recebida de terceiros - Exterior",
    "5351": "Aquisição de serviço de transporte",
    "5355": "Aquisição de serviço de locação",
    "5401": "Aquisição de energia elétrica",
    "5402": "Aquisição de energia elétrica - Exterior",
}

# ============================================================================
# MODELOS DE DADOS
# ============================================================================

@dataclass
class FiscalItem:
    """Item fiscal"""
    numero_item: int = 1
    codigo_produto: str = ""
    descricao_produto: str = ""
    quantidade: Decimal = Decimal("0")
    valor_unitario: Decimal = Decimal("0")
    valor_total: Decimal = Decimal("0")
    base_calculo: Optional[Decimal] = None
    aliquota: Optional[Decimal] = None
    valor_imposto: Optional[Decimal] = None
    cst: str = ""
    cfop: str = ""
    tipo_operacao: str = "entrada"
    original_base: Optional[Decimal] = None
    original_aliquota: Optional[Decimal] = None
    original_imposto: Optional[Decimal] = None
    modified: bool = False
    record_id: str = ""
    block: str = "C"
    
    def to_dict(self) -> dict:
        """Converte para dicionário"""
        return {
            'numero_item': self.numero_item,
            'codigo_produto': self.codigo_produto,
            'descricao_produto': self.descricao_produto,
            'quantidade': float(self.quantidade) if self.quantidade else 0,
            'valor_unitario': float(self.valor_unitario) if self.valor_unitario else 0,
            'valor_total': float(self.valor_total) if self.valor_total else 0,
            'base_calculo': float(self.base_calculo) if self.base_calculo else None,
            'aliquota': float(self.aliquota) if self.aliquota else None,
            'valor_imposto': float(self.valor_imposto) if self.valor_imposto else None,
            'cst': self.cst,
            'cfop': self.cfop,
            'tipo_operacao': self.tipo_operacao,
            'modified': self.modified,
            'block': self.block
        }

@dataclass
class FiscalDocument:
    """Documento fiscal"""
    chave_acesso: str = ""
    cnpj_emitente: str = ""
    cnpj_destinatario: str = ""
    valor_total: Decimal = Decimal("0")
    valor_base_calculo: Optional[Decimal] = None
    valor_icms: Optional[Decimal] = None
    aliquota_icms: Optional[Decimal] = None
    cst: str = ""
    cfop: str = ""
    tipo_operacao: str = "entrada"
    data_emissao: date = date.today()
    situacao: str = "normal"
    block: str = "C"
    record_type: str = "C100"
    items: List[FiscalItem] = field(default_factory=list)

@dataclass
class SpedRecord:
    """Registro SPED"""
    record_id: str
    record_type: str
    block: str
    line_number: int
    raw_data: str
    fields: Dict[str, Any]
    parent_record: Optional[str] = None
    children_records: List[str] = field(default_factory=list)
    modified: bool = False
    original_data: Dict[str, Any] = field(default_factory=dict)

@dataclass
class FiscalRule:
    """Regra fiscal"""
    rule_id: str
    cst: str
    cfop: str
    operation_type: str  # 'entrada', 'saida', 'ambos'
    requires_base: bool = True
    requires_aliquota: bool = True
    requires_imposto: bool = True
    default_base_formula: Optional[str] = None
    default_aliquota: Optional[Decimal] = None
    default_imposto_formula: Optional[str] = None
    valid_cfops: List[str] = field(default_factory=list)
    priority: int = 1
    enabled: bool = True

# ============================================================================
# MOTOR DE REGRAS
# ============================================================================

class RulesEngine:
    """Motor de regras fiscais"""
    
    def __init__(self):
        self.rules: List[FiscalRule] = []
        self.rule_cache: Dict[str, List[FiscalRule]] = {}
        self._load_default_rules()
    
    def _load_default_rules(self):
        """Carrega regras padrão"""
        default_rules = [
            # CST 00 - Tributada integralmente
            FiscalRule("R001", "00", "*", "ambos", True, True, True, 
                      "item_value", Decimal("18.00"), "base * aliquota / 100"),
            
            # CST 10 - Tributada com ST
            FiscalRule("R002", "10", "*", "ambos", True, True, True,
                      "item_value", Decimal("18.00"), "base * aliquota / 100"),
            
            # CST 20 - Com redução de base
            FiscalRule("R003", "20", "*", "ambos", True, True, True,
                      "item_value", Decimal("12.00"), "base * aliquota / 100"),
            
            # CST 30 - Isenta com ST
            FiscalRule("R004", "30", "*", "ambos", False, False, False,
                      None, None, None),
            
            # CST 40 - Isenta
            FiscalRule("R005", "40", "*", "ambos", False, False, False,
                      None, None, None),
            
            # CST 41 - Não tributada
            FiscalRule("R006", "41", "*", "ambos", False, False, False,
                      None, None, None),
            
            # CST 50 - Suspensão
            FiscalRule("R007", "50", "*", "ambos", False, False, False,
                      None, None, None),
            
            # CST 51 - Diferimento
            FiscalRule("R008", "51", "*", "ambos", False, False, False,
                      None, None, None),
            
            # CST 60 - ICMS cobrado anteriormente
            FiscalRule("R009", "60", "*", "ambos", False, False, False,
                      None, None, None),
            
            # CST 70 - Com redução e ST
            FiscalRule("R010", "70", "*", "ambos", True, True, True,
                      "item_value", Decimal("12.00"), "base * aliquota / 100"),
            
            # CST 90 - Outras
            FiscalRule("R011", "90", "*", "ambos", False, False, False,
                      None, None, None),
            
            # Regras específicas para D100 (CT-e)
            FiscalRule("R012", "00", "5351", "entrada", True, True, True,
                      "item_value", Decimal("12.00"), "base * aliquota / 100"),
            
            # Regras específicas para C500 (Energia)
            FiscalRule("R013", "00", "5401", "entrada", True, True, True,
                      "item_value", Decimal("25.00"), "base * aliquota / 100"),
            
            # Regras específicas para F100 (Aluguel)
            FiscalRule("R014", "00", "5355", "entrada", True, True, True,
                      "item_value", Decimal("18.00"), "base * aliquota / 100"),
        ]
        
        self.rules = default_rules
        self._build_cache()
    
    def _build_cache(self):
        """Constrói cache de regras"""
        self.rule_cache = {}
        for rule in self.rules:
            if rule.enabled:
                key = f"{rule.cst}|{rule.cfop}|{rule.operation_type}"
                if key not in self.rule_cache:
                    self.rule_cache[key] = []
                self.rule_cache[key].append(rule)
    
    def get_rule(self, cst: str, cfop: str, operation_type: str) -> Optional[FiscalRule]:
        """Obtém regra para CST, CFOP e tipo de operação"""
        # Busca exata
        key_exact = f"{cst}|{cfop}|{operation_type}"
        if key_exact in self.rule_cache and self.rule_cache[key_exact]:
            return self.rule_cache[key_exact][0]
        
        # Busca com wildcard
        key_wildcard = f"{cst}|*|{operation_type}"
        if key_wildcard in self.rule_cache and self.rule_cache[key_wildcard]:
            return self.rule_cache[key_wildcard][0]
        
        # Busca com operation_type 'ambos'
        key_ambos = f"{cst}|{cfop}|ambos"
        if key_ambos in self.rule_cache and self.rule_cache[key_ambos]:
            return self.rule_cache[key_ambos][0]
        
        # Busca com wildcard e ambos
        key_ambos_wildcard = f"{cst}|*|ambos"
        if key_ambos_wildcard in self.rule_cache and self.rule_cache[key_ambos_wildcard]:
            return self.rule_cache[key_ambos_wildcard][0]
        
        return None
    
    def validate_item(self, item: FiscalItem) -> Tuple[bool, List[str]]:
        """Valida item fiscal contra regras"""
        errors = []
        
        rule = self.get_rule(item.cst, item.cfop, item.tipo_operacao)
        if not rule:
            errors.append(f"Regra não encontrada para CST {item.cst}, CFOP {item.cfop}")
            return False, errors
        
        if rule.requires_base and (item.base_calculo is None or item.base_calculo <= 0):
            errors.append("Base de cálculo obrigatória não informada")
        
        if rule.requires_aliquota and (item.aliquota is None or item.aliquota <= 0):
            errors.append("Alíquota obrigatória não informada")
        
        if rule.requires_imposto and (item.valor_imposto is None or item.valor_imposto <= 0):
            errors.append("Valor do imposto obrigatório não informado")
        
        # Validações para CST isento
        if not rule.requires_base and item.base_calculo and item.base_calculo > 0:
            errors.append("CST isento com base de cálculo informada")
        
        if not rule.requires_aliquota and item.aliquota and item.aliquota > 0:
            errors.append("CST isento com alíquota informada")
        
        if not rule.requires_imposto and item.valor_imposto and item.valor_imposto > 0:
            errors.append("CST isento com valor de imposto informado")
        
        return len(errors) == 0, errors
    
    def suggest_correction(self, item: FiscalItem) -> Dict[str, Any]:
        """Sugere correções para um item fiscal"""
        suggestions = {}
        
        rule = self.get_rule(item.cst, item.cfop, item.tipo_operacao)
        if not rule:
            return {"error": "Regra não encontrada"}
        
        # Sugere base de cálculo
        if rule.requires_base and (item.base_calculo is None or item.base_calculo <= 0):
            if rule.default_base_formula == "item_value" and item.valor_total:
                suggestions["base_calculo"] = item.valor_total
        
        # Sugere alíquota
        if rule.requires_aliquota and (item.aliquota is None or item.aliquota <= 0):
            if rule.default_aliquota:
                suggestions["aliquota"] = rule.default_aliquota
        
        # Sugere valor do imposto
        if rule.requires_imposto and (item.valor_imposto is None or item.valor_imposto <= 0):
            base = suggestions.get("base_calculo", item.base_calculo)
            aliquota = suggestions.get("aliquota", item.aliquota)
            
            if base and aliquota:
                try:
                    imposto = (base * aliquota) / Decimal("100")
                    suggestions["valor_imposto"] = imposto.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                except Exception:
                    pass
        
        return suggestions
    
    def add_rule(self, rule: FiscalRule):
        """Adiciona regra"""
        self.rules.append(rule)
        self._build_cache()
    
    def remove_rule(self, rule_id: str):
        """Remove regra"""
        self.rules = [r for r in self.rules if r.rule_id != rule_id]
        self._build_cache()

# ============================================================================
# PARSERS
# ============================================================================

class SpedParser:
    """Parser para arquivos SPED"""
    
    def __init__(self):
        self.records: List[SpedRecord] = []
        self.documents: Dict[str, FiscalDocument] = {}
        self.items: List[FiscalItem] = []
        self.blocks: Dict[str, List[SpedRecord]] = {}
        self.summary: Dict[str, Any] = {}
        self.company_info: Dict[str, str] = {}
        self.period: str = ""
    
    def parse(self, content: str) -> Dict[str, Any]:
        """Parseia arquivo SPED"""
        try:
            lines = content.splitlines()
            self._parse_lines(lines)
            self._build_hierarchy()
            self._extract_documents()
            self._generate_summary()
            
            return {
                'records': {r.record_id: r for r in self.records},
                'documents': self.documents,
                'blocks': self.blocks,
                'items': self.items,
                'summary': self.summary,
                'period': self.period,
                'company_info': self.company_info
            }
        except Exception as e:
            raise Exception(f"Erro ao parsear arquivo: {str(e)}")
    
    def _parse_lines(self, lines: List[str]):
        """Parseia linhas do arquivo"""
        current_block = "0"
        current_parent = None
        
        for line_num, line in enumerate(lines, 1):
            if not line.strip():
                continue
            
            parts = line.split("|")
            if len(parts) < 2:
                continue
            
            record_type = parts[1].strip()
            
            # Identifica bloco
            block = self._identify_block(record_type)
            
            # Extrai campos
            fields = self._extract_fields(record_type, parts)
            
            # Cria registro
            record = SpedRecord(
                record_id=f"{record_type}_{line_num}",
                record_type=record_type,
                block=block,
                line_number=line_num,
                raw_data=line,
                fields=fields,
                original_data=copy.deepcopy(fields)
            )
            
            # Gerencia hierarquia
            if record_type in ["0000", "1000", "C100", "D100", "E100", "F100"]:
                current_parent = record
            elif current_parent:
                record.parent_record = current_parent.record_id
                current_parent.children_records.append(record.record_id)
            
            # Armazena registro
            self.records.append(record)
            
            # Armazena no bloco
            if block not in self.blocks:
                self.blocks[block] = []
            self.blocks[block].append(record)
            
            # Extrai informações da empresa
            if record_type == "0000":
                self.company_info = {
                    'cnpj': fields.get('cnpj', ''),
                    'nome': fields.get('nome_empresa', ''),
                    'ie': fields.get('ie', '')
                }
                self.period = f"{fields.get('data_inicial', '')} a {fields.get('data_final', '')}"
    
    def _identify_block(self, record_type: str) -> str:
        """Identifica bloco do registro"""
        if record_type.startswith("0"):
            return "0"
        elif record_type.startswith("1"):
            return "1"
        elif record_type.startswith("C"):
            return "C"
        elif record_type.startswith("D"):
            return "D"
        elif record_type.startswith("E"):
            return "E"
        elif record_type.startswith("F"):
            return "F"
        elif record_type.startswith("H"):
            return "H"
        elif record_type.startswith("9"):
            return "9"
        return "0"
    
    def _extract_fields(self, record_type: str, parts: List[str]) -> Dict[str, Any]:
        """Extrai campos do registro"""
        fields = {}
        
        # Registro 0000 - Abertura
        if record_type == "0000" and len(parts) >= 6:
            fields["codigo_versao"] = parts[2] if len(parts) > 2 else ""
            fields["tipo_arquivo"] = parts[3] if len(parts) > 3 else ""
            fields["cnpj"] = self._clean(parts[4] if len(parts) > 4 else "")
            fields["nome_empresa"] = self._clean(parts[5] if len(parts) > 5 else "")
            fields["ie"] = self._clean(parts[7] if len(parts) > 7 else "")
            fields["data_inicial"] = parts[9] if len(parts) > 9 else ""
            fields["data_final"] = parts[10] if len(parts) > 10 else ""
        
        # Registro C100 - Nota Fiscal
        elif record_type == "C100" and len(parts) >= 15:
            fields["tipo_documento"] = parts[2] if len(parts) > 2 else ""
            fields["serie"] = self._clean(parts[3] if len(parts) > 3 else "")
            fields["numero_documento"] = parts[4] if len(parts) > 4 else ""
            fields["chave_acesso"] = self._clean(parts[5] if len(parts) > 5 else "")
            fields["data_emissao"] = parts[6] if len(parts) > 6 else ""
            fields["cnpj_emitente"] = self._clean(parts[7] if len(parts) > 7 else "")
            fields["cnpj_destinatario"] = self._clean(parts[8] if len(parts) > 8 else "")
            fields["valor_total"] = self._parse_decimal(parts[9] if len(parts) > 9 else "0")
            fields["base_calculo"] = self._parse_decimal(parts[10] if len(parts) > 10 else None)
            fields["valor_icms"] = self._parse_decimal(parts[11] if len(parts) > 11 else None)
            fields["aliquota"] = self._parse_decimal(parts[12] if len(parts) > 12 else None)
            fields["cfop"] = self._clean(parts[13] if len(parts) > 13 else "")
            fields["cst"] = self._clean(parts[14] if len(parts) > 14 else "")
            fields["tipo_operacao"] = self._identify_operation_type(fields.get("cfop", ""))
        
        # Registro C170 - Itens da NF
        elif record_type == "C170" and len(parts) >= 13:
            fields["numero_item"] = int(parts[2]) if len(parts) > 2 and parts[2].isdigit() else 1
            fields["codigo_produto"] = self._clean(parts[3] if len(parts) > 3 else "")
            fields["descricao_produto"] = self._clean(parts[4] if len(parts) > 4 else "")
            fields["quantidade"] = self._parse_decimal(parts[5] if len(parts) > 5 else "0")
            fields["valor_unitario"] = self._parse_decimal(parts[6] if len(parts) > 6 else "0")
            fields["valor_total"] = self._parse_decimal(parts[7] if len(parts) > 7 else "0")
            fields["base_calculo"] = self._parse_decimal(parts[8] if len(parts) > 8 else None)
            fields["aliquota"] = self._parse_decimal(parts[9] if len(parts) > 9 else None)
            fields["valor_imposto"] = self._parse_decimal(parts[10] if len(parts) > 10 else None)
            fields["cst"] = self._clean(parts[11] if len(parts) > 11 else "")
            fields["cfop"] = self._clean(parts[12] if len(parts) > 12 else "")
            fields["tipo_operacao"] = self._identify_operation_type(fields.get("cfop", ""))
            fields["block"] = "C"
        
        # Registro D100 - Transporte (CT-e)
        elif record_type == "D100" and len(parts) >= 9:
            fields["tipo_operacao"] = self._clean(parts[2] if len(parts) > 2 else "")
            fields["cfop"] = self._clean(parts[3] if len(parts) > 3 else "")
            fields["cnpj_emitente"] = self._clean(parts[4] if len(parts) > 4 else "")
            fields["valor_servico"] = self._parse_decimal(parts[5] if len(parts) > 5 else "0")
            fields["base_calculo"] = self._parse_decimal(parts[6] if len(parts) > 6 else None)
            fields["aliquota"] = self._parse_decimal(parts[7] if len(parts) > 7 else None)
            fields["valor_imposto"] = self._parse_decimal(parts[8] if len(parts) > 8 else None)
            fields["block"] = "D"
        
        # Registro D101 - Complemento CT-e
        elif record_type == "D101" and len(parts) >= 6:
            fields["cfop"] = self._clean(parts[2] if len(parts) > 2 else "")
            fields["base_calculo"] = self._parse_decimal(parts[3] if len(parts) > 3 else None)
            fields["aliquota"] = self._parse_decimal(parts[4] if len(parts) > 4 else None)
            fields["valor_imposto"] = self._parse_decimal(parts[5] if len(parts) > 5 else None)
            fields["block"] = "D"
        
        # Registro D105 - Complemento CT-e
        elif record_type == "D105" and len(parts) >= 6:
            fields["cfop"] = self._clean(parts[2] if len(parts) > 2 else "")
            fields["base_calculo"] = self._parse_decimal(parts[3] if len(parts) > 3 else None)
            fields["aliquota"] = self._parse_decimal(parts[4] if len(parts) > 4 else None)
            fields["valor_imposto"] = self._parse_decimal(parts[5] if len(parts) > 5 else None)
            fields["block"] = "D"
        
        # Registro C500 - Energia Elétrica
        elif record_type == "C500" and len(parts) >= 9:
            fields["tipo_operacao"] = self._clean(parts[2] if len(parts) > 2 else "")
            fields["cfop"] = self._clean(parts[3] if len(parts) > 3 else "")
            fields["cnpj_emitente"] = self._clean(parts[4] if len(parts) > 4 else "")
            fields["valor_total"] = self._parse_decimal(parts[5] if len(parts) > 5 else "0")
            fields["base_calculo"] = self._parse_decimal(parts[6] if len(parts) > 6 else None)
            fields["aliquota"] = self._parse_decimal(parts[7] if len(parts) > 7 else None)
            fields["valor_imposto"] = self._parse_decimal(parts[8] if len(parts) > 8 else None)
            fields["block"] = "C"
        
        # Registro C501 - Complemento Energia
        elif record_type == "C501" and len(parts) >= 6:
            fields["cfop"] = self._clean(parts[2] if len(parts) > 2 else "")
            fields["base_calculo"] = self._parse_decimal(parts[3] if len(parts) > 3 else None)
            fields["aliquota"] = self._parse_decimal(parts[4] if len(parts) > 4 else None)
            fields["valor_imposto"] = self._parse_decimal(parts[5] if len(parts) > 5 else None)
            fields["block"] = "C"
        
        # Registro C505 - Complemento Energia
        elif record_type == "C505" and len(parts) >= 6:
            fields["cfop"] = self._clean(parts[2] if len(parts) > 2 else "")
            fields["base_calculo"] = self._parse_decimal(parts[3] if len(parts) > 3 else None)
            fields["aliquota"] = self._parse_decimal(parts[4] if len(parts) > 4 else None)
            fields["valor_imposto"] = self._parse_decimal(parts[5] if len(parts) > 5 else None)
            fields["block"] = "C"
        
        # Registro F100 - Aluguel
        elif record_type == "F100" and len(parts) >= 8:
            fields["tipo_operacao"] = self._clean(parts[2] if len(parts) > 2 else "")
            fields["cfop"] = self._clean(parts[3] if len(parts) > 3 else "")
            fields["cnpj_emitente"] = self._clean(parts[4] if len(parts) > 4 else "")
            fields["valor_total"] = self._parse_decimal(parts[5] if len(parts) > 5 else "0")
            fields["base_calculo"] = self._parse_decimal(parts[6] if len(parts) > 6 else None)
            fields["aliquota"] = self._parse_decimal(parts[7] if len(parts) > 7 else None)
            fields["valor_imposto"] = self._parse_decimal(parts[8] if len(parts) > 8 else None)
            fields["block"] = "F"
        
        return fields
    
    def _build_hierarchy(self):
        """Constrói hierarquia de registros"""
        # Já construída durante o parsing
        pass
    
    def _extract_documents(self):
        """Extrai documentos fiscais"""
        # Processa registros C100
        for record in self.records:
            if record.record_type == "C100":
                doc = self._create_document_from_c100(record)
                if doc.chave_acesso:
                    self.documents[doc.chave_acesso] = doc
            
            # Processa registros D100
            elif record.record_type == "D100":
                doc = self._create_document_from_d100(record)
                if doc.chave_acesso:
                    self.documents[doc.chave_acesso] = doc
            
            # Processa registros C500
            elif record.record_type == "C500":
                doc = self._create_document_from_c500(record)
                if doc.chave_acesso:
                    self.documents[doc.chave_acesso] = doc
            
            # Processa registros F100
            elif record.record_type == "F100":
                doc = self._create_document_from_f100(record)
                if doc.chave_acesso:
                    self.documents[doc.chave_acesso] = doc
        
        # Extrai itens
        self._extract_items()
    
    def _create_document_from_c100(self, record: SpedRecord) -> FiscalDocument:
        """Cria documento a partir de C100"""
        fields = record.fields
        return FiscalDocument(
            chave_acesso=fields.get('chave_acesso', f"C100_{record.line_number}"),
            cnpj_emitente=fields.get('cnpj_emitente', ''),
            cnpj_destinatario=fields.get('cnpj_destinatario', ''),
            valor_total=fields.get('valor_total', Decimal('0')),
            valor_base_calculo=fields.get('base_calculo'),
            valor_icms=fields.get('valor_icms'),
            aliquota_icms=fields.get('aliquota'),
            cst=fields.get('cst', ''),
            cfop=fields.get('cfop', ''),
            tipo_operacao=fields.get('tipo_operacao', 'entrada'),
            data_emissao=self._parse_date(fields.get('data_emissao', '')),
            block='C',
            record_type='C100'
        )
    
    def _create_document_from_d100(self, record: SpedRecord) -> FiscalDocument:
        """Cria documento a partir de D100"""
        fields = record.fields
        return FiscalDocument(
            chave_acesso=f"D100_{record.line_number}",
            cnpj_emitente=fields.get('cnpj_emitente', ''),
            cnpj_destinatario='',
            valor_total=fields.get('valor_servico', Decimal('0')),
            valor_base_calculo=fields.get('base_calculo'),
            valor_icms=fields.get('valor_imposto'),
            aliquota_icms=fields.get('aliquota'),
            cst='00',
            cfop=fields.get('cfop', ''),
            tipo_operacao=fields.get('tipo_operacao', 'entrada'),
            data_emissao=date.today(),
            block='D',
            record_type='D100'
        )
    
    def _create_document_from_c500(self, record: SpedRecord) -> FiscalDocument:
        """Cria documento a partir de C500"""
        fields = record.fields
        return FiscalDocument(
            chave_acesso=f"C500_{record.line_number}",
            cnpj_emitente=fields.get('cnpj_emitente', ''),
            cnpj_destinatario='',
            valor_total=fields.get('valor_total', Decimal('0')),
            valor_base_calculo=fields.get('base_calculo'),
            valor_icms=fields.get('valor_imposto'),
            aliquota_icms=fields.get('aliquota'),
            cst='00',
            cfop=fields.get('cfop', ''),
            tipo_operacao=fields.get('tipo_operacao', 'entrada'),
            data_emissao=date.today(),
            block='C',
            record_type='C500'
        )
    
    def _create_document_from_f100(self, record: SpedRecord) -> FiscalDocument:
        """Cria documento a partir de F100"""
        fields = record.fields
        return FiscalDocument(
            chave_acesso=f"F100_{record.line_number}",
            cnpj_emitente=fields.get('cnpj_emitente', ''),
            cnpj_destinatario='',
            valor_total=fields.get('valor_total', Decimal('0')),
            valor_base_calculo=fields.get('base_calculo'),
            valor_icms=fields.get('valor_imposto'),
            aliquota_icms=fields.get('aliquota'),
            cst='00',
            cfop=fields.get('cfop', ''),
            tipo_operacao=fields.get('tipo_operacao', 'entrada'),
            data_emissao=date.today(),
            block='F',
            record_type='F100'
        )
    
    def _extract_items(self):
        """Extrai itens dos registros"""
        # Extrai itens C170
        for record in self.records:
            if record.record_type == "C170":
                fields = record.fields
                item = FiscalItem(
                    numero_item=fields.get('numero_item', 1),
                    codigo_produto=fields.get('codigo_produto', ''),
                    descricao_produto=fields.get('descricao_produto', ''),
                    quantidade=fields.get('quantidade', Decimal('0')),
                    valor_unitario=fields.get('valor_unitario', Decimal('0')),
                    valor_total=fields.get('valor_total', Decimal('0')),
                    base_calculo=fields.get('base_calculo'),
                    aliquota=fields.get('aliquota'),
                    valor_imposto=fields.get('valor_imposto'),
                    cst=fields.get('cst', ''),
                    cfop=fields.get('cfop', ''),
                    tipo_operacao=fields.get('tipo_operacao', 'entrada'),
                    block='C',
                    record_id=record.record_id
                )
                self.items.append(item)
        
        # Extrai itens D100/D101/D105
        for record in self.records:
            if record.record_type in ["D100", "D101", "D105"]:
                fields = record.fields
                if fields.get('valor_servico') or fields.get('base_calculo'):
                    item = FiscalItem(
                        numero_item=1,
                        codigo_produto='TRANSPORTE',
                        descricao_produto='Serviço de Transporte',
                        quantidade=Decimal('1'),
                        valor_unitario=fields.get('valor_servico', Decimal('0')),
                        valor_total=fields.get('valor_servico', Decimal('0')),
                        base_calculo=fields.get('base_calculo'),
                        aliquota=fields.get('aliquota'),
                        valor_imposto=fields.get('valor_imposto'),
                        cst='00',
                        cfop=fields.get('cfop', ''),
                        tipo_operacao=fields.get('tipo_operacao', 'entrada'),
                        block='D',
                        record_id=record.record_id
                    )
                    self.items.append(item)
        
        # Extrai itens C500/C501/C505 (Energia)
        for record in self.records:
            if record.record_type in ["C500", "C501", "C505"]:
                fields = record.fields
                if fields.get('valor_total') or fields.get('base_calculo'):
                    item = FiscalItem(
                        numero_item=1,
                        codigo_produto='ENERGIA',
                        descricao_produto='Energia Elétrica',
                        quantidade=Decimal('1'),
                        valor_unitario=fields.get('valor_total', Decimal('0')),
                        valor_total=fields.get('valor_total', Decimal('0')),
                        base_calculo=fields.get('base_calculo'),
                        aliquota=fields.get('aliquota'),
                        valor_imposto=fields.get('valor_imposto'),
                        cst='00',
                        cfop=fields.get('cfop', ''),
                        tipo_operacao=fields.get('tipo_operacao', 'entrada'),
                        block='C',
                        record_id=record.record_id
                    )
                    self.items.append(item)
        
        # Extrai itens F100 (Aluguel)
        for record in self.records:
            if record.record_type == "F100":
                fields = record.fields
                if fields.get('valor_total') or fields.get('base_calculo'):
                    item = FiscalItem(
                        numero_item=1,
                        codigo_produto='ALUGUEL',
                        descricao_produto='Serviço de Locação',
                        quantidade=Decimal('1'),
                        valor_unitario=fields.get('valor_total', Decimal('0')),
                        valor_total=fields.get('valor_total', Decimal('0')),
                        base_calculo=fields.get('base_calculo'),
                        aliquota=fields.get('aliquota'),
                        valor_imposto=fields.get('valor_imposto'),
                        cst='00',
                        cfop=fields.get('cfop', ''),
                        tipo_operacao=fields.get('tipo_operacao', 'entrada'),
                        block='F',
                        record_id=record.record_id
                    )
                    self.items.append(item)
    
    def _generate_summary(self):
        """Gera resumo do arquivo"""
        self.summary = {
            'total_records': len(self.records),
            'blocks_count': len(self.blocks),
            'documents_count': len(self.documents),
            'items_count': len(self.items),
            'block_summary': {}
        }
        
        for block, records in self.blocks.items():
            record_types = {}
            for record in records:
                record_type = record.record_type
                if record_type not in record_types:
                    record_types[record_type] = 0
                record_types[record_type] += 1
            
            self.summary['block_summary'][block] = {
                'total_records': len(records),
                'record_types': record_types
            }
    
    def _clean(self, value: str) -> str:
        """Limpa string"""
        if not value:
            return ""
        return value.strip()
    
    def _parse_decimal(self, value: Any) -> Optional[Decimal]:
        """Converte para Decimal"""
        if value is None or value == "":
            return None
        try:
            # Remove caracteres não numéricos
            if isinstance(value, str):
                value = re.sub(r'[^\d.,-]', '', value)
                value = value.replace(',', '.')
            return Decimal(str(value))
        except:
            return None
    
    def _parse_date(self, value: str) -> date:
        """Converte para data"""
        if not value:
            return date.today()
        try:
            if len(value) == 8:  # AAAAMMDD
                return datetime.strptime(value, '%Y%m%d').date()
            elif '-' in value:
                return datetime.strptime(value, '%Y-%m-%d').date()
            elif '/' in value:
                return datetime.strptime(value, '%d/%m/%Y').date()
        except:
            pass
        return date.today()
    
    def _identify_operation_type(self, cfop: str) -> str:
        """Identifica tipo de operação pelo CFOP"""
        if not cfop:
            return "entrada"
        
        # CFOP de entrada: 1xxx, 2xxx, 3xxx
        if cfop.startswith("1") or cfop.startswith("2") or cfop.startswith("3"):
            return "entrada"
        # CFOP de saída: 5xxx, 6xxx, 7xxx
        elif cfop.startswith("5") or cfop.startswith("6") or cfop.startswith("7"):
            return "saida"
        return "entrada"

# ============================================================================
# SERVIÇOS
# ============================================================================

class ValidationService:
    """Serviço de validação fiscal"""
    
    def __init__(self, rules_engine: RulesEngine):
        self.rules_engine = rules_engine
        self.inconsistencies: List[Dict] = []
    
    def validate_all(self, items: List[FiscalItem]) -> List[Dict]:
        """Valida todos os itens"""
        self.inconsistencies = []
        
        for item in items:
            # Valida campos obrigatórios
            if not item.cst:
                self.inconsistencies.append({
                    'type': 'critical',
                    'item': item.descricao_produto,
                    'message': 'CST não informado',
                    'details': f'Produto: {item.descricao_produto}, CFOP: {item.cfop}',
                    'severity': 'alta'
                })
            
            if not item.cfop:
                self.inconsistencies.append({
                    'type': 'critical',
                    'item': item.descricao_produto,
                    'message': 'CFOP não informado',
                    'details': f'Produto: {item.descricao_produto}, CST: {item.cst}',
                    'severity': 'alta'
                })
            
            # Valida regras fiscais
            is_valid, errors = self.rules_engine.validate_item(item)
            if not is_valid:
                for error in errors:
                    severity = 'alta' if 'obrigatória' in error else 'media'
                    self.inconsistencies.append({
                        'type': 'critical' if severity == 'alta' else 'warning',
                        'item': item.descricao_produto,
                        'message': error,
                        'details': f'Produto: {item.descricao_produto}, CST: {item.cst}, CFOP: {item.cfop}',
                        'severity': severity
                    })
        
        return self.inconsistencies

class CorrectionService:
    """Serviço de correção fiscal"""
    
    def __init__(self, rules_engine: RulesEngine):
        self.rules_engine = rules_engine
        self.corrections: List[Dict] = []
    
    def correct_item(self, item: FiscalItem) -> Dict[str, Any]:
        """Corrige um item fiscal"""
        suggestions = self.rules_engine.suggest_correction(item)
        
        if "error" in suggestions:
            return suggestions
        
        changes = {}
        
        # Salva valores originais
        if not hasattr(item, 'original_base'):
            item.original_base = item.base_calculo
        if not hasattr(item, 'original_aliquota'):
            item.original_aliquota = item.aliquota
        if not hasattr(item, 'original_imposto'):
            item.original_imposto = item.valor_imposto
        
        # Aplica correções
        if "base_calculo" in suggestions:
            changes['base_calculo'] = {
                'old': item.base_calculo,
                'new': suggestions['base_calculo']
            }
            item.base_calculo = suggestions['base_calculo']
        
        if "aliquota" in suggestions:
            changes['aliquota'] = {
                'old': item.aliquota,
                'new': suggestions['aliquota']
            }
            item.aliquota = suggestions['aliquota']
        
        if "valor_imposto" in suggestions:
            changes['valor_imposto'] = {
                'old': item.valor_imposto,
                'new': suggestions['valor_imposto']
            }
            item.valor_imposto = suggestions['valor_imposto']
        
        if changes:
            item.modified = True
            self.corrections.append({
                'item': item.descricao_produto,
                'changes': changes,
                'timestamp': datetime.now().isoformat()
            })
        
        return changes
    
    def correct_mass(self, items: List[FiscalItem]) -> Dict[str, int]:
        """Corrige múltiplos itens em massa"""
        results = {
            'total': len(items),
            'corrected': 0,
            'base_corrected': 0,
            'aliquota_corrected': 0,
            'imposto_corrected': 0
        }
        
        for item in items:
            changes = self.correct_item(item)
            if changes:
                results['corrected'] += 1
                if 'base_calculo' in changes:
                    results['base_corrected'] += 1
                if 'aliquota' in changes:
                    results['aliquota_corrected'] += 1
                if 'valor_imposto' in changes:
                    results['imposto_corrected'] += 1
        
        return results

class ExportService:
    """Serviço de exportação"""
    
    def __init__(self):
        pass
    
    def export_sped(self, records: List[SpedRecord]) -> str:
        """Exporta SPED corrigido"""
        lines = []
        for record in records:
            # Reconstrói a linha com os campos atualizados
            if record.modified:
                # Se modificado, usa campos atuais
                parts = [record.record_type]
                # Adiciona campos na ordem correta
                for key in record.fields:
                    value = record.fields.get(key, '')
                    if value is None:
                        value = ''
                    parts.append(str(value))
                lines.append('|' + '|'.join(parts))
            else:
                # Usa dados originais
                lines.append(record.raw_data)
        
        return '\n'.join(lines)
    
    def export_excel(self, items: List[FiscalItem], inconsistencies: List[Dict], 
                     corrections: List[Dict], audit_log: List[Dict]) -> bytes:
        """Exporta relatório em Excel"""
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Aba: Itens
            if items:
                items_data = []
                for item in items:
                    items_data.append({
                        'Produto': item.descricao_produto,
                        'CST': item.cst,
                        'CFOP': item.cfop,
                        'Valor Total': float(item.valor_total) if item.valor_total else 0,
                        'Base': float(item.base_calculo) if item.base_calculo else 0,
                        'Alíquota': float(item.aliquota) if item.aliquota else 0,
                        'Imposto': float(item.valor_imposto) if item.valor_imposto else 0,
                        'Block': item.block,
                        'Modificado': 'Sim' if item.modified else 'Não'
                    })
                df_items = pd.DataFrame(items_data)
                df_items.to_excel(writer, sheet_name='Itens', index=False)
            
            # Aba: Inconsistências
            if inconsistencies:
                df_inc = pd.DataFrame(inconsistencies)
                df_inc.to_excel(writer, sheet_name='Inconsistencias', index=False)
            
            # Aba: Correções
            if corrections:
                corr_data = []
                for corr in corrections:
                    for field, change in corr.get('changes', {}).items():
                        corr_data.append({
                            'Item': corr.get('item', ''),
                            'Campo': field,
                            'Valor Antigo': change.get('old'),
                            'Valor Novo': change.get('new')
                        })
                if corr_data:
                    df_corr = pd.DataFrame(corr_data)
                    df_corr.to_excel(writer, sheet_name='Correcoes', index=False)
            
            # Aba: Auditoria
            if audit_log:
                df_audit = pd.DataFrame(audit_log)
                df_audit.to_excel(writer, sheet_name='Auditoria', index=False)
        
        return output.getvalue()
    
    def export_csv_inconsistencies(self, inconsistencies: List[Dict]) -> str:
        """Exporta inconsistências em CSV"""
        df = pd.DataFrame(inconsistencies)
        return df.to_csv(index=False, encoding='utf-8')

# ============================================================================
# UI COMPONENTS
# ============================================================================

def render_dashboard():
    """Renderiza dashboard"""
    st.markdown("""
        <style>
        .dashboard-header {
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            padding: 2rem;
            border-radius: 10px;
            color: white;
            margin-bottom: 2rem;
        }
        .metric-card {
            background: white;
            padding: 1.5rem;
            border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            border-left: 4px solid #2a5298;
        }
        .metric-value {
            font-size: 2rem;
            font-weight: bold;
            color: #1e3c72;
        }
        .metric-label {
            color: #6c757d;
            font-size: 0.9rem;
        }
        </style>
    """, unsafe_allow_html=True)
    
    st.markdown('<div class="dashboard-header"><h1>📊 Dashboard Fiscal</h1></div>', unsafe_allow_html=True)
    
    if not st.session_state.get('items'):
        st.info("Nenhum arquivo SPED carregado. Acesse a seção 'Upload' para começar.")
        return
    
    # Métricas principais
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{len(st.session_state.get('items', []))}</div>
                <div class="metric-label">📦 Total Itens</div>
            </div>
        """, unsafe_allow_html=True)
    
    with col2:
        problematic = len([i for i in st.session_state.get('items', []) 
                          if (i.base_calculo is None or i.base_calculo <= 0) or
                             (i.aliquota is None or i.aliquota <= 0) or
                             (i.valor_imposto is None or i.valor_imposto <= 0)])
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #dc3545;">
                <div class="metric-value" style="color: #dc3545;">{problematic}</div>
                <div class="metric-label">⚠️ Itens com Problemas</div>
            </div>
        """, unsafe_allow_html=True)
    
    with col3:
        total_docs = len(st.session_state.get('documents', {}))
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #28a745;">
                <div class="metric-value" style="color: #28a745;">{total_docs}</div>
                <div class="metric-label">📑 Notas Fiscais</div>
            </div>
        """, unsafe_allow_html=True)
    
    with col4:
        total_corrected = len(st.session_state.get('audit_log', []))
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #ffc107;">
                <div class="metric-value" style="color: #ffc107;">{total_corrected}</div>
                <div class="metric-label">✅ Correções Aplicadas</div>
            </div>
        """, unsafe_allow_html=True)
    
    # Gráficos
    st.markdown("---")
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Distribuição por CST")
        if st.session_state.get('items'):
            cst_data = {}
            for item in st.session_state.items:
                cst = item.cst or 'N/A'
                cst_data[cst] = cst_data.get(cst, 0) + 1
            
            if cst_data:
                df_cst = pd.DataFrame(list(cst_data.items()), columns=['CST', 'Quantidade'])
                fig = px.pie(df_cst, values='Quantidade', names='CST', 
                           title='Itens por CST', color_discrete_sequence=px.colors.qualitative.Set3)
                st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("Distribuição por CFOP")
        if st.session_state.get('items'):
            cfop_data = {}
            for item in st.session_state.items:
                cfop = item.cfop or 'N/A'
                cfop_data[cfop] = cfop_data.get(cfop, 0) + 1
            
            if cfop_data:
                df_cfop = pd.DataFrame(list(cfop_data.items()), columns=['CFOP', 'Quantidade'])
                fig = px.bar(df_cfop, x='CFOP', y='Quantidade', 
                           title='Itens por CFOP', color='CFOP')
                st.plotly_chart(fig, use_container_width=True)
    
    # Tabela de blocos
    st.subheader("📋 Resumo por Bloco")
    if st.session_state.get('records'):
        block_data = []
        for record in st.session_state.records:
            block_data.append({
                'Bloco': record.block,
                'Registro': record.record_type,
                'Linha': record.line_number,
                'Modificado': '✅' if record.modified else ''
            })
        
        df_blocks = pd.DataFrame(block_data)
        block_summary = df_blocks.groupby(['Bloco', 'Registro']).size().reset_index(name='Quantidade')
        st.dataframe(block_summary, use_container_width=True)

def render_upload():
    """Renderiza seção de upload"""
    st.markdown("""
        <style>
        .upload-box {
            border: 2px dashed #2a5298;
            border-radius: 10px;
            padding: 3rem;
            text-align: center;
            background: #f8f9fa;
        }
        </style>
    """, unsafe_allow_html=True)
    
    st.markdown('<h1>📤 Upload de Arquivo SPED</h1>', unsafe_allow_html=True)
    
    st.markdown("""
    ### Instruções
    Faça upload do arquivo SPED nos formatos:
    - **EFD ICMS/IPI** (blocos 0, C, D, E, H, 1, 9)
    - **EFD Contribuições** (blocos 0, C, D, F, 1, 9)
    
    O sistema processará automaticamente e estruturará os dados para análise.
    """)
    
    uploaded_file = st.file_uploader(
        "Selecione o arquivo TXT",
        type=['txt'],
        help="Arquivo SPED no formato TXT com codificação UTF-8"
    )
    
    if uploaded_file:
        try:
            content = uploaded_file.read().decode('utf-8')
            st.session_state.file_content = content
            st.session_state.file_name = uploaded_file.name
            
            with st.spinner("Processando arquivo..."):
                parser = SpedParser()
                data = parser.parse(content)
                
                # Armazena no session state
                st.session_state.sped_data = data
                st.session_state.records = data['records']
                st.session_state.documents = data['documents']
                st.session_state.items = data['items']
                st.session_state.blocks = data['blocks']
                st.session_state.summary = data['summary']
                st.session_state.period = data.get('period', '')
                st.session_state.company_info = data.get('company_info', {})
                
                # Inicializa serviços
                if 'rules_engine' not in st.session_state:
                    st.session_state.rules_engine = RulesEngine()
                
                # Valida dados
                validator = ValidationService(st.session_state.rules_engine)
                st.session_state.inconsistencies = validator.validate_all(st.session_state.items)
                
                # Inicializa logs
                if 'audit_log' not in st.session_state:
                    st.session_state.audit_log = []
                
                st.success(f"✅ Arquivo processado com sucesso!")
                
                # Exibe resumo
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Registros", len(st.session_state.records))
                with col2:
                    st.metric("Notas Fiscais", len(st.session_state.documents))
                with col3:
                    st.metric("Itens", len(st.session_state.items))
                with col4:
                    st.metric("Inconsistências", len(st.session_state.inconsistencies))
                
                # Informações da empresa
                if st.session_state.company_info:
                    with st.expander("🏢 Informações da Empresa"):
                        st.json(st.session_state.company_info)
                
                # Resumo dos blocos
                with st.expander("📊 Resumo dos Blocos"):
                    block_summary = []
                    for block, records in st.session_state.blocks.items():
                        record_types = {}
                        for rec in records:
                            record_types[rec.record_type] = record_types.get(rec.record_type, 0) + 1
                        
                        block_summary.append({
                            'Bloco': block,
                            'Total Registros': len(records),
                            'Tipos': ', '.join(record_types.keys())
                        })
                    
                    df_summary = pd.DataFrame(block_summary)
                    st.dataframe(df_summary, use_container_width=True)
        
        except Exception as e:
            st.error(f"❌ Erro ao processar arquivo: {str(e)}")
            st.exception(e)

def render_blocks():
    """Renderiza visualização de blocos"""
    st.markdown('<h1>📋 Visualização de Blocos</h1>', unsafe_allow_html=True)
    
    if not st.session_state.get('records'):
        st.info("Nenhum arquivo carregado. Acesse a seção 'Upload'.")
        return
    
    # Seleção de bloco
    blocks = sorted(st.session_state.blocks.keys())
    selected_block = st.selectbox("Selecione o Bloco", blocks)
    
    if selected_block:
        records = st.session_state.blocks.get(selected_block, [])
        st.subheader(f"Bloco {selected_block} - {len(records)} registros")
        
        # Filtros
        col1, col2 = st.columns(2)
        with col1:
            record_types = list(set(r.record_type for r in records))
            selected_type = st.selectbox("Tipo de Registro", ["Todos"] + sorted(record_types))
        
        with col2:
            search = st.text_input("Buscar", placeholder="Digite para filtrar...")
        
        # Aplica filtros
        filtered = records
        if selected_type != "Todos":
            filtered = [r for r in filtered if r.record_type == selected_type]
        if search:
            filtered = [r for r in filtered if search.lower() in r.raw_data.lower()]
        
        # Exibe registros
        if filtered:
            data = []
            for record in filtered[:1000]:
                row = {
                    'Tipo': record.record_type,
                    'Linha': record.line_number,
                    'Bloco': record.block,
                    'Campos': len(record.fields)
                }
                # Adiciona campos principais
                for key in list(record.fields.keys())[:3]:
                    row[key] = str(record.fields.get(key, ''))[:30]
                data.append(row)
            
            df = pd.DataFrame(data)
            st.dataframe(df, use_container_width=True)
        else:
            st.info("Nenhum registro encontrado.")

def render_records():
    """Renderiza visualização de registros"""
    st.markdown('<h1>📄 Visualização de Registros</h1>', unsafe_allow_html=True)
    
    if not st.session_state.get('records'):
        st.info("Nenhum arquivo carregado.")
        return
    
    records = list(st.session_state.records.values()) if isinstance(st.session_state.records, dict) else st.session_state.records
    
    if not records:
        st.info("Nenhum registro encontrado.")
        return
    
    # Filtros
    col1, col2, col3 = st.columns(3)
    with col1:
        record_types = list(set(r.record_type for r in records))
        selected_types = st.multiselect("Tipo de Registro", sorted(record_types))
    
    with col2:
        blocks = list(set(r.block for r in records))
        selected_blocks = st.multiselect("Bloco", sorted(blocks))
    
    with col3:
        search = st.text_input("Buscar", placeholder="Digite para filtrar...")
    
    # Aplica filtros
    filtered = records
    if selected_types:
        filtered = [r for r in filtered if r.record_type in selected_types]
    if selected_blocks:
        filtered = [r for r in filtered if r.block in selected_blocks]
    if search:
        filtered = [r for r in filtered if search.lower() in r.raw_data.lower()]
    
    st.info(f"📊 Registros encontrados: {len(filtered)}")
    
    # Paginação
    page_size = 50
    total_pages = (len(filtered) + page_size - 1) // page_size
    page = st.number_input("Página", min_value=1, max_value=total_pages or 1, value=1)
    
    start = (page - 1) * page_size
    end = min(start + page_size, len(filtered))
    page_records = filtered[start:end]
    
    # Exibe registros
    if page_records:
        data = []
        for record in page_records:
            data.append({
                'Tipo': record.record_type,
                'Bloco': record.block,
                'Linha': record.line_number,
                'Campos': len(record.fields),
                'Modificado': '✅' if record.modified else ''
            })
        
        df = pd.DataFrame(data)
        st.dataframe(df, use_container_width=True)
        
        st.caption(f"Mostrando {len(page_records)} de {len(filtered)} registros (Página {page} de {total_pages})")
        
        # Detalhes do registro selecionado
        st.markdown("---")
        selected_idx = st.selectbox("Selecione um registro para detalhes", range(len(page_records)))
        if selected_idx is not None:
            record = page_records[selected_idx]
            st.json({
                'Tipo': record.record_type,
                'Bloco': record.block,
                'Linha': record.line_number,
                'Campos': record.fields,
                'Modificado': record.modified
            })

def render_invoices():
    """Renderiza visualização de notas fiscais"""
    st.markdown('<h1>📑 Notas Fiscais</h1>', unsafe_allow_html=True)
    
    if not st.session_state.get('documents'):
        st.info("Nenhuma nota fiscal encontrada.")
        return
    
    documents = list(st.session_state.documents.values())
    
    st.info(f"📊 Notas fiscais encontradas: {len(documents)}")
    
    # Filtros
    col1, col2 = st.columns(2)
    with col1:
        cst_list = list(set(d.cst for d in documents if d.cst))
        selected_cst = st.multiselect("CST", sorted(cst_list))
    
    with col2:
        cfop_list = list(set(d.cfop for d in documents if d.cfop))
        selected_cfop = st.multiselect("CFOP", sorted(cfop_list))
    
    # Aplica filtros
    filtered = documents
    if selected_cst:
        filtered = [d for d in filtered if d.cst in selected_cst]
    if selected_cfop:
        filtered = [d for d in filtered if d.cfop in selected_cfop]
    
    # Exibe notas
    for doc in filtered[:20]:
        with st.expander(f"📄 {doc.record_type} - {doc.chave_acesso[:20]}..."):
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"**Emitente:** {doc.cnpj_emitente}")
                st.write(f"**CST:** {doc.cst}")
                st.write(f"**CFOP:** {doc.cfop}")
            with col2:
                st.write(f"**Valor Total:** R$ {float(doc.valor_total):,.2f}")
                if doc.valor_base_calculo:
                    st.write(f"**Base:** R$ {float(doc.valor_base_calculo):,.2f}")
                if doc.aliquota_icms:
                    st.write(f"**Alíquota:** {float(doc.aliquota_icms):.2f}%")
            
            if doc.items:
                st.markdown("**Itens:**")
                items_data = []
                for item in doc.items:
                    items_data.append({
                        'Item': item.numero_item,
                        'Produto': item.descricao_produto[:30] + '...' if len(item.descricao_produto) > 30 else item.descricao_produto,
                        'Valor': float(item.valor_total) if item.valor_total else 0,
                        'Base': float(item.base_calculo) if item.base_calculo else 0,
                        'Aliq': float(item.aliquota) if item.aliquota else 0,
                        'Imposto': float(item.valor_imposto) if item.valor_imposto else 0
                    })
                df_items = pd.DataFrame(items_data)
                st.dataframe(df_items, use_container_width=True)

def render_items():
    """Renderiza visualização de itens"""
    st.markdown('<h1>📦 Itens Fiscais</h1>', unsafe_allow_html=True)
    
    if not st.session_state.get('items'):
        st.info("Nenhum item encontrado.")
        return
    
    items = st.session_state.items
    
    # Filtros
    col1, col2, col3 = st.columns(3)
    with col1:
        cst_list = list(set(i.cst for i in items if i.cst))
        selected_cst = st.multiselect("CST", sorted(cst_list))
    
    with col2:
        cfop_list = list(set(i.cfop for i in items if i.cfop))
        selected_cfop = st.multiselect("CFOP", sorted(cfop_list))
    
    with col3:
        only_problems = st.checkbox("Apenas itens com problemas", value=False)
    
    # Aplica filtros
    filtered = items
    if selected_cst:
        filtered = [i for i in filtered if i.cst in selected_cst]
    if selected_cfop:
        filtered = [i for i in filtered if i.cfop in selected_cfop]
    if only_problems:
        filtered = [i for i in filtered if 
                   (i.base_calculo is None or i.base_calculo <= 0) or
                   (i.aliquota is None or i.aliquota <= 0) or
                   (i.valor_imposto is None or i.valor_imposto <= 0)]
    
    st.info(f"📊 Itens encontrados: {len(filtered)}")
    
    if filtered:
        # Exibe tabela
        data = []
        for item in filtered[:1000]:
            has_problem = (item.base_calculo is None or item.base_calculo <= 0) or \
                         (item.aliquota is None or item.aliquota <= 0) or \
                         (item.valor_imposto is None or item.valor_imposto <= 0)
            
            data.append({
                'Produto': item.descricao_produto[:40] + '...' if len(item.descricao_produto) > 40 else item.descricao_produto,
                'CST': item.cst,
                'CFOP': item.cfop,
                'Valor Total': float(item.valor_total) if item.valor_total else 0,
                'Base': float(item.base_calculo) if item.base_calculo else 0,
                'Alíquota': float(item.aliquota) if item.aliquota else 0,
                'Imposto': float(item.valor_imposto) if item.valor_imposto else 0,
                'Status': '⚠️' if has_problem else '✅',
                'Block': item.block
            })
        
        df = pd.DataFrame(data)
        st.dataframe(df, use_container_width=True)

def render_inconsistencies():
    """Renderiza seção de inconsistências"""
    st.markdown('<h1>⚠️ Inconsistências Fiscais</h1>', unsafe_allow_html=True)
    
    if not st.session_state.get('inconsistencies'):
        st.success("✅ Nenhuma inconsistência encontrada!")
        return
    
    inconsistencies = st.session_state.inconsistencies
    
    # Estatísticas
    critical = len([i for i in inconsistencies if i.get('severity') == 'alta'])
    warning = len([i for i in inconsistencies if i.get('severity') == 'media'])
    info = len([i for i in inconsistencies if i.get('severity') not in ['alta', 'media']])
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("🔴 Críticas", critical)
    with col2:
        st.metric("🟡 Avisos", warning)
    with col3:
        st.metric("🔵 Informações", info)
    
    # Filtros
    severity_filter = st.selectbox("Filtrar por Severidade", ["Todas", "alta", "media", "baixa"])
    
    filtered = inconsistencies
    if severity_filter != "Todas":
        filtered = [i for i in filtered if i.get('severity') == severity_filter]
    
    # Exibe inconsistências
    for inc in filtered:
        severity = inc.get('severity', 'baixa')
        if severity == 'alta':
            st.error(f"🔴 **{inc.get('item', '')}** - {inc.get('message', '')}")
        elif severity == 'media':
            st.warning(f"🟡 **{inc.get('item', '')}** - {inc.get('message', '')}")
        else:
            st.info(f"🔵 **{inc.get('item', '')}** - {inc.get('message', '')}")
        
        st.caption(inc.get('details', ''))

def render_mass_correction():
    """Renderiza seção de correção em massa"""
    st.markdown('<h1>🔧 Correções em Massa</h1>', unsafe_allow_html=True)
    
    if not st.session_state.get('items'):
        st.info("Nenhum item disponível para correção.")
        return
    
    items = st.session_state.items
    
    # Filtros
    st.subheader("🔍 Filtros")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        cst_filter = st.multiselect("CST", sorted(list(set(i.cst for i in items if i.cst))))
    
    with col2:
        cfop_filter = st.multiselect("CFOP", sorted(list(set(i.cfop for i in items if i.cfop))))
    
    with col3:
        only_problems = st.checkbox("Apenas itens com problemas", value=True)
    
    # Aplica filtros
    filtered = items
    if cst_filter:
        filtered = [i for i in filtered if i.cst in cst_filter]
    if cfop_filter:
        filtered = [i for i in filtered if i.cfop in cfop_filter]
    if only_problems:
        filtered = [i for i in filtered if 
                   (i.base_calculo is None or i.base_calculo <= 0) or
                   (i.aliquota is None or i.aliquota <= 0) or
                   (i.valor_imposto is None or i.valor_imposto <= 0)]
    
    st.info(f"📊 Itens selecionados: {len(filtered)}")
    
    if not filtered:
        st.success("✅ Nenhum item precisa de correção!")
        return
    
    # Ações em massa
    st.subheader("🛠️ Ações")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("📊 Corrigir Bases", type="primary"):
            with st.spinner("Corrigindo bases..."):
                corrected = 0
                for item in filtered:
                    if item.base_calculo is None or item.base_calculo <= 0:
                        if item.valor_total and item.valor_total > 0:
                            item.original_base = item.base_calculo
                            item.base_calculo = item.valor_total
                            item.modified = True
                            corrected += 1
                            
                            st.session_state.audit_log.append({
                                'timestamp': datetime.now().isoformat(),
                                'user': 'Sistema',
                                'operation': 'Correção em Massa - Base',
                                'item': item.descricao_produto,
                                'old_value': item.original_base,
                                'new_value': item.base_calculo
                            })
                
                st.success(f"✅ Bases corrigidas para {corrected} itens")
                st.rerun()
    
    with col2:
        if st.button("📊 Corrigir Alíquotas", type="primary"):
            with st.spinner("Corrigindo alíquotas..."):
                corrected = 0
                for item in filtered:
                    if item.aliquota is None or item.aliquota <= 0:
                        rule = st.session_state.rules_engine.get_rule(
                            item.cst, item.cfop, item.tipo_operacao
                        )
                        if rule and rule.default_aliquota:
                            item.original_aliquota = item.aliquota
                            item.aliquota = rule.default_aliquota
                            item.modified = True
                            corrected += 1
                            
                            st.session_state.audit_log.append({
                                'timestamp': datetime.now().isoformat(),
                                'user': 'Sistema',
                                'operation': 'Correção em Massa - Alíquota',
                                'item': item.descricao_produto,
                                'old_value': item.original_aliquota,
                                'new_value': item.aliquota
                            })
                
                st.success(f"✅ Alíquotas corrigidas para {corrected} itens")
                st.rerun()
    
    with col3:
        if st.button("💰 Recalcular Impostos", type="primary"):
            with st.spinner("Recalculando impostos..."):
                corrected = 0
                for item in filtered:
                    if item.base_calculo and item.aliquota:
                        if item.valor_imposto is None or item.valor_imposto <= 0:
                            try:
                                imposto = item.base_calculo * item.aliquota / 100
                                item.original_imposto = item.valor_imposto
                                item.valor_imposto = imposto.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                                item.modified = True
                                corrected += 1
                                
                                st.session_state.audit_log.append({
                                    'timestamp': datetime.now().isoformat(),
                                    'user': 'Sistema',
                                    'operation': 'Correção em Massa - Imposto',
                                    'item': item.descricao_produto,
                                    'old_value': item.original_imposto,
                                    'new_value': item.valor_imposto
                                })
                            except Exception as e:
                                st.warning(f"Erro ao recalcular imposto para {item.descricao_produto}")
                
                st.success(f"✅ Impostos recalculados para {corrected} itens")
                st.rerun()
    
    # Aplicar todas
    st.markdown("---")
    if st.button("🚀 Aplicar Todas as Correções", type="primary"):
        with st.spinner("Aplicando todas as correções..."):
            corrections = CorrectionService(st.session_state.rules_engine)
            results = corrections.correct_mass(filtered)
            
            # Atualiza logs
            for item in filtered:
                if item.modified:
                    st.session_state.audit_log.append({
                        'timestamp': datetime.now().isoformat(),
                        'user': 'Sistema',
                        'operation': 'Correção Completa',
                        'item': item.descricao_produto,
                        'old_base': getattr(item, 'original_base', None),
                        'new_base': item.base_calculo,
                        'old_aliquota': getattr(item, 'original_aliquota', None),
                        'new_aliquota': item.aliquota,
                        'old_imposto': getattr(item, 'original_imposto', None),
                        'new_imposto': item.valor_imposto
                    })
            
            st.success(f"""
                ✅ Correções aplicadas com sucesso!
                - Total itens: {results['total']}
                - Corrigidos: {results['corrected']}
                - Bases: {results['base_corrected']}
                - Alíquotas: {results['aliquota_corrected']}
                - Impostos: {results['imposto_corrected']}
            """)
            st.rerun()
    
    # Preview
    st.subheader("📋 Preview dos Itens Selecionados")
    preview_data = []
    for item in filtered[:50]:
        has_problem = (item.base_calculo is None or item.base_calculo <= 0) or \
                     (item.aliquota is None or item.aliquota <= 0) or \
                     (item.valor_imposto is None or item.valor_imposto <= 0)
        preview_data.append({
            'Produto': item.descricao_produto[:30] + '...' if len(item.descricao_produto) > 30 else item.descricao_produto,
            'CST': item.cst,
            'CFOP': item.cfop,
            'Base': float(item.base_calculo) if item.base_calculo else 0,
            'Alíquota': float(item.aliquota) if item.aliquota else 0,
            'Imposto': float(item.valor_imposto) if item.valor_imposto else 0,
            'Status': '⚠️' if has_problem else '✅'
        })
    
    if preview_data:
        df_preview = pd.DataFrame(preview_data)
        st.dataframe(df_preview, use_container_width=True)
        st.caption(f"Mostrando {len(preview_data)} de {len(filtered)} itens")

def render_manual_editor():
    """Renderiza editor manual"""
    st.markdown('<h1>✏️ Editor Manual</h1>', unsafe_allow_html=True)
    
    if not st.session_state.get('items'):
        st.info("Nenhum item disponível para edição.")
        return
    
    items = st.session_state.items
    
    # Seleção de item
    item_options = {f"{i.descricao_produto} - CST {i.cst} - CFOP {i.cfop}": i for i in items}
    selected_key = st.selectbox("Selecione o item para editar", sorted(item_options.keys()))
    
    if selected_key:
        item = item_options[selected_key]
        
        st.markdown("### ℹ️ Informações do Item")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.write(f"**Produto:** {item.descricao_produto}")
            st.write(f"**Código:** {item.codigo_produto}")
        with col2:
            st.write(f"**CST:** {item.cst}")
            st.write(f"**CFOP:** {item.cfop}")
        with col3:
            st.write(f"**Block:** {item.block}")
            st.write(f"**Modificado:** {'✅ Sim' if item.modified else '❌ Não'}")
        
        st.markdown("### ✏️ Edição")
        
        col1, col2 = st.columns(2)
        with col1:
            valor_total = st.number_input(
                "Valor Total",
                value=float(item.valor_total) if item.valor_total else 0.0,
                step=0.01,
                format="%.2f"
            )
            
            base_calculo = st.number_input(
                "Base de Cálculo",
                value=float(item.base_calculo) if item.base_calculo else 0.0,
                step=0.01,
                format="%.2f"
            )
        
        with col2:
            aliquota = st.number_input(
                "Alíquota (%)",
                value=float(item.aliquota) if item.aliquota else 0.0,
                step=0.01,
                format="%.2f"
            )
            
            valor_imposto = st.number_input(
                "Valor do Imposto",
                value=float(item.valor_imposto) if item.valor_imposto else 0.0,
                step=0.01,
                format="%.2f"
            )
        
        # Sugestões
        if st.button("💡 Sugerir Correções"):
            suggestions = st.session_state.rules_engine.suggest_correction(item)
            if "error" in suggestions:
                st.warning(suggestions["error"])
            else:
                st.json(suggestions)
        
        # Salvar
        if st.button("💾 Salvar Alterações", type="primary"):
            # Salva valores
            old_values = {
                'valor_total': item.valor_total,
                'base_calculo': item.base_calculo,
                'aliquota': item.aliquota,
                'valor_imposto': item.valor_imposto
            }
            
            item.valor_total = Decimal(str(valor_total))
            item.base_calculo = Decimal(str(base_calculo)) if base_calculo > 0 else None
            item.aliquota = Decimal(str(aliquota)) if aliquota > 0 else None
            item.valor_imposto = Decimal(str(valor_imposto)) if valor_imposto > 0 else None
            item.modified = True
            
            # Registra auditoria
            changes = []
            if old_values['valor_total'] != item.valor_total:
                changes.append(('valor_total', old_values['valor_total'], item.valor_total))
            if old_values['base_calculo'] != item.base_calculo:
                changes.append(('base_calculo', old_values['base_calculo'], item.base_calculo))
            if old_values['aliquota'] != item.aliquota:
                changes.append(('aliquota', old_values['aliquota'], item.aliquota))
            if old_values['valor_imposto'] != item.valor_imposto:
                changes.append(('valor_imposto', old_values['valor_imposto'], item.valor_imposto))
            
            for field, old, new in changes:
                st.session_state.audit_log.append({
                    'timestamp': datetime.now().isoformat(),
                    'user': 'Usuário',
                    'operation': 'Edição Manual',
                    'item': item.descricao_produto,
                    'field': field,
                    'old_value': old,
                    'new_value': new
                })
            
            st.success("✅ Alterações salvas com sucesso!")
            st.rerun()

def render_export():
    """Renderiza seção de exportação"""
    st.markdown('<h1>📥 Exportação</h1>', unsafe_allow_html=True)
    
    if not st.session_state.get('records'):
        st.info("Nenhum dado disponível para exportação.")
        return
    
    export_service = ExportService()
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📄 Exportar SPED")
        if st.button("Exportar SPED Corrigido"):
            with st.spinner("Gerando arquivo SPED..."):
                records = list(st.session_state.records.values()) if isinstance(st.session_state.records, dict) else st.session_state.records
                content = export_service.export_sped(records)
                st.download_button(
                    label="📥 Baixar SPED",
                    data=content,
                    file_name=f"sped_corrigido_{datetime.now().strftime('%Y%m%d')}.txt",
                    mime="text/plain"
                )
    
    with col2:
        st.subheader("📊 Exportar Excel")
        if st.button("Exportar Relatório Excel"):
            with st.spinner("Gerando relatório Excel..."):
                items = st.session_state.get('items', [])
                inconsistencies = st.session_state.get('inconsistencies', [])
                corrections = []
                audit_log = st.session_state.get('audit_log', [])
                
                excel_data = export_service.export_excel(
                    items, inconsistencies, corrections, audit_log
                )
                st.download_button(
                    label="📥 Baixar Excel",
                    data=excel_data,
                    file_name=f"relatorio_sped_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
    st.markdown("---")
    
    st.subheader("📋 CSV de Inconsistências")
    if st.button("Exportar CSV"):
        inconsistencies = st.session_state.get('inconsistencies', [])
        if inconsistencies:
            csv_data = export_service.export_csv_inconsistencies(inconsistencies)
            st.download_button(
                label="📥 Baixar CSV",
                data=csv_data,
                file_name=f"inconsistencias_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
        else:
            st.info("Nenhuma inconsistência para exportar.")

def render_audit():
    """Renderiza log de auditoria"""
    st.markdown('<h1>📜 Log de Auditoria</h1>', unsafe_allow_html=True)
    
    if not st.session_state.get('audit_log'):
        st.info("Nenhuma atividade registrada.")
        return
    
    audit_log = st.session_state.audit_log
    
    # Estatísticas
    st.metric("Total de Registros", len(audit_log))
    
    # Filtros
    col1, col2 = st.columns(2)
    with col1:
        operations = list(set(log.get('operation', '') for log in audit_log))
        selected_op = st.selectbox("Operação", ["Todas"] + sorted([op for op in operations if op]))
    
    with col2:
        search = st.text_input("Buscar", placeholder="Item...")
    
    # Aplica filtros
    filtered = audit_log
    if selected_op != "Todas":
        filtered = [log for log in filtered if log.get('operation') == selected_op]
    if search:
        filtered = [log for log in filtered if search.lower() in log.get('item', '').lower()]
    
    # Exibe logs
    for log in filtered[-50:]:  # Últimos 50
        with st.expander(f"📝 {log.get('timestamp', '')} - {log.get('operation', '')}"):
            st.json(log)

# ============================================================================
# MAIN APP
# ============================================================================

def main():
    """Função principal da aplicação"""
    
    # Configuração da página
    st.set_page_config(
        page_title="SPED Fiscal Analytics",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # CSS personalizado
    st.markdown("""
        <style>
        .stApp {
            background-color: #f8f9fa;
        }
        .css-1d391kg {
            background-color: #1e3c72;
        }
        .stButton > button {
            background-color: #2a5298;
            color: white;
            border-radius: 5px;
            border: none;
            padding: 0.5rem 1rem;
            font-weight: 500;
        }
        .stButton > button:hover {
            background-color: #1e3c72;
            color: white;
        }
        .stSelectbox > div > div {
            border-radius: 5px;
        }
        </style>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.image("https://img.icons8.com/color/96/000000/accounting.png", width=80)
        st.title("SPED Fiscal")
        st.markdown("---")
        
        # Menu
        menu = [
            "📊 Dashboard",
            "📤 Upload",
            "📋 Blocos",
            "📄 Registros",
            "📑 Notas",
            "📦 Itens",
            "⚠️ Inconsistências",
            "🔧 Correções",
            "✏️ Editor",
            "📥 Exportar",
            "📜 Auditoria"
        ]
        
        selected = st.radio("Navegação", menu)
        
        st.markdown("---")
        
        # Status
        if st.session_state.get('items'):
            st.success(f"✅ {len(st.session_state.items)} itens carregados")
        else:
            st.info("📤 Aguardando upload")
    
    # Renderiza seção selecionada
    if selected == "📊 Dashboard":
        render_dashboard()
    elif selected == "📤 Upload":
        render_upload()
    elif selected == "📋 Blocos":
        render_blocks()
    elif selected == "📄 Registros":
        render_records()
    elif selected == "📑 Notas":
        render_invoices()
    elif selected == "📦 Itens":
        render_items()
    elif selected == "⚠️ Inconsistências":
        render_inconsistencies()
    elif selected == "🔧 Correções":
        render_mass_correction()
    elif selected == "✏️ Editor":
        render_manual_editor()
    elif selected == "📥 Exportar":
        render_export()
    elif selected == "📜 Auditoria":
        render_audit()

# ============================================================================
# EXECUÇÃO
# ============================================================================

if __name__ == "__main__":
    main()