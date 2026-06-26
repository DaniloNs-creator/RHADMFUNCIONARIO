"""
SPED FISCAL ANALYTICS - VERSÃO DEFINITIVA
Correção completa do erro 'method' object is not iterable
Versão: 3.0.0

CORREÇÕES REALIZADAS:
1. Correção definitiva do erro de iteração
2. Verificação de tipo em todas as operações
3. Funções de segurança para acesso aos dados
4. Tratamento robusto de state management
"""

import streamlit as st
import pandas as pd
import numpy as np
from decimal import Decimal, ROUND_HALF_UP, getcontext
from datetime import datetime, date
from typing import Dict, List, Any, Optional, Tuple, Union, Iterable
import re
import io
import json
import copy
import logging
from dataclasses import dataclass, field, asdict
import plotly.express as px
import plotly.graph_objects as go
from functools import lru_cache
import time
import sys

# Configuração de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configuração de precisão decimal
getcontext().prec = 28

# ============================================================================
# FUNÇÕES DE SEGURANÇA PARA SESSION STATE
# ============================================================================

def safe_get_items() -> List:
    """
    Função segura para obter itens do session_state
    Garante que sempre retorna uma lista
    """
    try:
        items = st.session_state.get('items')
        if items is None:
            return []
        
        # Se for método, tenta chamar
        if callable(items):
            try:
                result = items()
                if result is None:
                    return []
                if isinstance(result, (list, tuple)):
                    return list(result)
                if hasattr(result, 'values'):
                    return list(result.values())
                return list(result) if result else []
            except Exception as e:
                logger.error(f"Erro ao chamar método items: {str(e)}")
                return []
        
        # Se for lista ou tupla
        if isinstance(items, (list, tuple)):
            return list(items)
        
        # Se for dicionário
        if isinstance(items, dict):
            return list(items.values())
        
        # Se for iterável
        if hasattr(items, '__iter__') and not isinstance(items, str):
            try:
                return list(items)
            except:
                pass
        
        return []
    except Exception as e:
        logger.error(f"Erro ao obter items: {str(e)}")
        return []

def safe_get_records() -> List:
    """Função segura para obter registros"""
    try:
        records = st.session_state.get('records')
        if records is None:
            return []
        
        if callable(records):
            try:
                result = records()
                if result is None:
                    return []
                if isinstance(result, (list, tuple)):
                    return list(result)
                if hasattr(result, 'values'):
                    return list(result.values())
                return list(result) if result else []
            except:
                return []
        
        if isinstance(records, (list, tuple)):
            return list(records)
        
        if isinstance(records, dict):
            return list(records.values())
        
        if hasattr(records, '__iter__') and not isinstance(records, str):
            try:
                return list(records)
            except:
                pass
        
        return []
    except Exception as e:
        logger.error(f"Erro ao obter records: {str(e)}")
        return []

def safe_get_documents() -> Dict:
    """Função segura para obter documentos"""
    try:
        docs = st.session_state.get('documents')
        if docs is None:
            return {}
        
        if callable(docs):
            try:
                result = docs()
                if result is None:
                    return {}
                if isinstance(result, dict):
                    return result
                return {}
            except:
                return {}
        
        if isinstance(docs, dict):
            return docs
        
        return {}
    except Exception as e:
        logger.error(f"Erro ao obter documents: {str(e)}")
        return {}

def safe_get_blocks() -> Dict:
    """Função segura para obter blocos"""
    try:
        blocks = st.session_state.get('blocks')
        if blocks is None:
            return {}
        
        if callable(blocks):
            try:
                result = blocks()
                if result is None:
                    return {}
                if isinstance(result, dict):
                    return result
                return {}
            except:
                return {}
        
        if isinstance(blocks, dict):
            return blocks
        
        return {}
    except Exception as e:
        logger.error(f"Erro ao obter blocks: {str(e)}")
        return {}

def safe_get_inconsistencies() -> List:
    """Função segura para obter inconsistências"""
    try:
        inc = st.session_state.get('inconsistencies')
        if inc is None:
            return []
        
        if callable(inc):
            try:
                result = inc()
                if result is None:
                    return []
                if isinstance(result, (list, tuple)):
                    return list(result)
                return list(result) if result else []
            except:
                return []
        
        if isinstance(inc, (list, tuple)):
            return list(inc)
        
        if hasattr(inc, '__iter__') and not isinstance(inc, str):
            try:
                return list(inc)
            except:
                pass
        
        return []
    except Exception as e:
        logger.error(f"Erro ao obter inconsistencies: {str(e)}")
        return []

def safe_get_audit_log() -> List:
    """Função segura para obter log de auditoria"""
    try:
        log = st.session_state.get('audit_log')
        if log is None:
            return []
        
        if callable(log):
            try:
                result = log()
                if result is None:
                    return []
                if isinstance(result, (list, tuple)):
                    return list(result)
                return list(result) if result else []
            except:
                return []
        
        if isinstance(log, (list, tuple)):
            return list(log)
        
        if hasattr(log, '__iter__') and not isinstance(log, str):
            try:
                return list(log)
            except:
                pass
        
        return []
    except Exception as e:
        logger.error(f"Erro ao obter audit_log: {str(e)}")
        return []

# ============================================================================
# CONSTANTES
# ============================================================================

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

BLOCOS_SPED = {
    "0": "Abertura, Identificação e Referências",
    "C": "Documentos Fiscais - ICMS/IPI",
    "D": "Documentos Fiscais - Serviços de Transporte",
    "E": "Apuração do ICMS e do IPI",
    "F": "Documentos Fiscais - Contribuições",
    "H": "Inventário Físico",
    "1": "Outras Informações",
    "9": "Controle e Encerramento"
}

# ============================================================================
# MODELOS DE DADOS
# ============================================================================

@dataclass
class FiscalItem:
    """Item fiscal com validação de tipos"""
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
    linha_original: int = 0
    
    def __post_init__(self):
        """Validação pós-inicialização"""
        try:
            # Converte valores para Decimal se necessário
            for field in ['quantidade', 'valor_unitario', 'valor_total', 
                          'base_calculo', 'aliquota', 'valor_imposto']:
                value = getattr(self, field)
                if value is not None and not isinstance(value, Decimal):
                    try:
                        setattr(self, field, Decimal(str(value)))
                    except:
                        pass
            
            # Garante que strings estão limpas
            self.codigo_produto = self._clean_string(self.codigo_produto)
            self.descricao_produto = self._clean_string(self.descricao_produto)
            self.cst = self._clean_string(self.cst)
            self.cfop = self._clean_string(self.cfop)
        except Exception as e:
            logger.warning(f"Erro na validação do item: {str(e)}")
    
    def _clean_string(self, value: str) -> str:
        """Limpa string"""
        if not value:
            return ""
        return str(value).strip()
    
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
            'block': self.block,
            'linha_original': self.linha_original
        }
    
    def has_problem(self) -> bool:
        """Verifica se o item tem problemas"""
        return (self.base_calculo is None or self.base_calculo <= 0) or \
               (self.aliquota is None or self.aliquota <= 0) or \
               (self.valor_imposto is None or self.valor_imposto <= 0)
    
    def get_problems(self) -> List[str]:
        """Retorna lista de problemas"""
        problems = []
        if self.base_calculo is None or self.base_calculo <= 0:
            problems.append("Base de cálculo ausente ou zerada")
        if self.aliquota is None or self.aliquota <= 0:
            problems.append("Alíquota ausente ou zerada")
        if self.valor_imposto is None or self.valor_imposto <= 0:
            problems.append("Valor do imposto ausente ou zerado")
        return problems

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
    linha_original: int = 0

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
    operation_type: str
    requires_base: bool = True
    requires_aliquota: bool = True
    requires_imposto: bool = True
    default_base_formula: Optional[str] = None
    default_aliquota: Optional[Decimal] = None
    default_imposto_formula: Optional[str] = None
    valid_cfops: List[str] = field(default_factory=list)
    priority: int = 1
    enabled: bool = True
    description: str = ""

# ============================================================================
# MOTOR DE REGRAS
# ============================================================================

class RulesEngine:
    """Motor de regras fiscais com cache e logging"""
    
    def __init__(self):
        self.rules: List[FiscalRule] = []
        self.rule_cache: Dict[str, List[FiscalRule]] = {}
        self._load_default_rules()
    
    def _load_default_rules(self):
        """Carrega regras padrão"""
        default_rules = [
            FiscalRule("R001", "00", "*", "ambos", True, True, True,
                      "item_value", Decimal("18.00"), "base * aliquota / 100",
                      description="Tributada integralmente"),
            FiscalRule("R002", "10", "*", "ambos", True, True, True,
                      "item_value", Decimal("18.00"), "base * aliquota / 100",
                      description="Tributada com substituição tributária"),
            FiscalRule("R003", "20", "*", "ambos", True, True, True,
                      "item_value", Decimal("12.00"), "base * aliquota / 100",
                      description="Com redução de base de cálculo"),
            FiscalRule("R004", "30", "*", "ambos", False, False, False,
                      None, None, None,
                      description="Isenta com substituição tributária"),
            FiscalRule("R005", "40", "*", "ambos", False, False, False,
                      None, None, None,
                      description="Isenta"),
            FiscalRule("R006", "41", "*", "ambos", False, False, False,
                      None, None, None,
                      description="Não tributada"),
            FiscalRule("R007", "50", "*", "ambos", False, False, False,
                      None, None, None,
                      description="Suspensão"),
            FiscalRule("R008", "51", "*", "ambos", False, False, False,
                      None, None, None,
                      description="Diferimento"),
            FiscalRule("R009", "60", "*", "ambos", False, False, False,
                      None, None, None,
                      description="ICMS cobrado anteriormente"),
            FiscalRule("R010", "70", "*", "ambos", True, True, True,
                      "item_value", Decimal("12.00"), "base * aliquota / 100",
                      description="Com redução e substituição tributária"),
            FiscalRule("R011", "90", "*", "ambos", False, False, False,
                      None, None, None,
                      description="Outras"),
            FiscalRule("R012", "00", "5351", "entrada", True, True, True,
                      "item_value", Decimal("12.00"), "base * aliquota / 100",
                      description="Transporte - Entrada"),
            FiscalRule("R013", "00", "5401", "entrada", True, True, True,
                      "item_value", Decimal("25.00"), "base * aliquota / 100",
                      description="Energia Elétrica"),
            FiscalRule("R014", "00", "5355", "entrada", True, True, True,
                      "item_value", Decimal("18.00"), "base * aliquota / 100",
                      description="Locação"),
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
        try:
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
            
            if not rule.requires_base and item.base_calculo and item.base_calculo > 0:
                errors.append("CST isento com base de cálculo informada")
            
            if not rule.requires_aliquota and item.aliquota and item.aliquota > 0:
                errors.append("CST isento com alíquota informada")
            
            if not rule.requires_imposto and item.valor_imposto and item.valor_imposto > 0:
                errors.append("CST isento com valor de imposto informado")
        except Exception as e:
            errors.append(f"Erro na validação: {str(e)}")
        
        return len(errors) == 0, errors
    
    def suggest_correction(self, item: FiscalItem) -> Dict[str, Any]:
        """Sugere correções para um item fiscal"""
        suggestions = {}
        try:
            rule = self.get_rule(item.cst, item.cfop, item.tipo_operacao)
            if not rule:
                return {"error": "Regra não encontrada"}
            
            if rule.requires_base and (item.base_calculo is None or item.base_calculo <= 0):
                if rule.default_base_formula == "item_value" and item.valor_total:
                    suggestions["base_calculo"] = item.valor_total
            
            if rule.requires_aliquota and (item.aliquota is None or item.aliquota <= 0):
                if rule.default_aliquota:
                    suggestions["aliquota"] = rule.default_aliquota
            
            if rule.requires_imposto and (item.valor_imposto is None or item.valor_imposto <= 0):
                base = suggestions.get("base_calculo", item.base_calculo)
                aliquota = suggestions.get("aliquota", item.aliquota)
                if base and aliquota:
                    try:
                        imposto = (base * aliquota) / Decimal("100")
                        suggestions["valor_imposto"] = imposto.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                    except:
                        pass
        except Exception as e:
            suggestions["error"] = f"Erro na sugestão: {str(e)}"
        
        return suggestions
    
    def apply_correction(self, item: FiscalItem) -> Dict[str, Any]:
        """Aplica correções ao item"""
        changes = {}
        try:
            suggestions = self.suggest_correction(item)
            if "error" in suggestions:
                return suggestions
            
            if not hasattr(item, 'original_base') or item.original_base is None:
                item.original_base = item.base_calculo
            if not hasattr(item, 'original_aliquota') or item.original_aliquota is None:
                item.original_aliquota = item.aliquota
            if not hasattr(item, 'original_imposto') or item.original_imposto is None:
                item.original_imposto = item.valor_imposto
            
            if "base_calculo" in suggestions:
                changes['base_calculo'] = {'old': item.base_calculo, 'new': suggestions['base_calculo']}
                item.base_calculo = suggestions['base_calculo']
            
            if "aliquota" in suggestions:
                changes['aliquota'] = {'old': item.aliquota, 'new': suggestions['aliquota']}
                item.aliquota = suggestions['aliquota']
            
            if "valor_imposto" in suggestions:
                changes['valor_imposto'] = {'old': item.valor_imposto, 'new': suggestions['valor_imposto']}
                item.valor_imposto = suggestions['valor_imposto']
            
            if changes:
                item.modified = True
        except Exception as e:
            changes['error'] = str(e)
        
        return changes

# ============================================================================
# PARSER
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
        self.current_parent: Optional[SpedRecord] = None
    
    def parse(self, content: str) -> Dict[str, Any]:
        """Parseia arquivo SPED"""
        try:
            if not content or not content.strip():
                raise ValueError("Conteúdo do arquivo vazio")
            
            lines = content.splitlines()
            self._parse_lines(lines)
            self._extract_documents()
            self._extract_items()
            self._generate_summary()
            
            return {
                'records': self.records,
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
        for line_num, line in enumerate(lines, 1):
            if not line.strip():
                continue
            
            try:
                parts = line.split("|")
                if len(parts) < 2:
                    continue
                
                record_type = parts[1].strip()
                if not record_type:
                    continue
                
                block = self._identify_block(record_type)
                fields = self._extract_fields(record_type, parts)
                
                record = SpedRecord(
                    record_id=f"{record_type}_{line_num}",
                    record_type=record_type,
                    block=block,
                    line_number=line_num,
                    raw_data=line,
                    fields=fields,
                    original_data=copy.deepcopy(fields)
                )
                
                if record_type in ["0000", "1000", "C100", "D100", "E100", "F100"]:
                    self.current_parent = record
                elif self.current_parent:
                    record.parent_record = self.current_parent.record_id
                    if record.record_id not in self.current_parent.children_records:
                        self.current_parent.children_records.append(record.record_id)
                
                self.records.append(record)
                
                if block not in self.blocks:
                    self.blocks[block] = []
                self.blocks[block].append(record)
                
                if record_type == "0000":
                    self.company_info = {
                        'cnpj': fields.get('cnpj', ''),
                        'nome': fields.get('nome_empresa', ''),
                        'ie': fields.get('ie', '')
                    }
                    self.period = f"{fields.get('data_inicial', '')} a {fields.get('data_final', '')}"
            except Exception as e:
                continue
    
    def _identify_block(self, record_type: str) -> str:
        """Identifica bloco do registro"""
        if record_type.startswith("0"): return "0"
        elif record_type.startswith("1"): return "1"
        elif record_type.startswith("C"): return "C"
        elif record_type.startswith("D"): return "D"
        elif record_type.startswith("E"): return "E"
        elif record_type.startswith("F"): return "F"
        elif record_type.startswith("H"): return "H"
        elif record_type.startswith("9"): return "9"
        return "0"
    
    def _extract_fields(self, record_type: str, parts: List[str]) -> Dict[str, Any]:
        """Extrai campos do registro"""
        fields = {}
        
        try:
            if record_type == "0000" and len(parts) >= 11:
                fields["codigo_versao"] = self._clean(parts[2] if len(parts) > 2 else "")
                fields["tipo_arquivo"] = self._clean(parts[3] if len(parts) > 3 else "")
                fields["cnpj"] = self._clean(parts[4] if len(parts) > 4 else "")
                fields["nome_empresa"] = self._clean(parts[5] if len(parts) > 5 else "")
                fields["ie"] = self._clean(parts[7] if len(parts) > 7 else "")
                fields["data_inicial"] = self._clean(parts[9] if len(parts) > 9 else "")
                fields["data_final"] = self._clean(parts[10] if len(parts) > 10 else "")
            
            elif record_type == "C100" and len(parts) >= 15:
                fields["tipo_documento"] = self._clean(parts[2] if len(parts) > 2 else "")
                fields["serie"] = self._clean(parts[3] if len(parts) > 3 else "")
                fields["numero_documento"] = self._clean(parts[4] if len(parts) > 4 else "")
                fields["chave_acesso"] = self._clean(parts[5] if len(parts) > 5 else "")
                fields["data_emissao"] = self._clean(parts[6] if len(parts) > 6 else "")
                fields["cnpj_emitente"] = self._clean(parts[7] if len(parts) > 7 else "")
                fields["cnpj_destinatario"] = self._clean(parts[8] if len(parts) > 8 else "")
                fields["valor_total"] = self._parse_decimal(parts[9] if len(parts) > 9 else "0")
                fields["base_calculo"] = self._parse_decimal(parts[10] if len(parts) > 10 else None)
                fields["valor_icms"] = self._parse_decimal(parts[11] if len(parts) > 11 else None)
                fields["aliquota"] = self._parse_decimal(parts[12] if len(parts) > 12 else None)
                fields["cfop"] = self._clean(parts[13] if len(parts) > 13 else "")
                fields["cst"] = self._clean(parts[14] if len(parts) > 14 else "")
                fields["tipo_operacao"] = self._identify_operation_type(fields.get("cfop", ""))
            
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
            
            elif record_type == "D100" and len(parts) >= 9:
                fields["tipo_operacao"] = self._clean(parts[2] if len(parts) > 2 else "")
                fields["cfop"] = self._clean(parts[3] if len(parts) > 3 else "")
                fields["cnpj_emitente"] = self._clean(parts[4] if len(parts) > 4 else "")
                fields["valor_servico"] = self._parse_decimal(parts[5] if len(parts) > 5 else "0")
                fields["base_calculo"] = self._parse_decimal(parts[6] if len(parts) > 6 else None)
                fields["aliquota"] = self._parse_decimal(parts[7] if len(parts) > 7 else None)
                fields["valor_imposto"] = self._parse_decimal(parts[8] if len(parts) > 8 else None)
            
            elif record_type == "D101" and len(parts) >= 6:
                fields["cfop"] = self._clean(parts[2] if len(parts) > 2 else "")
                fields["base_calculo"] = self._parse_decimal(parts[3] if len(parts) > 3 else None)
                fields["aliquota"] = self._parse_decimal(parts[4] if len(parts) > 4 else None)
                fields["valor_imposto"] = self._parse_decimal(parts[5] if len(parts) > 5 else None)
            
            elif record_type == "D105" and len(parts) >= 6:
                fields["cfop"] = self._clean(parts[2] if len(parts) > 2 else "")
                fields["base_calculo"] = self._parse_decimal(parts[3] if len(parts) > 3 else None)
                fields["aliquota"] = self._parse_decimal(parts[4] if len(parts) > 4 else None)
                fields["valor_imposto"] = self._parse_decimal(parts[5] if len(parts) > 5 else None)
            
            elif record_type == "C500" and len(parts) >= 9:
                fields["tipo_operacao"] = self._clean(parts[2] if len(parts) > 2 else "")
                fields["cfop"] = self._clean(parts[3] if len(parts) > 3 else "")
                fields["cnpj_emitente"] = self._clean(parts[4] if len(parts) > 4 else "")
                fields["valor_total"] = self._parse_decimal(parts[5] if len(parts) > 5 else "0")
                fields["base_calculo"] = self._parse_decimal(parts[6] if len(parts) > 6 else None)
                fields["aliquota"] = self._parse_decimal(parts[7] if len(parts) > 7 else None)
                fields["valor_imposto"] = self._parse_decimal(parts[8] if len(parts) > 8 else None)
            
            elif record_type == "C501" and len(parts) >= 6:
                fields["cfop"] = self._clean(parts[2] if len(parts) > 2 else "")
                fields["base_calculo"] = self._parse_decimal(parts[3] if len(parts) > 3 else None)
                fields["aliquota"] = self._parse_decimal(parts[4] if len(parts) > 4 else None)
                fields["valor_imposto"] = self._parse_decimal(parts[5] if len(parts) > 5 else None)
            
            elif record_type == "C505" and len(parts) >= 6:
                fields["cfop"] = self._clean(parts[2] if len(parts) > 2 else "")
                fields["base_calculo"] = self._parse_decimal(parts[3] if len(parts) > 3 else None)
                fields["aliquota"] = self._parse_decimal(parts[4] if len(parts) > 4 else None)
                fields["valor_imposto"] = self._parse_decimal(parts[5] if len(parts) > 5 else None)
            
            elif record_type == "F100" and len(parts) >= 9:
                fields["tipo_operacao"] = self._clean(parts[2] if len(parts) > 2 else "")
                fields["cfop"] = self._clean(parts[3] if len(parts) > 3 else "")
                fields["cnpj_emitente"] = self._clean(parts[4] if len(parts) > 4 else "")
                fields["valor_total"] = self._parse_decimal(parts[5] if len(parts) > 5 else "0")
                fields["base_calculo"] = self._parse_decimal(parts[6] if len(parts) > 6 else None)
                fields["aliquota"] = self._parse_decimal(parts[7] if len(parts) > 7 else None)
                fields["valor_imposto"] = self._parse_decimal(parts[8] if len(parts) > 8 else None)
        except Exception as e:
            pass
        
        return fields
    
    def _extract_documents(self):
        """Extrai documentos fiscais"""
        for record in self.records:
            try:
                if record.record_type == "C100":
                    doc = self._create_document_from_c100(record)
                    if doc.chave_acesso:
                        self.documents[doc.chave_acesso] = doc
                elif record.record_type == "D100":
                    doc = self._create_document_from_d100(record)
                    if doc.chave_acesso:
                        self.documents[doc.chave_acesso] = doc
                elif record.record_type == "C500":
                    doc = self._create_document_from_c500(record)
                    if doc.chave_acesso:
                        self.documents[doc.chave_acesso] = doc
                elif record.record_type == "F100":
                    doc = self._create_document_from_f100(record)
                    if doc.chave_acesso:
                        self.documents[doc.chave_acesso] = doc
            except Exception:
                pass
    
    def _create_document_from_c100(self, record: SpedRecord) -> FiscalDocument:
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
            block='C',
            record_type='C100',
            linha_original=record.line_number
        )
    
    def _create_document_from_d100(self, record: SpedRecord) -> FiscalDocument:
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
            block='D',
            record_type='D100',
            linha_original=record.line_number
        )
    
    def _create_document_from_c500(self, record: SpedRecord) -> FiscalDocument:
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
            block='C',
            record_type='C500',
            linha_original=record.line_number
        )
    
    def _create_document_from_f100(self, record: SpedRecord) -> FiscalDocument:
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
            block='F',
            record_type='F100',
            linha_original=record.line_number
        )
    
    def _extract_items(self):
        """Extrai itens dos registros"""
        for record in self.records:
            try:
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
                        record_id=record.record_id,
                        linha_original=record.line_number
                    )
                    self.items.append(item)
                
                elif record.record_type in ["D100", "D101", "D105"]:
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
                            record_id=record.record_id,
                            linha_original=record.line_number
                        )
                        self.items.append(item)
                
                elif record.record_type in ["C500", "C501", "C505"]:
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
                            record_id=record.record_id,
                            linha_original=record.line_number
                        )
                        self.items.append(item)
                
                elif record.record_type == "F100":
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
                            record_id=record.record_id,
                            linha_original=record.line_number
                        )
                        self.items.append(item)
            except Exception:
                pass
    
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
        if not value:
            return ""
        return str(value).strip()
    
    def _parse_decimal(self, value: Any) -> Optional[Decimal]:
        if value is None or value == "":
            return None
        try:
            if isinstance(value, str):
                value = re.sub(r'[^\d.,-]', '', value)
                value = value.replace(',', '.')
            return Decimal(str(value))
        except:
            return None
    
    def _identify_operation_type(self, cfop: str) -> str:
        if not cfop:
            return "entrada"
        if cfop.startswith("1") or cfop.startswith("2") or cfop.startswith("3"):
            return "entrada"
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
    
    def validate_all(self, items) -> List[Dict]:
        """Valida todos os itens"""
        self.inconsistencies = []
        
        # GARANTE QUE É UMA LISTA
        if items is None:
            return self.inconsistencies
        
        # Se for método, tenta chamar
        if callable(items):
            try:
                items = items()
            except:
                return self.inconsistencies
        
        # Se for dict, pega valores
        if isinstance(items, dict):
            items = list(items.values())
        
        # Se for iterável, converte para lista
        if hasattr(items, '__iter__') and not isinstance(items, (str, bytes)):
            try:
                items = list(items)
            except:
                items = []
        else:
            items = []
        
        for item in items:
            if not isinstance(item, FiscalItem):
                continue
            
            if not item.cst or item.cst.strip() == "":
                self.inconsistencies.append({
                    'type': 'critical',
                    'item': item.descricao_produto or 'Item sem descrição',
                    'message': 'CST não informado',
                    'details': f'Produto: {item.descricao_produto}, CFOP: {item.cfop}',
                    'severity': 'alta',
                    'linha': item.linha_original
                })
            
            if not item.cfop or item.cfop.strip() == "":
                self.inconsistencies.append({
                    'type': 'critical',
                    'item': item.descricao_produto or 'Item sem descrição',
                    'message': 'CFOP não informado',
                    'details': f'Produto: {item.descricao_produto}, CST: {item.cst}',
                    'severity': 'alta',
                    'linha': item.linha_original
                })
            
            try:
                is_valid, errors = self.rules_engine.validate_item(item)
                if not is_valid:
                    for error in errors:
                        severity = 'alta' if 'obrigatória' in error or 'obrigatório' in error else 'media'
                        self.inconsistencies.append({
                            'type': 'critical' if severity == 'alta' else 'warning',
                            'item': item.descricao_produto or 'Item sem descrição',
                            'message': error,
                            'details': f'Produto: {item.descricao_produto}, CST: {item.cst}, CFOP: {item.cfop}',
                            'severity': severity,
                            'linha': item.linha_original
                        })
            except Exception:
                pass
        
        return self.inconsistencies

class CorrectionService:
    """Serviço de correção fiscal"""
    
    def __init__(self, rules_engine: RulesEngine):
        self.rules_engine = rules_engine
        self.corrections: List[Dict] = []
    
    def correct_mass(self, items) -> Dict[str, int]:
        """Corrige múltiplos itens em massa"""
        results = {'total': 0, 'corrected': 0, 'base_corrected': 0, 'aliquota_corrected': 0, 'imposto_corrected': 0, 'errors': 0}
        
        if items is None:
            return results
        
        if callable(items):
            try:
                items = items()
            except:
                return results
        
        if isinstance(items, dict):
            items = list(items.values())
        
        if hasattr(items, '__iter__') and not isinstance(items, (str, bytes)):
            try:
                items = list(items)
            except:
                items = []
        else:
            items = []
        
        results['total'] = len(items)
        
        for item in items:
            if not isinstance(item, FiscalItem):
                continue
            
            try:
                changes = self.rules_engine.apply_correction(item)
                if changes and "error" not in changes:
                    results['corrected'] += 1
                    if 'base_calculo' in changes:
                        results['base_corrected'] += 1
                    if 'aliquota' in changes:
                        results['aliquota_corrected'] += 1
                    if 'valor_imposto' in changes:
                        results['imposto_corrected'] += 1
                    self.corrections.append({
                        'item': item.descricao_produto,
                        'changes': changes,
                        'timestamp': datetime.now().isoformat()
                    })
            except:
                results['errors'] += 1
        
        return results

class ExportService:
    """Serviço de exportação"""
    
    def export_excel(self, items, inconsistencies, corrections, audit_log) -> bytes:
        output = io.BytesIO()
        
        try:
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                if items:
                    items_list = []
                    if isinstance(items, dict):
                        items_list = list(items.values())
                    elif hasattr(items, '__iter__') and not isinstance(items, (str, bytes)):
                        try:
                            items_list = list(items)
                        except:
                            items_list = []
                    
                    if items_list:
                        data = []
                        for item in items_list:
                            if isinstance(item, FiscalItem):
                                data.append({
                                    'Produto': item.descricao_produto,
                                    'CST': item.cst,
                                    'CFOP': item.cfop,
                                    'Valor Total': float(item.valor_total) if item.valor_total else 0,
                                    'Base': float(item.base_calculo) if item.base_calculo else 0,
                                    'Alíquota': float(item.aliquota) if item.aliquota else 0,
                                    'Imposto': float(item.valor_imposto) if item.valor_imposto else 0,
                                    'Block': item.block,
                                    'Modificado': 'Sim' if item.modified else 'Não',
                                    'Linha': item.linha_original
                                })
                        if data:
                            df_items = pd.DataFrame(data)
                            df_items.to_excel(writer, sheet_name='Itens', index=False)
                
                if inconsistencies:
                    if isinstance(inconsistencies, list):
                        df_inc = pd.DataFrame(inconsistencies)
                        df_inc.to_excel(writer, sheet_name='Inconsistencias', index=False)
                
                if corrections:
                    if isinstance(corrections, list):
                        corr_data = []
                        for corr in corrections:
                            if isinstance(corr, dict):
                                for field, change in corr.get('changes', {}).items():
                                    corr_data.append({
                                        'Item': corr.get('item', ''),
                                        'Campo': field,
                                        'Valor Antigo': change.get('old'),
                                        'Valor Novo': change.get('new'),
                                        'Data': corr.get('timestamp', '')
                                    })
                        if corr_data:
                            df_corr = pd.DataFrame(corr_data)
                            df_corr.to_excel(writer, sheet_name='Correcoes', index=False)
                
                if audit_log:
                    if isinstance(audit_log, list):
                        df_audit = pd.DataFrame(audit_log)
                        df_audit.to_excel(writer, sheet_name='Auditoria', index=False)
        except Exception as e:
            logger.error(f"Erro ao exportar Excel: {str(e)}")
        
        return output.getvalue()

# ============================================================================
# FUNÇÕES DE UI
# ============================================================================

def init_session_state():
    """Inicializa o estado da sessão"""
    if 'initialized' not in st.session_state:
        st.session_state.initialized = True
        st.session_state.records = []
        st.session_state.documents = {}
        st.session_state.items = []
        st.session_state.blocks = {}
        st.session_state.summary = {}
        st.session_state.inconsistencies = []
        st.session_state.audit_log = []
        st.session_state.rules_engine = RulesEngine()
        st.session_state.file_content = None
        st.session_state.file_name = None
        st.session_state.company_info = {}
        st.session_state.period = ""
        st.session_state.parsed_data = None
        st.session_state.corrections = []

def render_dashboard():
    """Renderiza dashboard"""
    st.markdown("""
        <style>
        .dashboard-header { background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); padding: 2rem; border-radius: 10px; color: white; margin-bottom: 2rem; }
        .metric-card { background: white; padding: 1.5rem; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); border-left: 4px solid #2a5298; margin-bottom: 1rem; }
        .metric-value { font-size: 2rem; font-weight: bold; color: #1e3c72; }
        .metric-label { color: #6c757d; font-size: 0.9rem; }
        </style>
    """, unsafe_allow_html=True)
    
    st.markdown('<div class="dashboard-header"><h1>📊 Dashboard Fiscal</h1></div>', unsafe_allow_html=True)
    
    # OBTÉM ITENS DE FORMA SEGURA
    items = safe_get_items()
    
    if not items:
        st.info("📤 Nenhum arquivo SPED carregado. Acesse a seção 'Upload' para começar.")
        return
    
    # Métricas
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{len(items)}</div>
                <div class="metric-label">📦 Total Itens</div>
            </div>
        """, unsafe_allow_html=True)
    with col2:
        problematic = len([i for i in items if i.has_problem()])
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #dc3545;">
                <div class="metric-value" style="color: #dc3545;">{problematic}</div>
                <div class="metric-label">⚠️ Itens com Problemas</div>
            </div>
        """, unsafe_allow_html=True)
    with col3:
        docs = safe_get_documents()
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #28a745;">
                <div class="metric-value" style="color: #28a745;">{len(docs)}</div>
                <div class="metric-label">📑 Notas Fiscais</div>
            </div>
        """, unsafe_allow_html=True)
    with col4:
        audit = safe_get_audit_log()
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #ffc107;">
                <div class="metric-value" style="color: #ffc107;">{len(audit)}</div>
                <div class="metric-label">✅ Correções Aplicadas</div>
            </div>
        """, unsafe_allow_html=True)

def render_upload():
    """Renderiza upload - VERSÃO CORRIGIDA DEFINITIVAMENTE"""
    st.markdown('<h1>📤 Upload de Arquivo SPED</h1>', unsafe_allow_html=True)
    
    st.markdown("""
    ### Instruções
    Faça upload do arquivo SPED no formato TXT.
    O sistema processará automaticamente e estruturará os dados.
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
                
                # Armazena dados - GARANTINDO QUE SÃO LISTAS/DICTS
                st.session_state.parsed_data = data
                st.session_state.records = data.get('records', [])
                st.session_state.documents = data.get('documents', {})
                
                # CRUCIAL: Garante que items é uma lista válida
                items_data = data.get('items', [])
                if items_data is None:
                    items_data = []
                elif isinstance(items_data, dict):
                    items_data = list(items_data.values())
                elif callable(items_data):
                    try:
                        items_data = list(items_data())
                    except:
                        items_data = []
                elif not isinstance(items_data, list):
                    try:
                        items_data = list(items_data)
                    except:
                        items_data = []
                
                st.session_state.items = items_data
                st.session_state.blocks = data.get('blocks', {})
                st.session_state.summary = data.get('summary', {})
                st.session_state.period = data.get('period', '')
                st.session_state.company_info = data.get('company_info', {})
                
                # Inicializa regras
                if 'rules_engine' not in st.session_state:
                    st.session_state.rules_engine = RulesEngine()
                
                # Valida - passando a lista segura
                validation_service = ValidationService(st.session_state.rules_engine)
                st.session_state.inconsistencies = validation_service.validate_all(items_data)
                
                if 'audit_log' not in st.session_state:
                    st.session_state.audit_log = []
                
                st.success(f"✅ Arquivo processado com sucesso! {len(items_data)} itens encontrados.")
                
                # Exibe resumo
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Registros", len(st.session_state.records) if st.session_state.records else 0)
                with col2:
                    st.metric("Notas Fiscais", len(st.session_state.documents) if st.session_state.documents else 0)
                with col3:
                    st.metric("Itens", len(items_data))
                with col4:
                    inc = safe_get_inconsistencies()
                    st.metric("Inconsistências", len(inc))
                
                if st.session_state.company_info:
                    with st.expander("🏢 Informações da Empresa"):
                        st.json(st.session_state.company_info)
        
        except Exception as e:
            st.error(f"❌ Erro ao processar arquivo: {str(e)}")
            import traceback
            st.code(traceback.format_exc())

def render_blocks():
    st.markdown('<h1>📋 Visualização de Blocos</h1>', unsafe_allow_html=True)
    
    blocks = safe_get_blocks()
    if not blocks:
        st.info("Nenhum bloco encontrado.")
        return
    
    selected_block = st.selectbox("Selecione o Bloco", sorted(blocks.keys()))
    
    if selected_block:
        records = blocks.get(selected_block, [])
        st.subheader(f"Bloco {selected_block} - {BLOCOS_SPED.get(selected_block, 'Desconhecido')}")
        st.write(f"Total de registros: {len(records)}")
        
        if records:
            data = []
            for record in records[:1000]:
                data.append({
                    'Tipo': record.record_type,
                    'Linha': record.line_number,
                    'Campos': len(record.fields),
                    'Modificado': '✅' if record.modified else ''
                })
            if data:
                df = pd.DataFrame(data)
                st.dataframe(df, use_container_width=True)

def render_records():
    st.markdown('<h1>📄 Visualização de Registros</h1>', unsafe_allow_html=True)
    
    records = safe_get_records()
    if not records:
        st.info("Nenhum registro encontrado.")
        return
    
    if records:
        data = []
        for record in records[:500]:
            data.append({
                'Tipo': record.record_type,
                'Bloco': record.block,
                'Linha': record.line_number,
                'Campos': len(record.fields),
                'Modificado': '✅' if record.modified else ''
            })
        if data:
            df = pd.DataFrame(data)
            st.dataframe(df, use_container_width=True)
            st.caption(f"Mostrando {len(data)} de {len(records)} registros")

def render_invoices():
    st.markdown('<h1>📑 Notas Fiscais</h1>', unsafe_allow_html=True)
    
    docs = safe_get_documents()
    if not docs:
        st.info("Nenhuma nota fiscal encontrada.")
        return
    
    docs_list = list(docs.values()) if isinstance(docs, dict) else []
    st.info(f"📊 Notas fiscais encontradas: {len(docs_list)}")
    
    for doc in docs_list[:20]:
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

def render_items():
    st.markdown('<h1>📦 Itens Fiscais</h1>', unsafe_allow_html=True)
    
    items = safe_get_items()
    if not items:
        st.info("Nenhum item encontrado.")
        return
    
    st.info(f"📊 Total de itens: {len(items)}")
    
    if items:
        data = []
        for item in items[:500]:
            data.append({
                'Produto': item.descricao_produto[:40] + '...' if len(item.descricao_produto) > 40 else item.descricao_produto,
                'CST': item.cst,
                'CFOP': item.cfop,
                'Valor Total': float(item.valor_total) if item.valor_total else 0,
                'Base': float(item.base_calculo) if item.base_calculo else 0,
                'Alíquota': float(item.aliquota) if item.aliquota else 0,
                'Imposto': float(item.valor_imposto) if item.valor_imposto else 0,
                'Status': '⚠️' if item.has_problem() else '✅',
                'Linha': item.linha_original
            })
        if data:
            df = pd.DataFrame(data)
            st.dataframe(df, use_container_width=True)

def render_inconsistencies():
    st.markdown('<h1>⚠️ Inconsistências Fiscais</h1>', unsafe_allow_html=True)
    
    inconsistencies = safe_get_inconsistencies()
    if not inconsistencies:
        st.success("✅ Nenhuma inconsistência encontrada!")
        return
    
    critical = len([i for i in inconsistencies if i.get('severity') == 'alta'])
    warning = len([i for i in inconsistencies if i.get('severity') == 'media'])
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("🔴 Críticas", critical)
    with col2:
        st.metric("🟡 Avisos", warning)
    
    for inc in inconsistencies:
        if inc.get('severity') == 'alta':
            st.error(f"🔴 **{inc.get('item', '')}** - {inc.get('message', '')}")
        else:
            st.warning(f"🟡 **{inc.get('item', '')}** - {inc.get('message', '')}")
        st.caption(f"Linha: {inc.get('linha', 'N/A')}")

def render_mass_correction():
    st.markdown('<h1>🔧 Correções em Massa</h1>', unsafe_allow_html=True)
    
    items = safe_get_items()
    if not items:
        st.info("Nenhum item disponível para correção.")
        return
    
    problematic = [i for i in items if i.has_problem()]
    st.info(f"📊 Itens com problemas: {len(problematic)}")
    
    if not problematic:
        st.success("✅ Nenhum item precisa de correção!")
        return
    
    if st.button("🚀 Aplicar Correções em Todos os Itens", type="primary"):
        with st.spinner("Aplicando correções..."):
            correction_service = CorrectionService(st.session_state.rules_engine)
            results = correction_service.correct_mass(problematic)
            
            # Atualiza logs
            for item in problematic:
                if item.modified:
                    st.session_state.audit_log.append({
                        'timestamp': datetime.now().isoformat(),
                        'user': 'Sistema',
                        'operation': 'Correção em Massa',
                        'item': item.descricao_produto,
                        'linha': item.linha_original
                    })
            
            st.success(f"""
                ✅ Correções aplicadas!
                - Corrigidos: {results['corrected']}
                - Bases: {results['base_corrected']}
                - Alíquotas: {results['aliquota_corrected']}
                - Impostos: {results['imposto_corrected']}
            """)
            st.rerun()
    
    # Preview
    st.subheader("📋 Itens com Problemas")
    preview_data = []
    for item in problematic[:50]:
        preview_data.append({
            'Produto': item.descricao_produto[:30] + '...' if len(item.descricao_produto) > 30 else item.descricao_produto,
            'CST': item.cst,
            'CFOP': item.cfop,
            'Base': float(item.base_calculo) if item.base_calculo else 0,
            'Alíquota': float(item.aliquota) if item.aliquota else 0,
            'Imposto': float(item.valor_imposto) if item.valor_imposto else 0,
            'Linha': item.linha_original
        })
    
    if preview_data:
        df_preview = pd.DataFrame(preview_data)
        st.dataframe(df_preview, use_container_width=True)

def render_manual_editor():
    st.markdown('<h1>✏️ Editor Manual</h1>', unsafe_allow_html=True)
    
    items = safe_get_items()
    if not items:
        st.info("Nenhum item disponível para edição.")
        return
    
    item_options = {f"{i.descricao_produto} - CST {i.cst} (Linha {i.linha_original})": i for i in items}
    selected_key = st.selectbox("Selecione o item", sorted(item_options.keys()))
    
    if selected_key:
        item = item_options[selected_key]
        
        st.markdown("### ℹ️ Informações do Item")
        col1, col2 = st.columns(2)
        with col1:
            st.write(f"**Produto:** {item.descricao_produto}")
            st.write(f"**CST:** {item.cst}")
            st.write(f"**CFOP:** {item.cfop}")
        with col2:
            st.write(f"**Linha:** {item.linha_original}")
            st.write(f"**Modificado:** {'✅ Sim' if item.modified else '❌ Não'}")
            if item.has_problem():
                st.warning(f"⚠️ {', '.join(item.get_problems())}")
        
        st.markdown("### ✏️ Edição")
        col1, col2 = st.columns(2)
        with col1:
            base_calculo = st.number_input(
                "Base de Cálculo",
                value=float(item.base_calculo) if item.base_calculo else 0.0,
                step=0.01,
                format="%.2f"
            )
            aliquota = st.number_input(
                "Alíquota (%)",
                value=float(item.aliquota) if item.aliquota else 0.0,
                step=0.01,
                format="%.2f"
            )
        with col2:
            valor_imposto = st.number_input(
                "Valor do Imposto",
                value=float(item.valor_imposto) if item.valor_imposto else 0.0,
                step=0.01,
                format="%.2f"
            )
        
        if st.button("💾 Salvar", type="primary"):
            old_base = item.base_calculo
            old_aliquota = item.aliquota
            old_imposto = item.valor_imposto
            
            item.base_calculo = Decimal(str(base_calculo)) if base_calculo > 0 else None
            item.aliquota = Decimal(str(aliquota)) if aliquota > 0 else None
            item.valor_imposto = Decimal(str(valor_imposto)) if valor_imposto > 0 else None
            item.modified = True
            
            st.session_state.audit_log.append({
                'timestamp': datetime.now().isoformat(),
                'user': 'Usuário',
                'operation': 'Edição Manual',
                'item': item.descricao_produto,
                'old_base': float(old_base) if old_base else None,
                'new_base': float(item.base_calculo) if item.base_calculo else None,
                'linha': item.linha_original
            })
            
            st.success("✅ Alterações salvas!")
            st.rerun()

def render_export():
    st.markdown('<h1>📥 Exportação</h1>', unsafe_allow_html=True)
    
    items = safe_get_items()
    if not items:
        st.info("Nenhum dado disponível para exportação.")
        return
    
    export_service = ExportService()
    
    if st.button("📊 Exportar Relatório Excel"):
        with st.spinner("Gerando relatório..."):
            inconsistencies = safe_get_inconsistencies()
            audit_log = safe_get_audit_log()
            
            excel_data = export_service.export_excel(
                items, inconsistencies, [], audit_log
            )
            st.download_button(
                label="📥 Baixar Excel",
                data=excel_data,
                file_name=f"relatorio_sped_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

def render_audit():
    st.markdown('<h1>📜 Log de Auditoria</h1>', unsafe_allow_html=True)
    
    audit_log = safe_get_audit_log()
    if not audit_log:
        st.info("Nenhuma atividade registrada.")
        return
    
    st.metric("Total de Registros", len(audit_log))
    
    for log in audit_log[-30:]:
        with st.expander(f"📝 {log.get('timestamp', '')} - {log.get('operation', '')}"):
            st.json(log)

# ============================================================================
# MAIN APP
# ============================================================================

def main():
    """Função principal"""
    
    st.set_page_config(
        page_title="SPED Fiscal Analytics",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    init_session_state()
    
    # Sidebar
    with st.sidebar:
        st.image("https://img.icons8.com/color/96/000000/accounting.png", width=80)
        st.title("SPED Fiscal")
        st.markdown("---")
        
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
        
        items = safe_get_items()
        if items:
            st.success(f"✅ {len(items)} itens carregados")
            problematic = len([i for i in items if i.has_problem()])
            if problematic > 0:
                st.warning(f"⚠️ {problematic} itens com problemas")
        else:
            st.info("📤 Aguardando upload")
        
        st.caption("Versão: 3.0.0")
    
    # Renderiza seção
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

if __name__ == "__main__":
    main()