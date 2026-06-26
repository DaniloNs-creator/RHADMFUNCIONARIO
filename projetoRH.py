"""
SPED FISCAL ANALYTICS - SISTEMA COMPLETO CORRIGIDO
Versão: 2.0.0
Arquitetura: Streamlit + Pandas + Regras Fiscais

CORREÇÕES REALIZADAS:
1. Correção do erro 'method' object is not iterable
2. Melhorias na validação de tipos
3. Tratamento robusto de exceções
4. Otimização de performance
5. Cache inteligente
6. Logging completo
7. Interface aprimorada

Autor: Arquiteto de Software Sênior
Data: 2024
"""

import streamlit as st
import pandas as pd
import numpy as np
from decimal import Decimal, ROUND_HALF_UP, getcontext
from datetime import datetime, date
from typing import Dict, List, Any, Optional, Tuple, Union
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

# Configuração de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configuração de precisão decimal
getcontext().prec = 28

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
# MODELOS DE DADOS COM VALIDAÇÃO
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
    
    def to_dict(self) -> dict:
        """Converte para dicionário"""
        return {
            'chave_acesso': self.chave_acesso,
            'cnpj_emitente': self.cnpj_emitente,
            'cnpj_destinatario': self.cnpj_destinatario,
            'valor_total': float(self.valor_total) if self.valor_total else 0,
            'valor_base_calculo': float(self.valor_base_calculo) if self.valor_base_calculo else None,
            'valor_icms': float(self.valor_icms) if self.valor_icms else None,
            'aliquota_icms': float(self.aliquota_icms) if self.aliquota_icms else None,
            'cst': self.cst,
            'cfop': self.cfop,
            'tipo_operacao': self.tipo_operacao,
            'data_emissao': self.data_emissao.isoformat() if self.data_emissao else '',
            'situacao': self.situacao,
            'block': self.block,
            'record_type': self.record_type,
            'total_items': len(self.items)
        }

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
    
    def to_dict(self) -> dict:
        """Converte para dicionário"""
        return {
            'record_id': self.record_id,
            'record_type': self.record_type,
            'block': self.block,
            'line_number': self.line_number,
            'fields': self.fields,
            'modified': self.modified
        }

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
    description: str = ""
    
    def to_dict(self) -> dict:
        """Converte para dicionário"""
        return {
            'rule_id': self.rule_id,
            'cst': self.cst,
            'cfop': self.cfop,
            'operation_type': self.operation_type,
            'requires_base': self.requires_base,
            'requires_aliquota': self.requires_aliquota,
            'requires_imposto': self.requires_imposto,
            'default_aliquota': float(self.default_aliquota) if self.default_aliquota else None,
            'enabled': self.enabled,
            'description': self.description
        }

# ============================================================================
# MOTOR DE REGRAS MELHORADO
# ============================================================================

class RulesEngine:
    """Motor de regras fiscais com cache e logging"""
    
    def __init__(self):
        self.rules: List[FiscalRule] = []
        self.rule_cache: Dict[str, List[FiscalRule]] = {}
        self.logger = logging.getLogger(__name__)
        self._load_default_rules()
    
    def _load_default_rules(self):
        """Carrega regras padrão"""
        default_rules = [
            # CST 00 - Tributada integralmente
            FiscalRule("R001", "00", "*", "ambos", True, True, True,
                      "item_value", Decimal("18.00"), "base * aliquota / 100",
                      description="Tributada integralmente"),
            
            # CST 10 - Tributada com ST
            FiscalRule("R002", "10", "*", "ambos", True, True, True,
                      "item_value", Decimal("18.00"), "base * aliquota / 100",
                      description="Tributada com substituição tributária"),
            
            # CST 20 - Com redução de base
            FiscalRule("R003", "20", "*", "ambos", True, True, True,
                      "item_value", Decimal("12.00"), "base * aliquota / 100",
                      description="Com redução de base de cálculo"),
            
            # CST 30 - Isenta com ST
            FiscalRule("R004", "30", "*", "ambos", False, False, False,
                      None, None, None,
                      description="Isenta com substituição tributária"),
            
            # CST 40 - Isenta
            FiscalRule("R005", "40", "*", "ambos", False, False, False,
                      None, None, None,
                      description="Isenta"),
            
            # CST 41 - Não tributada
            FiscalRule("R006", "41", "*", "ambos", False, False, False,
                      None, None, None,
                      description="Não tributada"),
            
            # CST 50 - Suspensão
            FiscalRule("R007", "50", "*", "ambos", False, False, False,
                      None, None, None,
                      description="Suspensão"),
            
            # CST 51 - Diferimento
            FiscalRule("R008", "51", "*", "ambos", False, False, False,
                      None, None, None,
                      description="Diferimento"),
            
            # CST 60 - ICMS cobrado anteriormente
            FiscalRule("R009", "60", "*", "ambos", False, False, False,
                      None, None, None,
                      description="ICMS cobrado anteriormente"),
            
            # CST 70 - Com redução e ST
            FiscalRule("R010", "70", "*", "ambos", True, True, True,
                      "item_value", Decimal("12.00"), "base * aliquota / 100",
                      description="Com redução e substituição tributária"),
            
            # CST 90 - Outras
            FiscalRule("R011", "90", "*", "ambos", False, False, False,
                      None, None, None,
                      description="Outras"),
            
            # Regras específicas para D100 (CT-e)
            FiscalRule("R012", "00", "5351", "entrada", True, True, True,
                      "item_value", Decimal("12.00"), "base * aliquota / 100",
                      description="Transporte - Entrada"),
            
            # Regras específicas para C500 (Energia)
            FiscalRule("R013", "00", "5401", "entrada", True, True, True,
                      "item_value", Decimal("25.00"), "base * aliquota / 100",
                      description="Energia Elétrica"),
            
            # Regras específicas para F100 (Aluguel)
            FiscalRule("R014", "00", "5355", "entrada", True, True, True,
                      "item_value", Decimal("18.00"), "base * aliquota / 100",
                      description="Locação"),
        ]
        
        self.rules = default_rules
        self._build_cache()
        self.logger.info(f"Carregadas {len(self.rules)} regras")
    
    def _build_cache(self):
        """Constrói cache de regras"""
        self.rule_cache = {}
        for rule in self.rules:
            if rule.enabled:
                key = f"{rule.cst}|{rule.cfop}|{rule.operation_type}"
                if key not in self.rule_cache:
                    self.rule_cache[key] = []
                self.rule_cache[key].append(rule)
    
    @lru_cache(maxsize=1000)
    def get_rule_cached(self, cst: str, cfop: str, operation_type: str) -> Optional[tuple]:
        """Obtém regra com cache"""
        rule = self.get_rule(cst, cfop, operation_type)
        if rule:
            return (rule.rule_id, rule.requires_base, rule.requires_aliquota, 
                   rule.requires_imposto, rule.default_aliquota)
        return None
    
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
            
            # Validações para CST isento
            if not rule.requires_base and item.base_calculo and item.base_calculo > 0:
                errors.append("CST isento com base de cálculo informada")
            
            if not rule.requires_aliquota and item.aliquota and item.aliquota > 0:
                errors.append("CST isento com alíquota informada")
            
            if not rule.requires_imposto and item.valor_imposto and item.valor_imposto > 0:
                errors.append("CST isento com valor de imposto informado")
                
        except Exception as e:
            errors.append(f"Erro na validação: {str(e)}")
            self.logger.error(f"Erro ao validar item {item.descricao_produto}: {str(e)}")
        
        return len(errors) == 0, errors
    
    def suggest_correction(self, item: FiscalItem) -> Dict[str, Any]:
        """Sugere correções para um item fiscal"""
        suggestions = {}
        
        try:
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
                    except Exception as e:
                        self.logger.error(f"Erro ao calcular imposto: {str(e)}")
                        
        except Exception as e:
            suggestions["error"] = f"Erro na sugestão: {str(e)}"
            self.logger.error(f"Erro ao sugerir correção: {str(e)}")
        
        return suggestions
    
    def apply_correction(self, item: FiscalItem) -> Dict[str, Any]:
        """Aplica correções ao item"""
        changes = {}
        
        try:
            suggestions = self.suggest_correction(item)
            if "error" in suggestions:
                return suggestions
            
            # Salva valores originais se não existirem
            if not hasattr(item, 'original_base') or item.original_base is None:
                item.original_base = item.base_calculo
            if not hasattr(item, 'original_aliquota') or item.original_aliquota is None:
                item.original_aliquota = item.aliquota
            if not hasattr(item, 'original_imposto') or item.original_imposto is None:
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
                
        except Exception as e:
            changes['error'] = str(e)
            self.logger.error(f"Erro ao aplicar correção: {str(e)}")
        
        return changes

# ============================================================================
# PARSER MELHORADO
# ============================================================================

class SpedParser:
    """Parser para arquivos SPED com tratamento de erros robusto"""
    
    def __init__(self):
        self.records: List[SpedRecord] = []
        self.documents: Dict[str, FiscalDocument] = {}
        self.items: List[FiscalItem] = []
        self.blocks: Dict[str, List[SpedRecord]] = {}
        self.summary: Dict[str, Any] = {}
        self.company_info: Dict[str, str] = {}
        self.period: str = ""
        self.logger = logging.getLogger(__name__)
        self.current_parent: Optional[SpedRecord] = None
        self.line_number: int = 0
    
    def parse(self, content: str) -> Dict[str, Any]:
        """Parseia arquivo SPED"""
        start_time = time.time()
        self.logger.info("Iniciando parse do arquivo SPED")
        
        try:
            if not content or not content.strip():
                raise ValueError("Conteúdo do arquivo vazio")
            
            lines = content.splitlines()
            self.logger.info(f"Total de linhas: {len(lines)}")
            
            self._parse_lines(lines)
            self._build_hierarchy()
            self._extract_documents()
            self._extract_items()
            self._generate_summary()
            
            elapsed = time.time() - start_time
            self.logger.info(f"Parse concluído em {elapsed:.2f} segundos")
            
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
            self.logger.error(f"Erro no parse: {str(e)}")
            raise
    
    def _parse_lines(self, lines: List[str]):
        """Parseia linhas do arquivo"""
        for line_num, line in enumerate(lines, 1):
            self.line_number = line_num
            if not line.strip():
                continue
            
            try:
                parts = line.split("|")
                if len(parts) < 2:
                    continue
                
                record_type = parts[1].strip()
                if not record_type:
                    continue
                
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
                    self.current_parent = record
                elif self.current_parent:
                    record.parent_record = self.current_parent.record_id
                    if record.record_id not in self.current_parent.children_records:
                        self.current_parent.children_records.append(record.record_id)
                
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
                    
            except Exception as e:
                self.logger.warning(f"Erro na linha {line_num}: {str(e)}")
                continue
    
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
        
        try:
            # Registro 0000 - Abertura
            if record_type == "0000" and len(parts) >= 11:
                fields["codigo_versao"] = self._clean(parts[2] if len(parts) > 2 else "")
                fields["tipo_arquivo"] = self._clean(parts[3] if len(parts) > 3 else "")
                fields["cnpj"] = self._clean(parts[4] if len(parts) > 4 else "")
                fields["nome_empresa"] = self._clean(parts[5] if len(parts) > 5 else "")
                fields["ie"] = self._clean(parts[7] if len(parts) > 7 else "")
                fields["data_inicial"] = self._clean(parts[9] if len(parts) > 9 else "")
                fields["data_final"] = self._clean(parts[10] if len(parts) > 10 else "")
            
            # Registro C100 - Nota Fiscal
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
            elif record_type == "F100" and len(parts) >= 9:
                fields["tipo_operacao"] = self._clean(parts[2] if len(parts) > 2 else "")
                fields["cfop"] = self._clean(parts[3] if len(parts) > 3 else "")
                fields["cnpj_emitente"] = self._clean(parts[4] if len(parts) > 4 else "")
                fields["valor_total"] = self._parse_decimal(parts[5] if len(parts) > 5 else "0")
                fields["base_calculo"] = self._parse_decimal(parts[6] if len(parts) > 6 else None)
                fields["aliquota"] = self._parse_decimal(parts[7] if len(parts) > 7 else None)
                fields["valor_imposto"] = self._parse_decimal(parts[8] if len(parts) > 8 else None)
                fields["block"] = "F"
            
        except Exception as e:
            self.logger.warning(f"Erro ao extrair campos do registro {record_type}: {str(e)}")
        
        return fields
    
    def _build_hierarchy(self):
        """Constrói hierarquia de registros"""
        # Já construída durante o parsing
        pass
    
    def _extract_documents(self):
        """Extrai documentos fiscais"""
        try:
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
        except Exception as e:
            self.logger.error(f"Erro ao extrair documentos: {str(e)}")
    
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
            record_type='C100',
            linha_original=record.line_number
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
            record_type='D100',
            linha_original=record.line_number
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
            record_type='C500',
            linha_original=record.line_number
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
            record_type='F100',
            linha_original=record.line_number
        )
    
    def _extract_items(self):
        """Extrai itens dos registros"""
        try:
            # Extrai itens C170
            for record in self.records:
                if record.record_type == "C170":
                    fields = record.fields
                    try:
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
                    except Exception as e:
                        self.logger.warning(f"Erro ao criar item C170: {str(e)}")
            
            # Extrai itens D100/D101/D105
            for record in self.records:
                if record.record_type in ["D100", "D101", "D105"]:
                    fields = record.fields
                    if fields.get('valor_servico') or fields.get('base_calculo'):
                        try:
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
                        except Exception as e:
                            self.logger.warning(f"Erro ao criar item D100: {str(e)}")
            
            # Extrai itens C500/C501/C505 (Energia)
            for record in self.records:
                if record.record_type in ["C500", "C501", "C505"]:
                    fields = record.fields
                    if fields.get('valor_total') or fields.get('base_calculo'):
                        try:
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
                        except Exception as e:
                            self.logger.warning(f"Erro ao criar item C500: {str(e)}")
            
            # Extrai itens F100 (Aluguel)
            for record in self.records:
                if record.record_type == "F100":
                    fields = record.fields
                    if fields.get('valor_total') or fields.get('base_calculo'):
                        try:
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
                        except Exception as e:
                            self.logger.warning(f"Erro ao criar item F100: {str(e)}")
                            
        except Exception as e:
            self.logger.error(f"Erro ao extrair itens: {str(e)}")
    
    def _generate_summary(self):
        """Gera resumo do arquivo"""
        try:
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
        except Exception as e:
            self.logger.error(f"Erro ao gerar resumo: {str(e)}")
    
    def _clean(self, value: str) -> str:
        """Limpa string"""
        if not value:
            return ""
        return str(value).strip()
    
    def _parse_decimal(self, value: Any) -> Optional[Decimal]:
        """Converte para Decimal"""
        if value is None or value == "":
            return None
        try:
            if isinstance(value, str):
                # Remove caracteres não numéricos
                value = re.sub(r'[^\d.,-]', '', value)
                value = value.replace(',', '.')
            return Decimal(str(value))
        except Exception:
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
        except Exception:
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
# SERVIÇOS COM TRATAMENTO DE ERROS
# ============================================================================

class ValidationService:
    """Serviço de validação fiscal com tratamento robusto"""
    
    def __init__(self, rules_engine: RulesEngine):
        self.rules_engine = rules_engine
        self.inconsistencies: List[Dict] = []
        self.logger = logging.getLogger(__name__)
    
    def validate_all(self, items: List[FiscalItem]) -> List[Dict]:
        """Valida todos os itens"""
        self.inconsistencies = []
        
        if not items:
            self.logger.info("Nenhum item para validar")
            return self.inconsistencies
        
        # Garante que items é uma lista
        if not isinstance(items, list):
            self.logger.warning(f"items não é uma lista: {type(items)}")
            # Tenta converter para lista
            if hasattr(items, 'values'):
                items = list(items.values())
            else:
                items = list(items)
        
        self.logger.info(f"Iniciando validação de {len(items)} itens")
        
        for item in items:
            if not isinstance(item, FiscalItem):
                self.logger.warning(f"Item inválido: {type(item)}")
                continue
            
            # Valida campos obrigatórios
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
            
            # Valida regras fiscais
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
            except Exception as e:
                self.logger.error(f"Erro ao validar item {item.descricao_produto}: {str(e)}")
                self.inconsistencies.append({
                    'type': 'error',
                    'item': item.descricao_produto or 'Item sem descrição',
                    'message': f'Erro na validação: {str(e)}',
                    'details': f'Produto: {item.descricao_produto}',
                    'severity': 'alta',
                    'linha': item.linha_original
                })
        
        self.logger.info(f"Validação concluída. Encontradas {len(self.inconsistencies)} inconsistências")
        return self.inconsistencies

class CorrectionService:
    """Serviço de correção fiscal"""
    
    def __init__(self, rules_engine: RulesEngine):
        self.rules_engine = rules_engine
        self.corrections: List[Dict] = []
        self.logger = logging.getLogger(__name__)
    
    def correct_item(self, item: FiscalItem) -> Dict[str, Any]:
        """Corrige um item fiscal"""
        changes = {}
        
        try:
            suggestions = self.rules_engine.suggest_correction(item)
            if "error" in suggestions:
                return suggestions
            
            # Salva valores originais se não existirem
            if not hasattr(item, 'original_base') or item.original_base is None:
                item.original_base = item.base_calculo
            if not hasattr(item, 'original_aliquota') or item.original_aliquota is None:
                item.original_aliquota = item.aliquota
            if not hasattr(item, 'original_imposto') or item.original_imposto is None:
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
                
        except Exception as e:
            changes['error'] = str(e)
            self.logger.error(f"Erro ao corrigir item {item.descricao_produto}: {str(e)}")
        
        return changes
    
    def correct_mass(self, items: List[FiscalItem]) -> Dict[str, int]:
        """Corrige múltiplos itens em massa"""
        results = {
            'total': len(items) if items else 0,
            'corrected': 0,
            'base_corrected': 0,
            'aliquota_corrected': 0,
            'imposto_corrected': 0,
            'errors': 0
        }
        
        if not items:
            return results
        
        # Garante que é uma lista
        if not isinstance(items, list):
            items = list(items) if hasattr(items, 'values') else list(items)
        
        for item in items:
            try:
                changes = self.correct_item(item)
                if changes and "error" not in changes:
                    results['corrected'] += 1
                    if 'base_calculo' in changes:
                        results['base_corrected'] += 1
                    if 'aliquota' in changes:
                        results['aliquota_corrected'] += 1
                    if 'valor_imposto' in changes:
                        results['imposto_corrected'] += 1
                elif changes and "error" in changes:
                    results['errors'] += 1
            except Exception as e:
                results['errors'] += 1
                self.logger.error(f"Erro na correção em massa: {str(e)}")
        
        return results

class ExportService:
    """Serviço de exportação"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def export_sped(self, records: List[SpedRecord]) -> str:
        """Exporta SPED corrigido"""
        lines = []
        
        if not records:
            return ""
        
        # Garante que é uma lista
        if not isinstance(records, list):
            records = list(records.values()) if hasattr(records, 'values') else list(records)
        
        for record in records:
            try:
                if record.modified and record.fields:
                    # Reconstrói a linha com os campos atualizados
                    parts = [record.record_type]
                    # Adiciona campos na ordem correta
                    for key in sorted(record.fields.keys()):
                        value = record.fields.get(key, '')
                        if value is None:
                            value = ''
                        parts.append(str(value))
                    lines.append('|' + '|'.join(parts))
                else:
                    # Usa dados originais
                    lines.append(record.raw_data)
            except Exception as e:
                self.logger.warning(f"Erro ao exportar registro {record.record_id}: {str(e)}")
                lines.append(record.raw_data)
        
        return '\n'.join(lines)
    
    def export_excel(self, items: List[FiscalItem], inconsistencies: List[Dict], 
                     corrections: List[Dict], audit_log: List[Dict]) -> bytes:
        """Exporta relatório em Excel"""
        output = io.BytesIO()
        
        try:
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
                            'Modificado': 'Sim' if item.modified else 'Não',
                            'Linha': item.linha_original
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
                                'Valor Novo': change.get('new'),
                                'Data': corr.get('timestamp', '')
                            })
                    if corr_data:
                        df_corr = pd.DataFrame(corr_data)
                        df_corr.to_excel(writer, sheet_name='Correcoes', index=False)
                
                # Aba: Auditoria
                if audit_log:
                    df_audit = pd.DataFrame(audit_log)
                    df_audit.to_excel(writer, sheet_name='Auditoria', index=False)
                
                # Aba: Resumo
                summary_data = {
                    'Total Itens': len(items) if items else 0,
                    'Total Inconsistências': len(inconsistencies) if inconsistencies else 0,
                    'Total Correções': len(corrections) if corrections else 0,
                    'Total Auditoria': len(audit_log) if audit_log else 0,
                    'Data Exportação': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                df_summary = pd.DataFrame([summary_data])
                df_summary.to_excel(writer, sheet_name='Resumo', index=False)
                
        except Exception as e:
            self.logger.error(f"Erro ao exportar Excel: {str(e)}")
            raise
        
        return output.getvalue()
    
    def export_csv_inconsistencies(self, inconsistencies: List[Dict]) -> str:
        """Exporta inconsistências em CSV"""
        if not inconsistencies:
            return "Nenhuma inconsistência encontrada"
        
        try:
            df = pd.DataFrame(inconsistencies)
            return df.to_csv(index=False, encoding='utf-8-sig')
        except Exception as e:
            self.logger.error(f"Erro ao exportar CSV: {str(e)}")
            return f"Erro ao exportar: {str(e)}"

# ============================================================================
# FUNÇÕES DE UI - VERSÃO CORRIGIDA
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

# ============================================================================
# FUNÇÕES DE RENDERIZAÇÃO PRINCIPAIS
# ============================================================================

def render_dashboard():
    """Renderiza dashboard principal"""
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
            margin-bottom: 1rem;
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
    
    items = st.session_state.get('items', [])
    if not items:
        st.info("📤 Nenhum arquivo SPED carregado. Acesse a seção 'Upload' para começar.")
        return
    
    # Métricas principais
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
        cst_data = {}
        for item in items:
            cst = item.cst or 'N/A'
            cst_data[cst] = cst_data.get(cst, 0) + 1
        
        if cst_data:
            df_cst = pd.DataFrame(list(cst_data.items()), columns=['CST', 'Quantidade'])
            fig = px.pie(df_cst, values='Quantidade', names='CST', 
                       title='Itens por CST', color_discrete_sequence=px.colors.qualitative.Set3)
            st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("Distribuição por CFOP")
        cfop_data = {}
        for item in items:
            cfop = item.cfop or 'N/A'
            cfop_data[cfop] = cfop_data.get(cfop, 0) + 1
        
        if cfop_data:
            df_cfop = pd.DataFrame(list(cfop_data.items()), columns=['CFOP', 'Quantidade'])
            fig = px.bar(df_cfop, x='CFOP', y='Quantidade', 
                       title='Itens por CFOP', color='CFOP')
            st.plotly_chart(fig, use_container_width=True)
    
    # Resumo dos blocos
    st.subheader("📋 Resumo por Bloco")
    blocks = st.session_state.get('blocks', {})
    if blocks:
        block_data = []
        for block, records in blocks.items():
            block_data.append({
                'Bloco': block,
                'Descrição': BLOCOS_SPED.get(block, 'Desconhecido'),
                'Registros': len(records),
                'Itens': len([i for i in items if i.block == block])
            })
        df_blocks = pd.DataFrame(block_data)
        st.dataframe(df_blocks, use_container_width=True)

def render_upload():
    """Renderiza seção de upload - CORRIGIDA"""
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
                st.session_state.parsed_data = data
                st.session_state.records = data.get('records', {})
                st.session_state.documents = data.get('documents', {})
                st.session_state.items = data.get('items', [])
                st.session_state.blocks = data.get('blocks', {})
                st.session_state.summary = data.get('summary', {})
                st.session_state.period = data.get('period', '')
                st.session_state.company_info = data.get('company_info', {})
                
                # Inicializa regras se não existir
                if 'rules_engine' not in st.session_state:
                    st.session_state.rules_engine = RulesEngine()
                
                # IMPORTANTE: Validação corrigida - trata items como lista
                validation_service = ValidationService(st.session_state.rules_engine)
                items = st.session_state.items
                
                # Garante que items é uma lista
                if not isinstance(items, list):
                    if hasattr(items, 'values'):
                        items = list(items.values())
                    else:
                        items = list(items) if items else []
                    st.session_state.items = items
                
                # Valida os itens
                st.session_state.inconsistencies = validation_service.validate_all(items)
                
                # Inicializa logs
                if 'audit_log' not in st.session_state:
                    st.session_state.audit_log = []
                
                st.success(f"✅ Arquivo processado com sucesso!")
                
                # Exibe resumo
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Registros", len(st.session_state.records) if st.session_state.records else 0)
                with col2:
                    st.metric("Notas Fiscais", len(st.session_state.documents) if st.session_state.documents else 0)
                with col3:
                    st.metric("Itens", len(st.session_state.items) if st.session_state.items else 0)
                with col4:
                    st.metric("Inconsistências", len(st.session_state.inconsistencies) if st.session_state.inconsistencies else 0)
                
                # Informações da empresa
                if st.session_state.company_info:
                    with st.expander("🏢 Informações da Empresa"):
                        st.json(st.session_state.company_info)
                
                # Resumo dos blocos
                if st.session_state.blocks:
                    with st.expander("📊 Resumo dos Blocos"):
                        block_summary = []
                        for block, records in st.session_state.blocks.items():
                            record_types = {}
                            for rec in records:
                                record_types[rec.record_type] = record_types.get(rec.record_type, 0) + 1
                            
                            block_summary.append({
                                'Bloco': block,
                                'Descrição': BLOCOS_SPED.get(block, 'Desconhecido'),
                                'Total Registros': len(records),
                                'Tipos': ', '.join(record_types.keys())
                            })
                        
                        if block_summary:
                            df_summary = pd.DataFrame(block_summary)
                            st.dataframe(df_summary, use_container_width=True)
        
        except Exception as e:
            st.error(f"❌ Erro ao processar arquivo: {str(e)}")
            import traceback
            st.code(traceback.format_exc())

def render_blocks():
    """Renderiza visualização de blocos"""
    st.markdown('<h1>📋 Visualização de Blocos</h1>', unsafe_allow_html=True)
    
    records = st.session_state.get('records', {})
    if not records:
        st.info("Nenhum arquivo carregado. Acesse a seção 'Upload'.")
        return
    
    # Seleção de bloco
    blocks = sorted(st.session_state.get('blocks', {}).keys())
    if not blocks:
        st.info("Nenhum bloco encontrado.")
        return
    
    selected_block = st.selectbox("Selecione o Bloco", blocks)
    
    if selected_block:
        block_records = st.session_state.blocks.get(selected_block, [])
        st.subheader(f"Bloco {selected_block} - {BLOCOS_SPED.get(selected_block, 'Desconhecido')}")
        st.write(f"Total de registros: {len(block_records)}")
        
        # Filtros
        col1, col2 = st.columns(2)
        with col1:
            record_types = list(set(r.record_type for r in block_records))
            selected_type = st.selectbox("Tipo de Registro", ["Todos"] + sorted(record_types))
        
        with col2:
            search = st.text_input("Buscar", placeholder="Digite para filtrar...")
        
        # Aplica filtros
        filtered = block_records
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
                    'Campos': len(record.fields),
                    'Modificado': '✅' if record.modified else ''
                }
                # Adiciona campos principais
                for key in list(record.fields.keys())[:3]:
                    val = record.fields.get(key, '')
                    if val is None:
                        val = ''
                    row[key] = str(val)[:30]
                data.append(row)
            
            if data:
                df = pd.DataFrame(data)
                st.dataframe(df, use_container_width=True)
            else:
                st.info("Nenhum registro encontrado.")
        else:
            st.info("Nenhum registro encontrado para os filtros selecionados.")

def render_records():
    """Renderiza visualização de registros"""
    st.markdown('<h1>📄 Visualização de Registros</h1>', unsafe_allow_html=True)
    
    records = st.session_state.get('records', {})
    if not records:
        st.info("Nenhum arquivo carregado.")
        return
    
    # Converte para lista se for dicionário
    if isinstance(records, dict):
        records_list = list(records.values())
    else:
        records_list = records if isinstance(records, list) else []
    
    if not records_list:
        st.info("Nenhum registro encontrado.")
        return
    
    # Filtros
    col1, col2, col3 = st.columns(3)
    with col1:
        record_types = list(set(r.record_type for r in records_list))
        selected_types = st.multiselect("Tipo de Registro", sorted(record_types))
    
    with col2:
        blocks = list(set(r.block for r in records_list))
        selected_blocks = st.multiselect("Bloco", sorted(blocks))
    
    with col3:
        search = st.text_input("Buscar", placeholder="Digite para filtrar...")
    
    # Aplica filtros
    filtered = records_list
    if selected_types:
        filtered = [r for r in filtered if r.record_type in selected_types]
    if selected_blocks:
        filtered = [r for r in filtered if r.block in selected_blocks]
    if search:
        filtered = [r for r in filtered if search.lower() in r.raw_data.lower()]
    
    st.info(f"📊 Registros encontrados: {len(filtered)}")
    
    # Paginação
    page_size = 50
    total_pages = (len(filtered) + page_size - 1) // page_size if filtered else 1
    page = st.number_input("Página", min_value=1, max_value=total_pages, value=1)
    
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
        record_options = [f"{r.record_type} - Linha {r.line_number}" for r in page_records]
        if record_options:
            selected_idx = st.selectbox("Selecione um registro para detalhes", range(len(record_options)), 
                                      format_func=lambda i: record_options[i])
            if selected_idx is not None and selected_idx < len(page_records):
                record = page_records[selected_idx]
                st.json(record.to_dict())

def render_invoices():
    """Renderiza visualização de notas fiscais"""
    st.markdown('<h1>📑 Notas Fiscais</h1>', unsafe_allow_html=True)
    
    documents = st.session_state.get('documents', {})
    if not documents:
        st.info("Nenhuma nota fiscal encontrada.")
        return
    
    docs_list = list(documents.values()) if isinstance(documents, dict) else documents
    
    st.info(f"📊 Notas fiscais encontradas: {len(docs_list)}")
    
    # Filtros
    col1, col2 = st.columns(2)
    with col1:
        cst_list = list(set(d.cst for d in docs_list if d.cst))
        selected_cst = st.multiselect("CST", sorted(cst_list))
    
    with col2:
        cfop_list = list(set(d.cfop for d in docs_list if d.cfop))
        selected_cfop = st.multiselect("CFOP", sorted(cfop_list))
    
    # Aplica filtros
    filtered = docs_list
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
                st.write(f"**Linha:** {doc.linha_original}")
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
                        'Imposto': float(item.valor_imposto) if item.valor_imposto else 0,
                        'Status': '⚠️' if item.has_problem() else '✅'
                    })
                if items_data:
                    df_items = pd.DataFrame(items_data)
                    st.dataframe(df_items, use_container_width=True)

def render_items():
    """Renderiza visualização de itens"""
    st.markdown('<h1>📦 Itens Fiscais</h1>', unsafe_allow_html=True)
    
    items = st.session_state.get('items', [])
    if not items:
        st.info("Nenhum item encontrado.")
        return
    
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
        filtered = [i for i in filtered if i.has_problem()]
    
    st.info(f"📊 Itens encontrados: {len(filtered)}")
    
    if filtered:
        # Exibe tabela
        data = []
        for item in filtered[:1000]:
            data.append({
                'Produto': item.descricao_produto[:40] + '...' if len(item.descricao_produto) > 40 else item.descricao_produto,
                'CST': item.cst,
                'CFOP': item.cfop,
                'Valor Total': float(item.valor_total) if item.valor_total else 0,
                'Base': float(item.base_calculo) if item.base_calculo else 0,
                'Alíquota': float(item.aliquota) if item.aliquota else 0,
                'Imposto': float(item.valor_imposto) if item.valor_imposto else 0,
                'Status': '⚠️' if item.has_problem() else '✅',
                'Block': item.block,
                'Linha': item.linha_original
            })
        
        df = pd.DataFrame(data)
        st.dataframe(df, use_container_width=True)
        
        # Estatísticas
        st.markdown("---")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Itens", len(filtered))
        with col2:
            problematic = len([i for i in filtered if i.has_problem()])
            st.metric("Com Problemas", problematic)
        with col3:
            sem_base = len([i for i in filtered if i.base_calculo is None or i.base_calculo <= 0])
            st.metric("Sem Base", sem_base)
        with col4:
            sem_aliquota = len([i for i in filtered if i.aliquota is None or i.aliquota <= 0])
            st.metric("Sem Alíquota", sem_aliquota)

def render_inconsistencies():
    """Renderiza seção de inconsistências"""
    st.markdown('<h1>⚠️ Inconsistências Fiscais</h1>', unsafe_allow_html=True)
    
    inconsistencies = st.session_state.get('inconsistencies', [])
    if not inconsistencies:
        st.success("✅ Nenhuma inconsistência encontrada!")
        return
    
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
        
        st.caption(f"Linha: {inc.get('linha', 'N/A')} | {inc.get('details', '')}")

def render_mass_correction():
    """Renderiza seção de correção em massa"""
    st.markdown('<h1>🔧 Correções em Massa</h1>', unsafe_allow_html=True)
    
    items = st.session_state.get('items', [])
    if not items:
        st.info("Nenhum item disponível para correção.")
        return
    
    # Filtros
    st.subheader("🔍 Filtros")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        cst_list = sorted(list(set(i.cst for i in items if i.cst)))
        cst_filter = st.multiselect("CST", cst_list)
    
    with col2:
        cfop_list = sorted(list(set(i.cfop for i in items if i.cfop)))
        cfop_filter = st.multiselect("CFOP", cfop_list)
    
    with col3:
        only_problems = st.checkbox("Apenas itens com problemas", value=True)
    
    # Aplica filtros
    filtered = items
    if cst_filter:
        filtered = [i for i in filtered if i.cst in cst_filter]
    if cfop_filter:
        filtered = [i for i in filtered if i.cfop in cfop_filter]
    if only_problems:
        filtered = [i for i in filtered if i.has_problem()]
    
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
                                'old_value': float(item.original_base) if item.original_base else None,
                                'new_value': float(item.base_calculo) if item.base_calculo else None,
                                'linha': item.linha_original
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
                                'old_value': float(item.original_aliquota) if item.original_aliquota else None,
                                'new_value': float(item.aliquota) if item.aliquota else None,
                                'linha': item.linha_original
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
                                    'old_value': float(item.original_imposto) if item.original_imposto else None,
                                    'new_value': float(item.valor_imposto) if item.valor_imposto else None,
                                    'linha': item.linha_original
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
                        'old_base': float(item.original_base) if item.original_base else None,
                        'new_base': float(item.base_calculo) if item.base_calculo else None,
                        'old_aliquota': float(item.original_aliquota) if item.original_aliquota else None,
                        'new_aliquota': float(item.aliquota) if item.aliquota else None,
                        'old_imposto': float(item.original_imposto) if item.original_imposto else None,
                        'new_imposto': float(item.valor_imposto) if item.valor_imposto else None,
                        'linha': item.linha_original
                    })
            
            st.success(f"""
                ✅ Correções aplicadas com sucesso!
                - Total itens: {results['total']}
                - Corrigidos: {results['corrected']}
                - Bases: {results['base_corrected']}
                - Alíquotas: {results['aliquota_corrected']}
                - Impostos: {results['imposto_corrected']}
                - Erros: {results['errors']}
            """)
            st.rerun()
    
    # Preview
    st.subheader("📋 Preview dos Itens Selecionados")
    preview_data = []
    for item in filtered[:50]:
        preview_data.append({
            'Produto': item.descricao_produto[:30] + '...' if len(item.descricao_produto) > 30 else item.descricao_produto,
            'CST': item.cst,
            'CFOP': item.cfop,
            'Base': float(item.base_calculo) if item.base_calculo else 0,
            'Alíquota': float(item.aliquota) if item.aliquota else 0,
            'Imposto': float(item.valor_imposto) if item.valor_imposto else 0,
            'Status': '⚠️' if item.has_problem() else '✅',
            'Linha': item.linha_original
        })
    
    if preview_data:
        df_preview = pd.DataFrame(preview_data)
        st.dataframe(df_preview, use_container_width=True)
        st.caption(f"Mostrando {len(preview_data)} de {len(filtered)} itens")

def render_manual_editor():
    """Renderiza editor manual"""
    st.markdown('<h1>✏️ Editor Manual</h1>', unsafe_allow_html=True)
    
    items = st.session_state.get('items', [])
    if not items:
        st.info("Nenhum item disponível para edição.")
        return
    
    # Seleção de item
    item_options = {f"{i.descricao_produto} - CST {i.cst} - CFOP {i.cfop} (Linha {i.linha_original})": i for i in items}
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
            st.write(f"**Linha:** {item.linha_original}")
            st.write(f"**Modificado:** {'✅ Sim' if item.modified else '❌ Não'}")
        
        # Problemas atuais
        if item.has_problem():
            st.warning(f"⚠️ Problemas: {', '.join(item.get_problems())}")
        else:
            st.success("✅ Item sem problemas")
        
        st.markdown("### ✏️ Edição")
        
        col1, col2 = st.columns(2)
        with col1:
            valor_total = st.number_input(
                "Valor Total",
                value=float(item.valor_total) if item.valor_total else 0.0,
                step=0.01,
                format="%.2f",
                key=f"vt_{item.record_id}"
            )
            
            base_calculo = st.number_input(
                "Base de Cálculo",
                value=float(item.base_calculo) if item.base_calculo else 0.0,
                step=0.01,
                format="%.2f",
                key=f"bc_{item.record_id}"
            )
        
        with col2:
            aliquota = st.number_input(
                "Alíquota (%)",
                value=float(item.aliquota) if item.aliquota else 0.0,
                step=0.01,
                format="%.2f",
                key=f"al_{item.record_id}"
            )
            
            valor_imposto = st.number_input(
                "Valor do Imposto",
                value=float(item.valor_imposto) if item.valor_imposto else 0.0,
                step=0.01,
                format="%.2f",
                key=f"vi_{item.record_id}"
            )
        
        # Sugestões
        col1, col2 = st.columns(2)
        with col1:
            if st.button("💡 Sugerir Correções"):
                suggestions = st.session_state.rules_engine.suggest_correction(item)
                if "error" in suggestions:
                    st.warning(suggestions["error"])
                else:
                    st.json(suggestions)
        
        with col2:
            if st.button("🔄 Aplicar Sugestões"):
                changes = st.session_state.rules_engine.apply_correction(item)
                if "error" in changes:
                    st.error(changes["error"])
                else:
                    st.success("✅ Sugestões aplicadas!")
                    st.rerun()
        
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
                    'old_value': float(old) if old else None,
                    'new_value': float(new) if new else None,
                    'linha': item.linha_original
                })
            
            st.success("✅ Alterações salvas com sucesso!")
            st.rerun()

def render_export():
    """Renderiza seção de exportação"""
    st.markdown('<h1>📥 Exportação</h1>', unsafe_allow_html=True)
    
    records = st.session_state.get('records', {})
    if not records:
        st.info("Nenhum dado disponível para exportação.")
        return
    
    export_service = ExportService()
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📄 Exportar SPED")
        if st.button("Exportar SPED Corrigido"):
            with st.spinner("Gerando arquivo SPED..."):
                records_list = list(records.values()) if isinstance(records, dict) else records
                content = export_service.export_sped(records_list)
                st.download_button(
                    label="📥 Baixar SPED",
                    data=content,
                    file_name=f"sped_corrigido_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                    mime="text/plain"
                )
    
    with col2:
        st.subheader("📊 Exportar Excel")
        if st.button("Exportar Relatório Excel"):
            with st.spinner("Gerando relatório Excel..."):
                items = st.session_state.get('items', [])
                inconsistencies = st.session_state.get('inconsistencies', [])
                corrections = st.session_state.get('corrections', [])
                audit_log = st.session_state.get('audit_log', [])
                
                excel_data = export_service.export_excel(
                    items, inconsistencies, corrections, audit_log
                )
                st.download_button(
                    label="📥 Baixar Excel",
                    data=excel_data,
                    file_name=f"relatorio_sped_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
    st.markdown("---")
    
    st.subheader("📋 CSV de Inconsistências")
    inconsistencies = st.session_state.get('inconsistencies', [])
    if inconsistencies:
        if st.button("Exportar CSV"):
            csv_data = export_service.export_csv_inconsistencies(inconsistencies)
            st.download_button(
                label="📥 Baixar CSV",
                data=csv_data,
                file_name=f"inconsistencias_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
    else:
        st.info("Nenhuma inconsistência para exportar.")

def render_audit():
    """Renderiza log de auditoria"""
    st.markdown('<h1>📜 Log de Auditoria</h1>', unsafe_allow_html=True)
    
    audit_log = st.session_state.get('audit_log', [])
    if not audit_log:
        st.info("Nenhuma atividade registrada.")
        return
    
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
    
    st.info(f"📊 Registros encontrados: {len(filtered)}")
    
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
        page_title="SPED Fiscal Analytics - Profissional",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Inicializa estado
    init_session_state()
    
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
        items = st.session_state.get('items', [])
        if items:
            st.success(f"✅ {len(items)} itens carregados")
            problematic = len([i for i in items if i.has_problem()])
            if problematic > 0:
                st.warning(f"⚠️ {problematic} itens com problemas")
        else:
            st.info("📤 Aguardando upload")
        
        # Informações do sistema
        st.markdown("---")
        st.caption(f"Versão: 2.0.0")
        st.caption(f"Streamlit: {st.__version__}")
    
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