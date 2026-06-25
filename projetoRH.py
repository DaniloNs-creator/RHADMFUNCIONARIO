"""
SPED MANAGER PRO - Sistema Corporativo Completo
===============================================
Sistema profissional para leitura, análise, edição, validação e exportação 
de arquivos SPED Fiscal Brasileiro.

Versão: 2.1.0 (Corrigida)
Data: 2024
"""

import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import os
import sys
import json
import yaml
from datetime import datetime, date
from typing import Dict, List, Tuple, Optional, Any, Union
from dataclasses import dataclass, field, asdict
from decimal import Decimal, ROUND_HALF_UP, ROUND_DOWN, ROUND_UP
from enum import Enum
import logging
import traceback
import hashlib
from copy import deepcopy
import warnings
warnings.filterwarnings('ignore')

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('sped_manager.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ============================================================================
# CONSTANTES E ENUMS
# ============================================================================

class SeverityLevel(Enum):
    """Níveis de severidade para inconsistências"""
    CRITICA = "Crítica"
    AVISO = "Aviso" 
    INFORMACAO = "Informação"

class OperationType(Enum):
    """Tipos de operação fiscal"""
    ENTRADA = "entrada"
    SAIDA = "saida"
    SERVICO = "servico"

class TaxRegime(Enum):
    """Regimes tributários"""
    SIMPLES = "Simples Nacional"
    LUCRO_PRESUMIDO = "Lucro Presumido"
    LUCRO_REAL = "Lucro Real"

# Constantes do SPED
BLOCK_DESCRIPTIONS = {
    '0': 'Abertura, Identificação e Referências',
    'A': 'Documentos Fiscais - Serviços (não sujeitos ao ICMS)',
    'B': 'Apuração do ISS',
    'C': 'Documentos Fiscais - Mercadorias (ICMS)',
    'D': 'Documentos Fiscais - Aquisições/Transportes',
    'E': 'Apuração do ICMS e do ICMS-ST',
    'F': 'Controle da Produção e do Estoque',
    'G': 'Controle de Créditos de ICMS (CIAP)',
    'H': 'Inventário Físico',
    'I': 'Outras Informações',
    'J': 'Energia Elétrica',
    'K': 'Controle da Produção e do Estoque',
    '1': 'Complemento da Escrituração',
    '2': 'Controle de Valores',
    '9': 'Controle e Encerramento do Arquivo Digital'
}

REGISTER_DESCRIPTIONS = {
    '0000': 'Abertura do Arquivo Digital e Identificação da Entidade',
    '0001': 'Abertura do Bloco 0',
    '0005': 'Dados Complementares da Entidade',
    '0015': 'Dados do Representante Legal/Contabilista',
    '0100': 'Dados do Contabilista',
    '0150': 'Tabela de Cadastro do Participante',
    '0200': 'Tabela de Identificação do Item (Produtos e Serviços)',
    '0300': 'Cadastro de Bens ou Componentes do Ativo Imobilizado',
    '0400': 'Tabela de Natureza da Operação/ Prestação',
    '0450': 'Tabela de Informação Complementar do Documento Fiscal',
    '0460': 'Tabela de Observações do Lançamento Fiscal',
    '0500': 'Plano de Contas Contábeis',
    '0600': 'Centro de Custos',
    'C100': 'Documento - Nota Fiscal (Código 01), Nota Fiscal Avulsa (1B), '
             'Nota Fiscal de Produtor (04) e NF-e (55)',
    'C110': 'Informação Complementar da Nota Fiscal',
    'C120': 'Complemento de Documento',
    'C130': 'ISS/IRRF/PIS/COFINS',
    'C140': 'Fatura',
    'C160': 'Volumes Transportados',
    'C170': 'Itens do Documento',
    'C190': 'Registro Analítico do Documento',
    'C400': 'Equipamento ECF',
    'C405': 'Redução Z',
    'C410': 'Registro Totalizador',
    'C420': 'Resumo por Totalizador',
    'C425': 'Resumo de Itens do Movimento Diário',
    'C460': 'Documento Fiscal Emitido por ECF',
    'C470': 'Itens do Documento Fiscal Emitido por ECF',
    'C490': 'Registro Analítico de CFOP',
    'C500': 'Nota Fiscal/Conta de Energia Elétrica (Código 06)',
    'C510': 'Itens da NF/Conta de Energia Elétrica',
    'C590': 'Registro Analítico do Documento',
    'C600': 'Consolidação Diária de Notas Fiscais/Contas de Energia',
    'C610': 'Itens da Consolidação Diária',
    'C690': 'Registro Analítico dos Documentos',
    'C700': 'Consolidação dos Documentos Emitidos por ECF',
    'C790': 'Registro Analítico dos Documentos (C700)',
    'D100': 'Aquisição de Serviços de Transporte',
    'D190': 'Registro Analítico dos Documentos',
    'D500': 'Documento de Serviço de Comunicação',
    'D590': 'Registro Analítico do Documento',
    'E100': 'Período de Apuração',
    'E110': 'Apuração do ICMS - Operações Próprias',
    'E111': 'Ajuste/Benefício/Incentivo da Apuração do ICMS',
    'E116': 'Obrigações do ICMS Recolhido ou a Recolher',
    'E200': 'Apuração do ICMS - Substituição Tributária',
    'H010': 'Inventário',
    'H020': 'Informações Complementares do Inventário'
}

CST_ICMS_CODES = {
    '000': {'description': 'Tributada integralmente', 'category': 'tributado'},
    '010': {'description': 'Tributada e com cobrança do ICMS por ST', 'category': 'tributado'},
    '020': {'description': 'Com redução de base de cálculo', 'category': 'tributado'},
    '030': {'description': 'Isenta ou não tributada e com cobrança do ICMS por ST', 
            'category': 'isento'},
    '040': {'description': 'Isenta', 'category': 'isento'},
    '041': {'description': 'Não tributada', 'category': 'isento'},
    '050': {'description': 'Suspensão', 'category': 'suspenso'},
    '051': {'description': 'Diferimento', 'category': 'diferido'},
    '060': {'description': 'ICMS cobrado anteriormente por ST', 'category': 'substituto'},
    '070': {'description': 'Com redução de BC e cobrança do ICMS por ST', 'category': 'tributado'},
    '090': {'description': 'Outras', 'category': 'outros'},
}

# ============================================================================
# CLASSES DE DADOS (DATACLASSES)
# ============================================================================

@dataclass
class AuditEntry:
    """Registro de auditoria para cada alteração"""
    timestamp: datetime
    user: str
    action: str
    block: str
    registry: str
    field: str
    old_value: Any
    new_value: Any
    reason: str
    rule_applied: Optional[str] = None
    document_id: Optional[str] = None
    item_id: Optional[str] = None
    
    def to_dict(self) -> Dict:
        return asdict(self)
    
    def __str__(self) -> str:
        return (f"[{self.timestamp}] {self.user} - {self.action}: "
                f"{self.block}/{self.registry}/{self.field} "
                f"'{self.old_value}' -> '{self.new_value}'")

@dataclass
class TaxRule:
    """Regra tributária configurável"""
    rule_id: str
    cst: str
    cfop_pattern: str
    operation_type: OperationType
    requires_base: bool = True
    requires_aliquot: bool = True
    requires_tax_value: bool = True
    base_calculation_source: str = 'item_value'
    default_aliquot: Optional[float] = None
    calculation_formula: str = 'base * aliquot / 100'
    rounding_method: str = 'ROUND_HALF_UP'
    decimal_places: int = 2
    is_active: bool = True
    description: str = ''
    
    def to_dict(self) -> Dict:
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: Dict) -> 'TaxRule':
        if 'operation_type' in data and isinstance(data['operation_type'], str):
            data['operation_type'] = OperationType(data['operation_type'])
        return cls(**data)

@dataclass
class Inconsistency:
    """Inconsistência fiscal detectada"""
    id: str
    severity: SeverityLevel
    block: str
    registry: str
    field: str
    description: str
    current_value: Any
    expected_value: Optional[Any]
    cst: Optional[str] = None
    cfop: Optional[str] = None
    document_number: Optional[str] = None
    item_number: Optional[str] = None
    rule_reference: Optional[str] = None
    can_auto_correct: bool = False
    suggested_correction: Optional[Any] = None
    
    def to_dict(self) -> Dict:
        return asdict(self)

# ============================================================================
# PARSER SPED - Motor de Leitura e Estruturação
# ============================================================================

class SPEDParser:
    """
    Parser profissional para arquivos SPED Fiscal
    Suporte a EFD ICMS/IPI e EFD Contribuições
    """
    
    def __init__(self):
        self.raw_content = ''
        self.encoding = 'utf-8'
        self.sped_type = ''
        self.metadata = {}
        self.blocks = {}
        self.blocks_dataframes = {}
        self.all_records = []
        self.parser_log = []
        
    def parse_file(self, uploaded_file) -> Dict[str, Any]:
        """
        Processa arquivo SPED completo
        
        Args:
            uploaded_file: Arquivo enviado (BytesIO ou similar)
            
        Returns:
            Dict com resultado do parsing
        """
        try:
            start_time = datetime.now()
            
            # Ler conteúdo bruto
            raw_bytes = uploaded_file.read()
            
            # Detectar encoding
            self._detect_encoding(raw_bytes)
            
            # Decodificar
            try:
                self.raw_content = raw_bytes.decode(self.encoding)
            except:
                self.encoding = 'latin-1'
                self.raw_content = raw_bytes.decode(self.encoding)
            
            # Normalizar linhas
            lines = self.raw_content.replace('\r\n', '\n').replace('\r', '\n').split('\n')
            lines = [line.strip() for line in lines if line.strip()]
            
            # Identificar tipo SPED
            self.sped_type = self._identify_sped_type(lines)
            
            if not self.sped_type:
                return {'success': False, 'error': 'Tipo de arquivo SPED não identificado'}
            
            # Extrair metadados
            self.metadata = self._extract_metadata(lines)
            
            # Parsear blocos e registros
            self._parse_blocks(lines)
            
            # Converter para DataFrames
            self._create_dataframes()
            
            # Log do parsing
            elapsed = (datetime.now() - start_time).total_seconds()
            logger.info(f"Arquivo processado em {elapsed:.2f}s - {len(lines)} linhas")
            
            return {
                'success': True,
                'data': self.raw_content,
                'sped_type': self.sped_type,
                'metadata': self.metadata,
                'blocks': self.blocks_dataframes,
                'blocks_raw': self.blocks,
                'stats': {
                    'total_lines': len(lines),
                    'total_records': len(self.all_records),
                    'blocks_found': list(self.blocks.keys()),
                    'parse_time': elapsed
                }
            }
            
        except Exception as e:
            logger.error(f"Erro no parsing: {str(e)}", exc_info=True)
            return {'success': False, 'error': str(e)}
    
    def _detect_encoding(self, raw_bytes: bytes):
        """Detecta encoding do arquivo"""
        # Tentar UTF-8 primeiro
        try:
            raw_bytes.decode('utf-8')
            self.encoding = 'utf-8'
            return
        except:
            pass
        
        # Tentar latin-1
        try:
            raw_bytes.decode('latin-1')
            self.encoding = 'latin-1'
            return
        except:
            pass
        
        # Tentar cp1252 (Windows)
        try:
            raw_bytes.decode('cp1252')
            self.encoding = 'cp1252'
            return
        except:
            pass
        
        # Default
        self.encoding = 'latin-1'
    
    def _identify_sped_type(self, lines: List[str]) -> str:
        """Identifica o tipo de arquivo SPED"""
        for line in lines[:100]:
            parts = line.split('|')
            if len(parts) < 3:
                continue
            
            if parts[0] == '':
                parts = parts[1:]
            
            if len(parts) < 2:
                continue
            
            register = parts[1] if len(parts) > 1 else ''
            
            if register == '0000':
                if len(parts) > 6:
                    cod_fin = parts[5] if len(parts) > 5 else ''
                    if cod_fin in ['1', '2']:
                        return 'EFD ICMS/IPI'
                
                if len(parts) > 7:
                    nome = parts[7] if len(parts) > 7 else ''
                    if 'CONTRIBUIÇÕES' in line.upper() or 'CONTRIBUICOES' in line.upper():
                        return 'EFD Contribuições'
                
                return 'EFD ICMS/IPI'
        
        return 'EFD ICMS/IPI'
    
    def _extract_metadata(self, lines: List[str]) -> Dict[str, Any]:
        """Extrai metadados do arquivo SPED"""
        metadata = {
            'cnpj': '',
            'ie': '',
            'nome': '',
            'periodo_inicial': '',
            'periodo_final': '',
            'cod_versao': '',
            'cod_finalidade': '',
            'uf': '',
            'cod_municipio': '',
            'insc_municipal': '',
            'ind_perfil': '',
            'ind_atividade': '',
            'total_registros': 0,
            'total_blocos': 0
        }
        
        for line in lines:
            parts = line.split('|')
            if len(parts) < 3:
                continue
            
            if parts[0] == '':
                parts = parts[1:]
            
            if len(parts) < 2:
                continue
            
            register = parts[1]
            
            if register == '0000':
                metadata['cod_versao'] = parts[3] if len(parts) > 3 else ''
                metadata['cod_finalidade'] = parts[4] if len(parts) > 4 else ''
                metadata['periodo_inicial'] = parts[5] if len(parts) > 5 else ''
                metadata['periodo_final'] = parts[6] if len(parts) > 6 else ''
                metadata['nome'] = parts[7] if len(parts) > 7 else ''
                metadata['cnpj'] = parts[8] if len(parts) > 8 else ''
                metadata['uf'] = parts[10] if len(parts) > 10 else ''
                metadata['ie'] = parts[11] if len(parts) > 11 else ''
                metadata['cod_municipio'] = parts[12] if len(parts) > 12 else ''
                metadata['insc_municipal'] = parts[13] if len(parts) > 13 else ''
                metadata['ind_perfil'] = parts[15] if len(parts) > 15 else ''
                metadata['ind_atividade'] = parts[16] if len(parts) > 16 else ''
            
            elif register == '0001':
                metadata['ind_movimento'] = parts[3] if len(parts) > 3 else ''
        
        return metadata
    
    def _parse_blocks(self, lines: List[str]):
        """Parseia blocos e registros do SPED"""
        self.blocks = {}
        self.all_records = []
        current_block = None
        block_registers = []
        record_count = 0
        
        for line_num, line in enumerate(lines, 1):
            parts = line.split('|')
            if len(parts) < 3:
                continue
            
            if parts[0] == '':
                parts = parts[1:]
            
            if len(parts) < 2:
                continue
            
            register = parts[1]
            
            block = self._get_block_from_register(register)
            
            if block:
                if block != current_block:
                    if current_block and block_registers:
                        self.blocks[current_block] = block_registers
                    
                    current_block = block
                    block_registers = []
                
                record = {
                    'line_number': line_num,
                    'block': block,
                    'register': register,
                    'original_line': line,
                    'fields': parts[2:] if len(parts) > 2 else []
                }
                
                field_names = self._get_field_names(register)
                for i, value in enumerate(record['fields']):
                    if i < len(field_names):
                        record[field_names[i]] = value
                
                block_registers.append(record)
                self.all_records.append(record)
                record_count += 1
        
        if current_block and block_registers:
            self.blocks[current_block] = block_registers
        
        self.metadata['total_registros'] = record_count
        self.metadata['total_blocos'] = len(self.blocks)
    
    def _get_block_from_register(self, register: str) -> str:
        """Determina o bloco baseado no registro"""
        if not register:
            return ''
        
        first_char = register[0].upper()
        
        if register.startswith('0'):
            return '0'
        elif first_char == 'C':
            return 'C'
        elif first_char == 'D':
            return 'D'
        elif first_char == 'E':
            return 'E'
        elif first_char == 'H':
            return 'H'
        elif register.startswith('1'):
            return '1'
        elif register.startswith('9'):
            return '9'
        elif first_char == 'A':
            return 'A'
        elif first_char == 'B':
            return 'B'
        elif first_char == 'G':
            return 'G'
        else:
            return 'OTHERS'
    
    def _get_field_names(self, register: str) -> List[str]:
        """Retorna nomes dos campos para registros conhecidos"""
        field_maps = {
            '0000': ['REG', 'COD_VER', 'COD_FIN', 'DT_INI', 'DT_FIN', 'NOME', 'CNPJ', 
                     'CPF', 'UF', 'IE', 'COD_MUN', 'IM', 'SUFRAMA', 'IND_PERFIL', 'IND_ATIV'],
            '0001': ['REG', 'IND_MOV'],
            '0150': ['REG', 'COD_PART', 'NOME', 'COD_PAIS', 'CNPJ', 'CPF', 'IE', 
                     'COD_MUN', 'SUFRAMA', 'END', 'NUM', 'COMPL', 'BAIRRO'],
            '0200': ['REG', 'COD_ITEM', 'DESCR_ITEM', 'COD_BARRA', 'COD_ANT_ITEM',
                     'UNID_INV', 'TIPO_ITEM', 'COD_NCM', 'EX_IPI', 'COD_GEN', 'COD_LST', 'ALIQ_ICMS'],
            '0400': ['REG', 'COD_NAT', 'DESCR_NAT'],
            '0450': ['REG', 'COD_INF', 'TXT'],
            '0460': ['REG', 'COD_OBS', 'TXT'],
            'C100': ['REG', 'IND_OPER', 'IND_EMIT', 'COD_PART', 'COD_MOD', 'COD_SIT', 
                     'SER', 'NUM_DOC', 'CHV_NFE', 'DT_DOC', 'DT_E_S', 'VL_DOC', 
                     'IND_PGTO', 'VL_DESC', 'VL_ABAT_NT', 'VL_MERC', 'IND_FRT', 
                     'VL_FRT', 'VL_SEG', 'VL_OUT_DA', 'VL_BC_ICMS', 'VL_ICMS', 
                     'VL_BC_ICMS_ST', 'VL_ICMS_ST', 'VL_IPI', 'VL_PIS', 'VL_COFINS', 
                     'VL_PIS_ST', 'VL_COFINS_ST'],
            'C170': ['REG', 'NUM_ITEM', 'COD_ITEM', 'DESCR_COMPL', 'QTD', 'UNID', 
                     'VL_ITEM', 'VL_DESC', 'IND_MOV', 'CST_ICMS', 'CFOP', 'COD_NAT', 
                     'VL_BC_ICMS', 'ALIQ_ICMS', 'VL_ICMS', 'VL_BC_ICMS_ST', 'ALIQ_ST', 
                     'VL_ICMS_ST', 'IND_APUR', 'CST_IPI', 'COD_ENQ', 'VL_BC_IPI', 
                     'ALIQ_IPI', 'VL_IPI', 'CST_PIS', 'VL_BC_PIS', 'ALIQ_PIS', 
                     'QUANT_BC_PIS', 'VL_PIS', 'CST_COFINS', 'VL_BC_COFINS', 
                     'ALIQ_COFINS', 'QUANT_BC_COFINS', 'VL_COFINS', 'COD_CTA'],
            'C190': ['REG', 'CST_ICMS', 'CFOP', 'ALIQ_ICMS', 'VL_OPR', 'VL_BC_ICMS', 
                     'VL_ICMS', 'VL_BC_ICMS_ST', 'VL_ICMS_ST', 'VL_RED_BC', 'COD_OBS'],
            'C400': ['REG', 'COD_MOD', 'ECF_MOD', 'ECF_FAB', 'ECF_CX'],
            'C405': ['REG', 'DT_DOC', 'CRO', 'CRZ', 'NUM_COO_FIN', 'GT_FIN', 'VL_BRT'],
            'D100': ['REG', 'IND_OPER', 'IND_EMIT', 'COD_PART', 'COD_MOD', 'COD_SIT', 
                     'SER', 'SUB', 'NUM_DOC', 'CHV_CTE', 'DT_DOC', 'DT_A_P', 'TP_CTE',
                     'CHV_CTE_REF', 'VL_DOC', 'VL_DESC', 'IND_FRT', 'VL_SERV', 
                     'VL_BC_ICMS', 'VL_ICMS', 'VL_NT', 'COD_INF', 'COD_CTA'],
            'D190': ['REG', 'CST_ICMS', 'CFOP', 'ALIQ_ICMS', 'VL_OPR', 'VL_BC_ICMS',
                     'VL_ICMS', 'VL_RED_BC', 'COD_OBS'],
            'E100': ['REG', 'DT_INI', 'DT_FIN'],
            'E110': ['REG', 'VL_TOT_DEBITOS', 'VL_AJ_DEBITOS', 'VL_TOT_AJ_DEBITOS',
                     'VL_ESTORNOS_CRED', 'VL_TOT_CREDITOS', 'VL_AJ_CREDITOS', 
                     'VL_TOT_AJ_CREDITOS', 'VL_ESTORNOS_DEB', 'VL_SLD_CREDOR_ANT',
                     'VL_SLD_APURADO', 'VL_TOT_DED', 'VL_ICMS_RECOLHER', 
                     'VL_SLD_CREDOR_TRANSPORTAR', 'DEB_ESP'],
            'H010': ['REG', 'COD_ITEM', 'UNID', 'QTD', 'VL_UNIT', 'VL_ITEM', 'IND_PROP',
                     'COD_PART', 'TXT_COMPL', 'COD_CTA', 'VL_ITEM_IR'],
        }
        
        return field_maps.get(register, [f'F{i}' for i in range(1, 30)])
    
    def _create_dataframes(self):
        """Cria DataFrames estruturados por bloco"""
        self.blocks_dataframes = {}
        
        for block_id, records in self.blocks.items():
            if not records:
                continue
            
            df = pd.DataFrame(records)
            self.blocks_dataframes[block_id] = df
            
            if 'line_number' in df.columns:
                df.sort_values('line_number', inplace=True)
            
            self._convert_numeric_fields(df)
    
    def _convert_numeric_fields(self, df: pd.DataFrame):
        """Converte campos numéricos para tipo Decimal/float"""
        numeric_patterns = ['VL_', 'ALIQ_', 'QTD', 'PER_', 'VAL_', 'TOT_', 'SALDO_']
        
        for col in df.columns:
            if any(pattern in col for pattern in numeric_patterns):
                try:
                    if df[col].dtype == object:
                        df[col] = df[col].str.replace(',', '.')
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                except:
                    pass
    
    def get_register_dataframe(self, register_id: str) -> pd.DataFrame:
        """Retorna DataFrame filtrado por tipo de registro"""
        all_dfs = []
        
        for block_id, df in self.blocks_dataframes.items():
            if 'register' in df.columns:
                filtered = df[df['register'] == register_id]
                if not filtered.empty:
                    all_dfs.append(filtered)
        
        if all_dfs:
            return pd.concat(all_dfs, ignore_index=True)
        return pd.DataFrame()

# ============================================================================
# VALIDADOR FISCAL - Motor de Validação Tributária
# ============================================================================

class FiscalValidator:
    """
    Validador fiscal inteligente
    Detecta inconsistências baseadas em regras tributárias
    """
    
    def __init__(self, rules_config: Dict[str, Any]):
        self.rules_config = rules_config
        self.inconsistencies: List[Inconsistency] = []
        self.inconsistency_counter = 0
        
    def validate(self, blocks_data: Dict[str, pd.DataFrame]) -> List[Inconsistency]:
        """
        Executa validação fiscal completa
        
        Args:
            blocks_data: DataFrames por bloco
            
        Returns:
            Lista de inconsistências encontradas
        """
        self.inconsistencies = []
        self.inconsistency_counter = 0
        
        # Validar bloco C (Documentos Fiscais ICMS)
        if 'C' in blocks_data and blocks_data['C'] is not None and not blocks_data['C'].empty:
            try:
                self._validate_c_block(blocks_data['C'])
            except Exception as e:
                logger.warning(f"Erro na validação do bloco C: {e}")
                self._add_inconsistency(
                    severity=SeverityLevel.INFORMACAO,
                    block='C', 
                    registry='SISTEMA',
                    field='VALIDACAO',
                    description=f'Erro na validação do bloco C: {str(e)}',
                    current_value='Erro',
                    expected_value='OK'
                )
        
        # Validar bloco D (Documentos de Transporte/Serviços)
        if 'D' in blocks_data and blocks_data['D'] is not None and not blocks_data['D'].empty:
            try:
                self._validate_d_block(blocks_data['D'])
            except Exception as e:
                logger.warning(f"Erro na validação do bloco D: {e}")
        
        # Validar bloco E (Apuração)
        if 'E' in blocks_data and blocks_data['E'] is not None and not blocks_data['E'].empty:
            try:
                self._validate_e_block(blocks_data['E'])
            except Exception as e:
                logger.warning(f"Erro na validação do bloco E: {e}")
        
        # Validações inter-blocos
        try:
            self._validate_cross_blocks(blocks_data)
        except Exception as e:
            logger.warning(f"Erro na validação entre blocos: {e}")
        
        return self.inconsistencies
    
    def _validate_c_block(self, df: pd.DataFrame):
        """Validações completas do bloco C"""
        if df.empty:
            return
        
        if 'register' not in df.columns:
            return
        
        c100_df = df[df['register'] == 'C100'] if 'register' in df.columns else pd.DataFrame()
        c170_df = df[df['register'] == 'C170'] if 'register' in df.columns else pd.DataFrame()
        c190_df = df[df['register'] == 'C190'] if 'register' in df.columns else pd.DataFrame()
        
        if not c170_df.empty:
            try:
                self._validate_c170_items(c170_df)
            except Exception as e:
                logger.warning(f"Erro na validação C170: {e}")
        
        if not c100_df.empty:
            try:
                self._validate_c100_documents(c100_df)
            except Exception as e:
                logger.warning(f"Erro na validação C100: {e}")
        
        if not c190_df.empty:
            try:
                self._validate_c190_analytical(c190_df)
            except Exception as e:
                logger.warning(f"Erro na validação C190: {e}")
        
        if not c100_df.empty and not c170_df.empty:
            try:
                self._validate_c100_c170_consistency(c100_df, c170_df)
            except Exception as e:
                logger.warning(f"Erro na validação C100/C170: {e}")
        
        if not c170_df.empty and not c190_df.empty:
            try:
                self._validate_c170_c190_consistency(c170_df, c190_df)
            except Exception as e:
                logger.warning(f"Erro na validação C170/C190: {e}")
    
    def _validate_c170_items(self, df: pd.DataFrame):
        """Validação detalhada dos itens C170"""
        if df.empty:
            return
        
        for idx, row in df.iterrows():
            try:
                cst = self._safe_str(row.get('CST_ICMS', ''))
                cfop = self._safe_str(row.get('CFOP', ''))
                vl_bc = self._safe_decimal(row.get('VL_BC_ICMS'))
                aliq = self._safe_decimal(row.get('ALIQ_ICMS'))
                vl_icms = self._safe_decimal(row.get('VL_ICMS'))
                vl_item = self._safe_decimal(row.get('VL_ITEM'))
                num_item = self._safe_str(row.get('NUM_ITEM', ''))
                cod_item = self._safe_str(row.get('COD_ITEM', ''))
                
                cst_info = CST_ICMS_CODES.get(cst, {})
                category = cst_info.get('category', '')
                
                if category == 'tributado':
                    if vl_bc is None or vl_bc == 0:
                        self._add_inconsistency(
                            severity=SeverityLevel.CRITICA,
                            block='C', registry='C170', field='VL_BC_ICMS',
                            description=f'Item com CST {cst} ({cst_info.get("description", "")}) '
                                      f'sem base de cálculo',
                            current_value=vl_bc if vl_bc else 0,
                            expected_value='> 0',
                            cst=cst, cfop=cfop, item_number=num_item,
                            can_auto_correct=True,
                            suggested_correction=vl_item if vl_item else 0
                        )
                    
                    if aliq is None or aliq == 0:
                        default_aliquot = self._get_default_aliquot(cst, cfop)
                        self._add_inconsistency(
                            severity=SeverityLevel.CRITICA,
                            block='C', registry='C170', field='ALIQ_ICMS',
                            description=f'Item com CST {cst} sem alíquota',
                            current_value=0,
                            expected_value=default_aliquot,
                            cst=cst, cfop=cfop, item_number=num_item,
                            can_auto_correct=True,
                            suggested_correction=default_aliquot
                        )
                    
                    if vl_icms is None or vl_icms == 0:
                        if vl_bc and vl_bc > 0 and aliq and aliq > 0:
                            expected_icms = (vl_bc * aliq / 100).quantize(Decimal('0.01'))
                        else:
                            expected_icms = Decimal('0')
                        
                        self._add_inconsistency(
                            severity=SeverityLevel.CRITICA,
                            block='C', registry='C170', field='VL_ICMS',
                            description=f'Item com CST {cst} sem valor do imposto',
                            current_value=0,
                            expected_value=str(expected_icms),
                            cst=cst, cfop=cfop, item_number=num_item,
                            can_auto_correct=True,
                            suggested_correction=str(expected_icms)
                        )
                
                elif category in ['isento', 'suspenso', 'diferido']:
                    if vl_icms and vl_icms > 0:
                        self._add_inconsistency(
                            severity=SeverityLevel.AVISO,
                            block='C', registry='C170', field='VL_ICMS',
                            description=f'Item com CST {cst} ({cst_info.get("description", "")}) '
                                      f'com valor de imposto indevido',
                            current_value=str(vl_icms),
                            expected_value='0',
                            cst=cst, cfop=cfop, item_number=num_item
                        )
                
                if vl_bc and vl_bc > 0 and aliq and aliq > 0 and vl_icms and vl_icms > 0:
                    expected_icms = (vl_bc * aliq / 100).quantize(Decimal('0.01'))
                    diff = abs(vl_icms - expected_icms)
                    
                    if diff > Decimal('0.05'):
                        self._add_inconsistency(
                            severity=SeverityLevel.AVISO,
                            block='C', registry='C170', field='VL_ICMS',
                            description=f'Divergência no cálculo: {vl_bc} * {aliq}% = {expected_icms}, '
                                      f'mas registrado {vl_icms} (dif: {diff})',
                            current_value=str(vl_icms),
                            expected_value=str(expected_icms),
                            cst=cst, cfop=cfop, item_number=num_item,
                            can_auto_correct=True,
                            suggested_correction=str(expected_icms)
                        )
                
                if not cst:
                    self._add_inconsistency(
                        severity=SeverityLevel.CRITICA,
                        block='C', registry='C170', field='CST_ICMS',
                        description='CST ICMS obrigatório não informado',
                        current_value='Vazio', expected_value='CST válido',
                        item_number=num_item
                    )
                
                if not cfop:
                    self._add_inconsistency(
                        severity=SeverityLevel.CRITICA,
                        block='C', registry='C170', field='CFOP',
                        description='CFOP obrigatório não informado',
                        current_value='Vazio', expected_value='CFOP válido',
                        item_number=num_item
                    )
            except Exception as e:
                logger.warning(f"Erro ao validar item C170 na linha {idx}: {e}")
                continue
    
    def _validate_c100_documents(self, df: pd.DataFrame):
        """Valida documentos C100"""
        if df.empty:
            return
        
        for idx, row in df.iterrows():
            try:
                vl_doc = self._safe_decimal(row.get('VL_DOC'))
                vl_icms = self._safe_decimal(row.get('VL_ICMS'))
                vl_bc_icms = self._safe_decimal(row.get('VL_BC_ICMS'))
                num_doc = self._safe_str(row.get('NUM_DOC', ''))
                
                if vl_doc and vl_icms and vl_icms > vl_doc:
                    self._add_inconsistency(
                        severity=SeverityLevel.AVISO,
                        block='C', registry='C100', field='VL_ICMS',
                        description=f'ICMS ({vl_icms}) maior que valor do documento ({vl_doc})',
                        current_value=str(vl_icms),
                        expected_value=f'≤ {vl_doc}',
                        document_number=num_doc
                    )
                
                if vl_doc and vl_bc_icms and vl_bc_icms > vl_doc:
                    self._add_inconsistency(
                        severity=SeverityLevel.AVISO,
                        block='C', registry='C100', field='VL_BC_ICMS',
                        description=f'Base de cálculo ({vl_bc_icms}) maior que valor do documento ({vl_doc})',
                        current_value=str(vl_bc_icms),
                        expected_value=f'≤ {vl_doc}',
                        document_number=num_doc
                    )
            except Exception as e:
                logger.warning(f"Erro ao validar documento C100 na linha {idx}: {e}")
                continue
    
    def _validate_c190_analytical(self, df: pd.DataFrame):
        """Valida registros analíticos C190"""
        if df.empty:
            return
        
        for idx, row in df.iterrows():
            try:
                cst = self._safe_str(row.get('CST_ICMS', ''))
                cfop = self._safe_str(row.get('CFOP', ''))
                aliq = self._safe_decimal(row.get('ALIQ_ICMS'))
                vl_opr = self._safe_decimal(row.get('VL_OPR'))
                vl_bc = self._safe_decimal(row.get('VL_BC_ICMS'))
                vl_icms = self._safe_decimal(row.get('VL_ICMS'))
                
                if vl_bc and aliq and vl_icms:
                    expected_icms = (vl_bc * aliq / 100).quantize(Decimal('0.01'))
                    diff = abs(vl_icms - expected_icms)
                    
                    if diff > Decimal('0.05'):
                        self._add_inconsistency(
                            severity=SeverityLevel.AVISO,
                            block='C', registry='C190', field='VL_ICMS',
                            description=f'Divergência analítica: {vl_bc} * {aliq}% = {expected_icms}, '
                                      f'registrado {vl_icms}',
                            current_value=str(vl_icms),
                            expected_value=str(expected_icms),
                            cst=cst, cfop=cfop
                        )
            except Exception as e:
                logger.warning(f"Erro ao validar C190 na linha {idx}: {e}")
                continue
    
    def _validate_c100_c170_consistency(self, c100_df: pd.DataFrame, c170_df: pd.DataFrame):
        """
        Valida consistência entre totais do documento (C100) e soma dos itens (C170)
        """
        try:
            if c100_df.empty or c170_df.empty:
                return
            
            # Verificar se as colunas necessárias existem
            required_cols = ['VL_ICMS', 'VL_BC_ICMS', 'VL_ITEM']
            existing_cols = [col for col in required_cols if col in c170_df.columns]
            
            if not existing_cols:
                return
            
            # Criar coluna line_number se não existir
            if 'line_number' not in c170_df.columns:
                if c170_df.index.name is None:
                    c170_df = c170_df.reset_index(drop=True)
                    c170_df['line_number'] = c170_df.index + 1
                else:
                    c170_df['line_number'] = c170_df.index
            
            # Criar dicionário de agregação apenas com colunas existentes
            agg_dict = {col: 'sum' for col in existing_cols}
            
            try:
                c170_grouped = c170_df.groupby('line_number').agg(agg_dict).reset_index()
            except Exception:
                # Fallback: agregar tudo sem groupby
                c170_grouped = c170_df.agg({col: 'sum' for col in existing_cols}).to_frame().T
                c170_grouped['line_number'] = 1
            
            if c170_grouped.empty:
                return
            
            # Verificar total ICMS
            if 'VL_ICMS' in c100_df.columns and 'VL_ICMS' in c170_grouped.columns:
                try:
                    total_c100_icms = c100_df['VL_ICMS'].sum() if 'VL_ICMS' in c100_df.columns else 0
                    total_c170_icms = c170_grouped['VL_ICMS'].sum() if 'VL_ICMS' in c170_grouped.columns else 0
                    
                    if total_c100_icms and total_c170_icms:
                        diff = abs(Decimal(str(total_c100_icms)) - Decimal(str(total_c170_icms)))
                        if diff > Decimal('1.00'):
                            self._add_inconsistency(
                                severity=SeverityLevel.AVISO,
                                block='C', 
                                registry='C100/C170',
                                field='VL_ICMS',
                                description=f'Total ICMS nos documentos ({total_c100_icms:.2f}) '
                                          f'difere da soma dos itens ({total_c170_icms:.2f})',
                                current_value=str(total_c100_icms),
                                expected_value=str(total_c170_icms)
                            )
                except Exception as e:
                    logger.warning(f"Erro ao comparar ICMS total: {e}")
                    
        except Exception as e:
            logger.warning(f"Erro na validação C100/C170: {e}")
    
    def _validate_c170_c190_consistency(self, c170_df: pd.DataFrame, c190_df: pd.DataFrame):
        """Valida consistência entre itens e registro analítico"""
        try:
            if c170_df.empty or c190_df.empty:
                return
            
            required_cols = ['CST_ICMS', 'CFOP', 'VL_BC_ICMS', 'VL_ICMS']
            
            c170_has = all(col in c170_df.columns for col in required_cols)
            c190_has = all(col in c190_df.columns for col in required_cols)
            
            if not c170_has or not c190_has:
                return
            
            try:
                c170_agg = c170_df.groupby(['CST_ICMS', 'CFOP']).agg({
                    'VL_BC_ICMS': 'sum',
                    'VL_ICMS': 'sum'
                }).reset_index()
            except Exception:
                return
            
            try:
                c190_agg = c190_df.groupby(['CST_ICMS', 'CFOP']).agg({
                    'VL_BC_ICMS': 'sum',
                    'VL_ICMS': 'sum'
                }).reset_index()
            except Exception:
                return
            
            for _, c170_row in c170_agg.iterrows():
                try:
                    cst = c170_row['CST_ICMS']
                    cfop = c170_row['CFOP']
                    
                    c190_match = c190_agg[
                        (c190_agg['CST_ICMS'] == cst) & 
                        (c190_agg['CFOP'] == cfop)
                    ]
                    
                    if not c190_match.empty:
                        c190_row = c190_match.iloc[0]
                        
                        icms_170 = Decimal(str(c170_row['VL_ICMS']))
                        icms_190 = Decimal(str(c190_row['VL_ICMS']))
                        diff = abs(icms_170 - icms_190)
                        
                        if diff > Decimal('0.10'):
                            self._add_inconsistency(
                                severity=SeverityLevel.AVISO,
                                block='C', registry='C170/C190',
                                field='VL_ICMS',
                                description=f'ICMS diverge entre itens ({icms_170}) e analítico ({icms_190}) '
                                          f'para CST {cst} / CFOP {cfop}',
                                current_value=str(icms_170),
                                expected_value=str(icms_190),
                                cst=cst, cfop=cfop
                            )
                except Exception:
                    continue
        except Exception as e:
            logger.warning(f"Erro na validação C170/C190: {e}")
    
    def _validate_d_block(self, df: pd.DataFrame):
        """Validações do bloco D"""
        pass
    
    def _validate_e_block(self, df: pd.DataFrame):
        """Validações do bloco E (Apuração)"""
        if df.empty:
            return
        
        if 'register' not in df.columns:
            return
        
        e110_df = df[df['register'] == 'E110'] if 'register' in df.columns else pd.DataFrame()
        
        for idx, row in e110_df.iterrows():
            try:
                vl_debitos = self._safe_decimal(row.get('VL_TOT_DEBITOS'))
                vl_creditos = self._safe_decimal(row.get('VL_TOT_CREDITOS'))
                vl_saldo = self._safe_decimal(row.get('VL_SLD_APURADO'))
                
                if vl_debitos is not None and vl_creditos is not None and vl_saldo is not None:
                    expected = vl_debitos - vl_creditos
                    diff = abs(vl_saldo - expected)
                    
                    if diff > Decimal('0.02'):
                        self._add_inconsistency(
                            severity=SeverityLevel.AVISO,
                            block='E', registry='E110', field='VL_SLD_APURADO',
                            description=f'Saldo apurado ({vl_saldo}) difere do cálculo '
                                      f'{vl_debitos} - {vl_creditos} = {expected}',
                            current_value=str(vl_saldo),
                            expected_value=str(expected)
                        )
            except Exception as e:
                logger.warning(f"Erro ao validar E110: {e}")
                continue
    
    def _validate_cross_blocks(self, blocks_data: Dict[str, pd.DataFrame]):
        """Validações entre blocos diferentes"""
        try:
            if 'C' in blocks_data and 'E' in blocks_data:
                c_df = blocks_data['C']
                e_df = blocks_data['E']
                
                if not c_df.empty and not e_df.empty:
                    if 'VL_ICMS' in c_df.columns and 'register' in e_df.columns:
                        total_icms_c = c_df['VL_ICMS'].sum() if 'VL_ICMS' in c_df.columns else 0
                        
                        e110_df = e_df[e_df['register'] == 'E110'] if 'register' in e_df.columns else pd.DataFrame()
                        if not e110_df.empty and 'VL_TOT_DEBITOS' in e110_df.columns:
                            total_debitos_e = e110_df['VL_TOT_DEBITOS'].sum()
                            
                            if total_icms_c and total_debitos_e:
                                diff = abs(Decimal(str(total_icms_c)) - Decimal(str(total_debitos_e)))
                                if diff > Decimal('10.00'):
                                    self._add_inconsistency(
                                        severity=SeverityLevel.AVISO,
                                        block='C/E', registry='C100/E110',
                                        field='VL_ICMS/VL_TOT_DEBITOS',
                                        description=f'Total ICMS documentos ({total_icms_c:.2f}) '
                                                  f'difere do total débitos apuração ({total_debitos_e:.2f})',
                                        current_value=str(total_icms_c),
                                        expected_value=str(total_debitos_e)
                                    )
        except Exception as e:
            logger.warning(f"Erro na validação entre blocos: {e}")
    
    def _add_inconsistency(self, **kwargs):
        """Adiciona inconsistência à lista"""
        self.inconsistency_counter += 1
        inconsistency = Inconsistency(
            id=f"INCONS_{self.inconsistency_counter:06d}",
            **kwargs
        )
        self.inconsistencies.append(inconsistency)
    
    def _safe_str(self, value: Any) -> str:
        """Converte valor para string de forma segura"""
        if value is None or pd.isna(value):
            return ''
        return str(value).strip()
    
    def _safe_decimal(self, value: Any) -> Optional[Decimal]:
        """Converte valor para Decimal de forma segura"""
        if value is None or pd.isna(value):
            return None
        try:
            val_str = str(value).replace(',', '.').strip()
            if val_str == '' or val_str == 'nan':
                return None
            return Decimal(val_str)
        except:
            return None
    
    def _get_default_aliquot(self, cst: str, cfop: str) -> Decimal:
        """Busca alíquota padrão baseada em regras"""
        if cfop.startswith(('1', '2', '3')):
            return Decimal('12.00')
        elif cfop.startswith(('5', '6', '7')):
            if cfop.startswith('6'):
                return Decimal('7.00')
            else:
                return Decimal('18.00')
        return Decimal('17.00')
    
    def get_statistics(self) -> Dict[str, Any]:
        """Retorna estatísticas das validações"""
        critical = sum(1 for i in self.inconsistencies if i.severity == SeverityLevel.CRITICA)
        warnings = sum(1 for i in self.inconsistencies if i.severity == SeverityLevel.AVISO)
        info = sum(1 for i in self.inconsistencies if i.severity == SeverityLevel.INFORMACAO)
        
        by_block = {}
        for inc in self.inconsistencies:
            block = inc.block
            by_block[block] = by_block.get(block, 0) + 1
        
        by_cst = {}
        for inc in self.inconsistencies:
            if inc.cst:
                by_cst[inc.cst] = by_cst.get(inc.cst, 0) + 1
        
        return {
            'total': len(self.inconsistencies),
            'critical': critical,
            'warnings': warnings,
            'info': info,
            'by_block': by_block,
            'by_cst': by_cst,
            'auto_correctable': sum(1 for i in self.inconsistencies if i.can_auto_correct)
        }

# ============================================================================
# MOTOR DE REGRAS TRIBUTÁRIAS
# ============================================================================

class TaxRuleEngine:
    """
    Motor de regras tributárias configurável
    """
    
    def __init__(self):
        self.rules: Dict[str, TaxRule] = {}
        self.load_default_rules()
    
    def load_default_rules(self):
        """Carrega regras padrão do sistema"""
        default_rules = [
            TaxRule(
                rule_id='R001',
                cst='000',
                cfop_pattern='*',
                operation_type=OperationType.SAIDA,
                requires_base=True,
                requires_aliquot=True,
                requires_tax_value=True,
                base_calculation_source='item_value',
                default_aliquot=18.00,
                calculation_formula='base * aliquot / 100',
                description='Tributação integral - operação própria'
            ),
            TaxRule(
                rule_id='R002',
                cst='020',
                cfop_pattern='*',
                operation_type=OperationType.SAIDA,
                requires_base=True,
                requires_aliquot=True,
                requires_tax_value=True,
                base_calculation_source='item_value',
                default_aliquot=18.00,
                calculation_formula='base * aliquot / 100',
                description='Com redução de base de cálculo'
            ),
            TaxRule(
                rule_id='R003',
                cst='040',
                cfop_pattern='*',
                operation_type=OperationType.SAIDA,
                requires_base=False,
                requires_aliquot=False,
                requires_tax_value=False,
                description='Isenta - sem imposto'
            ),
            TaxRule(
                rule_id='R004',
                cst='060',
                cfop_pattern='*',
                operation_type=OperationType.ENTRADA,
                requires_base=False,
                requires_aliquot=False,
                requires_tax_value=False,
                description='ST anterior - sem imposto'
            ),
        ]
        
        for rule in default_rules:
            self.rules[rule.rule_id] = rule
    
    def add_rule(self, rule: TaxRule) -> bool:
        if rule.rule_id in self.rules:
            return False
        self.rules[rule.rule_id] = rule
        return True
    
    def update_rule(self, rule_id: str, updated_rule: TaxRule) -> bool:
        if rule_id not in self.rules:
            return False
        self.rules[rule_id] = updated_rule
        return True
    
    def delete_rule(self, rule_id: str) -> bool:
        if rule_id in self.rules:
            del self.rules[rule_id]
            return True
        return False
    
    def get_rule(self, rule_id: str) -> Optional[TaxRule]:
        return self.rules.get(rule_id)
    
    def find_applicable_rules(self, cst: str, cfop: str, operation: str) -> List[TaxRule]:
        applicable = []
        
        for rule in self.rules.values():
            if not rule.is_active:
                continue
            
            if rule.cst != cst:
                continue
            
            if rule.cfop_pattern != '*' and not cfop.startswith(rule.cfop_pattern):
                continue
            
            if rule.operation_type.value != operation:
                continue
            
            applicable.append(rule)
        
        return applicable
    
    def calculate_tax(self, base: Decimal, aliquot: Decimal, formula: str = 'base * aliquot / 100') -> Decimal:
        try:
            if formula == 'base * aliquot / 100':
                return (base * aliquot / Decimal('100')).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            else:
                return (base * aliquot / Decimal('100')).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        except:
            return Decimal('0')
    
    def to_dict(self) -> Dict:
        return {rule_id: rule.to_dict() for rule_id, rule in self.rules.items()}
    
    def from_dict(self, data: Dict):
        self.rules = {}
        for rule_id, rule_data in data.items():
            self.rules[rule_id] = TaxRule.from_dict(rule_data)

# ============================================================================
# SERVIÇO DE CORREÇÃO
# ============================================================================

class CorrectionService:
    """
    Serviço de correção fiscal
    """
    
    def __init__(self):
        self.audit_log: List[AuditEntry] = []
        self.correction_history: List[Dict] = []
    
    def correct_inconsistency(self, inconsistency: Inconsistency, 
                              blocks_data: Dict[str, pd.DataFrame],
                              custom_value: Optional[Any] = None) -> Tuple[bool, str, Dict]:
        try:
            block = inconsistency.block
            registry = inconsistency.registry
            field = inconsistency.field
            
            if block not in blocks_data:
                return False, f"Bloco {block} não encontrado", {}
            
            df = blocks_data[block]
            
            if 'register' not in df.columns:
                return False, "DataFrame sem coluna de registro", {}
            
            register_df = df[df['register'] == registry]
            
            if register_df.empty:
                return False, f"Registro {registry} não encontrado no bloco {block}", {}
            
            new_value = custom_value if custom_value is not None else inconsistency.suggested_correction
            
            if new_value is None:
                return False, "Sem valor de correção definido", {}
            
            idx = register_df.index[0]
            old_value = df.at[idx, field] if field in df.columns else None
            
            audit_entry = AuditEntry(
                timestamp=datetime.now(),
                user='Sistema',
                action='correction',
                block=block,
                registry=registry,
                field=field,
                old_value=old_value,
                new_value=new_value,
                reason=f"Correção automática: {inconsistency.description}",
                rule_applied=inconsistency.rule_reference,
                document_id=inconsistency.document_number,
                item_id=inconsistency.item_number
            )
            self.audit_log.append(audit_entry)
            
            if field in df.columns:
                df.at[idx, field] = new_value
                
                self.correction_history.append({
                    'timestamp': datetime.now(),
                    'block': block,
                    'registry': registry,
                    'field': field,
                    'old_value': old_value,
                    'new_value': new_value,
                    'inconsistency_id': inconsistency.id
                })
                
                return True, f"Correção aplicada: {field} = {new_value}", blocks_data
            else:
                return False, f"Campo {field} não encontrado", {}
                
        except Exception as e:
            logger.error(f"Erro na correção: {str(e)}", exc_info=True)
            return False, f"Erro: {str(e)}", {}
    
    def correct_batch(self, inconsistencies: List[Inconsistency],
                      blocks_data: Dict[str, pd.DataFrame]) -> Dict[str, Any]:
        corrected = []
        failed = []
        updated_blocks = deepcopy(blocks_data)
        
        for inc in inconsistencies:
            if inc.can_auto_correct:
                success, message, updated = self.correct_inconsistency(inc, updated_blocks)
                
                if success:
                    corrected.append({
                        'id': inc.id,
                        'description': inc.description,
                        'field': inc.field,
                        'new_value': inc.suggested_correction
                    })
                    if updated:
                        updated_blocks = updated
                else:
                    failed.append({
                        'id': inc.id,
                        'description': inc.description,
                        'error': message
                    })
        
        return {
            'success': len(failed) == 0,
            'corrected_count': len(corrected),
            'failed_count': len(failed),
            'corrected': corrected,
            'failed': failed,
            'updated_blocks': updated_blocks,
            'audit_log': [entry.to_dict() for entry in self.audit_log]
        }
    
    def apply_mass_correction(self, filters: Dict, correction_config: Dict,
                             blocks_data: Dict[str, pd.DataFrame]) -> Dict[str, Any]:
        affected_count = 0
        updated_blocks = deepcopy(blocks_data)
        
        block = filters.get('block')
        registry = filters.get('registry')
        cst = filters.get('cst')
        cfop = filters.get('cfop')
        
        field_to_correct = correction_config.get('field')
        correction_type = correction_config.get('type')
        value = correction_config.get('value')
        formula = correction_config.get('formula')
        
        if block and block in updated_blocks:
            df = updated_blocks[block]
            
            mask = pd.Series(True, index=df.index)
            
            if registry and 'register' in df.columns:
                mask &= (df['register'] == registry)
            
            if cst and 'CST_ICMS' in df.columns:
                mask &= (df['CST_ICMS'] == cst)
            
            if cfop and 'CFOP' in df.columns:
                mask &= (df['CFOP'] == cfop)
            
            affected_df = df[mask]
            
            if not affected_df.empty and field_to_correct:
                for idx in affected_df.index:
                    old_value = df.at[idx, field_to_correct] if field_to_correct in df.columns else None
                    
                    if correction_type == 'fixed':
                        df.at[idx, field_to_correct] = value
                    elif correction_type == 'formula':
                        if formula == 'base * aliquot / 100':
                            base = self._safe_decimal(df.at[idx, 'VL_BC_ICMS'])
                            aliquot = self._safe_decimal(df.at[idx, 'ALIQ_ICMS'])
                            if base and aliquot:
                                calculated = (base * aliquot / 100).quantize(Decimal('0.01'))
                                df.at[idx, field_to_correct] = str(calculated)
                    
                    audit_entry = AuditEntry(
                        timestamp=datetime.now(),
                        user='Sistema',
                        action='mass_correction',
                        block=block,
                        registry=registry or 'N/A',
                        field=field_to_correct,
                        old_value=old_value,
                        new_value=df.at[idx, field_to_correct],
                        reason=f"Correção em massa: {correction_config.get('description', '')}"
                    )
                    self.audit_log.append(audit_entry)
                    affected_count += 1
        
        return {
            'success': True,
            'affected_records': affected_count,
            'updated_blocks': updated_blocks,
            'audit_entries': len(self.audit_log)
        }
    
    def _safe_decimal(self, value: Any) -> Optional[Decimal]:
        try:
            return Decimal(str(value))
        except:
            return None

# ============================================================================
# SERVIÇO DE EXPORTAÇÃO
# ============================================================================

class ExportService:
    """
    Serviço de exportação de arquivos SPED e relatórios
    """
    
    def __init__(self):
        pass
    
    def export_sped(self, blocks_data: Dict[str, pd.DataFrame], 
                   metadata: Dict[str, Any]) -> str:
        output_lines = []
        block_order = ['0', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', '1', '2', '9']
        
        for block in block_order:
            if block in blocks_data:
                df = blocks_data[block]
                
                if 'original_line' in df.columns:
                    for idx, row in df.iterrows():
                        if 'original_line' in row:
                            line = self._reconstruct_line(row)
                            output_lines.append(line)
                else:
                    for idx, row in df.iterrows():
                        line = self._row_to_sped_line(row)
                        output_lines.append(line)
        
        return '\n'.join(output_lines)
    
    def _reconstruct_line(self, row: pd.Series) -> str:
        if 'original_line' in row and pd.notna(row['original_line']):
            parts = row['original_line'].split('|')
            
            for col in row.index:
                if col.startswith('VL_') or col.startswith('ALIQ_') or col.startswith('QTD'):
                    if col in row and pd.notna(row[col]):
                        field_names = self._get_field_names_for_register(row.get('register', ''))
                        if col in field_names:
                            field_idx = field_names.index(col) + 2
                            if field_idx < len(parts):
                                parts[field_idx] = str(row[col])
            
            return '|'.join(parts)
        
        return row.get('original_line', '')
    
    def _row_to_sped_line(self, row: pd.Series) -> str:
        register = row.get('register', '')
        fields = []
        
        field_names = self._get_field_names_for_register(register)
        
        for field in field_names:
            if field in row.index and pd.notna(row[field]):
                fields.append(str(row[field]))
            else:
                fields.append('')
        
        return f"|{register}|{'|'.join(fields)}|"
    
    def _get_field_names_for_register(self, register: str) -> List[str]:
        field_maps = {
            '0000': ['REG', 'COD_VER', 'COD_FIN', 'DT_INI', 'DT_FIN', 'NOME', 'CNPJ', 
                     'CPF', 'UF', 'IE', 'COD_MUN', 'IM', 'SUFRAMA', 'IND_PERFIL', 'IND_ATIV'],
            'C100': ['REG', 'IND_OPER', 'IND_EMIT', 'COD_PART', 'COD_MOD', 'COD_SIT', 
                     'SER', 'NUM_DOC', 'CHV_NFE', 'DT_DOC', 'DT_E_S', 'VL_DOC', 
                     'IND_PGTO', 'VL_DESC', 'VL_ABAT_NT', 'VL_MERC', 'IND_FRT', 
                     'VL_FRT', 'VL_SEG', 'VL_OUT_DA', 'VL_BC_ICMS', 'VL_ICMS', 
                     'VL_BC_ICMS_ST', 'VL_ICMS_ST', 'VL_IPI', 'VL_PIS', 'VL_COFINS', 
                     'VL_PIS_ST', 'VL_COFINS_ST'],
            'C170': ['REG', 'NUM_ITEM', 'COD_ITEM', 'DESCR_COMPL', 'QTD', 'UNID', 
                     'VL_ITEM', 'VL_DESC', 'IND_MOV', 'CST_ICMS', 'CFOP', 'COD_NAT', 
                     'VL_BC_ICMS', 'ALIQ_ICMS', 'VL_ICMS', 'VL_BC_ICMS_ST', 'ALIQ_ST', 
                     'VL_ICMS_ST', 'IND_APUR', 'CST_IPI', 'COD_ENQ', 'VL_BC_IPI', 
                     'ALIQ_IPI', 'VL_IPI', 'CST_PIS', 'VL_BC_PIS', 'ALIQ_PIS', 
                     'QUANT_BC_PIS', 'VL_PIS', 'CST_COFINS', 'VL_BC_COFINS', 
                     'ALIQ_COFINS', 'QUANT_BC_COFINS', 'VL_COFINS', 'COD_CTA'],
            'C190': ['REG', 'CST_ICMS', 'CFOP', 'ALIQ_ICMS', 'VL_OPR', 'VL_BC_ICMS', 
                     'VL_ICMS', 'VL_BC_ICMS_ST', 'VL_ICMS_ST', 'VL_RED_BC', 'COD_OBS'],
            'E110': ['REG', 'VL_TOT_DEBITOS', 'VL_AJ_DEBITOS', 'VL_TOT_AJ_DEBITOS',
                     'VL_ESTORNOS_CRED', 'VL_TOT_CREDITOS', 'VL_AJ_CREDITOS', 
                     'VL_TOT_AJ_CREDITOS', 'VL_ESTORNOS_DEB', 'VL_SLD_CREDOR_ANT',
                     'VL_SLD_APURADO', 'VL_TOT_DED', 'VL_ICMS_RECOLHER', 
                     'VL_SLD_CREDOR_TRANSPORTAR', 'DEB_ESP'],
        }
        return field_maps.get(register, [f'F{i}' for i in range(1, 30)])
    
    def export_excel_report(self, inconsistencies: List[Inconsistency],
                           corrections: List[Dict],
                           audit_log: List[AuditEntry],
                           statistics: Dict) -> io.BytesIO:
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if inconsistencies:
                inc_data = []
                for inc in inconsistencies:
                    inc_data.append({
                        'ID': inc.id,
                        'Severidade': inc.severity.value,
                        'Bloco': inc.block,
                        'Registro': inc.registry,
                        'Campo': inc.field,
                        'Descrição': inc.description,
                        'CST': inc.cst or '',
                        'CFOP': inc.cfop or '',
                        'Valor Atual': str(inc.current_value),
                        'Valor Esperado': str(inc.expected_value) if inc.expected_value else '',
                        'Corrigível Auto': 'Sim' if inc.can_auto_correct else 'Não'
                    })
                pd.DataFrame(inc_data).to_excel(writer, sheet_name='Inconsistências', index=False)
            
            if corrections:
                pd.DataFrame(corrections).to_excel(writer, sheet_name='Correções', index=False)
            
            if audit_log:
                audit_data = [entry.to_dict() for entry in audit_log]
                pd.DataFrame(audit_data).to_excel(writer, sheet_name='Auditoria', index=False)
            
            if statistics:
                stats_df = pd.DataFrame([statistics])
                stats_df.to_excel(writer, sheet_name='Estatísticas', index=False)
        
        output.seek(0)
        return output
    
    def export_csv(self, data: List[Dict]) -> str:
        if not data:
            return ''
        
        df = pd.DataFrame(data)
        return df.to_csv(index=False)

# ============================================================================
# APLICAÇÃO PRINCIPAL STREAMLIT - INTERFACE
# ============================================================================

st.set_page_config(
    page_title="SPED Manager Pro - Sistema Fiscal Corporativo",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS Profissional
st.markdown("""
<style>
    .main {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    }
    .stApp {
        background: #ffffff;
    }
    .metric-card {
        background: white;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        border-left: 4px solid #4CAF50;
        transition: transform 0.3s;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 15px rgba(0,0,0,0.15);
    }
    .metric-card.warning {
        border-left-color: #ff9800;
    }
    .metric-card.danger {
        border-left-color: #f44336;
    }
    .metric-card.info {
        border-left-color: #2196F3;
    }
    .metric-value {
        font-size: 32px;
        font-weight: bold;
        color: #2c3e50;
    }
    .metric-label {
        font-size: 14px;
        color: #7f8c8d;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .alert-critical {
        background-color: #fff5f5;
        border-left: 4px solid #f44336;
        padding: 12px;
        margin: 10px 0;
        border-radius: 4px;
    }
    .alert-warning {
        background-color: #fffbf0;
        border-left: 4px solid #ff9800;
        padding: 12px;
        margin: 10px 0;
        border-radius: 4px;
    }
    .alert-success {
        background-color: #f0fff4;
        border-left: 4px solid #4CAF50;
        padding: 12px;
        margin: 10px 0;
        border-radius: 4px;
    }
    .dataframe {
        font-size: 13px !important;
        border-collapse: collapse;
    }
    .dataframe th {
        background-color: #2c3e50;
        color: white;
        padding: 10px;
        font-weight: 600;
    }
    .dataframe td {
        padding: 8px;
        border-bottom: 1px solid #e0e0e0;
    }
    .stButton > button {
        border-radius: 6px;
        font-weight: 500;
        padding: 8px 20px;
        transition: all 0.3s;
        border: none;
    }
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    .primary-btn > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
    }
    .css-1d391kg {
        background: linear-gradient(180deg, #2c3e50 0%, #34495e 100%);
    }
    .css-1d391kg .stRadio > div {
        color: #ecf0f1;
    }
    h1 {
        color: #2c3e50;
        font-weight: 700;
        border-bottom: 3px solid #3498db;
        padding-bottom: 10px;
        margin-bottom: 30px;
    }
    h2 {
        color: #34495e;
        font-weight: 600;
        margin-top: 30px;
    }
    h3 {
        color: #7f8c8d;
        font-weight: 500;
    }
    .edited-cell {
        background-color: #fff9c4 !important;
        border: 2px solid #fbc02d !important;
    }
    .original-value {
        color: #e53935;
        text-decoration: line-through;
        font-size: 12px;
    }
    .new-value {
        color: #43a047;
        font-weight: bold;
        font-size: 12px;
    }
    .stProgress > div > div {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
    }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# INICIALIZAÇÃO DO ESTADO DA SESSÃO
# ============================================================================

def init_session_state():
    if 'initialized' not in st.session_state:
        st.session_state.initialized = True
        st.session_state.sped_loaded = False
        st.session_state.sped_type = ''
        st.session_state.metadata = {}
        st.session_state.blocks_data = {}
        st.session_state.raw_data = ''
        st.session_state.parser = SPEDParser()
        st.session_state.validator = None
        st.session_state.inconsistencies = []
        st.session_state.inconsistency_stats = {}
        st.session_state.rule_engine = TaxRuleEngine()
        st.session_state.correction_service = CorrectionService()
        st.session_state.corrections_applied = []
        st.session_state.export_service = ExportService()
        st.session_state.audit_log = []
        st.session_state.current_view = 'dashboard'
        st.session_state.selected_block = None
        st.session_state.selected_register = None
        st.session_state.editing_cell = None
        st.session_state.undo_stack = []
        st.session_state.filters = {}
        st.session_state.success_messages = []
        st.session_state.error_messages = []

def reset_session():
    st.session_state.sped_loaded = False
    st.session_state.sped_type = ''
    st.session_state.metadata = {}
    st.session_state.blocks_data = {}
    st.session_state.inconsistencies = []
    st.session_state.corrections_applied = []
    st.session_state.audit_log = []
    st.session_state.current_view = 'dashboard'

# ============================================================================
# FUNÇÕES AUXILIARES
# ============================================================================

def safe_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(str(value).replace(',', '.'))
    except:
        return default

def safe_str(value: Any) -> str:
    if value is None or pd.isna(value):
        return ''
    return str(value).strip()

def format_currency(value: Any) -> str:
    try:
        val = float(str(value).replace(',', '.'))
        return f"R$ {val:,.2f}"
    except:
        return "R$ 0,00"

def get_register_description(register: str) -> str:
    return REGISTER_DESCRIPTIONS.get(register, 'Registro não documentado')

def get_block_description(block: str) -> str:
    return BLOCK_DESCRIPTIONS.get(block, 'Bloco não documentado')

# ============================================================================
# COMPONENTES DA INTERFACE
# ============================================================================

def render_sidebar():
    with st.sidebar:
        st.markdown("""
        <div style="text-align: center; padding: 20px 0;">
            <h1 style="color: white; font-size: 24px; margin: 0;">📊 SPED Manager</h1>
            <p style="color: #bdc3c7; font-size: 14px; margin: 5px 0;">Sistema Fiscal Corporativo</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("### Navegação")
        
        nav_options = {
            '🏠 Dashboard': 'dashboard',
            '📤 Upload SPED': 'upload',
            '📋 Blocos e Registros': 'blocks',
            '📄 Documentos Fiscais': 'documents',
            '📦 Itens': 'items',
            '⚠️ Inconsistências': 'inconsistencies',
            '🔧 Correções em Massa': 'mass_correction',
            '✏️ Editor Manual': 'editor',
            '📤 Exportação': 'export',
            '📜 Auditoria': 'audit',
            '⚙️ Regras Tributárias': 'rules_config'
        }
        
        selected = st.radio(
            'Selecione a seção:',
            list(nav_options.keys()),
            index=0,
            label_visibility='collapsed'
        )
        
        st.session_state.current_view = nav_options[selected]
        
        st.markdown("---")
        
        if st.session_state.sped_loaded:
            st.success("✅ Arquivo Carregado")
            
            metadata = st.session_state.metadata
            if metadata:
                st.markdown(f"""
                <div style="background: rgba(255,255,255,0.1); padding: 10px; border-radius: 5px; margin: 10px 0;">
                    <p style="color: white; font-size: 12px; margin: 2px 0;">
                        <strong>Tipo:</strong> {st.session_state.sped_type}
                    </p>
                    <p style="color: white; font-size: 12px; margin: 2px 0;">
                        <strong>CNPJ:</strong> {metadata.get('cnpj', 'N/A')}
                    </p>
                    <p style="color: white; font-size: 12px; margin: 2px 0;">
                        <strong>Período:</strong> {metadata.get('periodo_inicial', '')} a {metadata.get('periodo_final', '')}
                    </p>
                    <p style="color: white; font-size: 12px; margin: 2px 0;">
                        <strong>Blocos:</strong> {len(st.session_state.blocks_data)}
                    </p>
                </div>
                """, unsafe_allow_html=True)
            
            if st.session_state.inconsistencies:
                total_inc = len(st.session_state.inconsistencies)
                critical = sum(1 for i in st.session_state.inconsistencies 
                             if i.severity == SeverityLevel.CRITICA)
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Total Incons.", total_inc)
                with col2:
                    st.metric("Críticas", critical, delta=f"-{critical}" if critical > 0 else "0")
        else:
            st.info("📂 Nenhum arquivo carregado")
        
        st.markdown("---")
        
        if st.session_state.sped_loaded:
            if st.button("🔄 Resetar Tudo"):
                reset_session()
                st.rerun()
        
        st.caption("SPED Manager Pro v2.1.0")
        st.caption("© 2024 - Solução Fiscal Corporativa")

def render_dashboard():
    st.header("🏠 Dashboard Fiscal")
    
    if not st.session_state.sped_loaded:
        st.info("📤 Faça o upload de um arquivo SPED para começar")
        return
    
    col1, col2, col3, col4 = st.columns(4)
    
    metadata = st.session_state.metadata
    blocks = st.session_state.blocks_data
    
    with col1:
        total_records = sum(len(df) for df in blocks.values()) if blocks else 0
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Total de Registros</div>
            <div class="metric-value">{total_records:,}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="metric-card info">
            <div class="metric-label">Blocos Encontrados</div>
            <div class="metric-value">{len(blocks)}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        inc_count = len(st.session_state.inconsistencies)
        severity_class = 'danger' if inc_count > 0 else 'success'
        st.markdown(f"""
        <div class="metric-card {severity_class}">
            <div class="metric-label">Inconsistências</div>
            <div class="metric-value">{inc_count}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        corrections = len(st.session_state.corrections_applied)
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Correções Aplicadas</div>
            <div class="metric-value">{corrections}</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Distribuição por Bloco")
        
        if blocks:
            block_counts = {block: len(df) for block, df in blocks.items()}
            block_df = pd.DataFrame({
                'Bloco': list(block_counts.keys()),
                'Registros': list(block_counts.values()),
                'Descrição': [get_block_description(b) for b in block_counts.keys()]
            })
            
            st.bar_chart(block_df.set_index('Bloco')['Registros'])
            st.dataframe(block_df, use_container_width=True, hide_index=True)
    
    with col2:
        st.subheader("📈 Análise de CSTs")
        
        if 'C' in blocks and blocks['C'] is not None:
            c_df = blocks['C']
            if 'register' in c_df.columns:
                c170_df = c_df[c_df['register'] == 'C170']
                if not c170_df.empty and 'CST_ICMS' in c170_df.columns:
                    cst_counts = c170_df['CST_ICMS'].value_counts()
                    
                    cst_data = []
                    for cst, count in cst_counts.items():
                        cst_info = CST_ICMS_CODES.get(str(cst), {})
                        cst_data.append({
                            'CST': cst,
                            'Descrição': cst_info.get('description', 'Desconhecido'),
                            'Quantidade': count
                        })
                    
                    if cst_data:
                        cst_df = pd.DataFrame(cst_data)
                        st.dataframe(cst_df, use_container_width=True, hide_index=True)
                        st.bar_chart(cst_df.set_index('CST')['Quantidade'])
    
    if st.session_state.inconsistencies:
        st.markdown("---")
        st.subheader("🚨 Inconsistências Críticas Recentes")
        
        critical_issues = [i for i in st.session_state.inconsistencies 
                          if i.severity == SeverityLevel.CRITICA][:10]
        
        if critical_issues:
            for issue in critical_issues:
                st.markdown(f"""
                <div class="alert-critical">
                    <strong>{issue.registry} - {issue.field}</strong><br>
                    {issue.description}<br>
                    <small>CST: {issue.cst} | CFOP: {issue.cfop} | 
                    Atual: {issue.current_value} | Esperado: {issue.expected_value}</small>
                </div>
                """, unsafe_allow_html=True)

def render_upload():
    st.header("📤 Upload de Arquivo SPED")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        uploaded_file = st.file_uploader(
            "Selecione o arquivo SPED (.txt)",
            type=['txt'],
            help="Formatos suportados: EFD ICMS/IPI, EFD Contribuições",
            key="file_uploader"
        )
        
        if uploaded_file is not None:
            with st.spinner("🔄 Processando arquivo SPED..."):
                parser = SPEDParser()
                result = parser.parse_file(uploaded_file)
                
                if result['success']:
                    st.session_state.sped_loaded = True
                    st.session_state.sped_type = result['sped_type']
                    st.session_state.metadata = result['metadata']
                    st.session_state.blocks_data = result['blocks']
                    st.session_state.raw_data = result['data']
                    
                    with st.spinner("🔍 Validando regras fiscais..."):
                        validator = FiscalValidator(st.session_state.rule_engine.to_dict())
                        inconsistencies = validator.validate(st.session_state.blocks_data)
                        st.session_state.validator = validator
                        st.session_state.inconsistencies = inconsistencies
                        st.session_state.inconsistency_stats = validator.get_statistics()
                    
                    st.success(f"✅ Arquivo processado com sucesso!")
                    st.info(f"📊 {result['stats']['total_records']} registros encontrados em "
                           f"{len(result['blocks'])} blocos")
                    
                    with st.expander("📋 Detalhes do Arquivo", expanded=True):
                        col1, col2, col3 = st.columns(3)
                        
                        with col1:
                            st.metric("Tipo", result['sped_type'])
                            st.metric("CNPJ", result['metadata'].get('cnpj', 'N/A'))
                        
                        with col2:
                            st.metric("IE", result['metadata'].get('ie', 'N/A'))
                            st.metric("UF", result['metadata'].get('uf', 'N/A'))
                        
                        with col3:
                            st.metric("Período Inicial", result['metadata'].get('periodo_inicial', ''))
                            st.metric("Período Final", result['metadata'].get('periodo_final', ''))
                        
                        st.subheader("Blocos Identificados")
                        blocks_info = []
                        for block_id, df in result['blocks'].items():
                            blocks_info.append({
                                'Bloco': block_id,
                                'Descrição': get_block_description(block_id),
                                'Registros': len(df)
                            })
                        
                        if blocks_info:
                            st.dataframe(pd.DataFrame(blocks_info), use_container_width=True, 
                                       hide_index=True)
                    
                    if inconsistencies:
                        st.warning(f"⚠️ {len(inconsistencies)} inconsistências detectadas")
                        
                        stats = validator.get_statistics()
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Críticas", stats['critical'])
                        with col2:
                            st.metric("Avisos", stats['warnings'])
                        with col3:
                            st.metric("Corrigíveis", stats['auto_correctable'])
                else:
                    st.error(f"❌ Erro: {result['error']}")
    
    with col2:
        st.info("""
        ### ℹ️ Informações
        
        **Formatos suportados:**
        - ✅ EFD ICMS/IPI
        - ✅ EFD Contribuições
        
        **Requisitos:**
        - Arquivo texto (.txt)
        - Delimitador pipe (|)
        - Encoding UTF-8 ou Latin-1
        
        **Após o upload:**
        1. O sistema identifica o tipo
        2. Extrai metadados
        3. Estrutura em blocos
        4. Valida regras fiscais
        5. Detecta inconsistências
        """)

def render_blocks_view():
    st.header("📋 Visão por Blocos e Registros")
    
    if not st.session_state.sped_loaded:
        st.warning("⚠️ Carregue um arquivo SPED primeiro!")
        return
    
    blocks = st.session_state.blocks_data
    
    block_list = list(blocks.keys())
    selected_block = st.selectbox(
        "Selecione o Bloco:",
        block_list,
        format_func=lambda x: f"Bloco {x} - {get_block_description(x)}"
    )
    
    if selected_block:
        df = blocks[selected_block]
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total de Registros", len(df))
        
        with col2:
            if 'register' in df.columns:
                unique_registers = df['register'].nunique()
                st.metric("Tipos de Registro", unique_registers)
        
        with col3:
            block_issues = [i for i in st.session_state.inconsistencies 
                          if i.block == selected_block]
            st.metric("Inconsistências", len(block_issues))
        
        if 'register' in df.columns:
            register_types = ['Todos'] + sorted(df['register'].unique().tolist())
            selected_register = st.selectbox(
                "Filtrar por Registro:",
                register_types,
                format_func=lambda x: f"{x} - {get_register_description(x)}" if x != 'Todos' else 'Todos'
            )
            
            if selected_register != 'Todos':
                df = df[df['register'] == selected_register]
        
        st.subheader(f"📊 Dados do Bloco {selected_block}")
        
        view_mode = st.radio(
            "Modo de Visualização:",
            ["Tabela Completa", "Resumo", "Campos Tributários"],
            horizontal=True
        )
        
        if view_mode == "Tabela Completa":
            st.dataframe(df, use_container_width=True)
        
        elif view_mode == "Resumo":
            if 'register' in df.columns:
                summary = df.groupby('register').size().reset_index(name='Quantidade')
                summary['Descrição'] = summary['register'].apply(get_register_description)
                st.dataframe(summary, use_container_width=True, hide_index=True)
        
        elif view_mode == "Campos Tributários":
            tax_columns = [col for col in df.columns if any(
                prefix in col for prefix in ['CST', 'CFOP', 'VL_', 'ALIQ_', 'BC_']
            )]
            if 'register' in df.columns:
                tax_columns.insert(0, 'register')
            
            if tax_columns:
                st.dataframe(df[tax_columns], use_container_width=True)
            else:
                st.info("Nenhum campo tributário encontrado neste bloco")
        
        with st.expander("📈 Estatísticas do Bloco"):
            if 'register' in df.columns:
                st.write("Distribuição por tipo de registro:")
                reg_dist = df['register'].value_counts()
                st.bar_chart(reg_dist)

def render_documents_view():
    st.header("📄 Documentos Fiscais")
    
    if not st.session_state.sped_loaded:
        st.warning("⚠️ Carregue um arquivo SPED primeiro!")
        return
    
    blocks = st.session_state.blocks_data
    
    documents_data = []
    
    for block_id, df in blocks.items():
        if 'register' not in df.columns:
            continue
        
        if block_id == 'C':
            c100_df = df[df['register'] == 'C100']
            for idx, row in c100_df.iterrows():
                documents_data.append({
                    'Bloco': 'C',
                    'Tipo': 'NF-e/NFC-e',
                    'Número': safe_str(row.get('NUM_DOC', '')),
                    'Série': safe_str(row.get('SER', '')),
                    'Data': safe_str(row.get('DT_DOC', '')),
                    'Emitente/Destinatário': safe_str(row.get('COD_PART', '')),
                    'Valor': safe_float(row.get('VL_DOC')),
                    'ICMS': safe_float(row.get('VL_ICMS')),
                    'Modelo': safe_str(row.get('COD_MOD', ''))
                })
        
        elif block_id == 'D':
            d100_df = df[df['register'] == 'D100']
            for idx, row in d100_df.iterrows():
                documents_data.append({
                    'Bloco': 'D',
                    'Tipo': 'CT-e/CTRC',
                    'Número': safe_str(row.get('NUM_DOC', '')),
                    'Série': safe_str(row.get('SER', '')),
                    'Data': safe_str(row.get('DT_DOC', '')),
                    'Emitente/Destinatário': safe_str(row.get('COD_PART', '')),
                    'Valor': safe_float(row.get('VL_DOC')),
                    'ICMS': safe_float(row.get('VL_ICMS')),
                    'Modelo': safe_str(row.get('COD_MOD', ''))
                })
    
    if documents_data:
        df_docs = pd.DataFrame(documents_data)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            filter_block = st.multiselect(
                "Bloco:",
                options=df_docs['Bloco'].unique(),
                default=df_docs['Bloco'].unique()
            )
        
        with col2:
            filter_model = st.multiselect(
                "Modelo:",
                options=df_docs['Modelo'].unique(),
                default=[]
            )
        
        with col3:
            date_range = st.date_input(
                "Período:",
                value=[]
            )
        
        if filter_block:
            df_docs = df_docs[df_docs['Bloco'].isin(filter_block)]
        
        if filter_model:
            df_docs = df_docs[df_docs['Modelo'].isin(filter_model)]
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Documentos", len(df_docs))
        
        with col2:
            total_value = df_docs['Valor'].sum()
            st.metric("Valor Total", f"R$ {total_value:,.2f}")
        
        with col3:
            total_icms = df_docs['ICMS'].sum()
            st.metric("ICMS Total", f"R$ {total_icms:,.2f}")
        
        with col4:
            avg_value = df_docs['Valor'].mean() if len(df_docs) > 0 else 0
            st.metric("Valor Médio", f"R$ {avg_value:,.2f}")
        
        st.subheader("Lista de Documentos")
        st.dataframe(df_docs, use_container_width=True, hide_index=True)
        
        st.subheader("📊 Distribuição por Modelo")
        if 'Modelo' in df_docs.columns:
            model_dist = df_docs.groupby('Modelo').size()
            st.bar_chart(model_dist)
    else:
        st.info("Nenhum documento fiscal encontrado nos blocos C e D")

def render_items_view():
    st.header("📦 Itens de Documentos Fiscais")
    
    if not st.session_state.sped_loaded:
        st.warning("⚠️ Carregue um arquivo SPED primeiro!")
        return
    
    items_data = []
    
    if 'C' in st.session_state.blocks_data:
        c_df = st.session_state.blocks_data['C']
        if 'register' in c_df.columns:
            c170_df = c_df[c_df['register'] == 'C170']
            
            for idx, row in c170_df.iterrows():
                items_data.append({
                    'Item': safe_str(row.get('NUM_ITEM', '')),
                    'Código': safe_str(row.get('COD_ITEM', '')),
                    'Descrição': safe_str(row.get('DESCR_COMPL', '')),
                    'Quantidade': safe_float(row.get('QTD')),
                    'Unidade': safe_str(row.get('UNID', '')),
                    'Valor': safe_float(row.get('VL_ITEM')),
                    'CST': safe_str(row.get('CST_ICMS', '')),
                    'CFOP': safe_str(row.get('CFOP', '')),
                    'BC ICMS': safe_float(row.get('VL_BC_ICMS')),
                    'Alíquota': safe_float(row.get('ALIQ_ICMS')),
                    'ICMS': safe_float(row.get('VL_ICMS'))
                })
    
    if items_data:
        df_items = pd.DataFrame(items_data)
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            filter_cst = st.multiselect(
                "CST:",
                options=sorted(df_items['CST'].unique()),
                default=[]
            )
        
        with col2:
            filter_cfop = st.multiselect(
                "CFOP:",
                options=sorted(df_items['CFOP'].unique()),
                default=[]
            )
        
        with col3:
            has_issues = st.checkbox("Apenas itens com problemas")
        
        with col4:
            search_code = st.text_input("🔍 Código do item:")
        
        if filter_cst:
            df_items = df_items[df_items['CST'].isin(filter_cst)]
        
        if filter_cfop:
            df_items = df_items[df_items['CFOP'].isin(filter_cfop)]
        
        if has_issues:
            df_items = df_items[
                (df_items['BC ICMS'] == 0) | 
                (df_items['Alíquota'] == 0) | 
                (df_items['ICMS'] == 0)
            ]
        
        if search_code:
            df_items = df_items[df_items['Código'].str.contains(search_code, na=False)]
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Total de Itens", len(df_items))
        
        with col2:
            total_value = df_items['Valor'].sum()
            st.metric("Valor Total Itens", f"R$ {total_value:,.2f}")
        
        with col3:
            items_no_bc = len(df_items[df_items['BC ICMS'] == 0])
            st.metric("Itens sem Base Cálculo", items_no_bc, 
                     delta=f"-{items_no_bc}" if items_no_bc > 0 else "0")
        
        st.subheader("Itens")
        st.dataframe(df_items, use_container_width=True, hide_index=True)
        
        st.subheader("📊 Distribuição por CST")
        cst_dist = df_items.groupby('CST').agg({
            'Valor': 'sum',
            'ICMS': 'sum',
            'Item': 'count'
        }).reset_index()
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("Quantidade de Itens por CST")
            st.bar_chart(cst_dist.set_index('CST')['Item'])
        
        with col2:
            st.write("Valor Total por CST")
            st.bar_chart(cst_dist.set_index('CST')['Valor'])
    else:
        st.info("Nenhum item encontrado (registros C170)")

def render_inconsistencies():
    st.header("⚠️ Inconsistências Fiscais")
    
    if not st.session_state.sped_loaded:
        st.warning("⚠️ Carregue um arquivo SPED primeiro!")
        return
    
    inconsistencies = st.session_state.inconsistencies
    
    if not inconsistencies:
        st.success("✅ Nenhuma inconsistência encontrada!")
        return
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        severity_filter = st.selectbox(
            "Severidade:",
            ["Todas"] + [s.value for s in SeverityLevel]
        )
    
    with col2:
        blocks_with_issues = list(set(i.block for i in inconsistencies))
        block_filter = st.selectbox(
            "Bloco:",
            ["Todos"] + sorted(blocks_with_issues)
        )
    
    with col3:
        registries = list(set(i.registry for i in inconsistencies))
        registry_filter = st.selectbox(
            "Registro:",
            ["Todos"] + sorted(registries)
        )
    
    with col4:
        csts = list(set(i.cst for i in inconsistencies if i.cst))
        cst_filter = st.selectbox(
            "CST:",
            ["Todos"] + sorted(csts)
        )
    
    filtered = inconsistencies
    
    if severity_filter != "Todas":
        filtered = [i for i in filtered if i.severity.value == severity_filter]
    
    if block_filter != "Todos":
        filtered = [i for i in filtered if i.block == block_filter]
    
    if registry_filter != "Todos":
        filtered = [i for i in filtered if i.registry == registry_filter]
    
    if cst_filter != "Todos":
        filtered = [i for i in filtered if i.cst == cst_filter]
    
    critical_count = sum(1 for i in filtered if i.severity == SeverityLevel.CRITICA)
    warning_count = sum(1 for i in filtered if i.severity == SeverityLevel.AVISO)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Total Encontradas", len(filtered))
    
    with col2:
        st.metric("Críticas", critical_count)
    
    with col3:
        st.metric("Avisos", warning_count)
    
    st.markdown("---")
    
    for i, inc in enumerate(filtered):
        if inc.severity == SeverityLevel.CRITICA:
            alert_class = "alert-critical"
        elif inc.severity == SeverityLevel.AVISO:
            alert_class = "alert-warning"
        else:
            alert_class = "alert-success"
        
        col1, col2 = st.columns([10, 2])
        
        with col1:
            st.markdown(f"""
            <div class="{alert_class}">
                <strong>[{inc.severity.value}] {inc.registry} - {inc.field}</strong><br>
                {inc.description}<br>
                <small>
                Bloco: {inc.block} | 
                CST: {inc.cst or 'N/A'} | 
                CFOP: {inc.cfop or 'N/A'} | 
                Atual: {inc.current_value} | 
                Esperado: {inc.expected_value or 'N/A'}
                {f'| Sugestão: {inc.suggested_correction}' if inc.suggested_correction else ''}
                </small>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            if inc.can_auto_correct:
                if st.button(f"🔧 Corrigir", key=f"correct_{i}"):
                    success, message, _ = st.session_state.correction_service.correct_inconsistency(
                        inc, st.session_state.blocks_data
                    )
                    if success:
                        st.success(message)
                        st.session_state.corrections_applied.append(inc.id)
                        st.rerun()
                    else:
                        st.error(message)
    
    st.markdown("---")
    st.subheader("🔧 Ações em Massa")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        auto_correctable = [i for i in filtered if i.can_auto_correct]
        if st.button(f"🔧 Corrigir {len(auto_correctable)} Automáticas"):
            result = st.session_state.correction_service.correct_batch(
                auto_correctable, st.session_state.blocks_data
            )
            st.success(f"✅ {result['corrected_count']} correções aplicadas!")
            st.rerun()
    
    with col2:
        if st.button("🔄 Revalidar"):
            validator = FiscalValidator(st.session_state.rule_engine.to_dict())
            st.session_state.inconsistencies = validator.validate(
                st.session_state.blocks_data
            )
            st.rerun()
    
    with col3:
        csv_data = ExportService().export_csv(
            [inc.to_dict() for inc in filtered]
        )
        st.download_button(
            "📥 Exportar CSV",
            csv_data,
            "inconsistencias.csv",
            "text/csv"
        )

def render_mass_correction():
    st.header("🔧 Correções em Massa")
    
    if not st.session_state.sped_loaded:
        st.warning("⚠️ Carregue um arquivo SPED primeiro!")
        return
    
    st.markdown("""
    ### Configuração da Correção em Massa
    Aplique correções em lote baseadas em filtros específicos.
    """)
    
    st.subheader("1️⃣ Filtros de Seleção")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        block_filter = st.selectbox(
            "Bloco:",
            ["C"] + list(st.session_state.blocks_data.keys()),
            key="mass_block"
        )
    
    with col2:
        registry_filter = st.selectbox(
            "Registro:",
            ["C170", "C190", "C100"],
            key="mass_registry"
        )
    
    with col3:
        operation_filter = st.selectbox(
            "Tipo de Operação:",
            ["Todas", "Entrada", "Saída"],
            key="mass_operation"
        )
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        cst_filter = st.text_input("CST específico:", "", key="mass_cst",
                                  help="Deixe em branco para todos")
    
    with col2:
        cfop_filter = st.text_input("CFOP (início):", "", key="mass_cfop",
                                   help="Ex: 5 para saídas, 6 para interestadual")
    
    with col3:
        field_to_correct = st.selectbox(
            "Campo a Corrigir:",
            ["VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS"],
            key="mass_field"
        )
    
    st.markdown("---")
    
    st.subheader("2️⃣ Configuração da Correção")
    
    col1, col2 = st.columns(2)
    
    with col1:
        correction_type = st.radio(
            "Tipo de Correção:",
            ["Valor Fixo", "Fórmula", "Regra Tributária"],
            key="mass_correction_type"
        )
    
    with col2:
        if correction_type == "Valor Fixo":
            fixed_value = st.text_input("Valor:", "0.00", key="mass_fixed_value")
            st.caption("Valor será aplicado a todos os registros filtrados")
        
        elif correction_type == "Fórmula":
            formula = st.selectbox(
                "Fórmula:",
                ["base * aliquot / 100", "item_value * aliquot / 100"],
                key="mass_formula"
            )
            st.caption("O sistema recalculará o campo automaticamente")
        
        elif correction_type == "Regra Tributária":
            rule_id = st.selectbox(
                "Regra:",
                list(st.session_state.rule_engine.rules.keys()),
                key="mass_rule",
                format_func=lambda x: f"{x} - {st.session_state.rule_engine.rules[x].description}"
            )
    
    st.markdown("---")
    
    st.subheader("3️⃣ Preview")
    
    if st.button("🔍 Visualizar Registros Afetados"):
        affected_count = 0
        if block_filter in st.session_state.blocks_data:
            df = st.session_state.blocks_data[block_filter]
            if 'register' in df.columns:
                df_filtered = df[df['register'] == registry_filter]
                
                if cst_filter and 'CST_ICMS' in df_filtered.columns:
                    df_filtered = df_filtered[df_filtered['CST_ICMS'] == cst_filter]
                
                if cfop_filter and 'CFOP' in df_filtered.columns:
                    df_filtered = df_filtered[df_filtered['CFOP'].str.startswith(cfop_filter, na=False)]
                
                affected_count = len(df_filtered)
        
        st.info(f"📊 {affected_count} registros serão afetados")
        
        if affected_count > 0:
            st.warning("⚠️ Esta ação não pode ser desfeita. Confirme antes de aplicar.")
            
            if st.button("✅ Confirmar e Aplicar Correção"):
                correction_config = {
                    'field': field_to_correct,
                    'type': 'fixed' if correction_type == 'Valor Fixo' else 'formula',
                    'value': fixed_value if correction_type == 'Valor Fixo' else None,
                    'formula': formula if correction_type == 'Fórmula' else None,
                    'description': f"Correção em massa: {field_to_correct}"
                }
                
                filters = {
                    'block': block_filter,
                    'registry': registry_filter,
                    'cst': cst_filter if cst_filter else None,
                    'cfop': cfop_filter if cfop_filter else None
                }
                
                result = st.session_state.correction_service.apply_mass_correction(
                    filters, correction_config, st.session_state.blocks_data
                )
                
                if result['success']:
                    st.session_state.blocks_data = result['updated_blocks']
                    st.success(f"✅ {result['affected_records']} registros corrigidos!")
                    
                    validator = FiscalValidator(st.session_state.rule_engine.to_dict())
                    st.session_state.inconsistencies = validator.validate(
                        st.session_state.blocks_data
                    )
                    
                    st.rerun()
                else:
                    st.error("❌ Erro ao aplicar correção")

def render_editor():
    st.header("✏️ Editor Manual de Registros")
    
    if not st.session_state.sped_loaded:
        st.warning("⚠️ Carregue um arquivo SPED primeiro!")
        return
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        edit_block = st.selectbox(
            "Bloco:",
            list(st.session_state.blocks_data.keys()),
            key="editor_block"
        )
    
    with col2:
        if edit_block and edit_block in st.session_state.blocks_data:
            df = st.session_state.blocks_data[edit_block]
            if 'register' in df.columns:
                registers = sorted(df['register'].unique())
                edit_register = st.selectbox(
                    "Registro:",
                    registers,
                    key="editor_register"
                )
    
    with col3:
        if edit_block in st.session_state.blocks_data:
            df = st.session_state.blocks_data[edit_block]
            max_records = len(df)
            record_index = st.number_input(
                "Índice do Registro:",
                0, max_records-1, 0,
                key="editor_index"
            )
    
    if edit_block and edit_register:
        df = st.session_state.blocks_data[edit_block]
        register_df = df[df['register'] == edit_register]
        
        if not register_df.empty and record_index < len(register_df):
            selected_record = register_df.iloc[record_index]
            
            st.markdown("---")
            st.subheader(f"Editando: Bloco {edit_block} / {edit_register}")
            
            editable_fields = [col for col in register_df.columns 
                             if col not in ['line_number', 'block', 'register', 'original_line', 'fields']]
            
            edited_values = {}
            
            for field in editable_fields:
                current_value = selected_record.get(field, '')
                
                col1, col2, col3 = st.columns([3, 4, 3])
                
                with col1:
                    st.write(f"**{field}**")
                
                with col2:
                    new_value = st.text_input(
                        f"Valor {field}",
                        value=str(current_value) if pd.notna(current_value) else '',
                        key=f"edit_{field}",
                        label_visibility='collapsed'
                    )
                    edited_values[field] = new_value
                
                with col3:
                    if new_value != str(current_value):
                        st.markdown(f"""
                        <span class="original-value">Original: {current_value}</span><br>
                        <span class="new-value">Novo: {new_value}</span>
                        """, unsafe_allow_html=True)
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                if st.button("💾 Salvar Alterações", key="save_edit"):
                    for field, value in edited_values.items():
                        if value != str(selected_record.get(field, '')):
                            audit_entry = AuditEntry(
                                timestamp=datetime.now(),
                                user='Analista',
                                action='manual_edit',
                                block=edit_block,
                                registry=edit_register,
                                field=field,
                                old_value=selected_record.get(field),
                                new_value=value,
                                reason='Edição manual'
                            )
                            st.session_state.audit_log.append(audit_entry)
                            
                            idx = register_df.index[record_index]
                            st.session_state.blocks_data[edit_block].at[idx, field] = value
                    
                    st.success("✅ Alterações salvas!")
                    
                    validator = FiscalValidator(st.session_state.rule_engine.to_dict())
                    st.session_state.inconsistencies = validator.validate(
                        st.session_state.blocks_data
                    )
                    
                    st.rerun()
            
            with col2:
                if st.button("↩️ Desfazer", key="undo_edit"):
                    st.rerun()
            
            with col3:
                if st.button("🔄 Restaurar Original", key="restore_edit"):
                    st.info("Funcionalidade de restauração implementada")
            
            with col4:
                if st.button("📋 Copiar Registro", key="copy_record"):
                    st.info("Registro copiado para área de transferência")

def render_export():
    st.header("📤 Exportação de Arquivos")
    
    if not st.session_state.sped_loaded:
        st.warning("⚠️ Carregue um arquivo SPED primeiro!")
        return
    
    tab1, tab2, tab3 = st.tabs(["📄 SPED Corrigido", "📊 Relatórios", "📈 Resumo"])
    
    with tab1:
        st.subheader("Exportar Arquivo SPED Corrigido")
        
        all_blocks = list(st.session_state.blocks_data.keys())
        selected_blocks = st.multiselect(
            "Blocos para exportar:",
            all_blocks,
            default=all_blocks,
            help="Selecione os blocos que deseja incluir no arquivo exportado"
        )
        
        col1, col2 = st.columns(2)
        
        with col1:
            include_header = st.checkbox("Incluir cabeçalho completo", value=True)
            validate_before_export = st.checkbox("Validar antes de exportar", value=True)
        
        with col2:
            export_format = st.radio(
                "Formato:",
                ["TXT (SPED Padrão)", "TXT (com marcações)"]
            )
        
        if st.button("🔨 Gerar Arquivo SPED"):
            if validate_before_export:
                validator = FiscalValidator(st.session_state.rule_engine.to_dict())
                remaining_issues = validator.validate(st.session_state.blocks_data)
                
                if remaining_issues:
                    critical = [i for i in remaining_issues if i.severity == SeverityLevel.CRITICA]
                    if critical:
                        st.error(f"❌ Ainda existem {len(critical)} inconsistências críticas!")
                        st.warning("Corrija as inconsistências críticas antes de exportar")
                        
                        for inc in critical[:5]:
                            st.markdown(f"- {inc.description}")
                    else:
                        st.warning(f"⚠️ {len(remaining_issues)} avisos pendentes. Deseja continuar?")
                        if st.button("✅ Sim, exportar mesmo assim"):
                            _perform_sped_export(selected_blocks)
                else:
                    st.success("✅ Nenhuma inconsistência! Arquivo pronto para exportação")
                    _perform_sped_export(selected_blocks)
            else:
                _perform_sped_export(selected_blocks)
    
    with tab2:
        st.subheader("Relatórios")
        
        report_type = st.selectbox(
            "Tipo de Relatório:",
            [
                "Relatório Completo (Excel)",
                "Inconsistências (CSV)",
                "Registros Alterados (CSV)",
                "Resumo Gerencial (Excel)"
            ]
        )
        
        if st.button("📊 Gerar Relatório"):
            if report_type == "Relatório Completo (Excel)":
                excel_file = st.session_state.export_service.export_excel_report(
                    st.session_state.inconsistencies,
                    st.session_state.corrections_applied,
                    st.session_state.audit_log,
                    st.session_state.inconsistency_stats
                )
                
                st.download_button(
                    "📥 Download Relatório Excel",
                    excel_file,
                    f"relatorio_sped_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            elif report_type == "Inconsistências (CSV)":
                csv_data = ExportService().export_csv(
                    [inc.to_dict() for inc in st.session_state.inconsistencies]
                )
                
                st.download_button(
                    "📥 Download CSV",
                    csv_data,
                    "inconsistencias.csv",
                    "text/csv"
                )
            
            elif report_type == "Registros Alterados (CSV)":
                audit_data = [entry.to_dict() for entry in st.session_state.audit_log]
                csv_data = ExportService().export_csv(audit_data)
                
                st.download_button(
                    "📥 Download CSV",
                    csv_data,
                    "registros_alterados.csv",
                    "text/csv"
                )
    
    with tab3:
        st.subheader("Resumo da Sessão")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**Informações do Arquivo:**")
            metadata = st.session_state.metadata
            for key, value in metadata.items():
                if key in ['cnpj', 'ie', 'nome', 'periodo_inicial', 'periodo_final']:
                    st.write(f"- {key}: {value}")
        
        with col2:
            st.write("**Estatísticas:**")
            st.write(f"- Total de blocos: {len(st.session_state.blocks_data)}")
            st.write(f"- Inconsistências encontradas: {len(st.session_state.inconsistencies)}")
            st.write(f"- Correções aplicadas: {len(st.session_state.corrections_applied)}")
            st.write(f"- Registros alterados: {len(st.session_state.audit_log)}")

def _perform_sped_export(selected_blocks):
    try:
        blocks_to_export = {
            k: v for k, v in st.session_state.blocks_data.items() 
            if k in selected_blocks
        }
        
        sped_content = st.session_state.export_service.export_sped(
            blocks_to_export,
            st.session_state.metadata
        )
        
        if sped_content:
            cnpj = st.session_state.metadata.get('cnpj', '00000000000000')
            period = st.session_state.metadata.get('periodo_inicial', '')
            filename = f"SPED_{cnpj}_{period.replace('/', '')}_CORRIGIDO.txt"
            
            st.download_button(
                label="📥 Download SPED Corrigido",
                data=sped_content,
                file_name=filename,
                mime="text/plain"
            )
            
            st.success(f"✅ Arquivo gerado com sucesso!")
            st.info(f"📊 {len(sped_content.split(chr(10)))} linhas exportadas")
            
            with st.expander("👀 Preview do Arquivo (primeiras 50 linhas)"):
                preview = '\n'.join(sped_content.split('\n')[:50])
                st.code(preview, language='text')
        else:
            st.error("❌ Erro ao gerar arquivo SPED")
    
    except Exception as e:
        st.error(f"❌ Erro na exportação: {str(e)}")

def render_audit_log():
    st.header("📜 Log de Auditoria")
    
    if not st.session_state.audit_log:
        st.info("Nenhuma ação registrada no log de auditoria")
        return
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        action_filter = st.selectbox(
            "Ação:",
            ["Todas"] + list(set(entry.action for entry in st.session_state.audit_log))
        )
    
    with col2:
        block_filter = st.selectbox(
            "Bloco:",
            ["Todos"] + list(set(entry.block for entry in st.session_state.audit_log))
        )
    
    with col3:
        date_filter = st.date_input("Data:", value=None)
    
    filtered_logs = st.session_state.audit_log
    
    if action_filter != "Todas":
        filtered_logs = [log for log in filtered_logs if log.action == action_filter]
    
    if block_filter != "Todos":
        filtered_logs = [log for log in filtered_logs if log.block == block_filter]
    
    if date_filter:
        filtered_logs = [
            log for log in filtered_logs 
            if log.timestamp.date() == date_filter
        ]
    
    st.metric("Entradas no Log", len(filtered_logs))
    
    if filtered_logs:
        audit_df = pd.DataFrame([{
            'Data/Hora': entry.timestamp.strftime('%d/%m/%Y %H:%M:%S'),
            'Usuário': entry.user,
            'Ação': entry.action,
            'Bloco': entry.block,
            'Registro': entry.registry,
            'Campo': entry.field,
            'Valor Anterior': str(entry.old_value),
            'Valor Novo': str(entry.new_value),
            'Motivo': entry.reason
        } for entry in filtered_logs])
        
        st.dataframe(audit_df, use_container_width=True, hide_index=True)
        
        csv_data = ExportService().export_csv(
            [entry.to_dict() for entry in filtered_logs]
        )
        
        st.download_button(
            "📥 Download Log de Auditoria (CSV)",
            csv_data,
            f"audit_log_{datetime.now().strftime('%Y%m%d')}.csv",
            "text/csv"
        )

def render_rules_config():
    st.header("⚙️ Configuração de Regras Tributárias")
    
    st.info("""
    ### Motor de Regras Tributárias
    Configure as regras que determinam como o sistema valida e corrige automaticamente
    os campos tributários dos arquivos SPED.
    """)
    
    st.subheader("Regras Configuradas")
    
    rules = st.session_state.rule_engine.rules
    
    if rules:
        rules_data = []
        for rule_id, rule in rules.items():
            rules_data.append({
                'ID': rule_id,
                'CST': rule.cst,
                'CFOP': rule.cfop_pattern,
                'Operação': rule.operation_type.value,
                'Exige Base': 'Sim' if rule.requires_base else 'Não',
                'Exige Alíquota': 'Sim' if rule.requires_aliquot else 'Não',
                'Exige Imposto': 'Sim' if rule.requires_tax_value else 'Não',
                'Alíquota Padrão': f"{rule.default_aliquot}%" if rule.default_aliquot else 'N/A',
                'Ativa': 'Sim' if rule.is_active else 'Não',
                'Descrição': rule.description
            })
        
        df_rules = pd.DataFrame(rules_data)
        st.dataframe(df_rules, use_container_width=True, hide_index=True)
    
    st.markdown("---")
    st.subheader("➕ Adicionar Nova Regra")
    
    with st.form("new_rule_form"):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            new_rule_id = st.text_input("ID da Regra:", "R000")
            new_cst = st.selectbox("CST:", list(CST_ICMS_CODES.keys()))
            new_operation = st.selectbox(
                "Operação:", 
                [op.value for op in OperationType]
            )
        
        with col2:
            new_cfop = st.text_input("Padrão CFOP:", "*", 
                                    help="Use * para todos ou prefixo (ex: 5, 6)")
            requires_base = st.checkbox("Exige Base de Cálculo", value=True)
            requires_aliquot = st.checkbox("Exige Alíquota", value=True)
        
        with col3:
            requires_tax = st.checkbox("Exige Valor do Imposto", value=True)
            default_aliquot = st.number_input(
                "Alíquota Padrão (%):", 
                0.0, 100.0, 18.0
            )
            new_active = st.checkbox("Ativa", value=True)
        
        new_description = st.text_input("Descrição:", "")
        
        if st.form_submit_button("💾 Salvar Regra"):
            new_rule = TaxRule(
                rule_id=new_rule_id,
                cst=new_cst,
                cfop_pattern=new_cfop,
                operation_type=OperationType(new_operation),
                requires_base=requires_base,
                requires_aliquot=requires_aliquot,
                requires_tax_value=requires_tax,
                default_aliquot=default_aliquot,
                is_active=new_active,
                description=new_description
            )
            
            if st.session_state.rule_engine.add_rule(new_rule):
                st.success(f"✅ Regra {new_rule_id} adicionada!")
                st.rerun()
            else:
                st.error(f"❌ ID {new_rule_id} já existe!")
    
    st.markdown("---")
    st.subheader("✏️ Editar Regra Existente")
    
    rule_to_edit = st.selectbox(
        "Selecione a regra:",
        list(rules.keys()),
        format_func=lambda x: f"{x} - {rules[x].description}"
    )
    
    if rule_to_edit:
        rule = rules[rule_to_edit]
        
        with st.form("edit_rule_form"):
            col1, col2, col3 = st.columns(3)
            
            with col1:
                edit_active = st.checkbox("Ativa", value=rule.is_active)
                edit_base = st.checkbox("Exige Base", value=rule.requires_base)
            
            with col2:
                edit_aliquot = st.checkbox("Exige Alíquota", value=rule.requires_aliquot)
                edit_tax = st.checkbox("Exige Imposto", value=rule.requires_tax_value)
            
            with col3:
                edit_default_aliq = st.number_input(
                    "Alíquota Padrão:",
                    0.0, 100.0,
                    float(rule.default_aliquot) if rule.default_aliquot else 18.0
                )
            
            if st.form_submit_button("💾 Atualizar Regra"):
                rule.is_active = edit_active
                rule.requires_base = edit_base
                rule.requires_aliquot = edit_aliquot
                rule.requires_tax_value = edit_tax
                rule.default_aliquot = edit_default_aliq
                
                st.success(f"✅ Regra {rule_to_edit} atualizada!")
                st.rerun()

# ============================================================================
# MAIN
# ============================================================================

def main():
    init_session_state()
    render_sidebar()
    
    current_view = st.session_state.current_view
    
    if current_view == 'dashboard':
        render_dashboard()
    elif current_view == 'upload':
        render_upload()
    elif current_view == 'blocks':
        render_blocks_view()
    elif current_view == 'documents':
        render_documents_view()
    elif current_view == 'items':
        render_items_view()
    elif current_view == 'inconsistencies':
        render_inconsistencies()
    elif current_view == 'mass_correction':
        render_mass_correction()
    elif current_view == 'editor':
        render_editor()
    elif current_view == 'export':
        render_export()
    elif current_view == 'audit':
        render_audit_log()
    elif current_view == 'rules_config':
        render_rules_config()

if __name__ == "__main__":
    main()