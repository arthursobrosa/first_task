from enum import Enum
from typing import Optional


class Agent(Enum):
    Concessionaria = 1
    Permissionaria = 2

    @property
    def type_name(self) -> str:
        match self:
            case Agent.Concessionaria:
                return "Concessionária"
            case Agent.Permissionaria:
                return "Permissionária"
            
    @property
    def path(self) -> str:
        match self:
            case Agent.Concessionaria:
                return "Concessionárias"
            case Agent.Permissionaria:
                return "Permissionárias"
            
    @property
    def cover_tab_name(self) -> str:
        return 'CAPA'
    
    @property
    def market_tab_name(self) -> str:
        match self:
            case Agent.Concessionaria:
                return 'Mercado_Receita'
            case Agent.Permissionaria:
                return 'BD MERCADO'
            
    @property
    def process_date_dn(self) -> str:
        match self:
            case Agent.Concessionaria:
                return 'LnkTxtDRPData'
            case Agent.Permissionaria:
                return 'DataRevRea'
            
    @property
    def process_dates_coord(self) -> list[str]:
        match self:
            case Agent.Concessionaria:
                return ['C10']
            case Agent.Permissionaria:
                return ['C15', 'C14']
            

    #TODO: Remover depois
    @property
    def concession_id_dn(self) -> Optional[str]:
        match self:
            case Agent.Concessionaria:
                return None
            case Agent.Permissionaria:
                return 'idsreag'
            
    @property
    def concession_ids_coord(self) -> list[str]:
        match self:
            case Agent.Concessionaria:
                return ['C23']
            case Agent.Permissionaria:
                return ['K10', 'C19']