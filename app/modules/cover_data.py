from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from .agent import Agent
from datetime import datetime, date


def get_process_date(workbook: Workbook, agent: Agent) -> date:
    cover = workbook[agent.cover_tab_name]

    if agent.process_date_dn:
        defined_name = workbook.defined_names[agent.process_date_dn]

        for tab_name, cell_ref in defined_name.destinations:
            tab_origin = workbook[tab_name]
            return tab_origin[cell_ref].value
    
    for coord in agent.process_dates_coord:
        value = cover[coord].value

        if isinstance(value, datetime):
            return value.date()
        elif isinstance(value, date):
            return value