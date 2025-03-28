print('Compilando bibliotecas... Por favor, aguarde')
import pandas as pd
import numpy as np
import xlwings as xw
import openpyxl
import os
import warnings

pd.set_option('display.max_columns', None)
warnings.filterwarnings("ignore", message="The behavior of DataFrame concatenation with empty or all-NA entries is deprecated")

repository_path = input('Ingresa la ruta de la carpeta "Controlling LATAM":\n')
final_repository = repository_path + r'\Reportes\SanityCheck'
pm_cpm_path = repository_path + ...
sanity_date = input('Ingresa la fecha en el formato YYYYMMDD (P. ej.: 20240930):\n')
sanity_per = input('Ingresa el periodo correspondiente:\n')

file_final_path = final_repository + f'\\P{sanity_per} - Sanity Check - {sanity_date}.xlsx'

pm_cpm = pd.read_excel(pm_cpm_path, usecols = ('B:G'))
pm_cpm.rename(columns = {'CPM_NAME': 'CPM', 'PM_NAME': 'PM'}, inplace = True)
pm_cpm['PROJECTDASH_ID'] = pm_cpm['PROJECTDASH_ID'].drop_duplicates()
pm_cpm_rename_dict = {
    'SEGMENT': 'Segment',
    'PORTFOLIO': 'Portfolio',
    'CATEGORY': 'Category'
}

sanity = pd.DataFrame()

# Row Level Formulas

def agregar_coma(value):
    value = str(value)
    if value != 'OK':
        new_value = value[:-2] + ',' + value[-2:]
        return new_value
    else:
        return value
    
def formato_fecha(value):
    value = str(value)
    if value == 'OK':
        return value
    else:
        new_value = value[:2] + '/' + value[2:4] + '/' + value[-4:]
        return new_value
    
def ab_cero(value):
    if value == 0:
        return 'Revisar'
    else:
        return 'OK'
    
def plancost(row):
    if row['Plan Cost'] < (row['Actual Cost'] + row['Committed Cost']):
        return 'Revisar'
    else:
        return 'OK'
    
def wip_ab(row):
    if row['WIP/UC'] > row['Order on Hand']:
        return 'Revisar'
    else:
        return 'OK'
    
def accrual(row):
    if (row['Order on Hand'] == 0) & ((row['Plan Cost'] - row['Actual Cost']) < row['Accrual']):
        return 'Revisar'
    else:
        return 'OK'
    
def wbs(value):
    aux = value[:4] + value[5:]
    return aux

def ae_negativo(row):
    if (row['New Order'] < 0) | (row['Order on Hand'] < 0) | (row['Sales'] < 0):
        return 'Revisar'
    else:
        return 'OK'

def proj_number(value):
    aux = value[0] + value[2:6] + value[7:]
    return aux

def to_date(value):
    value = str(value)
    if value[0] != 'n':   
        result = value[:2] + '/' + value[2:4] + '/' + value[4:]
        return result
    else:
        return value

def change_sign(value):
    if value == 0:
        return value
    else:
        value = value * -1
        return value
    
def proj_type(value):
    if value[:4] == '75OC':
        return 'CCM'
    else:
        return 'POC'
    
def negativo(row):
    if (row['Status'] != 'ENCE') & (row['Status'] != 'ENTE') & ((row['New Order'] < 0) | (row['Order on Hand'] < 0) | (row['Real - Revenue'] < 0) | (row['Real - Margin %'] < 0)):
        return 'Review'
    else:
        return 'OK'

def abcero(row):
    if (row['Project Type'] == 'CCM') & (row['Status'] != 'ENTE') & (row['Order on Hand'] == 0):
        return 'Review'
    else:
        return 'OK'

def margin(row):
    if (row['Status'] != 'ENCE') & (row['Status'] != 'ENTE') & (row['Real - Margin %'] > 70):
        return 'Review'
    else:
        return 'OK'

def plan_cost(row):
    if (row['Plan Cost'] < (row['Actual Cost'] + row['Committed Cost'])) & (row['Status'] != 'ENCE') & (row['Status'] != 'ENTE'):
        return 'Review'
    else:
        return 'OK'

def wip(row):
    if (row['WIP'] > row['Order on Hand']) & (row['Status'] != 'ENCE') & (row['Status'] != 'ENTE'):
        return 'Review'
    else:
        return 'OK'

def accrual_costs(row):
    if (row['Accrual'] > (row['Actual Cost'] + row['Committed Cost'])) & (row['Status'] != 'ENCE') & (row['Status'] != 'ENTE'):
        return 'Review'
    else:
        return 'OK'

def close_project(row):
    if (row['Order on Hand']==0) & (row['WIP']==0) & (row['Accrual']==0) & (row['Status'] == 'ENTE') & (row['Committed Cost']==0) & (row['Excess Cost']==0) & (row['Excess Bill']==0):
        return 'Review'
    else:
        return 'OK'

def poc_100(row):
    if (row['Project Type'] == 'POC') & (row['POC%'] == 100) & (row['Status'] != 'ENCE') & (row['Status'] != 'ENTE'):
        return 'Review'
    else:
        return 'OK'

def rev_planrev(row):
    if (row['Project Type']=='POC') & (row['Real - Revenue'] > row['New Order']) & (row['Status'] != 'ENCE') & (row['Status'] != 'ENTE'):
        return 'Review'
    else:
        return 'OK'

def high_cost(row):
    if row['New Order'] != 0:
        if (row['Project Type'] == 'POC') & (row['POC%'] - (row['Prog. Billing']/row['New Order']) > 30) & (row['Status'] != 'ENCE') & (row['Status'] != 'ENTE'):
            return 'Review'
        else:
            return 'OK'
    else:
        return 'OK'
    
# Sanity AAN

ar = repository_path + ...
bo = repository_path + ...
cl = repository_path + ...
pr = repository_path + ...
uy = repository_path + ...
co = repository_path + ...
ve = repository_path + ...

paises = [ar,bo,cl,pr,uy,co,ve]
country = ['ARG','BOL','CHL','PER','URU','COL','VEN']
j = 0

for i in paises:
    df = pd.read_excel(i, usecols=('A:S,V:AH'), decimal=',', thousands='.')
    df['Country'] = country[j]
    sanity = pd.concat([sanity, df])
    j += 1

sanity.reset_index(inplace = True)
sanity = sanity.drop('index', axis = 1)
sanity.rename(columns = {'Project Desc.': 'Description', 'Seg. Profit center': 'Segment',
                         'New Orders/OOH/Sales <0': 'AE/AB/Revenue < 0', 'New Orders(=0)': 'AE = 0'}, inplace = True)

sanity['GM>50%'] = sanity['GM>50%'].apply(lambda x: str(x).replace('.', ''))
sanity['Delay in completion date'] = sanity['Delay in completion date'].apply(lambda x: str(x).replace('.', ''))
sanity['GM>50%'] = sanity['GM>50%'].apply(agregar_coma)
sanity['Delay in completion date'] = sanity['Delay in completion date'].apply(formato_fecha)

last_col = sanity.columns[-1]
new_cols = [last_col] + [col for col in sanity.columns if col != last_col]

sanity = sanity.reindex(columns = new_cols)

sanity['OOH = 0'] = sanity['Order on Hand'].apply(ab_cero)
sanity['Plan Cost < Actual + Committed'] = sanity.apply(plancost, axis = 1)
sanity['WIP > OOH'] = sanity.apply(wip_ab, axis = 1)
sanity['Accrual > Plan Cost - Actual'] = sanity.apply(accrual, axis = 1)
sanity['Total Alertas'] = '' 
sanity['Comentarios'] = ''
sanity['Original Pln GM %'] = sanity['Original Pln GM %'] / 100
sanity['Pln. GM%'] = sanity['Pln. GM%'] / 100
sanity['Project Number'] = sanity['Project Number'].apply(wbs)

sanity_AAN = pd.merge(sanity, pm_cpm, left_on = 'Project Number', right_on = 'PROJECTDASH_ID', how = 'left')
sanity_AAN = sanity_AAN.drop(['PROJECTDASH_ID', 'Segment'], axis = 1)
sanity_AAN = sanity_AAN.rename(pm_cpm_rename_dict, axis = 1)

cols = list(sanity_AAN.columns)

sanity_AAN = sanity_AAN.reindex(columns = cols[0:1] + cols[38:40]+cols[1:5]+cols[40:44]+cols[5:6]+cols[30:33]+cols[29:30]+cols[33:38]+cols[6:29])
sanity_AAN['Total Alertas'] = sanity_AAN.iloc[:,9:16].apply(lambda x: (x != 'OK').sum(), axis=1)

sanity_SA = sanity_AAN[(sanity_AAN['Country'] != 'COL') & (sanity_AAN['Country'] != 'VEN')]
sanity_NA = sanity_AAN[(sanity_AAN['Country'] == 'COL') | (sanity_AAN['Country'] == 'VEN')]

# Sanity MEX

sanitymex = pd.read_excel(
    repository_path + r'\BasesDeDatos\03 - Sanity\Sanity MEX.XLSX', 
    sheet_name = 'Sheet1', 
    decimal = ',', 
    thousands = '.'
    )

sanitymex['Country'] = 'MEX'

sanitymex.drop(['Plan Cost','Finish Date', 'Accrual','Potential Loss Contract'], 
              axis=1, 
              inplace=True, 
              errors='ignore')

rename_dict = {
            'Project definition': 'Project Number',
            'NO OCC': 'New Order', 
            'OHH OCC': 'Order on Hand',
            'Sales OCC CCM': 'Sales', 
            'Plan cost RP': 'Plan Cost', 
            'Commited Cost OCC': 'Committed Cost',
            'Actual Cost OCC': 'Actual Cost', 
            'WIP OCC': 'WIP/UC', 
            'Accrual OCC': 'Accrual',
            'MRS name': 'Segment', 
            'Actual GM OCC': 'Actual GM', 
            'COS OCC': 'COS',
            'Finish date': 'Finish Date', 
            'Potential loss contract.1': 'Potential Loss Contract'
            }
sanitymex.rename(columns = rename_dict,inplace = True)

sanitymex = sanitymex[['Country',
                      'Project Number',
                      'Description',
                      'Profit Center',
                      'Segment',
                      'Status',
                      'New Order',
                      'Order on Hand',
                      'Sales',
                      'Plan Cost',
                      'Committed Cost', 
                      'Actual Cost',
                      'Accrual',
                      'COS', 
                      'WIP/UC',
                      'Planned GM',
                      'Actual GM',
                      'Created On',
                      'Finish Date', 
                      'Proj. is a Loss contract',
                      'Potential Loss Contract', 
                      'Plan Cost Updated', 
                      'GM>50%']]

sanitymex['Plan Cost Updated'] = sanitymex['Plan Cost Updated'].fillna(0)
sanitymex['GM>50%'] = sanitymex['GM>50%'].fillna(0).replace(0, 'OK')
sanitymex['AE/AB/Revenue < 0'] = sanitymex.apply(ae_negativo, axis = 1)
sanitymex['OOH = 0'] = sanitymex['Order on Hand'].apply(ab_cero)
sanitymex['Plan Cost < Actual + Committed'] = sanitymex.apply(plancost, axis = 1)
sanitymex['WIP > OOH'] = sanitymex.apply(wip_ab, axis = 1)
sanitymex['Accrual > Plan Cost - Actual'] = sanitymex.apply(accrual, axis = 1)
sanitymex['Total Alertas'] = ''
sanitymex['Comentarios'] = '' 

sanitymex['Project Number'] = sanitymex['Project Number'].apply(proj_number)

sanity_MEX = pd.merge(sanitymex, pm_cpm, left_on = 'Project Number', right_on = 'PROJECTDASH_ID', how = 'left')
sanity_MEX = sanity_MEX.drop(['PROJECTDASH_ID', 'Segment'], axis = 1)
sanity_MEX = sanity_MEX.rename(pm_cpm_rename_dict, axis = 1)
sanity_MEX['Total Alertas'] = sanity_MEX[['AE/AB/Revenue < 0', 
                                         'OOH = 0', 
                                         'GM>50%',
                                         'Plan Cost < Actual + Committed',
                                         'WIP > OOH',
                                         'Accrual > Plan Cost - Actual']].apply(lambda x: (x != 'OK').sum(), axis=1)

sanity_MEX = sanity_MEX[['Country',
                       'CPM',
                       'PM',
                       'Project Number',
                       'Description',
                       'Profit Center',
                       'Portfolio',
                       'Category',
                       'Segment',
                       'Status',
                       'AE/AB/Revenue < 0',
                       'OOH = 0',
                       'GM>50%',
                       'Plan Cost < Actual + Committed',
                       'WIP > OOH',
                       'Accrual > Plan Cost - Actual',
                       'Total Alertas',
                       'Comentarios',
                       'New Order',
                       'Order on Hand',
                       'Sales',
                       'Plan Cost',
                       'Committed Cost',
                       'Actual Cost',
                       'Accrual',
                       'COS',
                       'WIP/UC',
                       'Planned GM',
                       'Actual GM',
                       'Created On',
                       'Finish Date',
                       'Proj. is a Loss contract',
                       'Potential Loss Contract',
                       'Plan Cost Updated']]

# Sanity BR

sanitybra = pd.read_excel(
             repository_path + r'\BasesDeDatos\03 - Sanity\Sanity BR.XLSX', 
            decimal =',', 
            thousands = '.',
            usecols = ('A:H,J,L,Q,R,T,U,W,X,Z,AA:AE,AG'))

sanitybra['Country'] = 'BRA' 
sanitybra = sanitybra.loc[1:,:]
sanitybra.columns = [col.strip() for col in sanitybra.columns]

rename_dict = {
    'Project Number': 'WBS',
    'Project Desc.': 'CPM',
    'Plan. Ver. Tot.': 'Revenue Plan',
    'Plan. Cost Tot': 'Plan Cost',
    'Real - Cost': 'CoS',
    'Real Cost Tot': 'Actual Cost',
    'WIP': 'WIP',
    'Commited cost': 'Committed Cost',
    'Progr.Bill.': 'Prog. Billing',
    'Excess Cost': 'Excess Cost',
    'Excess Bill': 'Excess Bill',
    'POC%': 'POC%',
    'Accrual vf': 'Accrual',
    'Order Entry': 'New Order'
}
sanitybra.rename(rename_dict, axis = 1, inplace = True)

sanitybra = sanitybra.dropna(subset = ['Status'])

sanitybra['Project Type'] = sanitybra['WBS'].apply(proj_type)
sanitybra['Start date'] = sanitybra['Start date'].apply(to_date)
sanitybra['Finish date'] = sanitybra['Finish date'].apply(to_date)

columns_to_change = ['Plan Cost', 'New Order', 'Order on Hand', 'Real - Revenue', 
                    'WIP', 'Prog. Billing', 'Excess Bill']

for col in columns_to_change:
    sanitybra[col] = sanitybra[col].apply(change_sign)

sanitybra = sanitybra[[
    'Country',
    'WBS',
    'Customer name',
    'Profit Center',
    'Status',
    'Project Type',
    'New Order',
    'Order on Hand',
    'Real - Revenue',
    'Plan Cost',
    'Committed Cost',
    'Actual Cost',
    'CoS',
    'Revenue Plan',
    'Real - Margin %',
    'WIP',
    'Accrual',
    'POC%',
    'Prog. Billing',
    'Excess Cost',
    'Excess Bill',
    'Start date',
    'Finish date'
]]

sanitybra['AE/AB/Revenue/GM < 0'] = sanitybra.apply(negativo, axis = 1)
sanitybra['OOH = 0'] = sanitybra.apply(abcero, axis = 1)
sanitybra['G. Margin > 70%'] = sanitybra.apply(margin, axis = 1)
sanitybra['Plan Cost < Actual + Committed'] = sanitybra.apply(plan_cost, axis = 1)
sanitybra['WIP > OOH'] = sanitybra.apply(wip, axis = 1)
sanitybra['Accrual > Actual + Committed'] = sanitybra.apply(accrual_costs, axis = 1)
sanitybra['Close Projects'] = sanitybra.apply(close_project, axis = 1)
sanitybra['POC = 100%'] = sanitybra.apply(poc_100, axis = 1)
sanitybra['Revenue > Plan Revenue'] = sanitybra.apply(rev_planrev, axis = 1)
sanitybra['High Cost in Excess'] = sanitybra.apply(high_cost, axis = 1)
sanitybra['Total Alertas'] = ''
sanitybra['Comentarios'] = ''

sanity_BRA = pd.merge(sanitybra, pm_cpm, left_on = 'WBS', right_on = 'PROJECTDASH_ID', how = 'left')
sanity_BRA = sanity_BRA.drop('PROJECTDASH_ID', axis = 1)

sanity_BRA['Total Alertas'] = sanity_BRA[[
                                        'AE/AB/Revenue/GM < 0',
                                        'OOH = 0',
                                        'G. Margin > 70%',
                                        'Plan Cost < Actual + Committed',
                                        'WIP > OOH',
                                        'Accrual > Actual + Committed',
                                        'Close Projects',
                                        'POC = 100%',
                                        'Revenue > Plan Revenue',
                                        'High Cost in Excess']].apply(lambda x: (x != 'OK').sum(), axis=1)

sanity_BRA = sanity_BRA.rename(pm_cpm_rename_dict, axis = 1)

sanity_BRA = sanity_BRA[[
            'Country',
            'CPM',
            'PM',
            'WBS',
            'Customer name',
            'Profit Center',
            'Portfolio',
            'Category',
            'Segment',
            'Status',
            'Project Type',
            'AE/AB/Revenue/GM < 0',
            'OOH = 0',
            'G. Margin > 70%',
            'Plan Cost < Actual + Committed',
            'WIP > OOH',
            'Accrual > Actual + Committed',
            'Close Projects',
            'POC = 100%',
            'Revenue > Plan Revenue',
            'High Cost in Excess',
            'Total Alertas',
            'Comentarios',
            'New Order',
            'Order on Hand',
            'Real - Revenue',
            'Plan Cost',
            'Committed Cost',
            'Actual Cost',
            'CoS',
            'Revenue Plan',
            'Real - Margin %',
            'WIP',
            'Accrual',
            'POC%',
            'Prog. Billing',
            'Excess Cost',
            'Excess Bill',
            'Start date',
            'Finish date'
            ]]

# Resumens

def create_sanity_summary(sanity_df, id_column='Project Number', checks=None, use_shape=False):
    resumen = pd.DataFrame()
    
    resumen['CantProj'] = sanity_df.groupby('Country')[id_column].nunique()
    
    if 'Ord. Bckl. GM %' in sanity_df.columns:
        resumen['%PromAB'] = round(sanity_df.groupby('Country')['Ord. Bckl. GM %'].mean(), 2)
    
    if checks:
        for col in checks:
            filtered = sanity_df.loc[sanity_df[col] != 'OK', :]
            if use_shape:
                resumen[col] = filtered.shape[0]
            else:
                resumen[col] = filtered.groupby('Country')[id_column].nunique()
    
    resumen = resumen.fillna(0)
    resumen.reset_index(inplace=True)
    
    return resumen

cols_sa_na = [
    'AE/AB/Revenue < 0',
    'AE = 0',
    'OOH = 0',
    'GM>50%',
    'Plan Cost < Actual + Committed',
    'WIP > OOH',
    'Accrual > Plan Cost - Actual'
]

cols_mex = [
    'AE/AB/Revenue < 0',
    'OOH = 0',
    'GM>50%',
    'Plan Cost < Actual + Committed',
    'WIP > OOH',
    'Accrual > Plan Cost - Actual'
]

cols_bra = [
    'AE/AB/Revenue/GM < 0',
    'OOH = 0',
    'G. Margin > 70%',
    'Plan Cost < Actual + Committed',
    'WIP > OOH',
    'Accrual > Actual + Committed',
    'Close Projects',
    'POC = 100%',
    'Revenue > Plan Revenue',
    'High Cost in Excess'
]

resumen_SA = create_sanity_summary(sanity_SA, checks=cols_sa_na)
resumen_NA = create_sanity_summary(sanity_NA, checks=cols_sa_na)
resumen_mex = create_sanity_summary(sanity_MEX, checks=cols_mex, use_shape=True)
resumen_bra = create_sanity_summary(sanity_BRA, id_column='WBS', checks=cols_bra, use_shape=True)

# Excel File Handling

def format_title(sheet, cell, text, columns=4):
    range_str = f'{cell}:{chr(ord(cell[0]) + columns - 1)}{cell[1:]}'
    sheet.range(cell).value = text
    sheet.range(range_str).merge()
    sheet.range(cell).api.Font.Bold = True
    sheet.range(cell).api.Font.Size = 15
    sheet.range(cell).api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    sheet.range(cell).color = (96, 73, 122)
    sheet.range(cell).api.Font.ColorIndex = 2
    sheet.range(cell).columns.autofit()
    return sheet

def format_table(sheet, start_cell, df, columns):
    end_col = chr(ord(start_cell[0]) + columns - 1)
    end_row = int(start_cell[1:]) + df.shape[0]
    range_str = f'{start_cell}:{end_col}{end_row}'
    header_range = f'{start_cell}:{end_col}{start_cell[1:]}'
    
    sheet.range(start_cell).options(index=False).value = df
    sheet.range(header_range).color = (177, 160, 199)
    sheet.range(range_str).api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    sheet.range(range_str).api.Borders.LineStyle = 1
    sheet.range(range_str).api.Borders.Weight = 2
    for col in range(ord(start_cell[0]) - ord('A'), ord(end_col) - ord('A') + 1):
        sheet.api.Columns(col + 1).AutoFit()
    
    return sheet

def create_report(workbook):
    wb = xw.Book() if workbook is None else workbook
    
    sheet_names = ['Summary', 'Sanity SA', 'Sanity NA', 'Sanity MEX', 'Sanity BRA']
    wb.sheets[0].name = sheet_names[0]
    for name in sheet_names[1:]:
        wb.sheets.add(name=name, after=wb.sheets[-1])
    
    wb.sheets[0].range('A1:BR500').color = (255, 255, 255)
    
    titles = [
        ('A1', 'Resumen Alertas SA'),
        ('A10', 'Resumen Alertas NA'),
        ('A16', 'Resumen Alertas Brasil'),
        ('A21', 'Resumen Alertas MesoamÃ©rica')
    ]
    
    tables = [
        ('A3', resumen_SA, 10),
        ('A12', resumen_NA, 10),
        ('A18', resumen_bra, 12),
        ('A23', resumen_mex, 8)
    ]
    
    summary_sheet = wb.sheets[0]
    
    for cell, text in titles:
        format_title(summary_sheet, cell, text)
    
    for start_cell, df, columns in tables:
        format_table(summary_sheet, start_cell, df, columns)
    
    return wb

def format_sheet(workbook, sheet_name, df, title, exception_columns, border_col_range):
    sheet = workbook.sheets[sheet_name]
    
    sheet.range('A1').value = title
    sheet.range('A1:B1').merge()
    sheet.range('A1').api.Font.Bold = True
    sheet.range('A1').api.Font.Size = 15
    sheet.range('A1').color = (96, 73, 122)
    sheet.range('A1').api.Font.ColorIndex = 2
    sheet.range('A1').api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    start_range = sheet.range('A3')
    start_range.options(index=False).value = df
    end_col = chr(ord('A') + df.shape[1] - 1)
    end_row = 3 + df.shape[0]
    data_range = f'A3:{end_col}{end_row}'
    row3_range = sheet.range('3:3').expand('right')
    row3_range.api.Font.Italic = True
    row3_range.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    exception_start_col = None
    exception_end_col = None
    for cell in row3_range:
        if (cell.value in exception_columns or 
            cell.value == "Total Alertas" or 
            cell.value == "Comentarios"):
            if exception_start_col is None:
                exception_start_col = cell.column
            exception_end_col = cell.column
            cell.color = (96, 73, 122)
            cell.api.Font.ColorIndex = 2
        else:
            cell.color = (177, 160, 199)
    
    if exception_start_col is not None:
        row2_start = sheet.cells(2, exception_start_col)
        row2_end = sheet.cells(2, exception_end_col)
        row2_start.value = 'Alertas'
        row2_start.api.Font.ColorIndex = 2
        row2_range = f'{row2_start.address}:{row2_end.address}'
        sheet.range(row2_range).color = (96, 73, 122)
        sheet.range(row2_range).api.Merge()
        sheet.range(row2_range).api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter

    sheet.range(f'A3:{border_col_range}{len(df) + 3}').api.Borders.LineStyle = 1
    sheet.range(f'A3:{border_col_range}{len(df) + 3}').api.Borders.Weight = 2

    for ws in workbook.sheets:
        ws.autofit(axis='columns')
    return sheet

rep = create_report(None)

sheets_configs = {
    'Sanity SA': {'df': sanity_SA, 'title': 'Sanity Check SA', 'exception_columns': cols_sa_na,'border_col_range': 'AQ'},
    'Sanity NA': {'df': sanity_NA, 'title': 'Sanity Check NA', 'exception_columns': cols_sa_na,'border_col_range': 'AQ'},
    'Sanity MEX': {'df': sanity_MEX, 'title': 'Sanity Check MEX', 'exception_columns': cols_mex,'border_col_range': 'AH'},
    'Sanity BRA': {'df': sanity_BRA, 'title': 'Sanity Check BRA', 'exception_columns': cols_bra,'border_col_range': 'AN'}
}

for sheet_num, config in sheets_configs.items():
    format_sheet(
        workbook = rep,
        sheet_name = sheet_num,
        df = config['df'],
        title = config['title'],
        exception_columns = config['exception_columns'],
        border_col_range = config['border_col_range']
    )

rep.save(file_final_path)
rep.close()

wb = openpyxl.load_workbook(file_final_path)

ws = wb['Sanity SA']
ws1 = wb['Sanity NA']
ws2 = wb['Sanity BRA']
ws3 = wb['Sanity MEX']

cell = ws['G4']
cell1 = ws1['G4']
cell2 = ws2['F4']
cell3 = ws3['F4']

ws.freeze_panes = cell 
ws1.freeze_panes = cell1
ws2.freeze_panes = cell2
ws3.freeze_panes = cell3

wb.save(file_final_path)

os.startfile(final_repository)