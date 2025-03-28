import pandas as pd
import numpy as np
import datetime

df_feminism = pd.DataFrame(pd.read_excel(r'C:\Users\leoso\Projects\Amazon Scraping\Livros Amazon.xlsx', sheet_name = 'Feminism_DB'))
df_antifeminism = pd.DataFrame(pd.read_excel(r'C:\Users\leoso\Projects\Amazon Scraping\Livros Amazon.xlsx', sheet_name = 'Antifeminism_DB'))

def convert_date(df, coluna='Data'):
    if coluna not in df.columns:
        raise ValueError(f"A coluna '{coluna}' não foi encontrada no DataFrame.")
    df[coluna] = pd.to_datetime(df[coluna], format='%d/%m/%Y', errors='coerce')
    return df

feminism_filter = (
    (df_feminism['Tipo Livro 1'] == 'Kindle') &
    (
        (df_feminism['Tipo Livro 2'].isna()) |
        (~df_feminism['Tipo Livro 2'].str.contains('Capa', na=False))
    )
)

df_feminism_kindle_only = df_feminism[feminism_filter].reset_index(drop = True).drop([10, 27, 76, 89, 103, 108]).drop_duplicates(subset='Nome')
df_feminism_no_kindle_only = pd.concat([df_feminism, df_feminism_kindle_only, df_feminism_kindle_only]).drop_duplicates(keep=False).drop_duplicates(subset='Nome')
antifeminism_filter = (
    (df_antifeminism['Tipo Livro 1'] == 'Kindle') &
    (
        df_antifeminism['Tipo Livro 2'].isna() |
        (~df_antifeminism['Tipo Livro 2'].str.contains('Capa', na=False))
    )
)

df_antifeminism_kindle_only = df_antifeminism[antifeminism_filter].reset_index(drop = True).drop(10).drop_duplicates(subset='Nome')
df_antifeminism_no_kindle_only = pd.concat([df_antifeminism, df_antifeminism_kindle_only, df_antifeminism_kindle_only]).drop_duplicates(keep=False).drop_duplicates(subset='Nome')
df_duplicates_physical_media = pd.concat([df_feminism_no_kindle_only, df_antifeminism_no_kindle_only])
df_duplicates_physical_media['Is_Duplicate'] = df_duplicates_physical_media.duplicated(subset=['Nome'])
df_duplicates_physical_media = df_duplicates_physical_media[df_duplicates_physical_media['Is_Duplicate'] == True]
df_duplicates_physical_media = df_duplicates_physical_media.reset_index(drop=True)
df_duplicates_kindle_only = pd.concat([df_feminism_kindle_only, df_antifeminism_kindle_only])
df_duplicates_kindle_only['Is_Duplicate'] = df_duplicates_kindle_only.duplicated(subset=['Nome'])
df_duplicates_kindle_only = df_duplicates_kindle_only[df_duplicates_kindle_only['Is_Duplicate'] == True]
df_duplicates_kindle_only = df_duplicates_kindle_only.reset_index(drop=True)

to_excel_dict = {
    'Feminism_Physical_Media': df_feminism_no_kindle_only,
    'Feminism_Kindle_Only': df_feminism_kindle_only,
    'AntiFeminism_Physical_Media': df_antifeminism_no_kindle_only,
    'AntiFeminism_Kindle_Only': df_antifeminism_kindle_only,
    'Duplicates_Physical_Media': df_duplicates_physical_media,
    'Duplicates_Kindle_Only': df_duplicates_kindle_only
}

dfs_list = list(to_excel_dict.values())
for i in dfs_list:
    i = convert_date(i)

with pd.ExcelWriter('Livros Amazon.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    for sheet_name, df in to_excel_dict.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)