import openpyxl
import xlsxwriter
import pandas as pd
import numpy as np
import settings


def get_dataframe(file, worksheet=None, usecols=None) -> pd.core.frame.DataFrame:
    df = pd.read_excel(file, sheet_name=worksheet, usecols=usecols)
    df['CNPJ'] = df['CNPJ'].apply(lambda x: str(x) if type(x) == int else x)
    return df


def find_branches(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(subset=['CNPJ'])
    branches = df[df['CNPJ'].str.contains('^(?!.*/0001).*$', regex=True)]
    return branches


def find_headquarters(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(subset=['CNPJ'])
    df_branches = find_branches(df)
    df_headquarters = df.drop(index=df_branches.index)
    return df_headquarters


def find_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    return df.drop_duplicates(subset=['CNPJ'], keep='first')


def create_excel_file(dfs: dict, filename: str) -> pd.ExcelWriter:
    writer = pd.ExcelWriter(settings.DATA_OUTPUT / f'{filename}.xlsx', engine='xlsxwriter')
    for item in dfs.items():
        item[1].to_excel(writer, item[0])
    return writer
