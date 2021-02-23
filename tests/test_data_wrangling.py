from pathlib import Path

import pytest
import pandas as pd

import settings
from data_wrangling import get_dataframe, find_branches, find_headquarters, find_duplicates, create_excel_file

FILE = 'data/2021-02-13_GuiaPetroleo.xlsx'
USECOLS = ("A:O")



def test_read_file(file=FILE, usecols=USECOLS):
    df = get_dataframe(file, worksheet=0, usecols=usecols)
    assert isinstance(df, pd.core.frame.DataFrame)
    assert df.iloc[0]['Categoria'] == 'Produtos'
    assert df.shape[1]==15


@pytest.mark.parametrize('worksheet, result', (
    (0, 209),
    (1, 223)
))
def test_get_branches(worksheet, result):
    df = get_dataframe(FILE, worksheet=worksheet, usecols=USECOLS)
    df = find_branches(df)
    assert df.shape == (result, 15)



def test_get_headquarters():
    df = get_dataframe(FILE, worksheet=0, usecols=USECOLS)
    headquarters = find_headquarters(df)
    branches = headquarters[headquarters['CNPJ'].str.contains('^(?!.*/0001).*$', regex=True)]
    assert len(branches) == 0



def test_unique_headquarters():
    df = get_dataframe(FILE, worksheet=0, usecols=USECOLS)
    headquarters = find_headquarters(df)
    headquarters = find_duplicates(headquarters)
    assert headquarters.shape == (782, 15)


def test_save_excel():
    df1 = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
    df2 = pd.DataFrame({'C': [1, 2, 3], 'D': [4, 5, 6]})
    worksheets = {'dataframe1': df1, 'dataframe2': df2}
    filename = 'teste'
    excel_file = create_excel_file(worksheets, filename)
    assert excel_file.book.sheetnames.keys() == worksheets.keys()
    assert Path(excel_file.book.filename.name) == settings.DATA_OUTPUT / f'{filename}.xlsx'
