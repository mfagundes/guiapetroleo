# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
from pathlib import Path
import settings
from data_wrangling import get_dataframe, find_duplicates, find_branches, find_headquarters, create_excel_file

FILE = settings.DATA_INPUT / '2021-02-13_GuiaPetroleo.xlsx'


def save_dataframes_unique(file):
    df = get_dataframe(file, worksheet=0, usecols=("A:O"))
    branches = find_branches(df)
    branches = find_duplicates(branches)
    headquartes = find_headquarters(df)
    headquartes = find_duplicates(headquartes)
    worksheets = {
        'filiais': branches,
        'matrizes': headquartes
    }
    writer = create_excel_file(worksheets, filename='petroleo_empresas')
    writer.save()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    try:
        Path.mkdir(Path(settings.DATA_INPUT))
        Path.mkdir(Path(settings.DATA_OUTPUT))
        Path.mkdir(Path(settings.REPORTS_PATH))
        Path.mkdir(Path(settings.PROFILE_PATH))
    except FileExistsError as e:
        pass
    else:
        print("Pastas de saída criadas")

    save_dataframes_unique(FILE)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
