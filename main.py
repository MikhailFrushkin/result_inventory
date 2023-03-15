import pandas as pd
import numpy as np

storages_list = ['011_825', '012_825', 'A11_825', ]
groups_tg_list = ['30', '31', '32', '33', '34', '35', '36', '37', '38', '39', '40', '41', '42', '43',
                  '44', '45', '46', '47', '48', '49', 'Z7']


def read_check_file(file_name):
    df = pd.read_excel(file_name, usecols=['Зона', 'Склад', 'Ячейка', 'Артикул', 'Принятие решения'])

    df_none = df[df['Принятие решения'].isnull()]

    df = df[~(df['Принятие решения'].isnull())]
    df['Склад'] = df['Склад'].astype('string').str.lower()
    df['Ячейка'] = df['Ячейка'].astype('string').str.lower()
    df['Принятие решения'] = df['Принятие решения'].astype(np.int64)

    # df.groupby(['group_var'], as_index=False).agg({'string_var ': ' '.join})
    df_union_art = df.groupby(['Артикул'], as_index=False).agg({'Принятие решения': 'sum'})
    art_list_check = df['Артикул'].unique().tolist()

    dict_art_check = {}
    for art in art_list_check:
        for sklad in storages_list:
            temp_df = df[(df['Артикул'] == art) & (df['Склад'] == sklad)]
            temp_df = temp_df.groupby(['Артикул'], as_index=False).agg({'Принятие решения': 'sum'})
            if not temp_df.empty:
                dict_art_check[art] = {sklad: {
                    temp_df.to_dict
                }
                }
    print(dict_art_check)
    # with pd.ExcelWriter('объедененные артикула пересчета.xlsx', engine='xlsxwriter') as writer:
    #     (max_row, max_col) = df.shape
    #     workbook = writer.book
    #     cell_format = workbook.add_format({'align': 'center', 'valign': 'top', 'font_size': 14})
    #
    #     df_union_art.to_excel(writer, sheet_name='Таблица', index=False, header=True, na_rep='')
    #
    #     worksheet = writer.sheets['Таблица']


if __name__ == '__main__':
    read_check_file('/home/mikhail/Downloads/InventoryCountJournal (4).xlsx')
    # read_check_file('/home/mikhail/Downloads/6.1 Складские лоты 06_04_2022.xlsx')
