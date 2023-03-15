import pandas as pd
import numpy as np

storages_list = ['011_825', '012_825', 'A11_825', ]
groups_tg_list = ['30', '31', '32', '33', '34', '35', '36', '37', '38', '39', '40', '41', '42', '43',
                  '44', '45', '46', '47', '48', '49', 'Z7']


def read_check_file(file_name):
    df = pd.read_excel(file_name, usecols=['Код \nноменклатуры', 'Описание товара', 'ТГ', 'Количество 6.1',
                                           'Количество просчет', 'Разница'])
    tg_list = df['ТГ'].unique().tolist()
    list_art_tg = df['Код \nноменклатуры'].unique().tolist()

    with pd.ExcelWriter('Расхождения по тг.xlsx', engine='xlsxwriter') as writer:
        df_sklad = pd.read_excel('/home/mikhail/Downloads/mishafrishkinloh.xlsx', skiprows=14,
                                 usecols=['Склад', 'Местоположение', 'Код \nноменклатуры', 'Описание товара',
                                          'ТГ', 'НГ', 'Физические \nзапасы', 'Передано на доставку',
                                          'Продано', 'Зарезерви\nровано', 'Доступно'])

        for tg in tg_list:
            try:
                temp_df = df[df['ТГ'] == tg]

                temp_df.to_excel(writer, sheet_name=f'{tg}', index=False, header=True, na_rep='')
                worksheet = writer.sheets[f'{tg}']
                set_column(temp_df, worksheet)
                df_sklad_tg = df_sklad[(df_sklad['ТГ'] == tg) & (df_sklad['Код \nноменклатуры'].isin(list_art_tg))]

                df_sklad_tg.to_excel(writer, sheet_name=f'Складские лоты ТГ{tg}', index=False, header=True, na_rep='')
                worksheet = writer.sheets[f'Складские лоты ТГ{tg}']
                set_column2(df_sklad_tg, worksheet)
            except Exception as ex:
                print(f'{tg} {ex}')


def set_column(df, worksheet):
    (max_row, max_col) = df.shape
    worksheet.autofilter(0, 0, max_row, max_col - 1)
    worksheet.set_column('A:A', 17)
    worksheet.set_column('B:B', 60)
    worksheet.set_column('C:J', 20)

def set_column2(df, worksheet):
    (max_row, max_col) = df.shape
    worksheet.autofilter(0, 0, max_row, max_col - 1)
    worksheet.set_column('A:A', 10)
    worksheet.set_column('B:B', 20)
    worksheet.set_column('C:C', 15)
    worksheet.set_column('D:D', 50)
    worksheet.set_column('E:F', 10)
    worksheet.set_column('G:K', 20)


if __name__ == '__main__':
    read_check_file('/home/mikhail/Downloads/Инвентура общая V_Sales,011,012,A11.xlsx')
