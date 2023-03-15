import glob

import pandas as pd

storages_list = ['IN_825', 'R12_825']
groups_tg_list = ['30', '31', '32', '33', '34', '35', '36', '37', '38', '39', '40', '41', '42', '43',
                  '44', '45', '46', '47', '48', '49', 'Z7']


def read_check_file(dir_name):
    files = glob.glob(f'{dir_name}/*.xlsx')
    print(files)

    df_union = pd.DataFrame()

    for file in files:
        df_temp = pd.read_excel(file)
        df_temp = df_temp[~(df_temp['Склад'].isin(storages_list))]
        print(file, len(df_temp))
        df_union = pd.concat([df_union, df_temp])
    print(f'длина до объединения: {len(df_union)}')
    df_union_art = df_union.groupby(['Код номенклатуры'], as_index=False).agg({
        'Модель': 'first',
        'Товарная группа': 'first',
        'Инвентарная разница': 'sum',
        'Себестоимость': 'first',
    })
    df_union_art = df_union_art.assign(Сумма=df_union_art['Инвентарная разница'] * df_union_art['Себестоимость'])
    df_union_art = df_union_art[df_union_art['Инвентарная разница'] != 0]

    df_union_art['Товарная группа'] = df_union_art['Товарная группа'].astype('string')

    with pd.ExcelWriter('Сводка.xlsx', engine='xlsxwriter') as writer:
        df_union_art2 = df_union_art.groupby(['Товарная группа'], as_index=False).agg({
            'Сумма': 'sum',
        })
        df_union_art2 = df_union_art2.sort_values('Товарная группа')
        df_union_art2.to_excel(writer, sheet_name='Сумма по тг', index=False, header=True, na_rep='')
        worksheet = writer.sheets['Сумма по тг']
        (max_row, max_col) = df_union_art2.shape
        worksheet.autofilter(0, 0, max_row, max_col - 1)
        worksheet.set_column('A:A', 17)
        worksheet.set_column('B:B', 30)

        df_union_art.to_excel(writer, sheet_name='Общий список', index=False, header=True, na_rep='')
        worksheet = writer.sheets['Общий список']
        (max_row, max_col) = df_union_art.shape
        worksheet.autofilter(0, 0, max_row, max_col - 1)
        worksheet.set_column('A:F', 20)
        worksheet.set_column('B:B', 60)

    print(f'длина после объединения: {len(df_union_art)}')
    tg_list = df_union_art['Товарная группа'].unique().tolist()
    list_art_tg = df_union_art['Код номенклатуры'].unique().tolist()

    list_zal = ['A11_825', 'V_825']
    list_sklad = ['011_825', '012_825']

    with pd.ExcelWriter('Расхождения по тг Зал.xlsx', engine='xlsxwriter') as writer:
        df_sklad = pd.read_excel('/home/mikhail/Downloads/mishafrishkinloh.xlsx', skiprows=14,
                                 usecols=['Склад', 'Местоположение', 'Код \nноменклатуры', 'Описание товара',
                                          'ТГ', 'НГ', 'Физические \nзапасы', 'Передано на доставку',
                                          'Продано', 'Зарезерви\nровано', 'Доступно'])

        for tg in tg_list:
            try:
                temp_df = df_union_art[df_union_art['Товарная группа'] == tg]

                temp_df.to_excel(writer, sheet_name=f'{tg}', index=False, header=True, na_rep='')
                worksheet = writer.sheets[f'{tg}']
                set_column(temp_df, worksheet)

                df_sklad_tg = df_sklad[(df_sklad['ТГ'] == tg) & (df_sklad['Код \nноменклатуры'].isin(list_art_tg))
                                       & df_sklad['Склад'].isin(list_zal)]
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
    read_check_file('/home/mikhail/Downloads/инвентура/Журналы инвентуры')
