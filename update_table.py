'''
update vtl base
'''
import os
import logging
import datetime as dt
import shutil
import pandas as pd
import xlwings as xw
from utils.table import WorkDB, NewDB
from utils.macros import copy_macros

#перевод на кириллицу
# sys.stdout.reconfigure(encoding='utf-8')
log_name = dt.datetime.now().strftime("%d-%m-%y_%H-%M-%S")
logging.basicConfig(filename=f'logs/{log_name}.log', level=logging.INFO, encoding='utf-8')
logger = logging.getLogger(__name__)



def update_table_main(work_base_name, base_sheet,
                      update_base_name, new_sheet):
    '''
    main update cell func
    '''
    try:
        logger.info('Started')
        logger.info("Копирование фаила")
        source_file = f'{work_base_name[:-5]}.xlsm'
        destination_file = f'{work_base_name[:-5]}_bkp.xlsm'
        shutil.copyfile(source_file, destination_file)
        logger.info("Копирование успешно")
        #defining classes
        vtl = WorkDB(work_base_name, base_sheet, logger)
        contract_table = NewDB(update_base_name, new_sheet, logger)
        logger.info('Create DFs')
        #select need update contract
        mod_contract_list = contract_table.select_contract_row(contract_table.get_col_list,
                                                            contract_table.contract_col)
        select_contract_row = vtl.select_contract_row(vtl.get_col_list,  vtl.contract_col)
        #get index mod rows, except contracts
        select_contract_row_list = []
        true_value_list_contract = []
        select_contract_row_list_except = []
        #get pay list
        value_list = contract_table.get_pay()
        logger.info('get contract pay')
        true_value_list = []
        for index, x in enumerate(mod_contract_list):
            try:
                select_contract_row_list.append(select_contract_row.index(x))
                true_value_list.append(value_list[index])
                true_value_list_contract.append(x)
            except ValueError:
                select_contract_row_list_except.append(x)
        #clear dublicate
        select_contract_row_list_except = list(set(select_contract_row_list_except))
        if select_contract_row_list_except:
            logger.error('error contract number = %s',
                        select_contract_row_list_except)
        #update cell
        vtl.update_contract_value(select_contract_row_list,
                                true_value_list_contract,
                                true_value_list)
        update_main = vtl.select_db
        #drop static value
        drop_col_num  = vtl.get_index_2019_date
        update_main.drop(0, inplace=True)
        update_main.drop(update_main.columns[0:drop_col_num], axis=1, inplace=True)
        #update data format
        error_loading = contract_table.get_error_contracts(select_contract_row_list_except)
        with pd.ExcelWriter(work_base_name,
                            engine='openpyxl',
                            if_sheet_exists='overlay',
                            mode="a") as writer:
            logger.info('start rewrite excel file')
            update_main.to_excel(writer, sheet_name=base_sheet,
                                index=False,
                                header=False,
                                startrow=2,
                                startcol=vtl.get_index_2019_date)
        os.rename(f'{work_base_name[:-5]}.xlsm', f'{work_base_name[:-5]}.xlsx')
        #
        update_main.drop(update_main.index, inplace=True)
        count_row = update_main.shape[0]
        with pd.ExcelWriter('Абонентская База VTL плюс.xlsx',
                            engine='openpyxl',
                            if_sheet_exists='overlay',
                            mode="a") as writer:
            update_main.to_excel(writer, sheet_name='Список абонентов',
                                index=False,
                                header=False,
                                startcol=count_row-1,
                                startrow=2)
        # Используем with для автоматического закрытия файла
        with xw.App(visible=False) as app:  # Запускаем Excel в фоновом режиме
            workbook = app.books.open(f'{work_base_name[:-5]}.xlsx')
            workbook.save(f'{work_base_name[:-5]}.xlsm')
        # Удаляем временный файл
        os.remove(f"{work_base_name[:-5]}.xlsx")
        logger.info('end rewrite excel file %s', base_sheet)
        if select_contract_row_list_except:
            error_loading.to_excel('output.xls', index=False, engine='openpyxl')
        logger.info('копирование макроса')
        target_file = os.path.abspath(source_file)
        export_file = os.path.abspath(destination_file)
        copy_macros(export_file, target_file)
        logger.info('Finished')
        return f"Успешная загрузка, проблемные контракты {select_contract_row_list_except}"
    except FileNotFoundError as e:
        return f"Неверное имя фаила {str(e).split(sep=':', maxsplit=1)[-1]}"
if __name__ == '__main__':
    update_table_main('Абонентская База VTL плюс.xlsm', 'Список абонентов',
                      'оплата.xlsx', '1')