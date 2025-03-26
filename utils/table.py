'''
table main class
'''
import datetime as dt
import pandas as pd
class SelectDB():
    '''
    super class
    '''
    def __init__(self, db_name, db_sheet, logger):
        self.db_name = db_name
        self.db_sheet = db_sheet
        self.logger= logger
        self.select_db = pd.read_excel(self.db_name,
                                       self.db_sheet)
    def select_contract_row(self, get_col_list, contract_col):
        '''
        select need mod rows
        '''
        contract =  get_col_list.index(contract_col)
        col_list = self.select_db.iloc[:,contract]
        contract_list = []
        for x in col_list:
            try:
                contract_list.append(int(x))
            except ValueError:
                contract_list.append(0)
        return contract_list
    def get_person_data(self, get_col_list, contract_num_row, col_name):
        '''
        get person rate value
        '''
        rate = get_col_list.index(col_name)
        return self.select_db.iloc[contract_num_row, rate]
class WorkDB(SelectDB):
    '''
    table main class
    '''
    def __init__(self, db_name, db_sheet, logger):
        SelectDB.__init__(self, db_name,
                         db_sheet, logger)
        self.contract_col = '№ договора '
        self.rate_col = 'тариф'
        self.debt_col = 'Сумма долга'
        self.connect_date = 'Дата.подключения.'
        self.get_col_list = list(self.select_db.iloc[0,:])
        self.get_index_today = self.get_index_today_func()
        self.get_index_2019_date = self.get_index_date_func('01.01.2019')
        self.get_index_2023_date = self.get_index_date_func('01.01.2023')
        self.get_index_2025_date = self.get_index_date_func('01.01.2025')
    def update_contract_value(self, contract_list, true_value_list_contract,
                              value_list):
        '''
        update cell
        '''
        for index, x in enumerate(contract_list):
            self.cell_update(x, true_value_list_contract[index],
                             value_list[index])
        return self.select_db
    def cell_update(self, contract_num_row, true_value_list_contract,
                    value):
        '''
        pass
        '''
        #get rate value
        rate_cell = int(SelectDB.get_person_data(self, self.get_col_list,
                                    contract_num_row,
                                    self.rate_col))
        debt_cell = SelectDB.get_person_data(self, self.get_col_list,
                                    contract_num_row,
                                    self.debt_col)
        num_pay_col = self.get_last_activ_date(contract_num_row)
        last_pay_col = self.get_last_pay_col(contract_num_row, num_pay_col)
        last_pay_value = self.select_db.iloc[contract_num_row, last_pay_col]
        if  pd.isna(last_pay_value):
            last_pay_value = 0
        new_value = value + last_pay_value
        self.logger.info('pay = %s, %s', new_value, contract_num_row)
        self.update_str_pay(contract_num_row, last_pay_col, rate_cell, new_value)
        self.logger.info('contract num %s,  debt = %s, pay = %s',
        true_value_list_contract, debt_cell, value)
        return self.select_db
    def get_col_connect_date(self, contract_num_row):
        '''
        получить дату подключения
        '''
        connect_date = SelectDB.get_person_data(self, self.get_col_list,
                                    contract_num_row,
                                    self.connect_date)
        col_date = dt.datetime.combine(connect_date.replace(day = 1),
                                       dt.time())
        return self.get_col_list.index(col_date)
    def get_last_activ_date(self, contract_num_row):
        '''
        get date collumn stop pay
        '''
        num_date_col = self.get_col_connect_date(contract_num_row)
        cell_value = self.select_db.iloc[contract_num_row, num_date_col]
        while not pd.isna(cell_value):
            try:
                num_date_col+=1
                cell_value = self.select_db.iloc[contract_num_row, num_date_col]
            except IndexError:
                num_date_col-=1
                df = pd.DataFrame({f'{len(self.select_db.columns)}':[None]* len(self.select_db)})
                self.select_db = pd.concat([self.select_db, df], axis=1)
        return num_date_col-1
    def update_str_pay(self, contract_num_row, last_pay_col, rate, value):
        ''' 
        update pay cell
        '''
        temp = False
        while value >= rate:
            temp = True
            try:
                if self.select_db.iloc[contract_num_row,
                            last_pay_col] != 'zzz':
                    self.select_db.iat[contract_num_row,
                                last_pay_col] = rate
                    value-=rate
                last_pay_col+=1
            except IndexError:
                last_pay_col-=1
                value+=rate
                df = pd.DataFrame({f'{len(self.select_db.columns)}':[None]* len(self.select_db)})
                self.select_db = pd.concat([self.select_db, df], axis=1)
        last_pay_col-=1 if temp else 0
        try:
            last_pay_value = self.select_db.iloc[contract_num_row, last_pay_col]
        except IndexError:
            df = pd.DataFrame({f'{len(self.select_db.columns)}':[None]* len(self.select_db)})
            self.select_db = pd.concat([self.select_db, df], axis=1)
            last_pay_value = self.select_db.iloc[contract_num_row, last_pay_col]
        if pd.isna(last_pay_value):
            last_pay_value = 0
        self.select_db.iat[contract_num_row,
                            last_pay_col] = value + last_pay_value
        return self.select_db
    def get_last_pay_col(self, contract_num_row, num_pay_col):
        '''
        получить столбик последней оплаты
        '''
        base_num=num_pay_col
        if self.select_db.iloc[contract_num_row, num_pay_col] == 'zzz':
            num_pay_col-=1
            while self.select_db.iloc[contract_num_row, num_pay_col] == 'zzz':
                num_pay_col-=1
                if num_pay_col <= self.get_col_connect_date(contract_num_row):
                    return base_num+1
        return num_pay_col
    def get_index_today_func(self):
        '''
        get column today
        '''
        today = dt.datetime.today()
        col_date = dt.datetime.combine(today.replace(day = 1),
                                       dt.time())
        return self.get_col_list.index(col_date)
    def get_index_date_func(self, date):
        '''
        get column start date
        '''
        col_date = dt.datetime.strptime(date, '%d.%m.%Y')
        return self.get_col_list.index(col_date)
class NewDB(SelectDB):
    '''
    update excel file
    '''
    def __init__(self, db_name, db_sheet, logger):
        SelectDB.__init__(self, db_name,
                         db_sheet, logger)
        self.contract_col = 'Номер договора'
        self.pay_col = 'Сумма'
        self.get_col_list = list(self.select_db.columns)
    def get_pay(self):
        '''
        get pay list
        '''
        pay =  self.get_col_list.index(self.pay_col)
        col_list = self.select_db.iloc[:, pay]
        return list(col_list)
    def get_error_contracts(self, select_contract_row_list_except):
        '''
        new df error value
        '''
        need_check = pd.DataFrame()
        for x in select_contract_row_list_except:
            temp = self.select_db.loc[self.select_db[self.contract_col] == int(x)]
            need_check = pd.concat([need_check, temp])
        return need_check.reset_index(drop=True)
