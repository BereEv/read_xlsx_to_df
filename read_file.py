import glob

import openpyxl
import pandas as pd
import time

class Empty():

    dir_in = ''
    num_sheets = 0
    noms = 0
    flags = ''
    file_part = ''

    def read_con(self):
        start = time.time()
        files = glob.glob(f'{self.dir_in}/{self.file_part}*')
        df = pd.DataFrame()
        for file in files:
            with open(file, 'r', encoding='utf-8') as file:
                wb = openpyxl.load_workbook(
                    file.name,
                    data_only=True,
                    keep_vba=True,
                    read_only=True)
                ws = wb.worksheets[self.num_sheets]
            lst = []
            col_lst = []
            count = 0
            flag = self.flags
            nom = self.noms
            strt = 0
            stp = 0
            flag_row = ''
            for rows in ws.values:
                rows += tuple([file.name])
                lst.append(rows)
                count += 1
                if flag in rows:
                    stp = count + nom
                    strt = count
                    flag_row = rows
                    
            column = []
            if self.noms > 0:
                col_lst = [row for row in zip(*lst[strt - 1:stp])]

                for tup in col_lst:
                    column.append([item for item in tup if item is not None])

                for del_item in ['[', ']', "'", 'None']:
                    column = [str.replace(str(item), del_item, '')
                              for item in column]
                    column = [str.replace(str(item), ',', '.')
                              for item in column]
            else:
                column = [str(row) for row in flag_row]
                column = [str.replace(str(item), 'None', '')
                          for item in column]

            column.pop()
            column.append('Name')
            column = tuple(column)
            print('Наименование столбцов', column, files)
            
            if self.noms > 0:
                data_lst = (ls for ls in lst[stp:])
            else:
                data_lst = (ls for ls in lst[strt:])

            #data_lst = tuple(data_lst)

            dd = pd.DataFrame(columns=column, data=data_lst)

            df = pd.concat([df, dd])
            #df = df.dropna(axis=1, how='all')
            #df = df.fillna(0)
            df = df.rename(columns=lambda x: str(x).strip())
            
        #if self.file_part != '':
         #   df.to_excel(f'log/{self.file_part}.xlsx')
        #else:
         #   df.to_excel(f'log/{self.dir_in}.xlsx')
        print(f'Файл {self.dir_in, self.file_part} прочитан')
        stop = time.time()
        print(stop - start)
        yield df
