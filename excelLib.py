from collections import Counter
import re
import openpyexcel
import uExcelFunc
import read_ini


class ExcelData(object):
    def __init__(self):
        self.read_param = read_ini.ReadIni()
        #self.workbook = openpyexcel.load_workbook('./2613 Бюджет проекта.xlsm', data_only=True)
        #self.sheet = self.workbook['Лист1']
        self.workbook1 = openpyexcel.load_workbook(self.read_param.get_excel_param('budget_path'), data_only=True,keep_vba=True)
        self.sheet1 = self.workbook1[self.read_param.get_excel_param('name_of_list')]
        key_id = self.read_param.get_excel_key_id()
        data_id = self.read_param.get_excel_data_id()
        self.excel_key_column = []
        self.excel_data = []
        self.excel_data_column = []
        self.excel_key_column.append(uExcelFunc.findCell(self.workbook1, key_id).column)
        for i in data_id:
            self.excel_data.append(self.get_list(i))
            self.excel_data_column.append(uExcelFunc.findCell(self.workbook1, i).column)
        self.excel_key = self.get_list(self.read_param.get_excel_key_id())
        self.cIdColumn = []
        self.data_row = uExcelFunc.findCell(self.workbook1, data_id[0]).row

    def get_key_data2(self):
        max_row = self.sheet1.max_row
        cell_start = uExcelFunc.findCell(self.workbook1, self.read_param.get_excel_key_id())
        row_start = cell_start.row
        column_start = cell_start.column
        analyzed_id = []
        for i in range(row_start + 1, max_row + 1):
            cell = self.sheet1[str(column_start) + str(i)].value
            if cell is not None:
                if analyzed_id:
                    for j in analyzed_id:
                        if j == cell:
                            continue
                        else:
                            analyzed_id.append(cell)
                else:
                    analyzed_id.append(cell)
                self.cIdColumn.append(cell)
        print(self.cIdColumn)

    def get_key_data(self):
        max_row = self.sheet.max_row
        cell_start = uExcelFunc.findCell(self.workbook1, self.read_param.get_excel_key_id())
        row_start = cell_start.row
        column_start = cell_start.column
        analyzed_id = []
        for i in range(row_start + 1, max_row + 1):
            cell = self.sheet1[str(column_start) + str(i)].value
            if cell is not None:
                match_cipher = re.search(r'\d{4}.\d{3}.\d{3}', cell)
                if match_cipher:
                    if analyzed_id:
                        for j in analyzed_id:
                            if j == cell:
                                continue
                            else:
                                analyzed_id.append(cell)
                    else:
                        analyzed_id.append(cell)
                    self.cIdColumn.append(cell)
        print(self.cIdColumn)

    def get_wb(self):
        return self.workbook1

    def get_key(self):
        return self.cIdColumn

    def get_list(self, cell_id):
        cell_key = uExcelFunc.findCell(self.workbook1, cell_id)
        max_row = self.sheet1.max_row
        list_key = []
        for i in range(cell_key.row + 1, max_row):
            list_key.append(self.sheet1[cell_key.column + str(i)].value)
        return list_key

    def cycle3(self, g_data):
        list_repeat = []
        row_excel_result = 1
        key_count = 0
        key_none_count = 0
        count_of_data = len(g_data[0])-1
        number_of_gData = 0
        print(count_of_data)
        # self.excel_key[len(g_data)] = ''
        for i in self.excel_key:
            exc_repeat = excel_key_repeat(self.excel_key, i)
            if i is None:
                key_none_count += 1
            if i is not None:
                if is_key_repeat(list_repeat, i):
                    print(
                        'Повтор ключа в excel: ' + str(i) + ' ' + self.excel_data_column[0] + str(row_excel_result + 1))
                    row_excel_result += 1
                    continue

                for k in range(1, count_of_data):

                    if k == 1:
                        row_excel_memory = row_excel_result
                    if k > 1:
                        row_excel_result = row_excel_memory
                    # row_excel_memory = 0
                    repeat_count = 0
                    g_count = 0
                    g_none_count = 0
                    for j in g_data:
                        if j is None:
                            g_none_count += 1
                            row_excel_result +=1

                        several = j[0][0].find(i)
                        if several is not -1:
                        # if i == j[0][0]:
                            repeat_count += 1
                            print(len(g_data[g_count]))
                            if repeat_count > exc_repeat:
                                if k==1:
                                    # print(g_data[g_count][k][0])
                                    print('добавляю значение по координате с добавлением строки ' + str(self.excel_data_column[k - 1]) + str(row_excel_result + self.data_row + key_count) + ' ' +g_data[g_count][k][0])
                                    self.sheet1.insert_rows(idx=row_excel_result + self.data_row)

                                    # print('ячейка: '+ str(self.sheet1[self.excel_key_column[0] + str(row_excel_result + 1 + self.data_row)].value))
                                self.sheet1[self.excel_data_column[k - 1] + str(row_excel_result + self.data_row + key_count )].value = str(g_data[g_count][k][0])

                            else:
                                print('добавляю значение по координате без добавления строки ' + self.excel_data_column[k - 1] + str(row_excel_result  + self.data_row + key_count) + ' ' + str(g_data[g_count][k][0]))
                                self.sheet1[self.excel_data_column[k - 1] + str(row_excel_result + self.data_row + key_count)].value = str(g_data[g_count][k][0])
                            row_excel_result += 1
                        g_count += 1
                    k += 1
                g_data[number_of_gData][count_of_data] = True
                number_of_gData+=1
            #key_count += 1
        self.workbook1.save(self.read_param.get_excel_param('budget_path'))
        self.workbook1.close()

    def cycle4(self,g_data):
        count_of_columns = len(g_data[0])-1
        excel_key_count = 0
        for i in self.excel_key:
            if i is None:
                excel_key_count+=1
                continue
            g_data_count = 0
            repeat_count = 0
            for j in g_data:
                if g_data[g_data_count][count_of_columns] == False:
                    several = j[0][0].find(i)
                    if several is not -1:
                        repeat_count += 1
                        if repeat_count>1 and len(j[0][0])!=len(i):
                            continue
                        if repeat_count>1:
                            self.sheet1.insert_rows(idx=excel_key_count + self.data_row + repeat_count)
                        for k in range (1,count_of_columns):
                            self.sheet1[self.excel_data_column[k-1]+str(self.data_row+excel_key_count + repeat_count)].value = str(g_data[g_data_count][k][0])
                        g_data[g_data_count][count_of_columns] = True
                g_data_count += 1
                excel_key_count += 1
            else:
                excel_key_count += 1
        self.workbook1.save(self.read_param.get_excel_param('budget_path'))
        self.workbook1.close()




def excel_key_repeat(excel_keys, key):
    repeat = False
    repeat_count = 0
    for i in excel_keys:
        if key == i:
            if repeat_count == 0:
                repeat = True
                repeat_count = 1
                continue
            else:
                return repeat_count
        else:
            if repeat:
                if i is None:
                    repeat_count += 1
                else:
                    repeat = False
                    return repeat_count
    return 0


def is_key_repeat(list_score, a):
    score = 0
    if not list_score:
        list_score.append(a)
        return 0
    else:
        for i in list_score:
            if i == a:
                score = 1
            if score == 0:
                score = -1
        list_score.append(a)
        if score == 1:
            return True
        if score == -1:
            return False


def is_now_space(key, row_excel_result):
    if key[row_excel_result] == '':
        return True
    else:
        return False
