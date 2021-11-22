import re
import os
import uGoogleFunc
import read_ini


class Spreadsheets:
    def __init__(self):  # Констрктор, создающий в классе объект таблицы Google по её id, файлу credentials и token
        read_param = read_ini.ReadIni()
        self.importData = []
        self.keys = []
        self.repeat_list = []
        self.gData = {}
        self.ssId = read_param.get_url()
        startdir = os.path.abspath(os.path.curdir)
        self.name_of_list = read_param.get_google_param('name_of_list')
        #self.ssId = '1s1ghFOoQqvjnbsEI2AChuCayLFSvsoJWnwp5KUCho9M'
        self.ss = uGoogleFunc.init(startdir+'\\credentials', startdir+'\\token')
        read_coords = read_param.get_google_data_id()
        read_coords.insert(0, read_param.get_google_key_id())
        print(read_coords)
        self.coordinates = uGoogleFunc.getCoordOfNr(self.ss, self.ssId, read_coords)


    def get_coord_of_nr(self):  # функция, создающая в классе объект координат по объекту таблицы, её id и именованной ячейке

        return self.coordinates

    def get_column(self):  # функция, создающий в классе объект колонку (gData) по коордиатам её первой ячейки
        list_coordinates = []
        for i in self.coordinates:
            row, column = i.split(' ')
            list_coordinates.append(uGoogleFunc.rowcol_to_a1(row, column))

        print('Координаты ключевых ячеек: ')
        print(list_coordinates)

        for i in list_coordinates:
            match_coord = re.search(r'\b\D\d',i)
            if match_coord:
                letter = i[:1]
            else:
                letter = i[:2]
            request = self.ss.values().get(spreadsheetId=self.ssId,
                                           range= self.name_of_list +'!' + str(i) + ':' + str(letter)).execute()
            # column = [i]=[request.get('values',[])]
            if [request.get('values',[])] ==[]:
                self.gData[i] = ''
            else:
                self.gData[i] = [request.get('values', [])]

        for i in self.gData:
            print(i)

    def find_key(self, a):  # функция, создающая в классе список ключей(номеров колонок),
        key = [a]           # из которых далее вычитывается информация для excel
        score = 0
        # self.keys = []
        id_column = list(self.gData.keys())[0]
        key_column = self.gData[id_column][0]
        for k in key_column:
             score += 1
             switch = 0
             repeat_count = 0
             if [key[0]] == k:
                 if self.keys == []:
                    self.keys.append(str(score - 1))
                    # switch = 1
                 else:
                    for i in self.keys:
                        if i == str(score-1):
                            repeat_count += 1
                    if repeat_count == 0:
                        print('Ключ в строке: ' + str(score))
                        self.keys.append(str(score-1))
        if self.keys:
            # self.keys = key +self.keys
            print('Ключ: ')
            print(self.keys)

    def compare_keys(self, cellColumn):  # функция, запрашивающая ключи по шифру из Excel в Google Sheets
        print('Функция прохода по ключевому столбцу:')
        for i in cellColumn:
            print('проверяю значение ' + str(i))
            self.find_key(i)

    def get_data_column(self):  # функция, извлекающая из google sheets информацию по ключу, найденному в findKey
        print('попытка извлечения данных')
        name_of_column = []
        score = 0
        for i in self.gData:
            score += 1
            name_of_column.append(i)

        for j in self.keys:
            score = 0
            data = []
            for i in name_of_column:
                score += 1
                if score <= len(name_of_column):
                    if self.gData[i][0][int(j)] == []:
                        data.append([''])
                    else:
                        data.append(self.gData[i][0][int(j)])
            self.importData.append(data)

        for i in self.importData:
            print(i)
        return self.importData

    def is_google_key_repeat(self,repeat_list,key_string):
        if repeat_list == []:
            repeat_list.append(key_string)
            return False
        else:
            for i in repeat_list:
                if key_string == i:
                    return True
                else:
                    repeat_list.append(key_string)
                    return False
