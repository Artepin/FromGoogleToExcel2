import configparser
import re
import os
from sys import argv

class ReadIni:
    def __init__(self):
        self.path_ini = ''
        self.path_ini = argv
        self.config = configparser.ConfigParser()
        #self.config.read('test_ini.ini')
        #self.config.ad('params.ini')
        #path = "\\Fserv\e\CLIENTS\Info\p2707-1_Сатурндинамика_26стендКоршуновРСС\3. Документы\6. Закупки\config_2707.ini"
        #os.path.currentdir = "\\\\Fserv\\e\\CLIENTS\\Info\\p2707-1_Сатурндинамика_26стендКоршуновРСС\\3. Документы\\6. Закупки\\"
        '''if os.path.exists("config_2707.ini") == False:
            print('false')
            self.config.read('config.ini')
        else:'''
        path = self.parse_path()
        self.config.read(path)
        print('пытаюсь прочитать файл конфигурации: '+ path)
        print()

    def parse_path(self):
        path = str(self.path_ini[1]).split("cfg:")
        path = path[1]
        #path = path.split('"')
        #path = path[1]
        return path

    def get_excel_key_id(self):
        print(parse_param(self.config['Excel']['key_id']))
        return parse_param(self.config['Excel']['key_id'])

    def get_excel_data_id(self):
        id = parse_param(self.config['Excel']['data_id']).split(', ')
        print(id)
        return id

    def get_excel_param(self, name_of_param):
        return parse_param(self.config['Excel'][name_of_param])

    def get_google_key_id(self):
        return parse_param(self.config['Sheets']['key_id'])

    def get_google_data_id(self):
        id = parse_param(self.config['Sheets']['data_id']).split(', ')
        print(id)
        return id

    def get_google_param(self, name_of_param):
        return parse_param(self.config['Sheets'][name_of_param])

    def get_url(self):
        ssid = parse_param(self.config['Sheets']['url'])
        ssid = ssid.split('d/')
        ssid = ssid[1]
        print(ssid)
        return ssid


def parse_param(param):
    param = param.split('"')
    param = param[1]
    return param
