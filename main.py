import excelLib
import sheetsApi
import read_ini
read_ini.ReadIni()
data = excelLib.ExcelData()
wb = data.get_wb()
key_data = data.get_key_data2()
cell_column = data.get_key()
ssObj = sheetsApi.Spreadsheets()
coords = ssObj.get_coord_of_nr()
counm = ssObj.get_column()
resultColumn = ssObj.compare_keys(cell_column)
gData = ssObj.get_data_column()
data.cycle4(gData)