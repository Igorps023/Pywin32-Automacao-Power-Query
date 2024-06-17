# %%
import win32com.client
import time
import os
import pathlib 
from pathlib import Path
# %%
home = Path('~').expanduser()
project_folder = home / 'Desktop/Python_Proj_PQ/'
# %%
xl = win32com.client.DispatchEx("Excel.Application")
fileName = str(project_folder / 'consolidador/Consolidado.xlsx')
wb = xl.workbooks.open(fileName)
xl.Visible = False
wb.RefreshAll()
wb.Save()
wb.Close(True)
xl.Quit()
# os.system("taskkill /f /im excel.exe")
print('Done')