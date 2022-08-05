
import os
import datetime
import ctypes
from openpyxl import load_workbook
import openpyxl
import xlwings as xw
from PyQt5.QtCore import QObject, pyqtSignal
from pathlib import Path
import winreg as wreg
import pythoncom


class Misc(QObject):
    maximum = pyqtSignal(int)  # signals to communicate worker thread with main thread
    progress = pyqtSignal(int)
    finished = pyqtSignal()

    def RemSpcCov(self):
        today = datetime.date.today()  # Initialize date as last sat (for Ovt Hours)
        idx = (today.weekday() + 1) % 7
        sat = today - datetime.timedelta(idx + 1)
        last_sat = datetime.datetime.strptime(str(sat), '%Y-%m-%d').strftime('%m %d %y')
        cur_year = sat.year

        data_folder = Path(f"C:/Users/Austin.Kidwell/Desktop/WE {last_sat}/")
        #data_folder = Path(f"//CPROME/Eng_Share/System Engineering/Visual Management/Timesheets/"
        #                   f"{cur_year} COV Timesheets/WE {last_sat}/")
        target_files = []
        for root, dirs, files in os.walk(data_folder):  # Obtain files to retrieve data from
            for file in files:
                if last_sat in file:
                    if "SUMMARY" not in file and 'summary' not in file:
                        ftarget = os.path.join(root, file)
                        target_files.append(ftarget)
        #print(len(target_files))

        if target_files == []:
            self.finished.emit()
            ctypes.windll.user32.MessageBoxW(0, f"Couldn't locate file {data_folder}", "Info", 0)
            return

        row = [6, 7, 8, 9, 10, 11, 14, 15, 16, 17, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
               37, 38, 39, 40]
        pythoncom.CoInitialize()
        excel_app = xw.App(visible=False)
        excel_app.api.EnableEvents = False
        excel_app.screen_updating = False
        #excel_app.calculation = 'manual'
        #excel_app.display_alerts = False
        count = 0
        self.maximum.emit(len(target_files))
        for file in target_files:
            wb = xw.Book(file)
            for sheet in wb.sheets:
                sheet = sheet.name
                if sheet != "Setup":               # Look through all but setup tab
                    ws = wb.sheets[sheet]
                    for i in range(3, 10):  # loop trough C-I hours
                        for j in row:
                            temp = ws.cells(j, i).value
                            if temp is not None:
                                temp = " ".join(str(temp).split())
                                try:
                                    ws.cells(j, i).value = float(temp)
                                except Exception:
                                    ws.cells(j, i).value = None
            try:
                wb.save(file)
                wb.close()
            except Exception:
                excel_app.kill()
                self.progress.emit(0)
                self.finished.emit()
                pythoncom.CoUninitialize()
                ctypes.windll.user32.MessageBoxW(0, f"Close excel file {data_folder} to allow editing", "Info", 0)
                return
            count += 1
            self.progress.emit(count)
        excel_app.kill()
        self.progress.emit(0)
        self.finished.emit()
        pythoncom.CoUninitialize()
        ctypes.windll.user32.MessageBoxW(0, f"Spaces in {data_folder} are removed", "Info", 0)

    def RemSpcDav(self):
        today = datetime.date.today()  # Initialize date as last sat (for Ovt Hours)
        idx = (today.weekday() + 1) % 7
        sat = today - datetime.timedelta(idx + 1)
        last_sat = datetime.datetime.strptime(str(sat), '%Y-%m-%d').strftime('%m %d %y')
        cur_year = sat.year

        data_folder = f"C:/Users/Austin.Kidwell/Desktop/Over Time Hours/WE {last_sat}/"
        #data_folder = Path(f"//CPROME/Eng_Share/System Engineering/Visual Management/Timesheets/"
        #                   f"{cur_year} DAV Timesheets/WE {last_sat}/")
        target_files = []
        for root, dirs, files in os.walk(data_folder):  # Obtain files to retrieve data from
            for file in files:
                if last_sat in file:
                    if "SUMMARY" not in file and 'summary' not in file:
                        ftarget = os.path.join(root, file)
                        target_files.append(ftarget)
        #print(len(target_files))

        if target_files == []:
            self.finished.emit()
            ctypes.windll.user32.MessageBoxW(0, f"Couldn't locate file {data_folder}", "Info", 0)
            return

        row = [6, 7, 8, 9, 10, 11, 14, 15, 16, 17, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
               37, 38, 39, 40]
        pythoncom.CoInitialize()
        excel_app = xw.App(visible=False)
        excel_app.api.EnableEvents = False
        excel_app.screen_updating = False
        #excel_app.calculation = 'manual'
        #excel_app.display_alerts = False
        count = 0
        self.maximum.emit(len(target_files))
        for file in target_files:
            wb = xw.Book(file)
            for sheet in wb.sheets:
                sheet = sheet.name
                if sheet != "Setup":  # Look through all but setup tab
                    ws = wb.sheets[sheet]
                    for i in range(3, 10):  # loop trough C-I hours
                        for j in row:
                            temp = ws.cells(j, i).value
                            if temp is not None:
                                temp = " ".join(str(temp).split())
                                try:
                                    ws.cells(j, i).value = float(temp)
                                except Exception:
                                    ws.cells(j, i).value = None
            try:
                wb.save(file)
                wb.close()
            except Exception:
                excel_app.kill()
                self.progress.emit(0)
                self.finished.emit()
                pythoncom.CoUninitialize()
                ctypes.windll.user32.MessageBoxW(0, f"Close excel file {data_folder} to allow editing", "Info", 0)
                return
            count += 1
            self.progress.emit(count)
        excel_app.kill()
        self.progress.emit(0)
        self.finished.emit()
        pythoncom.CoUninitialize()
        ctypes.windll.user32.MessageBoxW(0, f"Spaces in {data_folder} are removed", "Info", 0)

    def WinRegIssue(self):
        path = wreg.HKEY_LOCAL_MACHINE

        def save_reg(k='ActivationFilterOverride', v=1):        # save
            try:
                key = wreg.OpenKey(path, r"SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\"
                                           r"Wow6432Node\\Microsoft\\Office\\16.0\\Common\\COM Compatibility\\"
                                           r"{B54F3741-5B07-11cf-A4B0-00AA004A55E8}\\", 0, wreg.KEY_SET_VALUE | wreg.KEY_WOW64_64KEY)
                wreg.SetValueEx(key, k, 0, wreg.REG_DWORD, v)
                ctypes.windll.user32.MessageBoxW(0, "Registry Value Added", "Info", 0)
                return True
            except Exception as e:
                try:
                    key = wreg.OpenKey(path, r"SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\"
                                             r"Wow6432Node\\Microsoft\\Office\\16.0\\Common\\", 0,
                                       wreg.KEY_SET_VALUE | wreg.KEY_WOW64_64KEY)
                    newKey = wreg.CreateKey(key, r"COM Compatibility\\{B54F3741-5B07-11cf-A4B0-00AA004A55E8}")
                    wreg.SetValueEx(newKey, k, 0, wreg.REG_DWORD, v)
                    if newKey:
                        wreg.CloseKey(newKey)
                    ctypes.windll.user32.MessageBoxW(0, "Registry Value Added", "Info", 0)
                    return True
                except Exception as e:
                    ctypes.windll.user32.MessageBoxW(0, e, "Info", 0)
                return False

        save_reg()

    def CloseAllXl(self):
        result = os.system('TASKKILL /F /IM excel.exe')
        if result == 128:
            msg = "No Excel Instance Existed"
        else:
            msg = "All Excel Instances Closed"
        ctypes.windll.user32.MessageBoxW(0, msg, "Info", 0)