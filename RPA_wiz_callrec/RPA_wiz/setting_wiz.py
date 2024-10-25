import os
from pathlib import Path
from datetime import datetime, timedelta

class Login:
    def __init__(self) -> None:
        pass
    
    def login(self):
        self.user = 'nichamot'
        self.password = 'Nummon@082024'
    
    def webdriver(self):
        self.wiz = "https://tha-crm.wiz.ai/#/login"
class find_file:
    def __init__(self,path):
        self.path = path
    def file_last_time(self):
        path_file = self.path
        file_excelfile = [file for file in os.listdir(path_file) if file.endswith('.xlsx')]
        if file_excelfile:
            file_ex_time = max(
                (os.path.join(path_file, file) for file in file_excelfile),
                key = os.path.getmtime
            )
            return file_ex_time

url = 'https://tha-crm.wiz.ai/#/login'
us = 'nichamot'
ps = 'Nummon@082024'

real_path = r"\\172.16.103.200\css_ctia_analy\Operation\05 - (RPA) Outbound Campaign (IM) by Last update Daily\Report Outbound Campaign (Accumulate)\Daily Report\ACARE Payment Remider\Report 2024\Raw Data\WIZ"
dir_path = os.path.dirname(os.path.abspath(__file__)) + os.sep
Data_path = dir_path + 'DATA DMR' + os.sep
Fill_rec_path = dir_path + 'DATA FILL REC' + os.sep
BC_path = dir_path + 'BC' + os.sep

Path(Data_path).mkdir(parents=True, exist_ok= True)
Path(Fill_rec_path).mkdir(parents=True, exist_ok= True)
Path(BC_path).mkdir(parents=True, exist_ok= True)


date = datetime.now() - timedelta(days=0)   # วันไฟล์ตรงนี้


current_date = date.strftime("%d-%m-%Y")    # เปลี่ยนวันที่ให้เหมาะสมกับชื่อโฟลเดอร์
date_folder_path = os.path.join(real_path, current_date)


Path(date_folder_path).mkdir(parents=True, exist_ok=True)
print(f"Create folder: {date_folder_path}")


folder_main1 = "Data Last Record"
folder_main2 = "BC"


sub1_folder = os.path.join(date_folder_path, folder_main1)
Path(sub1_folder).mkdir(parents=True, exist_ok=True)
print(f"Create folder: {sub1_folder}")


sub2_folder = os.path.join(date_folder_path, folder_main2)
Path(sub2_folder).mkdir(parents=True, exist_ok=True)
print(f"Create folder: {sub2_folder}")