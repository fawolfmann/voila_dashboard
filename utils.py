## Read staff excel, and return list of employes with code
import pandas as pd
import os

def staff_code():
    path = os.path.join('O:','Libovich','F','Estudio','Time Sheets','Staff.xls')
    staff_df = pd.read_excel(path)
    staff_df.columns = staff_df.iloc[5]
    staff_df = staff_df[6:]
    active_staff = staff_df[staff_df['ACTIVO'].notnull()].loc[:, ['CODE','NAME']]
    staff_codes = []
    for index, staff in active_staff.iterrows():
        staff_code = '{:02d}'.format(staff.CODE) , staff.NAME
        staff_codes.append(staff_code)
    return staff_codes