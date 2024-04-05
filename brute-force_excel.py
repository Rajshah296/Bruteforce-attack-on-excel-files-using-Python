import win32com.client
import time
import Progress_bar


password_file=r'C:\Users\RAJ\OneDrive\Documents\xato-net-10-million-passwords.txt'
excel_file=r'C:\Users\RAJ\OneDrive\Documents\list.xlsx'

excel=win32com.client.Dispatch("Excel.Application")

password_list=[]
#extract passwords from file and load to list object.
with open(password_file,'r',encoding='utf-8') as pwd:
    passwords= pwd.readlines()
    for password in passwords:
        password_list.append(password.replace('\n',''))

Progress_bar.progress_bar(0,len(password_list)) #initializing progress bar with 0 and total length of the task
for i,password in enumerate(password_list):
    try:
        wb= excel.Workbooks.Open(excel_file, False, True, None, password)
        wb.unprotect(password)
        print(f"Successfull password: {password}")
        excel.DisplayAlerts = False
        excel.Quit()
        time.sleep(1)
        quit()
    except:
        Progress_bar.progress_bar(i+1,len(password_list)) #show progress bar
        continue