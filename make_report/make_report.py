import win32com.client as win32
import win32gui
import win32con
import os

hwp=win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
current_dir = os.getcwd()
current_file = current_dir + '\\daily_report.hwp'

hwp.Open(current_file)

field_list = [ i for i in hwp.GetFieldList().split('\x02')]
field_replace = {"working_day":"2020-07-04","check_server":"서버점검: 점검중","check_mail":"메일 확인"}

for field in field_list :
    hwp.PutFieldText(f'{field}{{{{0}}}}',field_replace[field])

hwp.SaveAs(current_dir + '\\daily_report_new.hwp')
hwp.Quit()

