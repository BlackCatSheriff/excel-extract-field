from openpyxl import Workbook
import time
import random


def save_file(save_path, data):
    fields=['id', 'name', 'age', 'sex']
    # 增量添加,如果文件存在就打开追加
    wb = Workbook()
    sheet = wb.active
    sheet.append(fields)
    for d in data:
        sheet.append(d)
    wb.save(save_path)
    wb.close()


if __name__=="__main__":

    for i in range(100):
        data = list()
        for j in range(2000):
            tmp = list()
            tmp.append("000%d" % random.randint(1, 9998595695))
            tmp.append(time.time()*100000)
            tmp.append(random.randint(20, 90))
            tmp.append(random.choice(['男性', '女性', '中性']))
            data.append(tmp)
        save_file(r"C:\Users\www\Desktop\ttt\%s.xlsx" % int(round(time.time()*1000)), data)
    print("DONE")