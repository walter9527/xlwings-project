import xlwings as xw
import os
import time

wb = xw.Book()

sheet = wb.sheets[0]

sheet.range("A1").value = "线程数"
sheet.range("B1").value = "时间"
sheet.range("C1").value = "每秒平均执行次数"

nthread = list(range(1, 65))
run_time = []
avg_run_cnt = []

for n in nthread:
    avg = []
    for _ in range(3):
        start = time.time()
        os.system(f"a.exe {n}")
        end = time.time()
        avg.append(end - start)
    
    run_time.append(round(sum(avg)/len(avg), 3))
    avg_run_cnt.append(int(1e7//n*n//run_time[len(run_time) - 1]))


sheet.range("A2").options(transpose=True).value = nthread
sheet.range("B2").options(transpose=True).value = run_time
sheet.range("C2").options(transpose=True).value = avg_run_cnt

wb.save("1.xlsx")
wb.close()
