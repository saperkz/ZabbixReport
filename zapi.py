# ***************************
# App: GetVDC
# Description: Get report data from Vcloud API to excel
# Date: 25.05.2020
# Author: Olzhas Omarov
# Python version: 3.8
# Compatibility: Vcloud 9.5 and above (for older versions please contact me by email)
# Email: ooskenesary@gmail.com
# ***************************

import time
import os
import datetime
from pyzabbix.api import ZabbixAPI
import csv
import pandas
import openpyxl
import tkinter as tk
from tkinter import messagebox


def onError():
    tk.messagebox.showerror("Error", "Please input require data")


def onEnd():
    tk.messagebox.showinfo("OK", "Report is ready, please check your current folder")


def getmaindata():
    if len(edt_url.get()) == 0:
        onError()
    elif len(edt_login.get()) == 0:
        onError()
    elif len(edt_pass.get()) == 0:
        onError()
    else:
        btn_result.config(text="Get report again!")

        # ======Get server list from file servers.txt==========================
        filename = "servers.txt"
        with open(filename) as f:
            lines = f.read().splitlines()
        # ======Get server list from file servers.txt==========================

        # =======TXT result file=================
        fulldate = datetime.datetime.now()
        month_today = fulldate.strftime("%B")
        datestring = datetime.datetime.strftime(datetime.datetime.now(), '%Y-%m-%d_%H-%M-%S')
        resultfile = open('Zabbix_result_' + datestring + '.csv', 'w')
        resultfile_name = resultfile.name

        writer = csv.DictWriter(
            resultfile, fieldnames=["hostname", "Memory avg,%", "Memory max, %", "CPU load avg, %", "CPU load max, %",
                                    "HDD usage %"], delimiter=";")
        writer.writeheader()
        # =======end of TXT result file=================
        url = edt_url.get()
        login = edt_login.get()
        passwd = edt_pass.get()
        
        zapi = ZabbixAPI(url, user=login, password=passwd)
        
        servers = lines
        start_time_human = edt_start_date.get()
        end_time_human = edt_end_date.get()
        
        
        
        
        start_time = round(time.mktime(datetime.datetime.strptime(start_time_human, "%d/%m/%Y").timetuple()))
        end_time = round(time.mktime(datetime.datetime.strptime(end_time_human, "%d/%m/%Y").timetuple()))
        
        

        for hi in zapi.host.get(filter={"host": servers}, output="extend"):
            hid = hi['hostid']
            # print(hi['host'])
            count = 0
            memtotalprc = 0
            cpuloadvaluetotal = 0
            maxmem = 0
            maxcpu = 0
            for ii in zapi.item.get(filter={"hostid": hid}, search={"key_": "vm.memory.size[available]"},
                                    output="extend"):
                zitemid = ii['itemid']

                for hisid in zapi.trend.get(time_from=start_time, time_till=end_time, itemids=zitemid,
                                            output="extend"):

                    zz = zapi.item.get(filter={"hostid": hid}, search={"key_": "vm.memory.size[total]"},
                                       output="extend")
                    #print(zz)
                    memavailsize = int(hisid['value_avg'])
                    memtotalsize = int(zz[0]['lastvalue'])
                    memtotalsizeGB = memtotalsize / 1024 / 1024 / 1024
                    memresult = round(((memtotalsize - memavailsize) / 1024 / 1024 / 1024), 2)
                    memresultprc = round((memresult * 100 / memtotalsizeGB), 2)
                    count = count + 1
                    memtotalprc = memtotalprc + memresultprc
                    zzmemmax = 100 - (((int(hisid['value_min'])) / 1024 / 1024 / 1024) * 100 / memtotalsizeGB)
                    if maxmem <= zzmemmax:
                        maxmem = zzmemmax

            

            for icpu in zapi.item.get(filter={"hostid": hid}, search={"key_": "system.cpu.util[,idle]"},
                                      output="extend"):
                cpuitemid = icpu['itemid']
                for hicpusid in zapi.trend.get(time_from=start_time, time_till=end_time, itemids=cpuitemid,
                                               output="extend"):
                    # print(hicpusid['value'])
                    zzcpumax = hicpusid['value_max']
                    cpuidlevalue = hicpusid['value_avg']
                    cpuloadvalue = round(100 - (float(cpuidlevalue)), 2)
                    cpuloadvaluetotal = cpuloadvaluetotal + cpuloadvalue
                    # print(cpuloadvalue)
                    if maxcpu <= float(zzcpumax):
                        maxcpu = float(zzcpumax)

            for ifs in zapi.item.get(filter={"hostid": hid}, search={"key_": "last-space"}, output="extend"):
                fsvalue = ifs['lastvalue']
                # print(fsvalue)

            print(hi['host'], str(round((memtotalprc / count), 2)), maxmem, (cpuloadvaluetotal / count), maxcpu,
                  fsvalue, sep=';', file=resultfile)

        resultfile.close()
        # =======Excel result file=================
        df = pandas.read_csv(resultfile_name, sep=';')
        df.to_excel('Zabbix_result_' + datestring + '.xlsx', 'Sheet1')
        # =======end of Excel result file=================
        os.remove(resultfile_name)
        # ===========================================
        onEnd()


# ======================GUI========================================
rootwin = tk.Tk()
rootwin.title("Получение отчета с Zabbix")
rootwin.geometry("820x180")

lbl_url = tk.Label(rootwin, text="API URL:", font=("Arial Bold", 12))
lbl_url.grid(column=0, row=0)

edt_url = tk.Entry(rootwin, width=60)
edt_url.insert(0, "http://kaptmon.kazatomprom.kz/api_jsonrpc.php")
edt_url.grid(column=1, row=0)

lbl_login_url = tk.Label(rootwin, text=" example:  http://kaptmon.kazatomprom.kz/api_jsonrpc.php",
                         font=("Arial Bold", 10))
lbl_login_url.grid(column=2, row=0)

lbl_login = tk.Label(rootwin, text="Login:", font=("Arial Bold", 12))
lbl_login.grid(column=0, row=1)

edt_login = tk.Entry(rootwin, width=60)
edt_login.grid(column=1, row=1)

lbl_login_com = tk.Label(rootwin, font=("Arial Bold", 9))
lbl_login_com.grid(column=2, row=1)

lbl_pass = tk.Label(rootwin, text="Password:", font=("Arial Bold", 12))
lbl_pass.grid(column=0, row=2)

edt_pass = tk.Entry(rootwin, show='X', width=60)
edt_pass.grid(column=1, row=2)
login_str = edt_login.get()

# ====start date==========
lbl_start_date = tk.Label(rootwin, text="Start date:", font=("Arial Bold", 12))
lbl_start_date.grid(column=0, row=3)

edt_start_date = tk.Entry(rootwin, width=60)
edt_start_date.insert(0, "18/02/2021")
edt_start_date.grid(column=1, row=3)

lbl_login_start_date = tk.Label(rootwin, text=" format: dd/mm/yyyy, example:  21/02/2021", font=("Arial Bold", 10))
lbl_login_start_date.grid(column=2, row=3)

# =====end date============
lbl_end_date = tk.Label(rootwin, text="End date:", font=("Arial Bold", 12))
lbl_end_date.grid(column=0, row=4)

edt_end_date = tk.Entry(rootwin, width=60)
edt_end_date.insert(0, "19/02/2021")
edt_end_date.grid(column=1, row=4)

lbl_login_end_date = tk.Label(rootwin, text=" format: dd/mm/yyyy, example:  22/02/2021", font=("Arial Bold", 10))
lbl_login_end_date.grid(column=2, row=4)
# ===========================

btn_result = tk.Button(rootwin, text="Get Report!", width=30, font=("Arial Bold", 13), command=getmaindata)
btn_result.grid(column=1, row=5)

rootwin.mainloop()
