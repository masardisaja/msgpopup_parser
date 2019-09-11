# WinPopUp IP Messenger Log Parser
# Created by masardi
# release date : 2019-08-26

import re, os, xlsxwriter, calendar, time
ts = calendar.timegm(time.gmtime())
outpath = "C:\IPMsgParser"

if not os.path.exists(outpath):
    os.makedirs(outpath)

print("##########################################")
print("##### MsgPopUp Log Parser by masardi #####")
print("##########################################")
print("")
print("Masukkan lokasi direktori log yang akan dianalisis...")
path = input("Log directory path : ")
print("")
print("Masukkan nama file output yang diinginkan...")
restname = input("Output filename : ")

log = open(outpath + "/error.log","a+")
separator = "==============================================================\n"
thead = "SOURCE\tSTAT\tTIME\tFROM\tIP ADDR\tTO\tMESSAGE\tMODE\tATTACHMENT"

files = []
datas = []

for r, d, f in os.walk(path):
    for file in f:
        if ".txt" in file:
            files.append(file)

i = int(0)
for f in files:
    file = open(os.path.join(path,f),"r",encoding="cp850")
    contents = file.read()
    segmen = contents.split(separator)
    file.close()

    fname = f[0:8]

    if i == 0:
        datas.append(thead)
    i += 1
    
    n = int(0)
    for rows in segmen:
        row = rows.split("\n")
        ln = len(row)
        
        stat = ""
        frm = ""
        time = ""
        to = ""
        msg = ""
        err = ""
        grp = ""
        att = ""
        
        for line in row:
            n = n + 1
            if n == 1:
                if line == "* Receive":
                    stat = "Recv"
                elif line == "* Send":
                    stat = "Send"
                else:
                    stat = "Err"
                    msg = msg + "\n" + line
            elif stat != "":
                if n == 2:
                    # Baris ke-2 adalah data pengirim #
                    
                    # Validasi format IP Pengirim
                    ip = re.compile('(([2][5][0-5]\.)|([2][0-4][0-9]\.)|([0-1]?[0-9]?[0-9]\.)){3}'
                                    +'(([2][5][0-5])|([2][0-4][0-9])|([0-1]?[0-9]?[0-9]))')
                    match = ip.search(line)
                    if match:
                        sip = match.group()
                    else:
                        sip = ""

                    # Validasi format nama Pengirim
                    
                    if line.find(':')>1 and line.find(',') < 0:
                        frm = line.replace("[" + sip + "]","")
                        err = "X"
                    else:
                        frm = sip
                        msg = msg + "\n" + line
                    
                elif n == 3:
                    # Baris ke-tiga adalah timestamp
                    
                    # Validasi format timestamp
                    chk = re.compile("^\([0-9]{2}\/[0-9]{2}\([A-z]{3}\)$")
                    if chk.match(line[0:11]):
                        dy = line[7:10]
                        dd = line[4:6]
                        mm = line[1:3]
                        hh = line[13:15]
                        mi = line[16:18]
                        am = line[11:13]
                        time = dd + "-" + mm + " " + hh + ":" + mi + " " + am + " (" + dy + ")"
                        time_x = time
                    else:
                        time = time_x
                        msg = msg + "\n" + line

                elif n == 4:
                    # Baris ke-4 adalah nama Penerima

                    # Validasi format penerima
                    if line[0:13] == "Received with":
                        to = line.replace("Received with ","")
                        to_x = to
                    else:
                        to = to_x
                        msg = msg + "\n" + line

                    # Cek apakah pesan pribadi (japri) atau group
                    if to.count(",") > 1:
                        grp = "group"
                    else:
                        grp = "japri"
                    
                else:
                    msg = msg + "\n" + line
            else:
                msg = msg + "\n" + line

            if line[0:12] == "[Attachment]":
                att = line[13:]
            
            if n == ln:
                n = 0
                raw = fname + "\t" + stat + "\t" + time + "\t" + frm + "\t" + sip + "\t" + to + "\t" + msg.replace("\t"," ") + "\t" + grp + "\t" + att
                if stat != "Err":
                    datas.append(raw)

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook(outpath + '/' + restname + '_' + str(ts) + '.xlsx')
worksheet = workbook.add_worksheet('Conversation')
worksheet2 = workbook.add_worksheet('Users')

row = 0
col = 0

for data in datas:
    dt = data.split('\t')
    if len(dt)==9:
        worksheet.write(row, col,     dt[0])
        worksheet.write(row, col + 1, dt[1])
        worksheet.write(row, col + 2, dt[2])
        worksheet.write(row, col + 3, dt[3])
        worksheet.write(row, col + 4, dt[4])
        worksheet.write(row, col + 5, dt[5])
        worksheet.write(row, col + 6, dt[6])
        worksheet.write(row, col + 7, dt[7])
        worksheet.write(row, col + 8, dt[8])
        row += 1        
    else:
        log.write(dt)
        
workbook.close()
log.close()

print(str(i) + ' file(s) telah diproses.')
print('File output telah disimpan di ' + outpath + '\\' + restname + '_' + str(ts) + '.xlsx')

input('Press Any Key To Exit.')
