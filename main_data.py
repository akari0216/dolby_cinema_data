import os
import re
import datetime
import numpy as np
import pandas as pd
from ftplib import FTP
from time import sleep
from logger import logger
from mail_setting import send_mail

#记得修改路径
path = "C:\\Users\\xieminchao\\Desktop\\dolby"
os.chdir(path)
df_total = pd.DataFrame(columns = ["影城编码","影城名称","影片名称","影厅编码","影厅名称","影厅座位数","放映日期","总场次","总人数","总票房","平均票价","上座率"])
cinema_info = pd.DataFrame(columns = ["影城编码","影城名称","影厅名称"],data = [["44010601","金逸影城深圳中心城店","杜比厅-new"],["42170801","金逸影城武汉王家湾店","2号厅布局-1013"]])

today = datetime.date.today()
start_date = str(today - datetime.timedelta(days = 7))
end_date = str(today - datetime.timedelta(days = 1))

result_file = "杜比影院票房%s-%s.xlsx" % (start_date[5:].replace("-",""),end_date[5:].replace("-",""))

#获取实时报表数据
def ftp_run(date_list):
    ftp = FTP()
    ftp.connect(host = "192.168.9.94",port = 21 ,timeout = 30)
    ftp.login(user = "sjzx",passwd = "jy123456@")
    list = ftp.nlst()
    for each_date in date_list:
        each_date = each_date.replace("-","")
        filename = "SessionRevenue_"+each_date+".csv"
        for each_file in list:
            judge = re.match(filename,each_file)
            if judge:
                file_handle = open(filename,"wb+")
                ftp.retrbinary("RETR "+filename,file_handle.write)
                file_handle.close()
                print("%s file download success" % filename)
                logger.info("%s file download success" % filename)
                sleep(1)
    ftp.quit()

df_total = pd.DataFrame()    
date_list =[str(x)[0:10] for x in pd.date_range(start = start_date,end = end_date,freq = "D")]
ftp_run(date_list)
listdir = os.listdir()
#对每日csv进行数据清洗
for each_date in date_list:
    filename = "SessionRevenue_"+ each_date.replace("-","") +".csv"
    if filename in listdir:
        each_df = pd.read_csv(filename,encoding = "utf-8")
        showcount_time = each_df["场次时间"]
        time_list = []
        t1 = datetime.datetime.strptime(each_date + " 06:00:00","%Y-%m-%d %H:%M:%S")
        t2 = datetime.datetime.strptime(str(datetime.date(int(each_date[0:4]),int(each_date[5:7]),int(each_date[8:10])) + datetime.timedelta(days =1)) + " 05:59:59","%Y-%m-%d %H:%M:%S")
        for each_time in showcount_time:
            tmp_time = datetime.datetime.strptime(each_time,"%Y-%m-%d %H:%M:%S")
            delta1 = tmp_time - t1
            delta2 = t2 - tmp_time
            if delta1.days == 0 and delta2.days == 0:
                time_list.append(each_time)
        each_df = each_df[each_df["场次时间"].isin(time_list)]
        film = np.array(each_df["影片"])
        pat = "（数字）|（数字3D）|（数字IMAX）|（数字IMAX3D）|（中国巨幕）|（中国巨幕立体）|（IMAX3D）|（IMAX 3D）|（IMAX）|\s*"
        for i in range(len(film)):
            film[i] = re.sub(pat,"",film[i])
        each_df["影片"] = film
        film_time = np.array(each_df["场次时间"])
        pat_time = "\d+-\d+-\d+"
        for i in range(len(film_time)):
            film_time[i] = re.findall(pat_time,film_time[i])[0]
        each_df["场次时间"] = film_time
        each_df = each_df[each_df["场次状态"].isin(["开启"])]
        df_total = pd.concat([df_total,each_df],ignore_index = True)
        print("%s data process complete" % each_date)
        logger.info("%s data process complete" % each_date)
        os.remove(filename)
        
df_total2 = pd.DataFrame(columns = df_total.columns)
df_total2["场次"] = df_total["场次时间"]
#按固定影院影厅进行筛选
for each_index in cinema_info.index:
    cinema = cinema_info["影城名称"][each_index]
    hall = cinema_info["影厅名称"][each_index]
    each_df = df_total[df_total["影院"].isin([cinema]) & df_total["影厅"].isin([hall])]
    df_total2 = pd.concat([df_total2,each_df],ignore_index = True)

#做透视表
df_table = pd.pivot_table(df_total2,index = ["影院","影片","影厅","场次时间"],values = ["场次","人数","票房","总座位数"],aggfunc = {"场次":len,"人数":np.sum,"票房":np.sum,"总座位数":np.sum},fill_value = 0,margins = False)
df_table.reset_index(inplace = True)
df_table.sort_values(by = ["影院","场次时间"],ascending = [True,True],inplace = True)
df_table["影厅座位数"] = np.divide(df_table["总座位数"],df_table["场次"]).astype(int)
df_table["平均票价"] = np.round(np.divide(df_table["票房"],df_table["人数"]),2)
df_table.fillna(0,inplace = True)
df_table["上座率"] = np.round(np.divide(df_table["人数"],df_table["总座位数"]) *100,2).astype(str) + np.full(len(df_table),"%")
df_table["影厅名称2"] = np.full(len(df_table),"杜比影院")
df_table = pd.merge(left = cinema_info,right = df_table,left_on = "影城名称",right_on = "影院",how = "right")
df_table.drop(columns = ["总座位数","影院","影厅"],axis = 1,inplace = True)
df_table["影片编码"] = np.full(len(df_table),"")
df_table = df_table.reindex(columns = ["影城编码","影城名称","影片编码","影片","影厅名称","影厅名称2","影厅座位数","场次时间","场次","人数","票房","平均票价","上座率"])
df_table.rename(columns = {"影片":"影片名称","影厅名称":"影厅编码","影厅名称2":"影厅名称","场次时间":"放映日期","场次":"总场次(场)","人数":"总人数(人)","票房":"总票房(元)"},inplace = True)
print(df_table)
logger.info("data calculate completed")
df_table.to_excel(result_file,header = True,index = False)

#发送邮件
send_mail(result_file)