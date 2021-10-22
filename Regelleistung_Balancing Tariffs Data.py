from bs4 import BeautifulSoup
import requests
from datetime import date,timedelta
import io

def Others(fdate,tdate,tname,data_type,lis):
    #print(tname,data_type)
    r_date, t_from, t_to, b_neg, b_pos, q_neg, q_pos = lis
    fh.write(fdate+"\t"+tdate+"\t"+tname+"\t"+data_type+"\t"+r_date+"\t"+t_from+"\t"+t_to+"\t"+btr+"\t"+b_neg+"\t"+b_pos+"\t"+b_imp+"\t"+b_exp+"\t"+
            ql+"\t"+q_neg+"\t"+q_pos+"\t"+q_imp+"\t"+q_exp+"\t"+q_eu+"\t"+q_impeu+"\t"+q_expeu+"\t"+nrv+"\t"+mol+"\t"+exp+"\n")

def Rzsaldo(fdate,tdate,tname,data_type,lis):
    r_date = t_from = t_to = b_neg = b_pos = q_neg = q_pos = btr = ql = nrv = mol = exp = b_imp = b_exp = q_imp = q_exp = q_eu = q_impeu = q_expeu ="NA"
    if tname == "Netzregelverbund":
        #print(tname,data_type)
        r_date, t_from, t_to, btr, q_neg, q_pos, ql, nrv, mol, exp = lis
        fh.write(fdate+"\t"+tdate+"\t"+tname+"\t"+data_type+"\t"+r_date+"\t"+t_from+"\t"+t_to+"\t"+btr+"\t"+b_neg+"\t"+b_pos+"\t"+b_imp+"\t"+b_exp+"\t"+
            ql+"\t"+q_neg+"\t"+q_pos+"\t"+q_imp+"\t"+q_exp+"\t"+q_eu+"\t"+q_impeu+"\t"+q_expeu+"\t"+nrv+"\t"+mol+"\t"+exp+"\n")
    if tname != "Netzregelverbund":
        #print(tname,data_type)
        r_date, t_from, t_to, btr, ql = lis
        fh.write(fdate+"\t"+tdate+"\t"+tname+"\t"+data_type+"\t"+r_date+"\t"+t_from+"\t"+t_to+"\t"+btr+"\t"+b_neg+"\t"+b_pos+"\t"+b_imp+"\t"+b_exp+"\t"+
            ql+"\t"+q_neg+"\t"+q_pos+"\t"+q_imp+"\t"+q_exp+"\t"+q_eu+"\t"+q_impeu+"\t"+q_expeu+"\t"+nrv+"\t"+mol+"\t"+exp+"\n")

def Rebap(fdate,tdate,tname,data_type,lis):
    #print(tname,data_type)
    r_date, t_from, t_to, q_eu = lis
    fh.write(fdate+"\t"+tdate+"\t"+tname+"\t"+data_type+"\t"+r_date+"\t"+t_from+"\t"+t_to+"\t"+btr+"\t"+b_neg+"\t"+b_pos+"\t"+b_imp+"\t"+b_exp+"\t"+
            ql+"\t"+q_neg+"\t"+q_pos+"\t"+q_imp+"\t"+q_exp+"\t"+q_eu+"\t"+q_impeu+"\t"+q_expeu+"\t"+nrv+"\t"+mol+"\t"+exp+"\n")
        
def Ea(fdate,tdate,tname,data_type,lis):
    #print(tname,data_type)
    r_date, t_from, t_to, q_neg, q_pos = lis
    fh.write(fdate+"\t"+tdate+"\t"+tname+"\t"+data_type+"\t"+r_date+"\t"+t_from+"\t"+t_to+"\t"+btr+"\t"+b_neg+"\t"+b_pos+"\t"+b_imp+"\t"+b_exp+"\t"+
            ql+"\t"+q_neg+"\t"+q_pos+"\t"+q_imp+"\t"+q_exp+"\t"+q_eu+"\t"+q_impeu+"\t"+q_expeu+"\t"+nrv+"\t"+mol+"\t"+exp+"\n")

def Prenet(fdate,tdate,tname,data_type,lis):
    r_date, t_from, t_to, q_impeu, q_expeu = lis
    fh.write(fdate+"\t"+tdate+"\t"+tname+"\t"+data_type+"\t"+r_date+"\t"+t_from+"\t"+t_to+"\t"+btr+"\t"+b_neg+"\t"+b_pos+"\t"+b_imp+"\t"+b_exp+"\t"+
            ql+"\t"+q_neg+"\t"+q_pos+"\t"+q_imp+"\t"+q_exp+"\t"+q_eu+"\t"+q_impeu+"\t"+q_expeu+"\t"+nrv+"\t"+mol+"\t"+exp+"\n")
def Others1(fdate,tdate,tname,data_type,lis):
    r_date, t_from, t_to, b_imp,b_exp,q_imp,q_exp = lis
    fh.write(fdate+"\t"+tdate+"\t"+tname+"\t"+data_type+"\t"+r_date+"\t"+t_from+"\t"+t_to+"\t"+btr+"\t"+b_neg+"\t"+b_pos+"\t"+b_imp+"\t"+b_exp+"\t"+
            ql+"\t"+q_neg+"\t"+q_pos+"\t"+q_imp+"\t"+q_exp+"\t"+q_eu+"\t"+q_impeu+"\t"+q_expeu+"\t"+nrv+"\t"+mol+"\t"+exp+"\n")

fh = io.open("Regelleistung_Balancing Tariffs Data(22.Oct.2021).xls",'w',encoding='utf-8')
fh.write("From Date\tTo Date\tTSO\tData Type\tDate\tTime From\tTime To\tbetr. [MW]\tbetr. NEG [MW]\tbetr. POS [MW]\tbetr. import [MW]\tbetr. export [MW]\tqual. [MW]\t"
                "qual. NEG [MW]\tqual. POS [MW]\tqual. import [MW]\tqual. export [MW]\tqual. [EUR/MWh]\tqual. import [EUR/MWh]\tqual. export [EUR/MWh]\t"
                "NRV balance more than 80"+'%'+ "of the contracted control reserve\tMOL DeviationM\tExplanation\n")
url = "https://www.regelleistung.net/ext/data/"
headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'Accept-Encoding': 'gzip, deflate, br',
    'Accept-Language': 'en-US,en;q=0.9',
    'Cache-Control': 'max-age=0',
    'Connection': 'keep-alive',
    'Content-Length': '49',
    'Content-Type': 'application/x-www-form-urlencoded',
    'Cookie': 'current-locale=en; BIGipServer~VRF02_External_Application_Access~pool_www-ext.regelleistung.net_18080=rd2o00000000000000000000ffff0aa6e36ao18080; BIGipServer~VRF02_External_Application_Access~pool_www-public.regelleistung.net_8080=rd2o00000000000000000000ffff0aa6e36ao8080; BIGipServer~VRF02_External_Application_Access~pool_www.regelleistung.net_80=rd2o00000000000000000000ffff0aa6e384o80; XSRF-TOKEN=994a0636-b7ca-453b-b468-77fe26392cf3; iprl_gw_sid=NmJmZTIyMzItNDI5OC00NmQwLWIzNjQtZjhkNThlMzQ4OGIw',
    'Host': 'www.regelleistung.net',
    'Origin': 'https://www.regelleistung.net',
    'Referer': 'https://www.regelleistung.net/ext/data/',
    'sec-ch-ua': '"Chromium";v="94", "Microsoft Edge";v="94", ";Not A Brand";v="99"',
    'sec-ch-ua-mobile': '?1',
    'sec-ch-ua-platform': '"Android"',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'same-origin',
    'Sec-Fetch-User': '?1',
    'Upgrade-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/94.0.4606.81 Mobile Safari/537.36 Edg/94.0.992.50',
}
tday = date.today()
tdate = str(tday.strftime("%d.%m.%Y"))
yday = tday - timedelta(days=1)
fdate = str(yday.strftime("%d.%m.%Y"))
tos_val = ['4', '3', '2', '1', '6', '11']
tos_name = ['50Hertz', 'Amprion', 'TenneT', 'TransnetBW', 'Netzregelverbund', 'IGCC']
for t in range(len(tos_val)):
    js = requests.get(url+"products?uenbId="+str(tos_val[t]),headers=headers).json()
    tname=tos_name[t]
    for d_val in js:
        data_type=js[d_val]
        print(tname,data_type)
        data = {
            'from': fdate ,
            'to': tdate,
            '_download': 'on',
            'tsoId': tos_val[t],
            'dataType': d_val
            }

        value = requests.post(url,data=data,headers=headers)
        bs = BeautifulSoup(value.content,"html.parser")
        table = bs.find("table").find('tbody')
        for tr in table.find_all('tr'):
            r_date = t_from = t_to = b_neg = b_pos = q_neg = q_pos = btr = ql = nrv = mol = exp = b_imp = b_exp = q_imp = q_exp = q_eu = q_impeu = q_expeu ="NA"
            lis = [td.text for td in tr.find_all('td')]
            if data_type == "MR" or data_type == "SCR" or data_type == "emergency power" or data_type == "SCR to aFRR cooperation" or data_type == "SCR from aFRR cooperation": 
                Others(fdate,tdate,tname,data_type,lis)
            elif data_type == "RZ_SALDO":
                Rzsaldo(fdate,tdate,tname,data_type,lis)
            elif data_type == "REBAP" or data_type == "German IGCC settlement price":
                Rebap(fdate,tdate,tname,data_type,lis)
            elif data_type == "emergency assistance":
                Ea(fdate,tdate,tname,data_type,lis)
            elif data_type == "Pre-Netting price":
                Prenet(fdate,tdate,tname,data_type,lis)
            elif (data_type == "Transfer_APG" or data_type == "Transfer_Elia" or data_type == "Transfer_ESO" or data_type == "Transfer_Swissgrid" or 
                data_type == "Transfer_CEPS" or data_type == "Transfer_EnDK" or data_type == "Transfer_REE" or data_type == "Transfer_RTE" or data_type == "Transfer_Admie" or 
                data_type == "Transfer_HOPS" or data_type == "Transfer_MAVIR" or data_type == "Transfer_TERNA" or data_type == "Transfer_TennetNL" or 
                data_type == "Transfer_PSE" or data_type == "Transfer_REN" or data_type == "Transfer_Transelectrica" or data_type == "Transfer_EMS" or 
                data_type == "Transfer_ELES" or data_type == "Transfer_SEPS" or data_type == "Transfer_Germany" or data_type == "Pre-Netting with APG" ):
                    Others1(fdate,tdate,tname,data_type,lis)
            
fh.close()