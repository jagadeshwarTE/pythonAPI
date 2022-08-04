import requests
import json
import numpy
import csv
from geopy.geocoders import Nominatim
import xlsxwriter
import pandas as pd

from datetime import datetime
from datetime import date
from datetime import timedelta
now = datetime.now()
start_time = now.strftime("%D - %H:%M:%S")
today = date.today()
yesterday = today - timedelta(days = 1)
firstday = date(yesterday.year, yesterday.month, 1)
month = yesterday.strftime("%B")
print(firstday)
print(yesterday)

###########   INITIALIZATIOn   ############

from datetime import datetime
now = datetime.now()
start_time = now.strftime("%H:%M:%S")

#************* Envision details ************
begin_time = str(firstday)+" 00:00:00"
end_time = str(yesterday)+" 23:00:00"
#begin_time = "2022-07-01 00:00:00"
#end_time = "2022-07-31 23:00:00"
#month = "July"
ed1 = end_time.split()
ed2 = ed1[0]
ed3 = ed2.split('-')
ed = int(ed3[2])
print(ed)


emt_metrics = "EMT.APConsumedKWH,EMT.APProductionKWH"

#*********** Sorting order **********


file = open("E:\Python_Programs\Total\Modified1\Finalised\Python\sorting.csv")
csvreader = csv.reader(file)
sort = []
sort_list = ['','','']
for sot in csvreader:
    sort.append(sot[0])
    sort_list = numpy.vstack((sort_list,sot))
#del sort[0]
#print(rows[0])
file.close()
sort_list = numpy.delete(sort_list,0,0)


# ************ Create a workbook and add a worksheet. ************
s4c = s4r = 0
s3c = s3r = 0
s2c = s2r = 0
s1c = s1r = 0
s5c = s5r = 0
s7c = s7r = 0
s8c = s8r = 0
s9c = s9r = 0
s10c = s10r = 0


wb = xlsxwriter.Workbook('E:\Python_Programs\Total\Gen_Report_fghjfgyhjfguyf'+str(yesterday)+'.xlsx')
cell_format = wb.add_format()
cell_format.set_font_color('red')
cell_format1 = wb.add_format()
cell_format1.set_font_color('blue')
#cell_format.set_shrink()
worksheet1 = wb.add_worksheet('Gen Report')
head = ['Country', 'Site Name', 'Capacity', 'Budget Production', 'EM_Production', 'INV_Generation Recorded','Expected PR','Act Vs Budget Procution(Bd_PR)', 'Budget Irradiation','Satellite Irr', 'Irradiation', 'Sat-Act Irr','Actual Vs Budget Irradiation','Power Limitation', 'Eq_Fail', 'Power_Fail', 'Scheduled', 'Rq_shutdown', 'Startup', 'Unspecified', 'Budget PR', 'Wc PR']
for r in head:
   worksheet1.write(s1r, s1c, r)
   s1c = s1c + 1
s1r = s1r + 1




worksheet2 = wb.add_worksheet('Gen data')
head = ['Country', 'Site Name', 'Device Name','Capacity', 'Count', 'Min_Date', 'Min_Production_Read', 'Min_Consumption_read','Max_Date', 'Max_Prod_Read','Max_Consumption_Read' ,'Production_Recorded','Consumption','Net Production']
for r in head:
   worksheet2.write(s2r, s2c, r)
   s2c = s2c + 1
s2r = s2r + 1

worksheet3 = wb.add_worksheet('No EMT Data')
head = ['Country', 'Site Name','Capacity', 'Device Name','Date']
for r in head:
   worksheet3.write(s3r, s3c, r)
   s3c = s3c + 1
s3r = s3r + 1

worksheet4 = wb.add_worksheet('Data Freeze')
head = ['Country', 'Site Name','Capacity', 'Device Name','Date','Meter Production read', 'Meter Consumption Read']
for r in head:
   worksheet4.write(s4r, s4c, r)
   s4c = s4c + 1
s4r = s4r + 1

worksheet5 = wb.add_worksheet('Device Reset Info')
head = ['Country', 'Site Name','Capacity', 'Device Name','Reset Date','Meter Production read', 'Meter Consumption Read']
for r in head:
   worksheet5.write(s5r, s5c, r)
   s5c = s5c + 1
s5r = s5r + 1

s6r = s6c = 0
worksheet6 = wb.add_worksheet('Inv_Gen no data')
head = ['Site Name','Date']
for r in head:
   worksheet6.write(s6r, s6c, r)
   s6c = s6c + 1
s6r = s6r + 1

worksheet7 = wb.add_worksheet('No Sat Irr')
head = ['Site Name','Date']
for r in head:
   worksheet7.write(s7r, s7c, r)
   s7c = s7c + 1
s7r = s7r + 1

worksheet8 = wb.add_worksheet('No Irr')
head = ['Site Name','Date']
for r in head:
   worksheet8.write(s8r, s8c, r)
   s8c = s8c + 1
s8r = s8r + 1

worksheet9 = wb.add_worksheet('Less IRR')
head = ['Site Name','Date','Sattelite Irradiation', 'Site Irradiation', '% Difference']
for r in head:
   worksheet9.write(s9r, s9c, r)
   s9c = s9c + 1
s9r = s9r + 1

worksheet10 = wb.add_worksheet('No Budget')
head = ['Site Name','Budget']
for r in head:
   worksheet10.write(s10r, s10c, r)
   s10c = s10c + 1
s10r = s10r + 1


k=1
dates = ['','']
for i in range(20):
    i+=1
    et = end_time.split()
    et1 = et[0]
    et2 = et1.split('-')
    et3 = et2[2]
    et3 = int(et3)
    
    bt = begin_time.split()
    bt1 = bt[0]
    bt2 = bt1.split('-')
    bt3 = bt2[2]
      
    if i*3 == et3:
        et = bt2[0]+'-'+bt2[1]+'-'+str(i*3)+' '+et[1]
        st = bt2[0]+'-'+bt2[1]+'-'+str(k)+' '+bt[1]
        date_range = [st,et]
        dates = numpy.vstack((dates,date_range))
        break
    elif i*3 > et3:
        et = bt2[0]+'-'+bt2[1]+'-'+str(et3)+' '+et[1]
        st = bt2[0]+'-'+bt2[1]+'-'+str(k)+' '+bt[1]
        date_range = [st,et]
        dates = numpy.vstack((dates,date_range))
        #print(st,'to',et)
        break
    et = bt2[0]+'-'+bt2[1]+'-'+str(i*3)+' '+et[1]
    st = bt2[0]+'-'+bt2[1]+'-'+str(k)+' '+bt[1]
    date_range = [st,et]
    dates = numpy.vstack((dates,date_range))
    k = i*3+1
    #print(st,'tp',et,"**",et3,"*",i)
#print(dates)
dates = numpy.delete(dates,0,0)
#print(dates)


#***************** LOGIN *************#

def my_login():
    url = "https://app-portal-eu2.envisioniot.com/solar-api/v1.0/loginService/getOrganizationList"
    payload = json.dumps({
      "username": "jagadeshwar",
      "password": "nl07EA@4284"
    })
    headers = {
      'Content-Type': 'application/json'
    }
    response = requests.request("POST", url, headers=headers, data=payload)
    #print(response.text)
    data = json.loads(response.text)
    org = data["data"]
    orgidd = org["organizations"]
    #print(orgids[0])
    orgids = orgidd[0]
    token = org["accessToken"]
    print(token)
    orgid = orgids["id"]
    #print(orgid)

    ########Choose organization############
    url = "https://app-portal-eu2.envisioniot.com/solar-api/v1.0/loginService/chooseOrganization?token="+token+"&organizationId="+orgid
    payload={}
    headers = {}
    response = requests.request("GET", url, headers=headers, data=payload)
    #print(response.text)
    return [token, orgid]

login = my_login()
token = login[0]
orgid = login[1]


#### Get asset mdmids ##########
url = "https://app-portal-eu2.envisioniot.com/solar-api/v1.0/mdmService/getObjectStructure?type=102&token="+token
payload={}
headers = {}
response = requests.request("GET", url, headers=headers, data=payload)
#print(response.text)
data = json.loads(response.text)
data1 = data["data"]
#print(data["status"])

if data["status"]!=0:
    print('No internet or not able to login')

inv_gen = 0

site_gen = 0
siteids = [['']]
mdmids = ','
extra = 0
aaa = 0
site_info = ['','','','']
if aaa == 0:
    data2 = data1[33]
for data2 in data1:

    aaa = aaa + 1
    print(aaa)
    

    if aaa>264851:
        break
    
    site_mdmid = data2["mdmid"]
    data3 = data2["attributes"]
    cap = data3["capacity"]
    prod_source = data3["prodSource"]
    sitename = data3["name"]
    print(sitename)
    #sitename = "S.Kijchai Enterprise PCL"
    ###########find country code############

    if sitename in sort:
        country = sort_list[sort.index(sitename)][1]
    else:
        i = 1
        while i>0:
            try:
                
                geolocator = Nominatim(user_agent="geoapiExercises")
                # Latitude & Longitude input
                Latitude = data3["latitude"]
                Longitude = data3["longitude"]
                location = geolocator.reverse(Latitude+","+Longitude)
                address = location.raw['address']
                country = address["country_code"]
                #print(country)
                tt = True

            except (NameError, geopy.exc.GeocoderUnavailable):

                print("searching country...")
                tt = False

            if tt:
                break
    #country = 'III'
    site_data = [country,sitename,cap,site_mdmid]
    site_info = numpy.vstack((site_info, site_data))
    mdmids = site_mdmid+','+mdmids
    siteids = numpy.vstack((siteids,site_mdmid))
mdmids = mdmids[:-2]
#print(mdmids)
siteids = numpy.delete(siteids, 0, 0)
#print(siteids)

site_info = numpy.delete(site_info, 0, 0)
siteinfo = [['','','','']]
sn = ["Fong"]
for ssn in site_info:
    if ssn[1] == "Fong":
        #print(ssn[1],"-*-*-*-*-*-*-")
        ssn1 = [ssn[0], 'Fong_77', ssn[2], ssn[3]]
        #print(ssn1)
        siteinfo = numpy.vstack((siteinfo,ssn1))
        fid = ssn[3]
        ssn2 = [ssn[0], 'Fong_79', ssn[2], fid+fid]
        siteinfo = numpy.vstack((siteinfo,ssn2))
    else:       
        siteinfo = numpy.vstack((siteinfo,ssn))
#print(siteinfo)
site_info = siteinfo
site_info = numpy.delete(site_info, 0, 0)


site_ids = ['']
for ssd in site_info:
    mmd = ssd[3]
    site_ids = numpy.vstack((site_ids,mmd))
site_ids = numpy.delete(site_ids, 0, 0)    



    ################### GET EMT mdmids ################

types = "211,206,208"

url = "https://app-portal-eu2.envisioniot.com/solar-api/v1.0/mdmService/getObjects?mdmids="+mdmids+"&token="+token+"&organizationId="+orgid+"&types="+types
payload={}
headers = {}
response = requests.request("GET", url, headers=headers, data=payload)
#print(response.text)
data = json.loads(response.text)
data1 = data["data"]
        
        
status = data["status"]
if status != 0:

    login = my_login()
    token = login[0]
    orgid = login[1]
            
    url = "https://app-portal-eu2.envisioniot.com/solar-api/v1.0/mdmService/getObjects?mdmids="+mdmids+"&token="+token+"&organizationId="+orgid+"&types="+types
    payload={}
    headers = {}
    response = requests.request("GET", url, headers=headers, data=payload)
    #print(response.text)
    data = json.loads(response.text)
    data1 = data["data"]

emt_mdmids = ','
dev_info = ['','','']
z=0
noemt = ['']
for site in siteids:
    #print(site,"*/*************",site_info[z])
    z+=1
    data2 = data1[site[0]]
    data3 = data2["mdmobjects"]
    try:
        data4 = data3["211"]
    except KeyError:
        for zz in site_info:
            if zz[3] == site:
                noemt = numpy.hstack((noemt,zz[1]))
        data4 = []
        dev_data = [site,'No EMT','No EMT']
        #dev_info.loc[len(dev_info)] = dev_data
    for data5 in data4:
        data6 = data5["attributes"]
        site_id = data6["parentID"]
        dev_id = data6["objectID"]
        dev_name = data6["name"]
        dev_data = [site_id,dev_id,dev_name]
        dev_info = numpy.vstack((dev_info, dev_data))
        emt_mdmids = dev_id+','+emt_mdmids
#print(dev_info)
emt_mdmids = emt_mdmids[:-2]
print(noemt)
print('******')
noemt = numpy.delete(noemt, 0)
print(noemt)


##************* REMOVE Devices *****************
#print(dev_info)
dev_info = numpy.delete(dev_info, 0, 0)
remove_dev = ["HF4HxGnL","gyNUXk5S","wBhd8PQ5","fa6HiR51","HvK1fEkJ","fKgPGdq6","Q4qZ4vr7"]
ddv = ddd = 0
devinfo = dev_info
for dd in dev_info:
    #print(dd[1])
    if dd[1] in remove_dev:
        print(ddv,dev_info[ddv])
        devinfo = numpy.delete(devinfo, ddv-ddd, 0)
        ddd+=1
    ddv+=1
dev_info = devinfo

deviceinfo = [['','','']]

#print(dev_info)

for ddn in dev_info:
    if ddn[0] == fid:
        print(ddn[1],"-*-*-*-*-*-*-")
        if ddn[1] == "St8pLaXD":
            ddn1 = [fid+fid, ddn[1], ddn[2]]
            print(ddn1)
            deviceinfo = numpy.vstack((deviceinfo,ddn1))
        else:
            deviceinfo = numpy.vstack((deviceinfo,ddn))                
    else:       
        deviceinfo = numpy.vstack((deviceinfo,ddn))
#print(siteinfo)
dev_info = deviceinfo
dev_info = numpy.delete(dev_info, 0, 0)
#print(dev_info)






metrics = "EMT.APConsumedKWH,EMT.APProductionKWH"
em_read = ['','','','']
for dat in dates:
    begin_tim = dat[0]
    end_tim = dat[1]
    
    url = "https://app-portal-eu2.envisioniot.com/solar-api/v1.0/metricService/multiMetrics?mdmids="+emt_mdmids+"&token="+token+"&metrics="+metrics+"&begin_time="+begin_tim+"&end_time="+end_tim+"&time_group="+"H"
    payload={}
    headers = {}
    response = requests.request("GET", url, headers=headers, data=payload)
    #print(response.text)
    stat = json.loads(response.text)
    status = stat["status"]
    #print(status)
                
    if status != 0:
        print("...Re-Connecting...")

        login = my_login()
        token = login[0]
        orgid = login[1]

        url = "https://app-portal-eu2.envisioniot.com/solar-api/v1.0/metricService/multiMetrics?mdmids="+emt_mdmids+"&token="+token+"&metrics="+metrics+"&begin_time="+begin_tim+"&end_time="+end_tim+"&time_group="+"H"
        payload={}
        headers = {}
        response = requests.request("GET", url, headers=headers, data=payload)
        #print(response.text)
        stat = json.loads(response.text)
        status = stat["status"]
        print(stat)

    
    emt1 = stat["metrics"]
    for emt in emt1:
        devId = emt["mdmId"]
        dt = emt["timestamp"]
        try:
            cons = emt["EMT.APConsumedKWH"]
        except KeyError:
            cons = '-'
            
        try:
            
            prod = emt["EMT.APProductionKWH"]
        except KeyError:
            
            prod = '-'
        dev_dd = [devId,dt,prod,cons]
        em_read = numpy.vstack((em_read, dev_dd))
    print(end_tim)

#print(em_read,"*******")
#dev_info =numpy.delete(dev_info, 0, 0)



metrics_read = "SITE.PlannedPR,SITE.PlannedProduction,SITE.BudgetIrradiance,SITE.APProduction,SITE.TotSatIrradiancePerOrientation,SITE.RadiationACC,SITE.PowerLimitationLoss,SITE.DtEquipFailCauseLoss,SITE.DtPowerFailCauseLoss,SITE.DtCurtailmentCauseLoss,SITE.DtLateStartupCauseLoss,SITE.DtSchedMaintCauseLoss,SITE.DtOtherLoss"


url = "https://app-portal-eu2.envisioniot.com/solar-api/v1.0/metricService/multiMetrics?mdmids="+mdmids+"&token="+token+"&metrics="+metrics_read+"&begin_time="+begin_time+"&end_time="+end_time+"&time_group=D"
payload={}
headers = {}
response = requests.request("GET", url, headers=headers, data=payload)
#print(response.text)
#print(response.text)
stat = json.loads(response.text)
status = stat["status"]
#print(status)
            
if status != 0:
    print("...Re-connecting...")
    login = my_login()
    token = login[0]
    orgid = login[1]    


    url = "https://app-portal-eu2.envisioniot.com/solar-api/v1.0/metricService/multiMetrics?mdmids="+mdmids+"&token="+token+"&metrics="+metrics_read+"&begin_time="+begin_time+"&end_time="+end_time+"&time_group=D"
    payload={}
    headers = {}

    response = requests.request("GET", url, headers=headers, data=payload)
    

site_data = json.loads(response.text)
#print(site_data)
data1 = site_data["metrics"]
aakk = 0
extra = 0
wr = [0,1]
#print("************",numpy.where(site_info == "CM CP Feed Mill (CPPH)"),"*/*/*/*/*/*/*/*/*/*/")    
for site in site_info:
    d=0
    aakk+=1
    if aakk >98796511:
        break
        print("break...........")

#for site in site_info:
    
    #site = site_info[29]
    siteid = site[3]
    sitename = site[1]
    print(sitename)
    #sitename = "S.Kijchai Enterprise PCL"
    #data2 = data1[0]
    total_Site_Prod = 0
    total_irr = 0
    total_inv_prod = total_emt_prod = 0
    total_sitecurtail = 0
    total_eq_fail = 0
    tota_eq_fail = ''
    total_pw_fail = 0
    tota_pw_fail = ''
    total_curt_ci = 0
    total_schld = 0
    total_curt = 0
    total_un = 0
    total_sattirr = total_satirr =total_abs= 0
    total_curtailment = total_pw_limit = total_startup = total_rq_shut = total_bdprod = total_bdirr = 0
    tota_un = ''
    status = 'F'
    irs = satirs = 'F'
    print(sitename)
    #print(len(data1))

    
    for data2 in data1:

        if siteid == data2["mdmId"]:
        
            try:
                inv_prod = data2["SITE.APProduction"]
            except KeyError:
                inv_prod = 0
                status = 'T'
                no = [sitename, data2["timestamp"]]
                s6c = 0
                for noo in no:
                    worksheet6.write(s6r, s6c, noo)
                    s6c = s6c + 1
                s6r = s6r + 1
                sattirr = 0
            inv_prod = float(inv_prod)
            total_inv_prod = total_inv_prod + inv_prod



            timestamp = data2["timestamp"]

            try:
                satdata = data2["SITE.TotSatIrradiancePerOrientation"]
                satdata1 = json.loads(satdata)
                satdata2 = satdata1[0]
                #print(satdata2["irr"])       

                sattirr = satdata2["irr"]
                #print(sattirr)
            except KeyError:
                no = [sitename, data2["timestamp"]]
                s7c = 0
                #print(no,'*******sat')
                for noo in no:
                    worksheet7.write(s7r, s7c, noo)
                    s7c = s7c + 1
                s7r = s7r + 1
                sattirr = 0
                
                sattirs = 'T'
            #print(data2)
            sattirr = float(sattirr)/1000
            total_sattirr = (total_sattirr + sattirr)
            tota_sattirr = irs+'*'+str(total_sattirr)
            
            try:
                irr = data2["SITE.RadiationACC"]
            except KeyError:
                no = [sitename, data2["timestamp"]]
                s8c = 0
                #print(no,'*******irr')
                for noo in no:
                    worksheet8.write(s8r, s8c, noo)
                    s8c = s8c + 1
                s8r = s8r + 1
                irr = 0
                #print(data2)
                irs = 'T'
            irr = float(irr)/1000



            try:
                delta_irr = abs((sattirr-irr)/sattirr)*100
            except ZeroDivisionError:
                delta_irr = 0
                

            if delta_irr>20 and sattirr!=0 and irr!=0:
                less = [sitename, data2["timestamp"],sattirr,irr,delta_irr]
                s9c = 0
                #print(less,'*******less')
                for noo in less:
                    worksheet9.write(s9r, s9c, noo)
                    s9c = s9c + 1
                s9r = s9r + 1
                sattirr = 0
                
            total_irr = (total_irr + irr)
            tota_irr = irs+'*'+str(total_irr)

            try:
                bdirr = data2["SITE.BudgetIrradiance"]
            except KeyError:
                head = [sitename,'Irr']
                s10c = 0
                for r in head:
                    worksheet10.write(s10r, s10c, r)
                    s10c = s10c + 1
                s10r = s10r + 1
                bdirr = 0
                #print(data2)
                irs = 'T'
            bdirr = float(bdirr)/1000
            total_bdirr = total_bdirr + bdirr

            try:
                bdprod = data2["SITE.PlannedProduction"]
            except KeyError:
                head = [sitename,'Prod']
                s10c = 0
                for r in head:
                    worksheet10.write(s10r, s10c, r)
                    s10c = s10c + 1
                s10r = s10r + 1
                bdirr = 0
                bdprod = 0
                #print(data2)
                irs = 'T'
            bdprod = float(bdprod)
            total_bdprod = total_bdprod + bdprod

            try:
                thpr = data2["SITE.PlannedPR"]
            except KeyError:
                thpr = 0
            thpr = float(thpr)

                        

    ###

            try:
                pw_limit = data2["SITE.PowerLimitationLoss"]
            except KeyError:
                pw_limit = 0
                status = 'T'
            pw_limit = float(pw_limit)
            total_pw_limit = total_pw_limit + pw_limit

            status = 'F'
            try:
                eq_fail = data2["SITE.DtEquipFailCauseLoss"]
            except KeyError:
                eq_fail = 0
                status = 'T'
            eq_fail = float(eq_fail)
            total_eq_fail = total_eq_fail + eq_fail
            tota_eq_fail = status+'*'+str(total_eq_fail)
            status = 'F'
            #print(total_eq_fail)

            status = 'F'
            try:
                pw_fail = data2["SITE.DtPowerFailCauseLoss"]
            except KeyError:
                pw_fail = 0
                status = 'T'
            pw_fail = float(pw_fail)
            total_pw_fail = total_pw_fail + pw_fail
            tota_pw_fail = status+'*'+str(total_pw_fail)
            

            try:
                rq_shut = data2["SITE.DtCurtailmentCauseLoss"]
            except KeyError:
                rq_shut = 0
                status = 'T'
            rq_shut = float(rq_shut)
            total_rq_shut = total_rq_shut + rq_shut
            

            status = 'F'
            try:
                startup = data2["SITE.DtLateStartupCauseLoss"]
            except KeyError:
                startup = 0
                status = 'T'
            startup = float(startup)
            total_startup = total_startup + startup ##### + startup

            status = 'F'
            try:
                schld = data2["SITE.DtSchedMaintCauseLoss"]
            except KeyError:
                schld = 0
                status = 'T'
            schld = float(schld)        
            total_schld = total_schld + schld

            
            status = 'F'
            try:
                un_sp = data2["SITE.DtOtherLoss"]
            except KeyError:
                un_sp = 0
                status = 'T'
            un_sp = float(un_sp)
            total_un = total_un + un_sp
            tota_un = status+'*'+str(total_un)
            status = 'F'
     
    #gen_report = [ site[0], site[1], site[2], total_inv_prod, total_emt_prod, total_sattirr,total_irr,total_pw_limit,total_eq_fail,total_pw_fail,total_curt_ci,total_schld,total_rq_shut,total_startup,total_un]
    #print(gen_report)    






    total_prod = total_cons = total_net = 0
    for dev in dev_info:
        d+=1
        rst = False
        #print(dev[0],"*/*/*/*/*/*/*/*/*/*/")
        if site[3] == dev[0]:
            r=0
            emt_data = [['','','','','','','']]
            emt_data_read = ['','','','','']
            for read in em_read:
                r+=1
                if dev[1] == read[0]:
                    #print(read)
                    prod = read[2]
                    cons = read[3]
                    if prod == '-':
                        em_data = [site[0],site[1],site[2],dev[2],read[1]]
                        em_data1 = [site[0],site[1],dev[2],read[1],'']
                        s3c = 0
                        for noem in em_data:
                            worksheet3.write(s3r, s3c, noem)
                            s3c = s3c + 1
                        s3r = s3r + 1
                    elif cons == '-' and prod != '-':
                        em_data = [site[0],site[1],site[2],dev[2],read[1],read[2],0]
                        em_data1 = [site[0],site[1],dev[2],read[1],read[2]]
                        emt_data = numpy.vstack((emt_data,em_data))     
                    else:
                        em_data = [site[0],site[1],site[2],dev[2],read[1],read[2],read[3]]
                        em_data1 = [site[0],site[1],dev[2],read[1],read[2]]
                        emt_data = numpy.vstack((emt_data,em_data))
                    emt_data_read = numpy.vstack((emt_data_read,em_data1))
            #print(emt_data,"******")
            path = "E:\Python_Programs\Total\Modified1\Finalised\Python\BackUP\\"+site[1]+"_"+dev[2]+"_"+month+".csv"
            with open(path, 'w', newline='') as file:
                writer = csv.writer(file)
                #writer.writerow(emt_data_read)
                # Use writerows() not writerow()
                writer.writerows(emt_data_read)
            emt_data = numpy.delete(emt_data, 0, 0)
            count = len(emt_data)
            #print(count)
            aa = bb = chk = 0
            flem = ['','','','','','','']
            #print("*********",emt_data,"********")
            if count != 0:

                for row in emt_data:
                    
                    dtt = row[4]
                    #print(dtt)
                    hr1 = dtt.split()
                    hr = int(hr1[1])
                    #print(row)
                    pw = float(row[5])
                    aa = pw
                    if aa == bb:
                        xyz = 0
                    else:
                        flem = numpy.vstack((flem,row))

                    cc = aa-bb
                    s5c = 0
                    if cc < 0:
                        for rr in row:
                            worksheet5.write(s5r, s5c, rr)
                            s5c = s5c+1
                        s5r = s5r+1
                        rst = True
                        print("**********RESET**********")

                    if hr in range(6,18):
                        if aa == bb and aa!=0 :
                            if chk == 0:
                                s4c = 0
                                for rr in row:
                                    worksheet4.write(s4r, s4c, rr)
                                    s4c = s4c + 1
                                s4r = s4r + 1
                            else:
                                chk = 1
                    bb = pw
                flem = numpy.delete(flem, 0, 0)

                m_read = flem[0]
                #print("****111111****")
                
                M_read = flem[len(flem)-1]
                
                dev_prod = float(M_read[5])-float(m_read[5])
                dev_cons = float(M_read[6])-float(m_read[6])
                dev_net = dev_prod - dev_cons

            else:
                flem = [[site[0],site[1],site[2],dev[2],read[1],'0','0']]
                m_read = flem[0]
                #print("****222222222****")
                
                M_read = flem[len(flem)-1]
                
                dev_prod = float(M_read[5])-float(m_read[5])
                dev_cons = float(M_read[6])-float(m_read[6])
                dev_net = dev_prod - dev_cons
            
            gen_data = [flem[0][0],flem[0][1],flem[0][2],flem[0][3],str(count),m_read[4],str(m_read[5]),str(m_read[6]),M_read[4],str(M_read[5]),str(M_read[6]),str(dev_prod),str(dev_cons),str(dev_net)]
            #print("****************",flem[0])
            #print(gen_data)
            #print("**********")
            #for rz in gen_data:
                #print(type(rz))
                #print("*****////****")
                #print(rz)
            #print(M_read)
            s2c = 0
            for rr in gen_data:
                t1 = M_read[4]
                tt1 = m_read[4]
                #print(t1,"******",tt1,"********",rr)
                if t1 == rr:
                    #print("yesss")
                    try:
                        t2 = t1.split()
                        th3 = t2[0]
                        th4 = th3.split('-')
                        td = int(th4[2])
                        try:
                            t3 = int(t2[1])
                            #print(type(t3),t3)
                        except IndexError:
                            t3 = 18
                    except AttributeError:
                        t3 = 18
                        th = ed
                    #print(t3)
                elif tt1 ==rr:
                    
                    try:
                        tt2 = tt1.split()
                        tt3 = tt2[0]
                        tt4 = tt3.split('-')
                        ttd = int(tt4[2])
                        #print("***********",ttd,"************")
                        try:
                            tth = int(tt2[1])
                            #print(type(tth),tth)
                        except IndexError:
                            tth = 0
                            ttd = 1
                    except AttributeError:
                        tth = 0
                        ttd = 1
                    #print(t3)
                        
                else:
                    t3 = 18
                    ttd = 1
                    tth = 0
                    td = ed
                #print(t3)
                #print(t3,'*',td,'*',tth,'*',ttd,'*',rst)
                if rst :
                    worksheet2.write(s2r, s2c, rr, cell_format)
                else:
                    if (t3 not in range (17,20) or td != ed or ttd != 1 or tth not in range(0,6)):
                        worksheet2.write(s2r, s2c, rr, cell_format)
                        t3 = 18
                    else:
                        worksheet2.write(s2r, s2c, rr)
                        #print("llll")
                s2c = s2c + 1
            s2r = s2r + 1


            total_prod = total_prod + dev_prod
            total_cons = total_cons + dev_cons
            total_net = total_net + dev_net

    #gen_report = [ site[0], site[1], site[2], total_inv_prod, total_emt_prod, total_sattirr,total_irr,total_sitecurtail,total_eq_fail,total_pw_fail,total_curt_ci,total_schld,total_un]
    #total_pw_limit,total_eq_fail,total_pw_fail,total_curt_ci,total_schld,total_rq_shut,total_startup,total_un]

    if site[1] in noemt:
        try:
            exppr = ((total_inv_prod+total_pw_limit+total_eq_fail+total_pw_fail+total_schld+total_rq_shut+total_startup-(total_irr*float(site[2])*1000))/(total_irr*float(site[2])*1000))*100
        except ZeroDivisionError :
            acpr = 0
            
        try:
            bdac_prod = ((total_inv_prod+total_pw_limit+total_eq_fail+total_pw_fail+total_schld+total_rq_shut+total_startup-total_bdprod)/(total_bdprod))*100
        except ZeroDivisionError :
            bdac_prod = 0
        
        try:
            satac_irr = ((total_irr-total_sattirr)/total_sattirr)*100
        except ZeroDivisionError :
            satac_irr = 0

        try:
            bdac_irr = ((total_irr-total_bdirr)/total_bdirr)*100
        except ZeroDivisionError :
            satac_irr = 0
            
            
    else:
        try:
            exppr = ((total_prod+total_pw_limit+total_eq_fail+total_pw_fail+total_schld+total_rq_shut+total_startup-(total_irr*float(site[2])*1000))/(total_irr*float(site[2])*1000))*100
        except ZeroDivisionError :
            acpr = 0

        try:
            bdac_prod = ((total_prod+total_pw_limit+total_eq_fail+total_pw_fail+total_schld+total_rq_shut+total_startup-total_bdprod)/(total_bdprod))*100
        except ZeroDivisionError :
            bdac_prod = 0

        
        try:
            satac_irr = ((total_irr-total_sattirr)/total_sattirr)*100
        except ZeroDivisionError :
            bdac_irr = 0

        try:
            bdac_irr = ((total_irr-total_bdirr)/total_bdirr)*100
        except ZeroDivisionError :
            satac_irr = 0
        
    #Alliance - PSAT,


    site_report = [ site[0], site[1], float(site[2]), total_bdprod,total_prod, total_inv_prod,exppr,bdac_prod,total_bdirr,(total_sattirr)*0.98,total_irr,satac_irr,bdac_irr,total_pw_limit,total_eq_fail,total_pw_fail,total_schld,total_rq_shut,total_startup,total_un, thpr,exppr]

    s1r = 1
    brk = False
    #print(len(sort),"/*/*/*/*/")
    for sn in sort:
        s1r = s1r + 1
        if sn == sitename:
            brk = True
            break
            

    if not brk:
        extra = extra+1
        s1r = s1r + extra
        #print(s1r)
              
    wr = numpy.hstack((wr,s1r))

    s1c = 0
    for rr in site_report:
        if abs(thpr-exppr)>10 and s1c>18:
            if str(rr) in noemt:
                worksheet1.write(s1r, s1c, rr, cell_format1)
            else:
                worksheet1.write(s1r, s1c, rr, cell_format)
        else:
            if str(rr) in noemt:
                worksheet1.write(s1r, s1c, rr, cell_format1)
            else:
                worksheet1.write(s1r, s1c, rr)
        s1c = s1c + 1
    s1r = s1r + 1

lr = max(wr)
s1c = 0
for i in range(0,lr):
    if i in wr:
            
        abcd = 0
    else:
        #print('yesssss',i,'-',sort_list[i-2][2])
        site_re = [ sort_list[i-2][1], sort_list[i-2][2], 0, 0,0,0, 0,0,0,0,0,0,0,0,0,0,0,0, 0,0]
        #print(site_report)
        s1c = 0
        for rrr in site_re:
            worksheet1.write(i, s1c, rrr)
            s1c = s1c + 1
                
                

            

wb.close()

now = datetime.now()

current_time = now.strftime("%H:%M:%S")
print("Start time =", start_time)
print("Current Time =", current_time)


##########
############
############
          
































