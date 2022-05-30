import requests
import pandas as pd
from io import StringIO
from io import BytesIO
import time
from datetime import date
from selenium import webdriver
import matplotlib.pyplot as plt
from matplotlib.pyplot import figure
from PIL import Image, ImageDraw, ImageFont
from arabic_reshaper import reshape
from bidi.algorithm import get_display
import numpy as np

def Day_list():  
    yaer = ['1399','1400','1401']
    mon = ['01','02','03','04','05','06','07','08','09','10','11','12']
    day = []
    listday= []
    for x in range(32):
        if x != 0:
            if len(str(x))==1:
                day.append('0'+str(x))
            else:
                day.append(str(x))
    for y in yaer:
        for m in mon:
            if int(m)<7:
                for d in day:
                    listday.append(y+m+d)
            if int(m)>=7 and int(m) != 12:
                for d in day[:-1]:
                    listday.append(y+m+d)
            if int(m) == 12 and y!='1399':
                for d in day[:-2]:
                    listday.append(y+m+d)
            if int(m) == 12 and y=='1399':
                for d in day[:-1]:
                    listday.append(y+m+d)
    listday = [int(x) for x in listday]
    return listday

def enum(text):
    return get_display(reshape(text))


def render_mpl_table(data, col_width=3.0, row_height=0.625, font_size=14,
                     header_color='#40466e', row_colors=['#f1f1f2', 'w'], edge_color='w',
                     bbox=[0, 0, 1, 1], header_columns=0,
                     ax=None, **kwargs):
    if ax is None:
        size = (np.array(data.shape[::-1]) + np.array([0, 1])) * np.array([col_width, row_height])
        fig, ax = plt.subplots(figsize=size)
        ax.axis('off')
    mpl_table = ax.table(cellText=data.values, bbox=bbox, colLabels=data.columns, **kwargs)
    mpl_table.auto_set_font_size(False)
    mpl_table.set_fontsize(font_size)

    for k, cell in mpl_table._cells.items():
        cell.set_edgecolor(edge_color)
        if k[0] == 0 or k[1] < header_columns:
            cell.set_text_props(weight='bold', color='w')
            cell.set_facecolor(header_color)
        else:
            cell.set_facecolor(row_colors[k[0]%len(row_colors) ])
    return ax.get_figure(), ax

def Table_Price_Nav(Count):
    dfp = pd.read_excel('Data/histori.xlsx').sort_values(by=['<DTYYYYMMDD>'],ascending=False).reset_index()[['<DTYYYYMMDD>','<CLOSE>']][:Count]
    dfn = pd.read_excel('Data/nav histori.xlsx').sort_values(by=['date'],ascending=False).reset_index().reset_index()[['NAV','date']]
    dfp = dfp[dfp['<DTYYYYMMDD>']<=dfn['date'].max()]
    df = dfp.set_index('<DTYYYYMMDD>').join(dfn.set_index('date')).sort_index()
    
    df['diff'] = df['<CLOSE>'] - df['NAV']
    df['diffr'] = ((df['<CLOSE>'] / df['NAV'])-1)*10000
    print(df['NAV'])
    df['diffr'] = [round(x)/100 for x in df['diffr']]
    df['diffr'] = [str(x)+'%' for x in df['diffr']]
    df.columns = [enum('پایانی'),enum('ابطال'),enum('اختلاف'),enum('نرخ اختلاف')]
    dt = {}
    for i in range(len(df.index)):
        d = list(df.index)[i]
        b = str(d)[:4]+'/'+str(d)[4:6]+'/'+str(d)[6:]
        l = list(df[df.index==d].iloc[0])
        dt.update({b:l})
    dff = pd.DataFrame(index=(df.columns),data = dt).reset_index()
    fig,ax = render_mpl_table(dff, header_columns=0, col_width=2.0)
    fig.savefig("Plot/PrcNavWek.png")


def Table_Volume(Count):
    df = pd.read_excel('Data/histori.xlsx').sort_values(by=['<DTYYYYMMDD>'],ascending=True).reset_index()[['<DTYYYYMMDD>','<VOL>']]
    df['mean7'] = df['<VOL>'].rolling(7).mean()
    df['mean30'] = df['<VOL>'].rolling(30).mean()
    df = df.sort_values(by=['<DTYYYYMMDD>'],ascending=False)[:Count]
    df['mean7'] = [round(x) for x in df['mean7']]
    df['mean30'] = [round(x) for x in df['mean30']]
    
    df['diff7'] = df['<VOL>'] - df['mean7']
    df['diff30'] = df['<VOL>'] - df['mean30']
    
    df['diffr7'] = ((df['<VOL>'] / df['mean7'])-1)*10000
    df['diffr7'] = [round(x)/100 for x in df['diffr7']]
    df['diffr7'] = [str(x)+'%' for x in df['diffr7']]

    df['diffr30'] = ((df['<VOL>'] / df['mean30'])-1)*10000
    df['diffr30'] = [round(x)/100 for x in df['diffr30']]
    df['diffr30'] = [str(x)+'%' for x in df['diffr30']]
    
    df['<VOL>'] = [str(round(x/10000)/100)+' M' for x in df['<VOL>']]
    df['mean7'] = [str(round(x/10000)/100)+' M' for x in df['mean7']]
    df['mean30'] = [str(round(x/10000)/100)+' M' for x in df['mean30']]
    
    df['diff7'] = [str(round(x/10000)/100)+' M' for x in df['diff7']]
    df['diff30'] = [str(round(x/10000)/100)+' M' for x in df['diff30']]
    
    df = df.set_index('<DTYYYYMMDD>').sort_index()
    df.columns = [enum('حجم معاملات'),enum('میانگین 7 روزه'), enum('میانگین 30 روزه'),enum('مقدار اختلاف با م 7 روزه'), enum('مقدار اختلاف با م 30 روزه'),enum('نرخ اختلاف با م 7 روزه'),enum('نرخ اختلاف با م 30 روزه')]
    dt = {}
    for i in range(len(df.index)):
        d = list(df.index)[i]
        b = str(d)[:4]+'/'+str(d)[4:6]+'/'+str(d)[6:]
        l = list(df[df.index==d].iloc[0])
        dt.update({b:l})
    dff = pd.DataFrame(index=(df.columns),data = dt).reset_index()
    fig,ax = render_mpl_table(dff, header_columns=0, col_width=2.20)
    fig.savefig("Plot/VolWek.png")  

def Table_Roa():
    dl = pd.DataFrame(index=(Day_list()))
    df = pd.read_excel('Data/histori.xlsx').drop(columns='Unnamed: 0').set_index('<DTYYYYMMDD>')[['<CLOSE>']]
    df = dl.join(df).sort_index(ascending=False)
    df['<CLOSE>']= df['<CLOSE>'].fillna(method='bfill')
    df = df.dropna()
    df = df[df.index<=pd.read_excel('Data/histori.xlsx')['<DTYYYYMMDD>'].max()].sort_index(ascending=True)
    
    df['roc1'] = ((df['<CLOSE>']/df['<CLOSE>'].shift(1))-1)*10000
    df['roc7'] = ((df['<CLOSE>']/df['<CLOSE>'].shift(7))-1)*10000
    df['roc14'] = ((df['<CLOSE>']/df['<CLOSE>'].shift(14))-1)*10000
    df['roc30'] = ((df['<CLOSE>']/df['<CLOSE>'].shift(30))-1)*10000
    df['roc90'] = ((df['<CLOSE>']/df['<CLOSE>'].shift(90))-1)*10000
    df['roc180'] = ((df['<CLOSE>']/df['<CLOSE>'].shift(180))-1)*10000
    df['roc365'] = ((df['<CLOSE>']/df['<CLOSE>'].shift(365))-1)*10000
    df = df.dropna()
    df['roc1'] = [round(x)/100 for x in df['roc1']]
    df['roc7'] = [round(x)/100 for x in df['roc7']]
    df['roc14'] = [round(x)/100 for x in df['roc14']]
    df['roc30'] = [round(x)/100 for x in df['roc30']]
    df['roc90'] = [round(x)/100 for x in df['roc90']]
    df['roc180'] = [round(x)/100 for x in df['roc180']]
    df['roc365'] = [round(x)/100 for x in df['roc365']]
    df = df[df.index==df.index.max()].drop(columns='<CLOSE>')
    dt = {}
    for i in range(len(df.index)):
        d = list(df.index)[i]
        b = 'today'
        l = list(df[df.index==d].iloc[0])
        dt.update({b:l})
    df.columns = [enum('روزانه'),enum('7روزه'),enum('14روزه'),enum('30روزه'),enum('90روزه'),enum('180روزه'),enum('365روزه')]
    
    dff = pd.DataFrame(index=(df.columns),data = dt)
    target = ((1+0.204)**(1/365))-1
    dff['target'] = [target, (((target+1)**7)-1), (((target+1)**14)-1), (((target+1)**30)-1), (((target+1)**90)-1), (((target+1)**180)-1), (((target+1)**365)-1)]
    dff['target'] = [round(x*10000)/100 for x in dff['target']]
    
    dff['diff'] = dff['today'] - dff['target']
    dff['diff'] = [round(x*100)/100 for x in dff['diff']]
    dff['qqe'] = [1,7,14,30,90,180,365]
    dff['ydm'] =  (((dff['today']/100) +1) ** (365/dff['qqe']))-1
    dff['ydm'] = [round(x*10000)/100 for x in dff['ydm']]
    dff['today'] = [str(x)+'%' for x in dff['today']]
    dff['target'] = [str(x)+'%' for x in dff['target']]
    dff['diff'] = [str(x)+'%' for x in dff['diff']]
    dff['ydm'] = [str(x)+'%' for x in dff['ydm']]
    dff = dff.drop(columns='qqe')

    dff.columns = [enum('نقطه به نطقه'),enum('هدف'),enum('اختلاف'),enum('سالانه شده')]
    dff = dff.reset_index()
    clmn = list(dff.columns)
    clmn = [clmn[0],clmn[2],clmn[4],clmn[1],clmn[3]]
    dff = dff.reindex(columns=clmn)
    fig,ax = render_mpl_table(dff, header_columns=0, col_width=3)
    fig.savefig("Plot/ROA.png")  
    
    
def texttopng_title(text,name , W, H , R , G , B, tr, tg, tb):
    img = Image.new('RGB', (W, H), color = (R, G, B))
    d = ImageDraw.Draw(img)
    fnt = ImageFont.truetype('Supplementary/BTraffic.ttf', 100)
    d.text((10,10), enum(text), font=fnt, fill=(tr, tg, tb))
    img.save(f'Plot/{name}.png')
    
    
def texttopng(text,name):
    img = Image.new('RGB', (400, 30), color = (255, 255, 255))
    d = ImageDraw.Draw(img)
    fnt = ImageFont.truetype('Supplementary/BTraffic.ttf', 15)
    d.text((10,10), enum(text), font=fnt, fill=(0, 0, 0))
    img.save(f'Plot/{name}.png')
    
    
def Get_nav():
    web = webdriver.Chrome('chromedriver.exe')
    web.set_window_size(10, 10)
    web.get('http://etf.isatispm.com/')
    
    time.sleep(1)
    
    nav = web.find_element_by_xpath('/html/body/div[4]/div[3]/div[3]/div/span[2]').text
    nav = int(nav.replace('ریال','').replace(',',''))
    
    dt = web.find_element_by_xpath('/html/body/div[4]/div[3]/div[1]/div/span[2]').text
    dt = int(dt.replace('/',''))
    
    web.close()
    return {'date':int(dt), 'NAV':int(nav)}

def Nav_Update():
    df_nav = pd.read_excel('Data/nav histori.xlsx').drop(columns='Unnamed: 0')

    row = Get_nav()
    
    df_nav = df_nav[df_nav['date']<row['date']]

    df_nav = df_nav.append(row,ignore_index=True)
    df_nav.to_excel('Data/nav histori.xlsx')
    print('NAV Updated')

        
def gregorian_to_jalali(GDate):
    GDate = str(GDate)
    gy = int(GDate[:4])
    gm = int(GDate[4:6])
    gd = int(GDate[6:])
    g_d_m = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334]
    if (gm > 2):
        gy2 = gy + 1
    else:
        gy2 = gy
    days = 355666 + (365 * gy) + ((gy2 + 3) // 4) - ((gy2 + 99) // 100) + ((gy2 + 399) // 400) + gd + g_d_m[gm - 1]
    jy = -1595 + (33 * (days // 12053))
    days %= 12053
    jy += 4 * (days // 1461)
    days %= 1461
    if (days > 365):
        jy += (days - 1) // 365
        days = (days - 1) % 365
    if (days < 186):
        jm = 1 + (days // 31)
        jd = 1 + (days % 31)
    else:
        jm = 7 + ((days - 186) // 30)
        jd = 1 + ((days - 186) % 30)
    if len(str(jm))==1:
        jm = '0'+str(jm)
    if len(str(jd))==1:
        jd = '0'+str(jd)
    JDate = str(jy)+str(jm)+str(jd)
    JDate = int(JDate)
    return JDate

def histori_Update():
    dff = pd.read_excel('Data/histori.xlsx').drop(columns='Unnamed: 0')
    Last_update = dff['<DTYYYYMMDD>'].max()
    if Last_update>15000000:
        dff['<DTYYYYMMDD>'] = [gregorian_to_jalali(x) for x in dff['<DTYYYYMMDD>']]
    dff = dff[dff['<DTYYYYMMDD>']<dff['<DTYYYYMMDD>'].max()]
    Last_update = dff['<DTYYYYMMDD>'].max()
     
    df = StringIO(requests.get(url='http://www.tsetmc.com/tsev2/data/Export-txt.aspx?t=i&a=1&b=0&i=18865325633315847').text)
    df = pd.read_csv(df,sep=",")
    df['<DTYYYYMMDD>'] = [gregorian_to_jalali(x) for x in df['<DTYYYYMMDD>']]
    df = df[df['<DTYYYYMMDD>']>Last_update]
    if len(df)>0:
        dff = dff.append(df)

    dff.to_excel('Data/histori.xlsx')
    print('histori Update')

def histori_Update_Kara():
    dff = pd.read_excel('Data/kara.xlsx').drop(columns='Unnamed: 0')
    Last_update = dff['<DTYYYYMMDD>'].max()
    if Last_update>15000000:
        dff['<DTYYYYMMDD>'] = [gregorian_to_jalali(x) for x in dff['<DTYYYYMMDD>']]
    dff = dff[dff['<DTYYYYMMDD>']<dff['<DTYYYYMMDD>'].max()]
    Last_update = dff['<DTYYYYMMDD>'].max()
     
    df = StringIO(requests.get(url='http://www.tsetmc.com/tsev2/data/Export-txt.aspx?t=i&a=1&b=0&i=71843282162462661').text)
    df = pd.read_csv(df,sep=",")
    df['<DTYYYYMMDD>'] = [gregorian_to_jalali(x) for x in df['<DTYYYYMMDD>']]
    df = df[df['<DTYYYYMMDD>']>Last_update]
    if len(df)>0:
        dff = dff.append(df)
    
    dff.to_excel('Data/Kara.xlsx')
    print('Kara histori Update')

def histori_Update_Etemad():
    dff = pd.read_excel('Data/Etemad.xlsx').drop(columns='Unnamed: 0')
    Last_update = dff['<DTYYYYMMDD>'].max()
    if Last_update>15000000:
        dff['<DTYYYYMMDD>'] = [gregorian_to_jalali(x) for x in dff['<DTYYYYMMDD>']]
        
    dff = dff[dff['<DTYYYYMMDD>']<dff['<DTYYYYMMDD>'].max()]
    Last_update = dff['<DTYYYYMMDD>'].max()
     
    df = StringIO(requests.get(url='http://www.tsetmc.com/tsev2/data/Export-txt.aspx?t=i&a=1&b=0&i=66818022341772870').text)
    df = pd.read_csv(df,sep=",")
    df['<DTYYYYMMDD>'] = [gregorian_to_jalali(x) for x in df['<DTYYYYMMDD>']]
    df = df[df['<DTYYYYMMDD>']>Last_update]
    if len(df)>0:
        dff = dff.append(df)
    
    dff.to_excel('Data/Etemad.xlsx')
    print('Etemad histori Update')

    
def histori_Update_Kamand():
    dff = pd.read_excel('Data/Kamand.xlsx').drop(columns='Unnamed: 0')
    Last_update = dff['<DTYYYYMMDD>'].max()
    if Last_update>15000000:
        dff['<DTYYYYMMDD>'] = [gregorian_to_jalali(x) for x in dff['<DTYYYYMMDD>']]
    dff = dff[dff['<DTYYYYMMDD>']<dff['<DTYYYYMMDD>'].max()]
    Last_update = dff['<DTYYYYMMDD>'].max()
     
    df = StringIO(requests.get(url='http://www.tsetmc.com/tsev2/data/Export-txt.aspx?t=i&a=1&b=0&i=34718633636164421').text)
    df = pd.read_csv(df,sep=",")
    df['<DTYYYYMMDD>'] = [gregorian_to_jalali(x) for x in df['<DTYYYYMMDD>']]
    df = df[df['<DTYYYYMMDD>']>Last_update]
    if len(df)>0:
        dff = dff.append(df)
    
    dff.to_excel('Data/Kamand.xlsx')
    print('Kamand histori Update')

def Histor_update_all():
    histori_Update()
    histori_Update_Kara()
    histori_Update_Etemad()
    histori_Update_Kamand()
    
def Plot_Nav_Close():
    nav =pd.read_excel('Data/nav histori.xlsx').drop(columns='Unnamed: 0').set_index('date')
    df = pd.read_excel('Data/histori.xlsx').drop(columns='Unnamed: 0')[['<DTYYYYMMDD>','<CLOSE>']].set_index('<DTYYYYMMDD>')
    df = df.join(nav)
    df = df.sort_index(ascending=False).reset_index()
    df = df[df.index<=60]
    df = df.sort_index(ascending=False)
    
    figure(figsize=(15, 8), dpi=100)
    x = [str(x)[:4]+'/'+str(x)[4:6]+'/'+str(x)[6:] for x in list(df['<DTYYYYMMDD>'])]
    plt.plot(x, list(df['NAV']),label='NAV' ,linestyle=':', color='#a863e0')
    plt.legend()
    plt.plot(x, list(df['<CLOSE>']),label='Close' ,linestyle='dashed', color='#fc3d3d')
    plt.legend()
    plt.xticks(rotation=90)
    plt.savefig("Plot/Close NAV.png", dpi=100, quality=100) 
    plt.close()
    
    
def Plot_Nav2Close():
    nav =pd.read_excel('Data/nav histori.xlsx').drop(columns='Unnamed: 0').set_index('date')
    df = pd.read_excel('Data/histori.xlsx').drop(columns='Unnamed: 0')[['<DTYYYYMMDD>','<CLOSE>']].set_index('<DTYYYYMMDD>')
    df = df.join(nav)
    df = df.sort_index(ascending=False).reset_index()
    df = df[df.index<=60]
    df = df.sort_index(ascending=False)
    df['rate'] = ((df['<CLOSE>'] / df['NAV'])-1)*10000
    df = df[df['rate']> -9999999]
    df['rate'] = [round(x) for x in df['rate']]
    df['rate'] = df['rate']/100
    figure(figsize=(15, 8), dpi=100)
    x = [str(x)[:4]+'/'+str(x)[4:6]+'/'+str(x)[6:] for x in list(df['<DTYYYYMMDD>'])]
    plt.plot(x, list(df['rate']),label='Close/NAV %' ,linestyle='dashed', color='#a863e0')
    plt.legend()
    plt.xticks(rotation=90)
    plt.savefig("Plot/Close2NAV.png", dpi=100, quality=100) 
    plt.close()
    
def Plot_Volume():
    df = pd.read_excel('Data/histori.xlsx').drop(columns='Unnamed: 0')[['<DTYYYYMMDD>','<VOL>']].set_index('<DTYYYYMMDD>')
    df = df.sort_index(ascending=False).reset_index()
    df = df[df.index<=60]
    df = df.sort_index(ascending=False)
    figure(figsize=(15, 8), dpi=100)
    x = [str(x)[:4]+'/'+str(x)[4:6]+'/'+str(x)[6:] for x in list(df['<DTYYYYMMDD>'])]
    plt.plot(x, list(df['<VOL>']),label='Volume' ,linestyle='dashed', color='#a863e0')
    plt.legend()
    plt.xticks(rotation=90)
    plt.savefig("Plot/Volume.png", dpi=100, quality=100) 
    plt.close()


def Plot_date():
    df = pd.read_excel('Data/histori.xlsx')['<DTYYYYMMDD>'].max()
    df = str(df)[:4]+'/'+str(df)[4:6]+'/'+str(df)[6:]
    return df




def Table_Roa_collation():
    dl = pd.DataFrame(index=(Day_list()))
    khtm = pd.read_excel('Data/histori.xlsx').drop(columns='Unnamed: 0').set_index('<DTYYYYMMDD>')[['<CLOSE>']]
    etmd = pd.read_excel('Data/Etemad.xlsx').drop(columns='Unnamed: 0').set_index('<DTYYYYMMDD>')[['<CLOSE>']]
    kmnd = pd.read_excel('Data/Kamand.xlsx').drop(columns='Unnamed: 0').set_index('<DTYYYYMMDD>')[['<CLOSE>']]
    kara = pd.read_excel('Data/Kara.xlsx').drop(columns='Unnamed: 0').set_index('<DTYYYYMMDD>')[['<CLOSE>']]
    
    df = dl.join(khtm).join(etmd).join(kmnd).join(kara).sort_index(ascending=False)
    df['<CLOSE>']= df['<CLOSE>'].fillna(method='bfill')
    # df = df.dropna()
    df = df[df.index<=pd.read_excel('Data/histori.xlsx')['<DTYYYYMMDD>'].max()].sort_index(ascending=True)
    
    df['roc1'] = ((df['<CLOSE>']/df['<CLOSE>'].shift(1))-1)*10000
    df['roc7'] = ((df['<CLOSE>']/df['<CLOSE>'].shift(7))-1)*10000
    df['roc14'] = ((df['<CLOSE>']/df['<CLOSE>'].shift(14))-1)*10000
    df['roc30'] = ((df['<CLOSE>']/df['<CLOSE>'].shift(30))-1)*10000
    df['roc90'] = ((df['<CLOSE>']/df['<CLOSE>'].shift(90))-1)*10000
    df['roc180'] = ((df['<CLOSE>']/df['<CLOSE>'].shift(180))-1)*10000
    df['roc365'] = ((df['<CLOSE>']/df['<CLOSE>'].shift(365))-1)*10000
    df = df.dropna()
    df['roc1'] = [round(x)/100 for x in df['roc1']]
    df['roc7'] = [round(x)/100 for x in df['roc7']]
    df['roc14'] = [round(x)/100 for x in df['roc14']]
    df['roc30'] = [round(x)/100 for x in df['roc30']]
    df['roc90'] = [round(x)/100 for x in df['roc90']]
    df['roc180'] = [round(x)/100 for x in df['roc180']]
    df['roc365'] = [round(x)/100 for x in df['roc365']]
    df = df[df.index==df.index.max()].drop(columns='<CLOSE>')
    dt = {}
    for i in range(len(df.index)):
        d = list(df.index)[i]
        b = 'today'
        l = list(df[df.index==d].iloc[0])
        dt.update({b:l})
    df.columns = [enum('روزانه'),enum('7روزه'),enum('14روزه'),enum('30روزه'),enum('90روزه'),enum('180روزه'),enum('365روزه')]
    
    dff = pd.DataFrame(index=(df.columns),data = dt)
    target = ((1+0.204)**(1/365))-1
    dff['target'] = [target, (((target+1)**7)-1), (((target+1)**14)-1), (((target+1)**30)-1), (((target+1)**90)-1), (((target+1)**180)-1), (((target+1)**365)-1)]
    dff['target'] = [round(x*10000)/100 for x in dff['target']]
    
    dff['diff'] = dff['today'] - dff['target']
    dff['diff'] = [round(x*100)/100 for x in dff['diff']]
    dff['qqe'] = [1,7,14,30,90,180,365]
    dff['ydm'] =  (((dff['today']/100) +1) ** (365/dff['qqe']))-1
    dff['ydm'] = [round(x*10000)/100 for x in dff['ydm']]
    dff['today'] = [str(x)+'%' for x in dff['today']]
    dff['target'] = [str(x)+'%' for x in dff['target']]
    dff['diff'] = [str(x)+'%' for x in dff['diff']]
    dff['ydm'] = [str(x)+'%' for x in dff['ydm']]
    dff = dff.drop(columns='qqe')

    dff.columns = [enum('نقطه به نطقه'),enum('هدف'),enum('اختلاف'),enum('سالانه شده')]
    dff = dff.reset_index()
    clmn = list(dff.columns)
    clmn = [clmn[0],clmn[2],clmn[4],clmn[1],clmn[3]]
    dff = dff.reindex(columns=clmn)
    fig,ax = render_mpl_table(dff, header_columns=0, col_width=3)
    fig.savefig("Plot/ROA.png")  


def Roa_All_point():
    dl = pd.DataFrame(index=(Day_list()))
    khtm = pd.read_excel('Data/histori.xlsx').drop(columns='Unnamed: 0').set_index('<DTYYYYMMDD>')[['<CLOSE>']]
    khtm.columns = ['khtm']
    etmd = pd.read_excel('Data/Etemad.xlsx').drop(columns='Unnamed: 0').set_index('<DTYYYYMMDD>')[['<CLOSE>']]
    etmd.columns = ['etmd']
    kmnd = pd.read_excel('Data/Kamand.xlsx').drop(columns='Unnamed: 0').set_index('<DTYYYYMMDD>')[['<CLOSE>']]
    kmnd.columns = ['kmnd']
    kmnd = kmnd[kmnd.index>13990900]
    kmnd = kmnd.sort_index()
    kmnd['c'] = kmnd['kmnd'].diff()
    kmnd['kmnd'] = kmnd['c'] / kmnd['kmnd']
    kmnd['kmnd'] = ((kmnd['kmnd']>0) * kmnd['kmnd'])+1
    kmnd = kmnd[['kmnd']]
    kara = pd.read_excel('Data/Kara.xlsx').drop(columns='Unnamed: 0').set_index('<DTYYYYMMDD>')[['<CLOSE>']]
    kara.columns = ['kara']
    
    df = dl.join(khtm).join(etmd).join(kara).sort_index(ascending=False)
    df= df.fillna(method='bfill')
    df = df.join(kmnd)
    df= df.fillna(1)
    # df = df.dropna()
    df = df[df.index<=pd.read_excel('Data/histori.xlsx')['<DTYYYYMMDD>'].max()].sort_index(ascending=True)
    
    df['roc1_Khtm'] = ((df['khtm']/df['khtm'].shift(1))-1)*10000
    df['roc7_khtm'] = ((df['khtm']/df['khtm'].shift(7))-1)*10000
    df['roc14_khtm'] = ((df['khtm']/df['khtm'].shift(14))-1)*10000
    df['roc30_khtm'] = ((df['khtm']/df['khtm'].shift(30))-1)*10000
    df['roc90_khtm'] = ((df['khtm']/df['khtm'].shift(90))-1)*10000
    df['roc180_khtm'] = ((df['khtm']/df['khtm'].shift(180))-1)*10000
    df['roc365_khtm'] = ((df['khtm']/df['khtm'].shift(365))-1)*10000
    
    df['roc1_etmd'] = ((df['etmd']/df['etmd'].shift(1))-1)*10000
    df['roc7_etmd'] = ((df['etmd']/df['etmd'].shift(7))-1)*10000
    df['roc14_etmd'] = ((df['etmd']/df['etmd'].shift(14))-1)*10000
    df['roc30_etmd'] = ((df['etmd']/df['etmd'].shift(30))-1)*10000
    df['roc90_etmd'] = ((df['etmd']/df['etmd'].shift(90))-1)*10000
    df['roc180_etmd'] = ((df['etmd']/df['etmd'].shift(180))-1)*10000
    df['roc365_etmd'] = ((df['etmd']/df['etmd'].shift(365))-1)*10000
    
    df['roc1_kara'] = ((df['kara']/df['kara'].shift(1))-1)*10000
    df['roc7_kara'] = ((df['kara']/df['kara'].shift(7))-1)*10000
    df['roc14_kara'] = ((df['kara']/df['kara'].shift(14))-1)*10000
    df['roc30_kara'] = ((df['kara']/df['kara'].shift(30))-1)*10000
    df['roc90_kara'] = ((df['kara']/df['kara'].shift(90))-1)*10000
    df['roc180_kara'] = ((df['kara']/df['kara'].shift(180))-1)*10000
    df['roc365_kara'] = ((df['kara']/df['kara'].shift(365))-1)*10000
    
    df = df.fillna(0)
    df['roc1_Khtm'] = [round(x)/100 for x in df['roc1_Khtm']]
    df['roc7_khtm'] = [round(x)/100 for x in df['roc7_khtm']]
    df['roc14_khtm'] = [round(x)/100 for x in df['roc14_khtm']]
    df['roc30_khtm'] = [round(x)/100 for x in df['roc30_khtm']]
    df['roc90_khtm'] = [round(x)/100 for x in df['roc90_khtm']]
    df['roc180_khtm'] = [round(x)/100 for x in df['roc180_khtm']]
    df['roc365_khtm'] = [round(x)/100 for x in df['roc365_khtm']]
    
    df['roc1_etmd'] = [round(x)/100 for x in df['roc1_etmd']]
    df['roc7_etmd'] = [round(x)/100 for x in df['roc7_etmd']]
    df['roc14_etmd'] = [round(x)/100 for x in df['roc14_etmd']]
    df['roc30_etmd'] = [round(x)/100 for x in df['roc30_etmd']]
    df['roc90_etmd'] = [round(x)/100 for x in df['roc90_etmd']]
    df['roc180_etmd'] = [round(x)/100 for x in df['roc180_etmd']]
    df['roc365_etmd'] = [round(x)/100 for x in df['roc365_etmd']]
    
    df['roc1_kara'] = [round(x)/100 for x in df['roc1_kara']]
    df['roc7_kara'] = [round(x)/100 for x in df['roc7_kara']]
    df['roc14_kara'] = [round(x)/100 for x in df['roc14_kara']]
    df['roc30_kara'] = [round(x)/100 for x in df['roc30_kara']]
    df['roc90_kara'] = [round(x)/100 for x in df['roc90_kara']]
    df['roc180_kara'] = [round(x)/100 for x in df['roc180_kara']]
    df['roc365_kara'] = [round(x)/100 for x in df['roc365_kara']]
    
    kmnd = {}
    for i in [1,7,14,30,90,180,365]:
        gf = df.reset_index()
        gf = gf[gf.index>gf.index.max()-i]
        gf = gf['kmnd'].cumprod()
        gf = round((list(gf[gf.index==gf.index.max()])[0]-1)*10000)/100
        kmnd.update({f'roc{i}_kmnd':gf})
        
        
    last = df[df.index==df.index.max()].reset_index()

    h = {'khtm':[last['roc1_Khtm'][0],last['roc7_khtm'][0],last['roc14_khtm'][0],last['roc30_khtm'][0],last['roc90_khtm'][0],last['roc180_khtm'][0],last['roc365_khtm'][0]],
         'etmd':[last['roc1_etmd'][0],last['roc7_etmd'][0],last['roc14_etmd'][0],last['roc30_etmd'][0],last['roc90_etmd'][0],last['roc180_etmd'][0],last['roc365_etmd'][0]],
         'kara':[last['roc1_kara'][0],last['roc7_kara'][0],last['roc14_kara'][0],last['roc30_kara'][0],last['roc90_kara'][0],last['roc180_kara'][0],last['roc365_kara'][0]],
         'kmnd':[kmnd['roc1_kmnd'],kmnd['roc7_kmnd'],kmnd['roc14_kmnd'],kmnd['roc30_kmnd'],kmnd['roc90_kmnd'],kmnd['roc180_kmnd'],kmnd['roc365_kmnd']]
         }
    
    df = pd.DataFrame(index=(['r1','r7','r14','r30','r90','r180','r365']),data=h)
    df = df.where(df <100 , "-")
    for i in df.columns:
        df[i] = [str(x)+' %' for x in df[i]]
    df = df.replace('- %', '-')
    
    df.columns = [enum('خاتم'),enum('اعتماد'),enum('کارا'),enum('کمند')]
    df.index = [enum('روزانه'),enum('7 روزه'),enum('14 روزه'),enum('30 روزه'),enum('90 روزه'),enum('180 روزه'),enum('365 روزه')]
    df = df.reset_index()
    fig,ax = render_mpl_table(df, header_columns=0, col_width=3)
    fig.savefig("Plot/ROA_point_all.png")  

def Roa_All_point_YDM():
    dl = pd.DataFrame(index=(Day_list()))
    khtm = pd.read_excel('Data/histori.xlsx').drop(columns='Unnamed: 0').set_index('<DTYYYYMMDD>')[['<CLOSE>']]
    khtm.columns = ['khtm']
    etmd = pd.read_excel('Data/Etemad.xlsx').drop(columns='Unnamed: 0').set_index('<DTYYYYMMDD>')[['<CLOSE>']]
    etmd.columns = ['etmd']
    kmnd = pd.read_excel('Data/Kamand.xlsx').drop(columns='Unnamed: 0').set_index('<DTYYYYMMDD>')[['<CLOSE>']]
    kmnd.columns = ['kmnd']
    kmnd = kmnd[kmnd.index>13990900]
    kmnd = kmnd.sort_index()
    kmnd['c'] = kmnd['kmnd'].diff()
    kmnd['kmnd'] = kmnd['c'] / kmnd['kmnd']
    kmnd['kmnd'] = ((kmnd['kmnd']>0) * kmnd['kmnd'])+1
    kmnd = kmnd[['kmnd']]
    kara = pd.read_excel('Data/Kara.xlsx').drop(columns='Unnamed: 0').set_index('<DTYYYYMMDD>')[['<CLOSE>']]
    kara.columns = ['kara']
    
    df = dl.join(khtm).join(etmd).join(kara).sort_index(ascending=False)
    df= df.fillna(method='bfill')
    df = df.join(kmnd)
    df= df.fillna(1)
    # df = df.dropna()
    df = df[df.index<=pd.read_excel('Data/histori.xlsx')['<DTYYYYMMDD>'].max()].sort_index(ascending=True)
    
    df['roc1_Khtm'] = (((df['khtm']/df['khtm'].shift(1))**(365/1))-1)*10000
    df['roc7_khtm'] = (((df['khtm']/df['khtm'].shift(7))**(365/7))-1)*10000
    df['roc14_khtm'] = (((df['khtm']/df['khtm'].shift(14))**(365/14))-1)*10000
    df['roc30_khtm'] = (((df['khtm']/df['khtm'].shift(30))**(365/30))-1)*10000
    df['roc90_khtm'] = (((df['khtm']/df['khtm'].shift(90))**(365/90))-1)*10000
    df['roc180_khtm'] = (((df['khtm']/df['khtm'].shift(180))**(365/180))-1)*10000
    df['roc365_khtm'] = (((df['khtm']/df['khtm'].shift(365))**(365/365))-1)*10000
    
    df['roc1_etmd'] = (((df['etmd']/df['etmd'].shift(1))**(365/1))-1)*10000
    df['roc7_etmd'] = (((df['etmd']/df['etmd'].shift(7))**(365/7))-1)*10000
    df['roc14_etmd'] = (((df['etmd']/df['etmd'].shift(14))**(365/14))-1)*10000
    df['roc30_etmd'] = (((df['etmd']/df['etmd'].shift(30))**(365/30))-1)*10000
    df['roc90_etmd'] = (((df['etmd']/df['etmd'].shift(90))**(365/90))-1)*10000
    df['roc180_etmd'] = (((df['etmd']/df['etmd'].shift(180))**(365/180))-1)*10000
    df['roc365_etmd'] = (((df['etmd']/df['etmd'].shift(365))**(365/365))-1)*10000
    
    df['roc1_kara'] = (((df['kara']/df['kara'].shift(1))**(365/1))-1)*10000
    df['roc7_kara'] = (((df['kara']/df['kara'].shift(7))**(365/7))-1)*10000
    df['roc14_kara'] = (((df['kara']/df['kara'].shift(14))**(365/14))-1)*10000
    df['roc30_kara'] = (((df['kara']/df['kara'].shift(30))**(365/30))-1)*10000
    df['roc90_kara'] = (((df['kara']/df['kara'].shift(90))**(365/90))-1)*10000
    df['roc180_kara'] = (((df['kara']/df['kara'].shift(180))**(365/180))-1)*10000
    df['roc365_kara'] = (((df['kara']/df['kara'].shift(365))**(365/365))-1)*10000
    
    df = df.fillna(0)
    df = df.replace(np.inf,0)
    
    df['roc1_Khtm'] = [round(x)/100 for x in df['roc1_Khtm']]
    df['roc7_khtm'] = [round(x)/100 for x in df['roc7_khtm']]
    df['roc14_khtm'] = [round(x)/100 for x in df['roc14_khtm']]
    df['roc30_khtm'] = [round(x)/100 for x in df['roc30_khtm']]
    df['roc90_khtm'] = [round(x)/100 for x in df['roc90_khtm']]
    df['roc180_khtm'] = [round(x)/100 for x in df['roc180_khtm']]
    df['roc365_khtm'] = [round(x)/100 for x in df['roc365_khtm']]
    
    df['roc1_etmd'] = [round(x)/100 for x in df['roc1_etmd']]
    df['roc7_etmd'] = [round(x)/100 for x in df['roc7_etmd']]
    df['roc14_etmd'] = [round(x)/100 for x in df['roc14_etmd']]
    df['roc30_etmd'] = [round(x)/100 for x in df['roc30_etmd']]
    df['roc90_etmd'] = [round(x)/100 for x in df['roc90_etmd']]
    df['roc180_etmd'] = [round(x)/100 for x in df['roc180_etmd']]
    df['roc365_etmd'] = [round(x)/100 for x in df['roc365_etmd']]
    
    df['roc1_kara'] = [round(x)/100 for x in df['roc1_kara']]
    df['roc7_kara'] = [round(x)/100 for x in df['roc7_kara']]
    df['roc14_kara'] = [round(x)/100 for x in df['roc14_kara']]
    df['roc30_kara'] = [round(x)/100 for x in df['roc30_kara']]
    df['roc90_kara'] = [round(x)/100 for x in df['roc90_kara']]
    df['roc180_kara'] = [round(x)/100 for x in df['roc180_kara']]
    df['roc365_kara'] = [round(x)/100 for x in df['roc365_kara']]
    
    kmnd = {}
    for i in [1,7,14,30,90,180,365]:
        gf = df.reset_index()
        gf = gf[gf.index>gf.index.max()-i]
        gf = gf['kmnd'].cumprod()**(365/i)
        gf = round((list(gf[gf.index==gf.index.max()])[0]-1)*10000)/100
        kmnd.update({f'roc{i}_kmnd':gf})
        
        
    last = df[df.index==df.index.max()].reset_index()

    h = {'khtm':[last['roc1_Khtm'][0],last['roc7_khtm'][0],last['roc14_khtm'][0],last['roc30_khtm'][0],last['roc90_khtm'][0],last['roc180_khtm'][0],last['roc365_khtm'][0]],
         'etmd':[last['roc1_etmd'][0],last['roc7_etmd'][0],last['roc14_etmd'][0],last['roc30_etmd'][0],last['roc90_etmd'][0],last['roc180_etmd'][0],last['roc365_etmd'][0]],
         'kara':[last['roc1_kara'][0],last['roc7_kara'][0],last['roc14_kara'][0],last['roc30_kara'][0],last['roc90_kara'][0],last['roc180_kara'][0],last['roc365_kara'][0]],
         'kmnd':[kmnd['roc1_kmnd'],kmnd['roc7_kmnd'],kmnd['roc14_kmnd'],kmnd['roc30_kmnd'],kmnd['roc90_kmnd'],kmnd['roc180_kmnd'],kmnd['roc365_kmnd']]
         }
    
    df = pd.DataFrame(index=(['r1','r7','r14','r30','r90','r180','r365']),data=h)
    df = df.where(df <100 , "-")
    for i in df.columns:
        df[i] = [str(x)+' %' for x in df[i]]
    df = df.replace('- %', '-')
    
    df.columns = [enum('خاتم'),enum('اعتماد'),enum('کارا'),enum('کمند')]
    df.index = [enum('روزانه'),enum('7 روزه'),enum('14 روزه'),enum('30 روزه'),enum('90 روزه'),enum('180 روزه'),enum('365 روزه')]
    df = df.reset_index()
    fig,ax = render_mpl_table(df, header_columns=0, col_width=3)
    fig.savefig("Plot/ROA_point_all_YDM.png")  


def Update_Rezer():
    df = pd.read_excel('data/rezerv.xlsx').drop(columns='Unnamed: 0')
    today = pd.read_excel('data/nav histori.xlsx')['date'].max()
    df = df[df['date']<today]
    dic = {}
    web = webdriver.Chrome('chromedriver.exe')
    web.set_window_size(200, 400)
    url = {'khtm':'http://www.tsetmc.com/Loader.aspx?ParTree=151311&i=18865325633315847',
           'etmd':'http://www.tsetmc.com/Loader.aspx?ParTree=151311&i=66818022341772870',
           'kara':'http://www.tsetmc.com/Loader.aspx?ParTree=151311&i=71843282162462661',
           'kmnd':'http://www.tsetmc.com/Loader.aspx?ParTree=151311&i=34718633636164421'} 
    for u in url:
        web.get(url[u])
        time.sleep(2)
        web.find_element_by_xpath('//*[@id="tabs"]/div/ul/li[9]/a').click()
        time.sleep(4)
        rezerv = web.find_element_by_xpath('//*[@id="PureData"]/div[2]/div[2]/table/tbody').get_attribute('innerHTML').split('</tr>')
        for i in rezerv:
            if 'رزرو' in i:
                rezerv1 = i
    
        try:rezerv1 = float(rezerv1.split('<td>')[3].split('</td>')[0])
        except:rezerv1=0
        dic.update({u:rezerv1})
        print({u:rezerv1})
    dic.update({'date':today})
    web.close()
    df = pd.concat([df, pd.DataFrame.from_records([dic])], ignore_index=True)
    df.to_excel('data/rezerv.xlsx')




def Rezerv_Code_Plot():
    df = pd.read_excel('data/rezerv.xlsx').drop(columns='Unnamed: 0')
    df = df.set_index('date')
    df = df.sort_index(ascending=False).reset_index()
    df = df[df.index<=60]
    df = df.sort_index(ascending=False)
    figure(figsize=(15, 8), dpi=100)
    x = [str(x)[4:6]+'/'+str(x)[6:] for x in list(df['date'])]
    plt.plot(x, list(df['khtm']),label=enum('خاتم') , color='#001BFF',linewidth=2.0)
    plt.legend()
    plt.plot(x,list(df['etmd']),label=enum('اعتماد') , color='#FF0000',linewidth=2.0)
    plt.legend()
    plt.plot(x, list(df['kara']),label=enum('کارا') , color='#CCDA00',linewidth=2.0)
    plt.legend()
    plt.plot(x,list(df['kmnd']),label=enum('کمند') , color='#11B611',linewidth=2.0)
    plt.legend()
    plt.xlabel(enum('تاریخ'))
    plt.ylabel(enum('% درصد'))
    plt.xticks(rotation=35)
    plt.savefig("Plot/Rezerv.png", dpi=100, quality=100) 
    plt.close()













