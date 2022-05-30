import fnc
import pandas as pd 
from fpdf import FPDF

fnc.Nav_Update()
fnc.Histor_update_all()
fnc.Table_Price_Nav(7)
fnc.Table_Volume(7)
fnc.Table_Roa()
fnc.Roa_All_point()
fnc.Roa_All_point_YDM()
fnc.Update_Rezer()
fnc.Rezerv_Code_Plot()

# fnc.Plot_Nav_Close()
# fnc.Plot_Nav2Close()


pdf = FPDF()
pdf.add_page()
pdf.set_font("Arial", size = 15)

fnc.texttopng_title('صندوق س خاتم ايساتيس پويا-ثابت (خاتم)','Title',1500, 120, 60,100,200,255,255,255)
pdf.image('Plot/Title.png',x=30 ,y=5 ,w=150, type = 'png')

fnc.texttopng_title(fnc.Plot_date(),'plotDate',500, 100, 60,100,200,255,255,255)
pdf.image('Plot/plotDate.png',x=90 ,y=18 ,w=30, type = 'png')

fnc.texttopng_title('گزارش هفت روز قیمت و ابطال','weekly_price_nav',1125, 120, 130,25,25,255,255,255)
pdf.image('Plot/PrcNavWek.png',x=0 ,y=31 ,w=210, type = 'png')
pdf.image('Plot/weekly_price_nav.png',x=129 ,y=27 ,w=60, type = 'png')

fnc.texttopng_title('گزارش هفت روز حجم معاملات','weekly_Vol',1125, 120, 130,25,25,255,255,255)
pdf.image('Plot/VolWek.png',x=0 ,y=79 ,w=210, type = 'png')
pdf.image('Plot/weekly_Vol.png',x=129 ,y=75 ,w=60, type = 'png')


fnc.texttopng_title('گزارش روزانه بازدهی','DailyRoa',800, 110, 130,25,25,255,255,255)
pdf.image('Plot/ROA.png',x=0 ,y=141 ,w=210, type = 'png')
pdf.image('Plot/DailyRoa.png',x=129 ,h=6.7, y=137 ,w=60, type = 'png')

fnc.texttopng_title('مقایسه بازدهی نقطه به نقطه','Pointtopoint',1070, 120, 130,25,25,255,255,255)
pdf.image('Plot/ROA_point_all.png',x=0 ,y=212 ,w=210, type = 'png')
pdf.image('Plot/Pointtopoint.png',x=119 , y=208 ,w=70, type = 'png')
fnc.texttopng_title('*صندوق کمند بدون لحاظ تاخیر در خرید مجدد از سود پرداختی و کارمزد آن محاسبه شده است','notekamand',3415, 120, 255,255,255,150,0,0)
pdf.image('Plot/notekamand.png',x=98 , y=275 ,w=90, type = 'png')

pdf.add_page()
fnc.texttopng_title('مقایسه بازدهی سالانه شده نقطه به نقطه','PointtopointYDM',1450, 120, 130,25,25,255,255,255)
pdf.image('Plot/ROA_point_all_YDM.png',x=0 ,y=15 ,w=210, type = 'png')
pdf.image('Plot/PointtopointYDM.png',x=105 , y=10 ,w=85, type = 'png')
fnc.texttopng_title('*صندوق کمند بدون لحاظ تاخیر در خرید مجدد از سود پرداختی و کارمزد آن محاسبه شده است','notekamand',3415, 120, 255,255,255,150,0,0)
pdf.image('Plot/notekamand.png',x=98 , y=77 ,w=90, type = 'png')

fnc.texttopng_title('نسبت دارایی های کد رزور صندوق ها به کل واحد ها','fundRezervCode',1870, 120, 130,25,25,255,255,255)
pdf.image("Plot/Rezerv.png",x=0 ,y=87 ,w=210, type = 'png')
pdf.image('Plot/fundRezervCode.png',x=105 , y=87 ,w=85, type = 'png')

# fnc.texttopng('قیمت و قیمت ابطال 60 روز کاری گذشته','work60dauclosenav')
# pdf.image('Plot/work60dauclosenav.png',x=70 ,y=5 ,w=100, type = 'png')
# pdf.image('Plot/Close NAV.png',x=0 ,y=15 ,w=200, type = 'png')

# fnc.texttopng('اختلاف قیمت و قیمت ابطال 60 روز کاری گذشته','Ratework60dauclosenav')
# pdf.image('Plot/Ratework60dauclosenav.png',x=70 ,y=140 ,w=100, type = 'png')
# pdf.image('Plot/Close2NAV.png',x=0 ,y=150 ,w=200, type = 'png')

# pdf.add_page()

# fnc.texttopng('حجم معاملات 60 روز کاری گذشته','volume60dauclosenav')
# pdf.image('Plot/volume60dauclosenav.png',x=5 ,y=140 ,w=100, type = 'png')
# pdf.image('Plot/Volume.png',x=0 ,y=15 ,w=200, type = 'png')

pdf.output("Khatam.pdf")