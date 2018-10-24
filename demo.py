from PIL import Image,ImageFont,ImageDraw
from openpyxl import load_workbook

wb = load_workbook("就业信息.xlsx")
#获取当前活跃的worksheet,默认就是第一个worksheet
#ws = wb.active

#当然也可以使用下面的方法

#获取所有表格(worksheet)的名字
sheets = wb.get_sheet_names()
#第一个表格的名称
sheet_first = sheets[0]
#获取特定的worksheet
ws = wb.get_sheet_by_name(sheet_first)

rows = ws.rows
columns = ws.columns

#迭代所有的行
index=0
data = []
for row in rows:
    if index==0:
        index+=1
        continue
    else:
        data.append([col.value for col in row])


# 字体
def change(classname,name,zhuanye,city,fuli,filename):

    im = Image.open("原版.png")
    width, height = im.size
    dw = ImageDraw.Draw(im)
    font1 = ImageFont.truetype(r'C:\Windows\Fonts\Arial.ttf',120)
    font2 = ImageFont.truetype(r'C:\Windows\Fonts\simsun.ttc',240)
    font3 = ImageFont.truetype(r'C:\Windows\Fonts\simsun.ttc',70)
    font4 = ImageFont.truetype(r'C:\Windows\Fonts\simsun.ttc',120)
    font5 = ImageFont.truetype(r'C:\Windows\Fonts\simsun.ttc',90)

    z_width,z_height = font3.getsize(zhuanye)
    c_width,c_height = font4.getsize(city)
    f_width,f_height = font5.getsize(fuli)


    dw.text((1000,1300),classname,font=font1,fill=(229,218,151))
    dw.text((1800,1200),name,font=font2,fill=(229,218,151))
    dw.text((width/2-z_width/2,1480),zhuanye,font=font3,fill=(229,218,151))
    dw.text((width/2-c_width/2,1600),city,font=font4,fill=(229,218,151))
    dw.text((width/2-f_width/2,2000),fuli,font=font5,fill=(229,218,151))
    im.save("./img/%s.png"%filename)



for item in data:
    print("正在制作%s的喜报.."%item[11])
    change(item[10],item[11][0]+"同学",str(item[16])+"-"+str(item[18])+"-"+str(item[19]),"成功面试"+item[20]+"（%s）"%item[7],item[-3],item[11])