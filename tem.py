import xlwings as xl
import datetime as d
timenow=d.datetime.now()
filedate=str(timenow.year)+str(timenow.month).zfill(2)+str(timenow.day).zfill(2)#利用如果日期是个位数补零
filename="21级6班每日健康监测"+filedate+'.xlsx'
#----------固定数据区域---------------------------------------------------
studentnames=['陈睿涵','陈重澔','初怡萍','丁浩文','窦可依','付锦杰','高斌','高嘉程','葛昊文','姜佳琪','姜俊彦','姜沛仪','孔庆杰','孔泽成','李赫','李书寒','李翔宇','李宜颖','林东晓','刘新玉','刘禹廷','马灵儿','马明宇','乾欣悦','曲倩','曲彦柯','孙佳怡','孙声赫','陶亚楠','王琳','王馨平','王宜栋','王媛媛','辛佳奇','修雨涵','徐承志','杨玉磊','于春皓','于享鹭','余涵','元誉锦','张国辉','张皓阳','张文','赵世林','郑凯元','周运霞','邹浩然','邹梦瑜','邹心瑶']
titlenames1=['序号','姓名','体温℃\n7：00','体温℃\n12：00','体温℃\n21：00','本人及共同居住人员是否有发热、干咳、乏力、咽痛、嗅（味）觉减退、腹泻等症','异常处置措','本人及共同居住人员是否为阳性或混管阳性']
examples01=['','张正常','36.1','36.4','36.6','无','','否']
examples02=['','李异常','36.7','36.4','36.5','父亲有发热情况','到XX医院就诊','是']
titlenames2=['A','B','C','D','E','F','G','H']
#-----------------------------------------------------------------------
app=xl.App(visible=True,add_book=False)
wb=app.books.add()
ws=wb.sheets['sheet1']
#-----------------------------------------------------------------------

#-----------完成第1行表格------------------------------------------------
string01='2021级6班学生每日身体情况统计表'
ws.range('A1:H1').api.merge
ws.range('A1').value=string01
ws.range('A1:H54').api.HorizontalAlignment=-4108 #所有单元格居中
ws.range('f5:f54').value='无'
ws.range('h5:h54').value='否'
ws.range('C5:E54').formula='=36+round(rand(),1)'
#-----------------------------------------------------------------------
#-----------完成第2-4行表格-----------------------------------------------
ws.range('F4').column_width=33
ws.range('G4').column_width=13
ws.range('F2','H2').api.WrapText=True
ws.range('F2','H2').api.Font.Size=8
ws.range("A3:H4").api.Font.Color=0x0000FF #第三行和第四行字体为红色

for i in range(2,5):
	for j in range(1,9):
		rangename=titlenames2[j-1]+str(i)
		if i==2:			
			ws.range(rangename).value=titlenames1[j-1]
		if i==3:			
			ws.range(rangename).value=examples01[j-1]
		if i==4:			
			ws.range(rangename).value=examples02[j-1]
#-----------------------------------------------------------------------
#-----------完成第5-54行表格----------------------------------------------
for i in range(5,55):
	ws.range('A'+str(i)).value=i-4  #填充序号
	ws.range('B'+str(i)).value=studentnames[i-5] #填充学生姓名
	for j in range(2,5):
		ws.range(titlenames2[j]+str(i)).value=ws.range(titlenames2[j]+str(i)).value 
	#将体温处计算得到的数值覆盖公式
wb.save(filename)
wb.close()
app.quit()
