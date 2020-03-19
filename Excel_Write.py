# -*- coding: utf-8 -*-
import xlwt

#设置表格样式
def set_style(name,height,bold=False):
	style = xlwt.XFStyle()
	font = xlwt.Font()
	font.name = name
	font.bold = bold
	font.color_index = 4
	font.height = height
	style.font = font
	return style

#写Excel
def write_excel():
	f = xlwt.Workbook()
	sheet1 = f.add_sheet('学生',cell_overwrite_ok=True)
	row0 = ["姓名","年龄","出生日期","爱好"]
	colum0 = ["张三","李四","王五","小明","小红","无名"]
	colum1 = ["1","2","3","4","5","6"]
	colum2 = ["2000/07/20","2000/07/20","2000/07/20","2000/07/20","2000/07/20","2000/07/20"]
	colum3 = ["打篮球","打篮球","打篮球","打篮球","打篮球","打篮球"]
	#写第一行
	for i in range(0,len(row0)):
		sheet1.write(0,i,row0[i],set_style('Times New Roman',220,True))
	#写第一列
	for i in range(0,len(colum0)):
		sheet1.write(i+1,0,colum0[i],set_style('Times New Roman',220,True))
	#写第二列
	for i in range(0,len(colum1)):
		sheet1.write(i+1,1,colum1[i],set_style('Times New Roman',220,True))
	#写第三列
	for i in range(0,len(colum2)):
		sheet1.write(i+1,2,colum2[i],set_style('Times New Roman',220,True))
	#写第四列
	for i in range(0,len(colum3)):
		sheet1.write(i+1,3,colum3[i],set_style('Times New Roman',220,True))

	# 合并第6行到第7行的第4列到第5列，里面所有的参数都是以0开始计算的。

	sheet1.write_merge(5, 6, 3, 3, '唱Rap')

	#合并第2行第3列到第4列
	sheet1.write_merge(1,1,2,3,'未知')


	f.save('test.xls')

if __name__ == '__main__':
	write_excel()