# -*- coding: utf-8 -*-
import xlrd
import xlsxwriter as wx
from Tkinter import *
def duibi(l1,l2):#给出两个列表值，2包含1则flag=1，否则为0
	coun=0
	flag=0
	for i in l1:
		if i not in l2:
			break
		coun=coun+1
	if coun==len(l1):
		flag=1
		return flag
class jiance():
	houxuan=[]
	sheet_name1=[]
	sheet_name2=[]
	re=[]
	nn=0
	data_all=[]
	yuce=[]
	def read_data(self,path):
		data=open(path,'r')
		self.s=data.read().split('!')
		self.houxuan=self.s[0].split(';')
		del self.s[0]
		self.yuce=self.s
		self.changdu=len(self.houxuan)
		self.width=len(self.yuce)
		for i in range(self.width):
			self.yuce[i]=self.yuce[i].split(';')
	def read_data2(self,path,yuce_changdu,fangf):
		self.fangf=fangf
		workbook = xlrd.open_workbook(path)
		sheet2 = workbook.sheet_by_index(0)
		if self.fangf==0:
			self.houxuan=sheet2.col_values(2)[12:(12+78)]
		if self.fangf==1:
			self.houxuan=sheet2.col_values(2)[2:(2+78)]
		self.yuce=[0]*yuce_changdu
		for i in range(yuce_changdu):
			if self.fangf==0:
				self.yuce[i] = sheet2.col_values(i+4)[1:12]
			if self.fangf==1:
				self.yuce[i] = sheet2.col_values(i+4)[1:2]
		self.changdu=len(self.houxuan)
		self.width=len(self.yuce)
	# def read_data3(self,path,yuce_changdu,fangf):
	# 	self.fangf=fangf
	# 	workbook = load_workbook(path)
	# 	sheet2 = workbook.sheet_by_index(0)
	# 	if self.fangf==0:
	# 		self.houxuan=sheet2.col_values(2)[12:(12+78)]
	# 	if self.fangf==1:
	# 		self.houxuan=sheet2.col_values(2)[2:(2+78)]
	# 	self.yuce=[0]*yuce_changdu
	# 	for i in range(yuce_changdu):
	# 		if self.fangf==0:
	# 			self.yuce[i] = sheet2.col_values(i+4)[1:12]
	# 		if self.fangf==1:
	# 			self.yuce[i] = sheet2.col_values(i+4)[1]
	# 	self.changdu=len(self.houxuan)
	# 	self.width=len(self.yuce)
	def jiance(self):
		method=[6,8]
		method1=[9,11]
		fangfa=method[self.fangf]
		digui=method1[self.fangf]
		self.nn=self.changdu/fangfa
		self.sheet_name1=[100]*self.nn
		self.sheet_name2=[100]*self.nn
		for xx in range(self.nn):
			self.sheet_name1[xx]=1+xx*fangfa
			self.sheet_name2[xx]=self.sheet_name1[xx]+digui
			if self.sheet_name2[xx]>self.changdu:
				del self.sheet_name2[-1]
				del self.sheet_name1[-1]
				self.nn=self.nn-1
				break		
		self.re=[0]*self.nn
		self.data_all=[0]*self.width
		if self.fangf==0:
			for k in range(self.width):#循环预测列
				rec=[100]*self.changdu
				for j in range(self.changdu):#循环候选行
					l2=self.houxuan[j].split()#每一行后选变成列表形式
					for l1 in self.yuce[k]:#11列数据分别去匹配,只要该行可以即终止。for 循环自身是一个分离列表过程。
						l1=l1.split()
						flag=duibi(l1,l2)
						if flag:
							rec[j]=j+1
							break
				self.data_all[k]=rec
			for i in range(self.nn):#候选行分块取
				tmp1=self.sheet_name1[i]-1
				tmp2=self.sheet_name2[i]
				tmp=[0]*len(self.data_all)
				for j in range(len(self.data_all)):
					tmp[j]=min(self.data_all[j][tmp1:tmp2])
				self.re[i]=tmp
		if self.fangf==1:
			for k in range(self.width):#循环预测列
				rec=[100]*self.changdu
				for j in range(self.changdu):#循环候选行
					l2=self.yuce[k][0].split()
					l1=self.houxuan[j].split()
					flag=duibi(l1,l2)
					if flag:
						rec[j]=j+1						
						break
				self.data_all[k]=rec
			for i in range(self.nn):#候选行分块取
				tmp1=self.sheet_name1[i]-1
				tmp2=self.sheet_name2[i]
				tmp=[0]*len(self.data_all)
				for j in range(len(self.data_all)):
					tmp[j]=min(self.data_all[j][tmp1:tmp2])
				self.re[i]=tmp

if __name__ == '__main__':
		aa=jiance()
		# aa.read_data(r'G:\BaiduYunDownload\wtx\data111')
		fangfa=int(raw_input('type0 or 1\n'))
		changdu=raw_input('length\n')
		if fangfa==0:
			wenjian='data1-10.xlsx'
		if fangfa==1:
			wenjian='data1-12.xlsx'
		aa.read_data2(wenjian,int(changdu),fangfa)
		aa.jiance()
		# aa.read_data2('data1-10.xls',15,0)
		# write_excel(aa)
		out=[]
		f = wx.Workbook(wenjian+'_shuju.xlsx')
		for i in range(aa.nn):
			name1=aa.sheet_name1[i]
			name2=aa.sheet_name2[i]
			sheet1 = f.add_worksheet(str(name1)+'-'+str((name1+name2-1)/2)+'-'+str(name2)) #创建sheet
			col=0
			while True:
				inputs = raw_input('di'+str(i+1)+'zu or q exit:').split()
				if 'q' in inputs: 
					del inputs[-1]
					break
			for j in inputs:
				if fangfa==0:
					for k in range(len(aa.yuce[int(j)-1])):
						sheet1.write(0,col,j)
						sheet1.write(1+k,col,aa.yuce[int(j)-1][k])
					col=col+1
				if fangfa==1:
					sheet1.write(0,col,j)
					sheet1.write(1,col,aa.yuce[int(j)-1][0])
					col=col+1
		f.close()
		raw_input('press any key to exit')
		# print aa.re
		# tem111=input('press any key to exit')
#列表增值，不用赋值的形式
		# print out[1]
		# tem111=raw_input('press any key to exit')
		# tem222=raw_input('press any key to exit')		
		# out1=[[1,12,3],[2,3,4]]
		# out2=[[1,2,3],[4,5,6]]
		# out=[out1,out2]
		# print out[0]
		# def button_act():
		# 	out=[]
		# 	for j in range(2):
		# 		out1=[]
		# 		while True:
		# 			inputs = raw_input('Input no.1 numbers or q:').split()
		# 			inputs=[1,12,13,'q']
		# 			if 'q' in inputs: 
		# 				del inputs[-1]
		# 				break
		# 		for i in inputs:
		# 			out1.append(aa.yuce[int(i)])
		# 		out.append(out1)
		# 	label.config(text=(str(out[0])+'\n'+str(out[1])))
		# top=Tk()
		# label=Label(top,text='press button')
		# label.pack()
		# button=Button(top,text='hello',command=button_act)#新标签
		# button.pack()
		# mainloop()
