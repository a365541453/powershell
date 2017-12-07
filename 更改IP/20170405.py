# -*- coding:utf-8 -*- 

import telnetlib
import time
import re
import xlsxwriter
import os
import xlrd
import tkFileDialog
import subprocess

def get_net_name(old_ip,port,name,password):
	network_name = []
	host = telnetlib.Telnet(old_ip,port)
	host.set_debuglevel(2)
	time.sleep(10)
	host.read_until(b'login:',20)
	host.write(name + '\r')
	host.read_until(b'password',20)
	host.write(password + '\r')
	host.read_until('>',10)
	host.write(b'netsh interface show interface' + '\r') 
	time.sleep(5)
	back = host.read_very_eager().split('\r\n')
	while '' in back:
		back.remove('')
	for row in back:
		a = u'已启用'
		str1 = a.encode('gb2312')
		resu = re.match(r'^%s'%str1,row)
		if resu:
			row = re.sub(r'(  )+',r'/',row)#两个空格替换成一个‘/’，方便分割文本 
			name = row.split('/')[3]
			name = re.sub(r'^( )+',r'',name)#网卡名称前面可能会有一个空格，把这个空格去掉
			network_name.append(name)
		host.close()
	return network_name

def set_ip(old_ip,port,name,password,net_name,new_ip,netmask,gateway):
	host = telnetlib.Telnet(old_ip,port)
	host.set_debuglevel(1)
	time.sleep(5)
	host.read_until(b'login:',20)
	host.write(name + '\r')
	host.read_until(b'password',20)
	host.write(password + '\r')
	host.read_until('>',10)
	time.sleep(5)
	net_name = net_name.decode("gb2312")
	net_name = bytearray(net_name,encoding='gbk')
	net_name = 'netsh interface ip set address \"%s\" static %s %s %s 1'%(net_name,new_ip,netmask,gateway)
	host.read_until(b'>',10)
	host.write(net_name + '\r')
	host.read_until(b'>',10)
	net_name = 'netsh interface ipv4 set address \"%s\" static %s %s %s 1'%(net_name,new_ip,netmask,gateway)
	host.read_until(b'>',10)
	host.write(net_name + '\r')
	host.read_until(b'>',10)
	time.sleep(5)
	host.close()

def wait_restart(old_ip,port):
	restart = 0
	a = 0
	while restart == 0:
		try:
			host = telnetlib.Telnet(old_ip,port)
			host.close()
			restart = 1
		except Exception, e:
			restart = 0

def up_all_network(old_ip,port,name,password):
	host = telnetlib.Telnet(old_ip,port)
	host.set_debuglevel(2)
	time.sleep(10)
	host.read_until(b'login:',20)
	host.write(name + '\r')
	host.read_until(b'password',20)
	host.write(password + '\r')
	host.read_until('>',10)
	host.write(b'netsh interface show interface' + '\r') 
	time.sleep(5)
	back = host.read_very_eager().split('\r\n')
	for row in back:
		a = u'已禁用'
		str1 = a.encode('gb2312')
		resu = re.match(r'^%s'%str1,row)
		if resu:
			row = re.sub(r'(  )+',r'/',row)#两个空格替换成一个‘/’，方便分割文本 
			name = row.split('/')[3]
			name = re.sub(r'^( )+',r'',name)#网卡名称前面可能会有一个空格，把这个空格去掉
			name = name.decode("gb2312")
			name = bytearray(name,encoding='gbk')
			name = 'netsh interface set interface \"%s\" enabled'%name
			host.write(name + '\r')
			host.read_until(b'>',10)
			time.sleep(5)
	host.close()		

ip_excle = tkFileDialog.askopenfilename() 
ip_sheet = xlrd.open_workbook(ip_excle).sheet_by_index(0)
ip_tolnumber = ip_sheet.nrows

for r in xrange(1,ip_tolnumber):
	vm_name = ip_sheet.cell_value(r,0).encode('gbk')	#虚拟机名字
	old_ip = ip_sheet.cell_value(r,1).encode('gb2312')	#虚拟机原有IP
	new_ip = ip_sheet.cell_value(r,2).encode('gb2312')	#新添加的IP
	netmask = ip_sheet.cell_value(r,3).encode('gb2312')	#新网卡掩码
	gateway = ip_sheet.cell_value(r,4).encode('gb2312')	#新网卡网关
	name = ip_sheet.cell_value(r,5)	.encode('gb2312')	#操作系统用户名
	password = ip_sheet.cell_value(r,6).encode('gb2312')	#操作系统密码
	
	network_name_old_list = get_net_name(old_ip,23,name,password) #返回的是操作系统中正在使用的网卡名称列表

	vmware = subprocess.Popen('powershell ./add_network.ps1 -vm_name %s' %vm_name)
	vmware.wait()
	
	hahaha = tkFileDialog.askopenfilename()

	wait_restart(old_ip,23)

	up_all_network(old_ip,23,name,password)

	network_name_new_list = get_net_name(old_ip,23,name,password)#返回新想网卡列表
	
	for net_name in network_name_new_list:
		if net_name not in network_name_old_list:
			set_ip(old_ip,'23',name,password,net_name,new_ip,netmask,gateway)






