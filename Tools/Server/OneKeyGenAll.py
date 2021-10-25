# -*- coding:utf-8 -*-
# python 2.7
import xlrd
import sys
import time
from Excel2Tars import *
import os
import csv
import codecs
import shutil
import platform
import ExcelCheck

xlsm_path = "../../Excels/"
bfile_path = "../../Output/DotB/"
csv_path = "../../Output/ServerCsv/"
tars_struct_path = "../../Output/Struct/"
t2py_path = "./"
server_genlog_str=""
gen_bfile_err_str=""
file_namelist = []

def is_suitable_region(region_value):
	if region_value != "*" and region_value != "*#" and region_value != "#*":
		return 0
	else:
		return 1

def get_xlsm_filelist(xlsm_namelist):
	file_path = os.getcwd() + "/"+xlsm_path
	dirs = os.listdir(file_path)
	for file_name in dirs:
		if file_name[0] == "~":
			continue    
		if(os.path.splitext(file_name)[1]) == ".xlsm":
			xlsm_namelist.append(file_name)
			print file_name	
	print file_path
	#exit()

def xlsm2tars(file_name):
	#xlsm文件名
	file_name =  unicode(file_name, "gbk")
	#生成的tars文件名
	tars_namespace = os.path.splitext(file_name)[0]
	tars_filename = tars_namespace + ".tars"
	
	print "xlsm2tars:"+xlsm_path+file_name
	global server_genlog_str
	server_genlog_str = server_genlog_str + "*************************************\n"
	server_genlog_str = server_genlog_str+ xlsm_path+file_name+"\n"
	excel_data = open_excel(xlsm_path+file_name)

	jce_file = open(tars_filename, "w")
	#jce_file.write("//AutoMake @ [%s]\n" %(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))))

	jce_begin_struct = []
	print_jce_module_begin(jce_begin_struct, tars_namespace, 0)
	jce_begin_str = "".join(jce_begin_struct)
	
	print jce_begin_str
	
	jce_file.write(jce_begin_str)
	

	#sheet表遍历
	sheetname_list = []
	server_genlog_str = server_genlog_str + getsheetname_list(excel_data, sheetname_list)

	print sheetname_list
	for sheet_name in sheetname_list:
		f=open("DotB2Table.log","a+")
		f.write(tars_namespace+":"+sheet_name)
		f.write('\n')
		jce_struct = []
		sheet_data = excel_data.sheet_by_name(sheet_name)	
		#导出csv
		xlssheet_to_csv(csv_path+sheet_name+".csv", sheet_data)
		
		#tars结构
		print_jce_struct_begin(jce_struct, sheet_name, 2)
    
		i = 0
		for col in range(1, sheet_data.ncols):
			if is_suitable_region(sheet_data.cell(0, col).value)==0:
				continue
			if sheet_data.cell(1, col).value.lower() == "text":
				print_jce_field(jce_struct, i, "string",
							sheet_data.cell(2, col).value, sheet_data.cell(4, col).value, 3)
			else:
				print_jce_field(jce_struct, i, sheet_data.cell(1, col).value.lower(),
							sheet_data.cell(2, col).value, sheet_data.cell(4, col).value, 3)
			i += 1
		print_jce_struct_end(jce_struct, 2)
    
		jce_str = "".join(jce_struct)
		#print jce_str
		jce_file.write(jce_str.encode("gbk"))

	jce_struct = []
	print_jce_struct_begin(jce_struct, tars_namespace+"Info", 2)
	i = 0;
	for name in sheetname_list:
		print_jce_field(jce_struct, i, "map<int, T" + name + ">", "map" + name, "", 3)
		i+= 1
	print_jce_struct_end(jce_struct, 2)
	all_cfg_jce = "".join(jce_struct)
	print all_cfg_jce
	jce_file.write(all_cfg_jce)


	jce_end_struct = []
	print_jce_module_end(jce_end_struct, 0)
	jce_end_str = "".join(jce_end_struct)
	print jce_end_str
	jce_file.write(jce_end_str.encode("gbk"))

	jce_file.close()

def xlssheet_to_csv(csv_filename, xls_table):

	with codecs.open(csv_filename, 'w', encoding='gbk') as fileobj:
		write = csv.writer(fileobj)
		for row_num in range(xls_table.nrows):
		
			write_rowvalue = []
			row_value = xls_table.row_values(row_num)
			if is_suitable_region(xls_table.cell(row_num, 0).value) == 0:

				continue

			for col in range(1, xls_table.ncols):				
				if is_suitable_region(xls_table.cell(0, col).value) == 0:
					continue
				if str(xls_table.cell(row_num, col).value) == "TEXT":
				    write_rowvalue.append("STRING")
				else:
					write_rowvalue.append(str(xls_table.cell(row_num, col).value))
			write.writerow(write_rowvalue)
		fileobj.close()
		
def getsheetname_list(excel_data, sheetname_list):
    
	sheet_str = ""
	sheet_data = excel_data.sheet_by_name("INDEX")
	for row in range(2,sheet_data.nrows):
		#INDEX页配置的sheet才会导出
		
		if is_suitable_region(sheet_data.cell(row, 0).value) == 0:
			continue
		sheetname_list.append(sheet_data.cell(row, 2).value)
		sheet_str= sheet_str+sheet_data.cell(row, 2).value + "========" + sheet_data.cell(row, 5).value + "\n"      
	sheet_str = sheet_str + "*************************************\n"
	return sheet_str	

def xlsm2BFile(file_namelist):

	
	for file_name in file_namelist:
		#xlsm文件名
		print "gen BFile for " + file_name + "...\n"
		file_name =  unicode(file_name, "gbk")
		tars_namespace = os.path.splitext(file_name)[0]
		str_lines = []
		str_lines.append("# -*- coding:gbk -*-\n")
		str_lines.append("from " + tars_namespace + " import *\n")
		str_lines.append("from CSV2TARSData import *\n")
		str_lines.append("DataFilePath = \""+csv_path+"\"\n")
		bfile_name = bfile_path+tars_namespace + ".b"
		bfile_struct = "T"+tars_namespace
		str_lines.append("BFile = DotBFile(\"" + bfile_name + "\",\'" + tars_namespace + "\', \'"+bfile_struct + "Info\')\n")
		excel_data = open_excel(xlsm_path+file_name)

		#sheet表遍历
		sheetname_list = []
		getsheetname_list(excel_data, sheetname_list)
		i=0
		for sheet_name in sheetname_list:

			i=i+1
			str_lines.append("\n")
			str_lines.append("csv_Info"+str(i) + " = CSVFile(DataFilePath, \"" + sheet_name + "\")\n")
			str_lines.append("BFile.csvFiles.append(" + "csv_Info"+str(i) +")\n")
			
			
		str_lines.append("iret = BFile.generateData()\n")
		str_lines.append("exit(iret)\n")
		file_obj = open("./tmpgen.py",'w')
		file_obj.writelines(str_lines)
		file_obj.close()
		
		iret = os.system('python tmpgen.py')
		if iret == 1:
			exit(1)

def UsePlatform():
  osstr = platform.system()
  if(osstr =="Windows"):
    return 1
  elif(osstr == "Linux"):
    return 0
  else:
    return 1

if __name__ == '__main__':
	#防止文件夹不存在
	if not os.path.exists(bfile_path):
	    os.mkdir(bfile_path)

	if not os.path.exists(csv_path):
	    os.mkdir(csv_path)
	
	if not os.path.exists(tars_struct_path):
	    os.mkdir(tars_struct_path)
	
	# 清空目录
	for src in os.listdir(bfile_path):
	    os.remove(bfile_path+src)

	for src in os.listdir(csv_path):
	    os.remove(csv_path+src)

	for src in os.listdir(tars_struct_path):
	    os.remove(tars_struct_path+src)

	#预检查
	if not ExcelCheck.ExcelPreCheck():
	    exit("")

	#1.开始将xlsm文件转表到tars, 导出csv
	#统一编码

	server_genlog = open("server_gen.log", "w")

	reload(sys)
	sys.setdefaultencoding('gbk')


	print "start xlsm to  tars ...\n"
	file_namelist = []
	file_namelist_pre = []
	get_xlsm_filelist(file_namelist_pre)

	for tmp_name in file_namelist_pre:
		sheetname_list_pre = []
		excel_data = open_excel(xlsm_path+tmp_name)
		getsheetname_list(excel_data, sheetname_list_pre)
		if len(sheetname_list_pre) == 0:
			continue
		file_namelist.append(tmp_name)

	for tmp_name in file_namelist:
		print "..." + tmp_name + "...\n"
		xlsm2tars(tmp_name)

	print "end xlsm to  tars !\n"

	#2.tars2python将tars转成python编解码文件.py, 并copy到当前目录
	print "start tars2python  ...\n"


	ostype=UsePlatform()
	print "os type:" +str(ostype)
	if ostype==0:
	    os.system(t2py_path + "tars2python ./*.tars")



	for tmp_name in file_namelist:
	    tars_namespace = os.path.splitext(tmp_name)[0]
	    if ostype==0:
		    os.system('cp com/qq/'+tars_namespace+'/'+tars_namespace+'.py ./')
	    else:
		    os.system('tars2xx\\tars2python.exe ' + tars_namespace+ '.tars')
		    os.system('copy com\\qq\\'+tars_namespace+'\\'+tars_namespace+'.py ')

	print "end tars2python !  \n"

	#3.生成*.b

	print "start gen BFile...\n"
	xlsm2BFile(file_namelist)

	server_genlog.write(server_genlog_str)
	server_genlog.close()

	#Linux下日志有点问题
	if ostype!=0:
	    print server_genlog_str

	#Windows下删除临时文件
	shutil.rmtree('com')
	os.remove("tmpgen.py")
	for tmp_name in file_namelist:
		os.remove(''+os.path.splitext(tmp_name)[0]+'.py')
		os.remove(''+os.path.splitext(tmp_name)[0]+'.pyc')
		shutil.move(''+os.path.splitext(tmp_name)[0]+'.tars', tars_struct_path+os.path.splitext(tmp_name)[0]+'.tars')
	
	print "end gen BFile!\n"

	print('\033[1;32m ============= Generate Succeed =============\033[0m')










