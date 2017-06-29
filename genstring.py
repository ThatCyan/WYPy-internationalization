# coding: utf-8

import sys
import codecs
import csv
import xlutils.copy
import xlrd

reload(sys)
sys.setdefaultencoding('utf8')

register_language = ['iOS KVN','Android KVN','页面','备注','zh-Hans','zh-Hant','en','ko','ja']

def main():

	mode = int(raw_input("\n请选择脚本工作模式\n\n1.载入 iOS Localizable.string 文件然后导出 language.cvs.\n2.导出 Localozable.string For iOS.\n3.纠正比对\n"))
	if mode == 1:
		load_localizable_dump_cvs()
	elif mode == 2:
		dump_localizable_for_ios()
	elif mode == 3:
		correct_sth()
	else:
		print("错误:指令有误")
		exit(0)

def load_localizable_dump_cvs():
	print("\n注意事项:\n1.请将Localizable.strings放置于当前脚本所在目录下\n2.请确认Locallizable.string文件中键值对的形式均为 \"key\" = \"value\";的形式 以英文双引号开头\n,且键值对中不出现'='3.每行只有一个键值对\n\n\n\n")
	print("开始读取文件")
	lzs = None
	try:
		lzs = file('Localizable.strings')
	except Exception as e:
		print('找不到Localizable.string文件')
		exit(0)

	print("加载文件成功")
	filter_file(lzs)

def dump_localizable_for_ios():
	print("\n注意事项:\n1.请将language.cvs文件置于脚本所在目录下\n")
	print("开始读取文件")
	# csvfile = file('language.csv', 'rb')
	# reader = csv.reader(csvfile)
	data = xlrd.open_workbook("language.xls")
	table = data.sheets()[0]
	reader = table.nrows
	for i in range(len(register_language)-4):
		with open(register_language[4+i]+'_Localizable.strings','w') as f:

			for j in range(1,reader):
				cell_key = table.cell(j,0).value
				cell_value = table.cell(j,4+i).value
				f.write("\"%s\" = \"%s\";\n"%(cell_key,cell_value))
			f.close()
			print(register_language[4+i]+'_Localizable.strings')
def correct_sth():
	#1 加载
	data = xlrd.open_workbook('android.xls')
	table = data.sheets()[0]
	dic = {}
	for i in range(table.nrows):
		cell_text = table.cell(i,5).value
		rr = dic.get(cell_text)

		if rr == None:
			dic[cell_text] = table.cell(i,8).value
		else:
			print("重复" + rr)
	print("共有:"+str(len(dic)))

	data2 = xlrd.open_workbook('language.xls')
	table2 = data2.sheets()[0]
	wb = xlutils.copy.copy(data2)
	ws = wb.get_sheet(0)
	for i in range(table2.nrows):
		cell_text = table2.cell(i,4).value
		k = dic.get(cell_text)
		print(k)
		if k!=None:
			ws.write(i, 7, k)
	wb.save('language.xls')


def filter_file(language_file):
	file_lines = language_file.readlines()
	true_lines = []
	for line in file_lines:
		result = filter_line(line)
		if result[0] == 0:
			pass
			#完全错误格式
		elif result[0] == 1:
			pass
			#类正确格式错误
		elif result[0] == 2:
			true_line = result[1]
			true_lines.append(true_line)
			#类正确格式正确
		elif result[0] == 3:
			exit(0)
	if len(true_lines)==0:
		print("加载文件失败,可用行为空")
		exit(0)
	else:

		csv_file = file('language.csv','wb')
		
		csv_file.write(codecs.BOM_UTF8)
		writer = csv.writer(csv_file)
		writer.writerow(register_language)
		for item in true_lines:
			writer.writerow(item)
		csv_file.close()
		print("写入完成,共有%d个键值对生成"%(len(true_lines)))	


def filter_line(line):
	clip = line.split('=')
	is_kv_line = False
	for char in clip[0]:
		if char == ' ':
			continue
		elif char == '\"':
			is_kv_line = True
			break
		else:
			is_kv_line = False
			return (0,None)
	if len(clip) == 2:
		key = get_info_inside_kv(clip[0])
		value = get_info_inside_kv(clip[1])
		return (2,[key,"","","",value,"" ,"","",""])

	else:
		print("键值内对中不该出现 '=' ")
def get_info_inside_kv(kv):
	start = 0
	end = 0

	for i in range(len(kv)):
		char = kv[0+i:1+i]
		if char == ' ' or char == '\n':
			continue
		elif char == '\"':
			start = i
			break
	for i in range(len(kv)):
		length = len(kv)
		char = kv[length-1-i:length-i]

		if char == ' ' or char == '\n' or char == ';':
			continue
		elif char == '\"':
			end = len(kv) - i
			break
	return kv[start+1:end-1]

if __name__ == '__main__':
	main()