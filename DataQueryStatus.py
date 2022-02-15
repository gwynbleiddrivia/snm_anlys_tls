import pandas as pd 
import os 
import openpyxl
from openpyxl import load_workbook as load 
dic = {0:'A',1:'B',2:'C',3:'D',4:'E',5:'F',6:'G',7:'H',8:'I',9:'J',10:'K',11:'L',12:'M',13:'N',14:'O',15:'P',16:'Q',17:'R',18:'S',19:'T',20:'U',21:'V',22:'W',23:'X',24:'Y',25:'Z',26:'AA',27:'AB',28:'AC',29:'AD',30:'AE',31:'AF',32:'AG',33:'AH',34:'AI',35:'AJ',36:'AK',37:'AL',38:'AM',39:'AN',40:'AO',41:'AP',42:'AQ',43:'AR',44:'AS',45:'AT',46:'AU',47:'AV',48:'AW',49:'AX',50:'AY',51:'AZ',52:'BA'}
path1 = '/home/serenity/Dropbox/Farah Checks/Pending/Core Questions Queries/'
path2 = '/home/serenity/Dropbox/Farah Checks/Pending/Special Questions Queries/'
output = '/home/serenity/MyNotebooks/QExperiment/'
file_addr1 = os.listdir('/home/serenity/Dropbox/Farah Checks/Pending/Core Questions Queries')
file_addr2 = os.listdir('/home/serenity/Dropbox/Farah Checks/Pending/Special Questions Queries')

# file_select = [nam for nam in file_addr if nam[-17:] == '(all months).xlsx'] 
# file_select = [nam for nam in file_addr if nam[-20:] == 'Inconsistencies.xlsx' or nam[-17:] == '(all months).xlsx']
# file_select = [nam for nam in file_addr if nam[-20:] == 'Inconsistencies.xlsx' or nam[-17:] == 'Data Queries.xlsx' or nam[-17:] == '(all months).xlsx']
file_select1 = [path1 + nam for nam in file_addr1 if nam[-20:] == 'Inconsistencies.xlsx' or nam[-12:] == 'Queries.xlsx' or nam[-17:] == '(all months).xlsx' or nam[-9:] == 'Name.xlsx']
file_select2 = [path2 + nam for nam in file_addr2 if nam[-20:] == 'Inconsistencies.xlsx' or nam[-12:] == 'Queries.xlsx' or nam[-17:] == '(all months).xlsx' or nam[-9:] == 'Name.xlsx']
file_select = file_select1 + file_select2

tabularasa = pd.DataFrame()
for file in file_select:
	xls = pd.ExcelFile(file)
	workbook = load(file, data_only = True)
	for tab in xls.sheet_names:
		df = pd.read_excel(file,sheet_name = tab)
		sheet = workbook[tab]
		answercolumn = dic[df.columns.get_loc('answer')]
		# answer1column = dic[df.columns.get_loc('answer_1')]
		# answer2column = dic[df.columns.get_loc('answer_2')]
		# answer3column = dic[df.columns.get_loc('answer_3')]
		statuscolumn = dic[df.columns.get_loc('status')]
		length = int(len(df['status'])) + 1
		pendingbutanswered = 0
		pendingbutnotanswered = 0
		queriedbutanswered1sttime = 0
		queried1stbutnotanswered = 0
		queriedbutanswered2ndtime = 0
		queried2ndbutnotanswered = 0
		queriedbutanswered3rdtime = 0
		queried3rdbutnotanswered = 0
		veri = 0
		for cell in range(1,length):
			if sheet[statuscolumn + str(cell)].value == 'solved' or sheet[statuscolumn + str(cell)].value == 'verified':
				veri = veri + 1
			if sheet[answercolumn + str(cell)].fill.start_color.index != "00000000" and sheet[statuscolumn + str(cell)].value == 'pending':
				pendingbutanswered = pendingbutanswered + 1
			if sheet[answercolumn + str(cell)].fill.start_color.index == "00000000" and sheet[statuscolumn + str(cell)].value == 'pending':
				pendingbutnotanswered = pendingbutnotanswered + 1
	
			if 'answer_1' in df.columns:
				if (sheet[statuscolumn + str(cell)].value == "queried" and sheet[dic[df.columns.get_loc('status')] + str(cell)].fill.start_color.index == "FFFFFF00" and sheet[dic[df.columns.get_loc('answer_1')] + str(cell)].fill.start_color.index == "FF00B0F0"):
					queriedbutanswered1sttime = queriedbutanswered1sttime + 1
				if (sheet[statuscolumn + str(cell)].value == "queried" and sheet[dic[df.columns.get_loc('status')] + str(cell)].fill.start_color.index == "FFFFC000" and sheet[dic[df.columns.get_loc('answer_1')] + str(cell)].fill.start_color.index == "FF00B0F0"):
					queriedbutanswered2ndtime = queriedbutanswered2ndtime + 1
				if (sheet[statuscolumn + str(cell)].value == "queried" and sheet[dic[df.columns.get_loc('status')] + str(cell)].fill.start_color.index == "FFC00000" and sheet[dic[df.columns.get_loc('answer_1')] + str(cell)].fill.start_color.index == "FF00B0F0"):
					queriedbutanswered3rdtime = queriedbutanswered3rdtime + 1
				if (sheet[statuscolumn + str(cell)].value == "queried" and sheet[dic[df.columns.get_loc('status')] + str(cell)].fill.start_color.index == "FFFFFF00" and sheet[dic[df.columns.get_loc('answer_1')] + str(cell)].fill.start_color.index != "FF00B0F0"):
					queried1stbutnotanswered = queried1stbutnotanswered + 1
				if (sheet[statuscolumn + str(cell)].value == "queried" and sheet[dic[df.columns.get_loc('status')] + str(cell)].fill.start_color.index == "FFFFC000" and sheet[dic[df.columns.get_loc('answer_1')] + str(cell)].fill.start_color.index != "FF00B0F0"):
					queried2ndbutnotanswered = queried2ndbutnotanswered + 1
				if (sheet[statuscolumn + str(cell)].value == "queried" and sheet[dic[df.columns.get_loc('status')] + str(cell)].fill.start_color.index == "FFC00000" and sheet[dic[df.columns.get_loc('answer_1')] + str(cell)].fill.start_color.index != "FF00B0F0"):
					queried3rdbutnotanswered = queried3rdbutnotanswered + 1			
			# if 'answer_2' in df.columns:
			# 	if (sheet[statuscolumn + str(cell)].value == "queried" and sheet[dic[df.columns.get_loc('answer_2')] + str(cell)].fill.start_color.index == "FF00B0F0") or (sheet[statuscolumn + str(cell)].value == "double-queried" and sheet[dic[df.columns.get_loc('answer_2')] + str(cell)].fill.start_color.index == "FF00B0F0"):
			# 		queriedbutanswered2ndtime = queriedbutanswered2ndtime + 1
			# 	if (sheet[statuscolumn + str(cell)].value == "queried" and sheet[dic[df.columns.get_loc('answer_2')] + str(cell)].fill.start_color.index != "FF00B0F0") or (sheet[statuscolumn + str(cell)].value == "double-queried" and sheet[dic[df.columns.get_loc('answer_2')] + str(cell)].fill.start_color.index != "FF00B0F0"):
			# 		queried2ndbutnotanswered = queried2ndbutnotanswered + 1
			# if 'answer_3' in df.columns:
			# 	if (sheet[statuscolumn + str(cell)].value == "queried" and sheet[dic[df.columns.get_loc('answer_3')] + str(cell)].fill.start_color.index == "FF00B0F0") or (sheet[statuscolumn + str(cell)].value == "triple-queried" and sheet[dic[df.columns.get_loc('answer_3')] + str(cell)].fill.start_color.index == "FF00B0F0"):
			# 		queriedbutanswered3rdtime = queriedbutanswered3rdtime + 1
			# 	if (sheet[statuscolumn + str(cell)].value == "queried" and sheet[dic[df.columns.get_loc('answer_3')] + str(cell)].fill.start_color.index != "FF00B0F0") or (sheet[statuscolumn + str(cell)].value == "triple-queried" and sheet[dic[df.columns.get_loc('answer_3')] + str(cell)].fill.start_color.index != "FF00B0F0"):
			# 		queried3rdbutnotanswered = queried3rdbutnotanswered + 1

		if path1 in file:
			file_temp = file.replace(path1,"")		     
		elif path2 in file:
			file_temp = file.replace(path2,"")
		

		tabulaseries = pd.Series({
						'File Name':file_temp,
						'Sheet Name': tab,
						'Verified Among Unsolved Files': veri,
						'Pending but Answered': pendingbutanswered,
						'Queried for the 1st time but Answered' : queriedbutanswered1sttime,
						'Queried for the 2nd time but Answered': queriedbutanswered2ndtime,
						'Queried for the 3rd time but Answered': queriedbutanswered3rdtime,
						'Pending but Not Answered': pendingbutnotanswered,
						'Queried for the 1st time but Not Answered' : queried1stbutnotanswered,
						'Queried for the 2nd time but Not Answered': queried2ndbutnotanswered,
						'Queried for the 3rd time but Not Answered': queried3rdbutnotanswered
						})
	
			
		tabularasa = tabularasa.append(tabulaseries,ignore_index = True)	
		print(file,tab,pendingbutanswered, pendingbutnotanswered, queriedbutanswered1sttime, queried1stbutnotanswered, queriedbutanswered2ndtime, queriedbutanswered2ndtime, queried2ndbutnotanswered, queriedbutanswered3rdtime, queried3rdbutnotanswered)
		# print(tabulaseries)
		# tabularasa.append(tabulaseries,ignore_index = True)

# print(tabularasa)
tabularasa.to_excel("Data Query Status.xlsx",encoding = 'utf-08',index=False)

		
#BLUE = FF00B0F0
#WHITE = 00000000
#YELLOW = FFFFFF00
#ORANGE = FFFFC000
#RED = FFC00000
