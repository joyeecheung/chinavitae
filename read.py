from openpyxl import Workbook
from openpyxl import load_workbook
import re

# 
asciiPattern = ur'[%s]+' % ''.join(chr(i) for i in range(32,127))
chinesePattern = ur'[\u4e00-\u9fff. ]+'

def split_name(name):
	matches = re.match('(%s) (%s)' % (asciiPattern, chinesePattern), name)
	return matches.group(1), matches.group(2)

inwb = load_workbook('Chinavitae_1_500.xlsm')
outwb = Workbook()
infoSheet = outwb.create_sheet('info')
careerSheet = outwb.create_sheet('career')
travelSheet = outwb.create_sheet('travel')

for sheetName in inwb.getsheets:
	if not sheetName.isdigit():
		pass
	sheet = inwb[sheetName]
	# sheet = inwb['1']
	colA, colB, colC = sheet.columns[:3]
	for idx, cell in enumerate(colA):
		if unicode(cell.value).startswith(u'Biography Revised:'):
			revisedTime = cell.value.partition(':')[-1].strip()
		if unicode(cell.value).startswith(u'Career Data Updated:'):
			updatedTime = cell.value.partition(':')[-1].strip()
		if unicode(cell.value).startswith(u'Born:'):
			birthYear = cell.value.partition(':')[-1].strip()
			engName, chName = split_name(colA[idx-2].value)
		if unicode(cell.value).startswith(u'Birthplace:'):
			birthPlace = cell.value.partition(':')[-1].strip()
		if unicode(cell.value) == u'Biography':
			bioLineIdx = idx
		if unicode(cell.value) == u'Career Data':
			careerIdx = idx
		if unicode(cell.value).startswith(u'Recent Travel'):
			travelIdx = idx
		if unicode(cell.value).startswith(u'Compare'):
			compIdx = idx
	cv = {
		'revisedTime': revisedTime,
		'updatedTime': updatedTime,
		'birthYear': birthYear,
		'engName': engName,
		'chName': chName,
		'birthPlace': birthPlace,
		'bioLine': colA[bioLineIdx+2].value,
		'careerData': [(colA[i].value, colB[i].value, colC[i].value) for i in range(careerIdx+2, travelIdx-1)],
		'travelData': [(colA[i].value, colB[i].value, colC[i].value) for i in range(travelIdx+2, compIdx)],
		'id': sheetName
	}
