from openpyxl import Workbook
from openpyxl import load_workbook
import re

dot = u'\u00b7'
dash = u'\u2014'
emph = u'\u2022'
dot2 = u'\u2027'

seps = (u'.', dot, dash, emph, dot2)

# regex pattern matching all ascii characters
asciiPattern = ur'[%s]+' % ''.join(chr(i) for i in range(32,127))
# regex pattern matching all common Chinese characters
chinesePattern = ur'[\u4e00-\u9fff. %s]+' % (''.join(seps))

def get_clean_ch_name(chName):
    cleanName = chName
    for sep in seps:
        cleanName = cleanName.replace(sep, u' ')
    return cleanName


def split_name(name):
    matches = re.match('(%s) (%s)' % (asciiPattern, chinesePattern), name)
    if matches:
        return matches.group(1).strip(), matches.group(2).strip()
    else:
        matches = re.findall('(%s)' % (chinesePattern), name)
        matches = ''.join(matches).strip()
        if matches:
            return None, matches
        else:
            matches = re.findall('(%s)' % (asciiPattern), name)
            return ''.join(matches).strip(), None


def split_ch_name(chName):
    if len(chName) < 4:
        chFirstName = chName[0]
        chLastName = chName[1:]
    elif len(chName) == 4:
        chFirstName = chName[:2]
        chLastName = chName[2:]
    else:
        cleanName = get_clean_ch_name(chName)
        print u' '.join(cleanName.split())
        chFirstName, chLastName = cleanName.split()[:2]
    return chFirstName, chLastName


def split_eng_name(engName):
    nameParts = engName.split()
    if len(nameParts) < 2:
        return nameParts[0], None
    engFirstName, engLastName = engName.split()[:2]
    return engFirstName, engLastName


def guess_bio(bioLine):
    # guessing gender
    gender = None
    if re.search(ur'\bmale\b', bioLine):
        gender = 'Male'
    elif re.search(ur'\bfemale\b', bioLine):
        gender = 'Female'
    # guessing nationality
    nationality = re.findall(ur'(\w+) nationality', bioLine)
    nationality = nationality[0] if len(nationality) > 0 else None
    return gender, nationality


def guess_career(line):
    role = entity = None
    # guess role and entity
    role, temp, entity = line[1].partition(', ')
    return role, entity


def guess_travel(line):
    nature = 'News'
    # guess nature
    if line[2].startswith(u'Travelled to') or line[2].startswith(u'Was in'):
        nature = 'Location'
    elif line[2].startswith(emph):
        nature = 'Meeting'
    return nature


# input workbook
inwb = load_workbook('Chinavitae_1_500.xlsm')

# output workbook
outwb = Workbook()

# create sheets and heads
travelSheet = outwb.create_sheet(0, 'travel')
travelSheet.append(('IndividualID', 'Date', 'Type', 'Description', 'Nature'))


careerSheet = outwb.create_sheet(0, 'career')
careerSheet.append(('IndividualID', 'Derived', 'Date', 'SourceInformation',
    'RoleTitle', 'EntityName', 'Comments'))

infoSheet = outwb.create_sheet(0, 'info')
infoSheet.append(('BioLastRevised', 'CareerLastUpdate', 
    'ChineseName', 'EnglishFirstName', 'EnglishLastName', 'ChineseFirstName',
    'ChineseLastName', 'YearOfBirth', 'BirthPlace', 'Gender', 'Nationality', 'Biography'))

for sheetName in inwb.get_sheet_names():
    if not sheetName.isdigit():
        continue
    sheet = inwb[sheetName]
    print "------------------------------------------------------"
    print "Processing sheet", sheetName, "........"
    print ' '

    if len(sheet.columns) < 2:
        continue

    colA, colB = sheet.columns[:2]

    if len(sheet.columns) <= 2:
        alen = len(colA)
        for i in range(1, alen):
            sheet.cell('C%s'%(i)).value = None

    colC = sheet.columns[2]
    
    revisedTime = updatedTime = birthYear = engName = chName = None
    birthPlace = bioLineIdx = careerIdx = travelIdx = compIdx = None
    engFirstName = engLastName = chFirstName = chLastName = None
    bioLine = gender = nationality = None

    # process column A to get all available fields and indexes
    for idx, cell in enumerate(colA):
        if unicode(cell.value).startswith(u'Biography Revised:'):
            revisedTime = cell.value.partition(':')[-1].strip()
        if unicode(cell.value).startswith(u'Career Data Updated:'):
            updatedTime = cell.value.partition(':')[-1].strip()
        if unicode(cell.value).startswith(u'Born:'):
            birthYear = cell.value.partition(':')[-1].strip()
        if unicode(cell.value).startswith(u'PHOTO:'):
            engName, chName = split_name(colA[idx+1].value)
        if unicode(cell.value).startswith(u'Birthplace:'):
            birthPlace = cell.value.partition(':')[-1].strip()
        if unicode(cell.value) == u'Biography':
            bioLineIdx = idx
        if unicode(cell.value) == u'Career Data':
            careerIdx = idx
        if unicode(cell.value).startswith(u'Recent Travel'):
            travelIdx = idx
        if unicode(cell.value) == u'Compare':
            compIdx = idx

    # english name
    if engName:
        engFirstName, engLastName = split_eng_name(engName)

    # chinese name
    if chName:
        chFirstName, chLastName = split_ch_name(chName)

    # biography
    if bioLineIdx:
        bioLine = colA[bioLineIdx+2].value
        gender, nationality = guess_bio(bioLine)

    cv = {
            'revisedTime': revisedTime, 
            'updatedTime': updatedTime, 
            'chName': chName, 
            'engFirstName': engFirstName, 
            'engLastName': engLastName, 
            'chFirstName': chFirstName, 
            'chLastName': chLastName, 
            'birthYear': birthYear, 
            'birthPlace': birthPlace, 
            'gender': gender, 
            'nationality': nationality, 
            'bioLine': bioLine, 
            'id': sheetName
        }

    name = get_clean_ch_name(chName) if chName else engName
    infoSheet.append((cv['revisedTime'], cv['updatedTime'], 
        cv['chName'], cv['engFirstName'], cv['engLastName'], cv['chFirstName'],
        cv['chLastName'], cv['birthYear'], cv['birthPlace'], cv['gender'],
        cv['nationality'], cv['bioLine']))
    print "Insert profile for", cv['id'], name

    if careerIdx:
        if travelIdx:
            print "careerIdx, travelIdx", careerIdx, travelIdx
            cv['careerData'] = [(colA[i].value, colB[i].value, colC[i].value) for i in range(careerIdx+2, travelIdx-1)]
        else:
            print "careerIdx, compIdx", careerIdx, compIdx
            cv['careerData'] = [(colA[i].value, colB[i].value, colC[i].value) for i in range(careerIdx+2, compIdx)]
        for idx, line in enumerate(cv['careerData']):
            if not line[0] and not line[1] and not line[2]:
                continue
            role, entity = guess_career(line)
            # print ' | '.join(unicode(i) for i in (cv['id'], None, line[0], line[1], role, entity, line[2]))
            careerSheet.append((cv['id'], None, line[0], line[1], 
                role, entity, line[2]))
        print "Inserted career data for", cv['id'], name

    if travelIdx:
        cv['travelData'] = [[colA[i].value, colB[i].value, colC[i].value] for i in range(travelIdx+2, compIdx)]
        for idx, line in enumerate(cv['travelData']):
            if not line[0] and not line[1] and not line[2]:
                continue
            if type(line[0]) is unicode and line[0].startswith("The most recent"):
                continue
            nature = guess_travel(line)
            if not line[0]:
                line[0] = cv['travelData'][idx-1][0] if idx > 0 else None
            if not line[1]:
                line[1] = cv['travelData'][idx-1][1] if idx > 0 else None
            # print ' | '.join(unicode(i) for i in (cv['id'], line[0], line[1], line[2], nature))
            travelSheet.append((cv['id'], line[0], line[1], line[2], nature))
        print "Inserted travel data for", cv['id'], name

outwb.save("out.xlsx")