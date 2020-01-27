import xlrd
import xlwt
def read_excel(fileName):
    #https://www.cnblogs.com/zhoujie/p/python18.html
    #https://www.cnblogs.com/lhj588/archive/2012/01/06/2314181.html
    workbook = xlrd.open_workbook(fileName)
    provinceSheet = workbook.sheet_by_index(1)
    rowNum = provinceSheet.nrows
    colNum = provinceSheet.ncols

    for i in range(1, rowNum):
        row = provinceSheet.row_values(i)
        SQL = "insert into ggtreecode (codetreetype, codetreecode, codetreename, uppercode, " \
              "displayno, creatorcode, createtime, updatercode, updatetime, validind) " \
              "values ('District', '" + str(row[2])[0:-2] + \
              "', '" + row[3] + "', '" + str(row[0])[0:-2] + "', '" + str(row[4])[0:-2] + \
              "', 'admin', localtimestamp(0), 'admin', localtimestamp(0), '1');"
        print(SQL)
def read_excel_cityDistrict(fileName, sheetName, area):
    workbook = xlrd.open_workbook(fileName)
    displayno = 58
    subDistrict = ""
    for j in sheetName:
        sheet = workbook.sheet_by_name(j)
        rowNum = sheet.nrows
        colNum = sheet.ncols
        for i in range(2, rowNum):
            row = sheet.row_values(i)
            if subDistrict == row[5]:
                #print(subDistrict)
                continue
            else:
                subDistrict = row[5]
                SQL = "insert into ggtreecode (codetreetype, codetreecode, codetreename, uppercode, " \
                      "displayno, creatorcode, createtime, updatercode, updatetime, validind) " \
                      "values (" + "'" + area + "'" + ", '" + str(row[4])[0:-2] + \
                      "', '" + row[5] + "', '" + str(row[2])[0:-2] + "', '" + str(displayno) + \
                      "', 'admin', localtimestamp(0), 'admin', localtimestamp(0), '1');"
                displayno = displayno + 1
            print(SQL)
def read_excel_urbanVilage(file, sheet, area):
    f = open(r"C:\Users\HP\Desktop\sqlKelurahan_Jawa_Tengah.txt","w")
    workbook = xlrd.open_workbook(file)
    displayno = 717
    for j in sheet:
        sheet = workbook.sheet_by_name(j)
        rowNum = sheet.nrows
        for i in range(2, rowNum):
            #print(j)
            row = sheet.row_values(i)
            print(row)
            SQL = "insert into ggtreecode (codetreetype, codetreecode, codetreename, uppercode, " \
                  "displayno, creatorcode, createtime, updatercode, updatetime, validind) " \
                  "values (" + "'" + area + "'" + ", '" + str(row[6])[0:-2] + \
                  "', '" + row[7] + "', '" + str(row[4])[0:-2] + "', '" + str(displayno) + \
                  "', 'admin', localtimestamp(0), 'admin', localtimestamp(0), '1');"
            displayno = displayno + 1
            print(SQL, file = f)

def read_excel_import_ggblock(file, sheets):
    #get the city/district code from province and distric or city sheet
    f = open(r"C:\Users\HP\Desktop\sqlGgblock.sql", "w")
    workbook = xlrd.open_workbook(file)
    districOrCitySheet = workbook.sheet_by_name('Province and Distric or City')
    districOrCityRowNum = districOrCitySheet.nrows
    # City/Distric Code and Name match dict
    districOrCityMap = {}
    postCode = set()


    for i in range(1, districOrCityRowNum):
        districOrCityRow = districOrCitySheet.row_values(i)
        districOrCityMap[districOrCityRow[2]] = districOrCityRow[3]

    for j in sheets:
        sheet = workbook.sheet_by_name(j)
        rowNum = sheet.nrows
        for m in range(2, rowNum):
        #for m in range(1, 3):
            row = sheet.row_values(m)
            flag = postCodeLoop(postCode, row[7])
            if flag == 1:
                continue
            else:
                postCode.add(row[7])
                city = districOrCityMap.get(row[2])
                insertSql = "insert into ggblock (postalcode, countrycode, countryname, provincecode, provincename, " \
                            "districtcode, districtname, creatorcode, createtime, updatercode, updatetime, validdate, validind" \
                            ", kecamatancode, kecamatan) values " \
                            "('" + str(row[7])[0:-2] + "', 'IDN', 'Indonesia', '" + str(row[0])[0:-2] + "', '" + row[1] + "', '" + str(row[2])[0:-2] + "', '" + city + "', 'taosy', localtimestamp(0), " \
                            "'taosy', localtimestamp(0), localtimestamp(0), 't', '" + str(row[4])[0:-2] + "', '" + row[5] + "');"
                        #print(row)
                        #print(row[4],type(row[4]), row[5],type(row[5]), row[6],type(row[6]), row[7],type([7]))
                        #postCode = postCode + 1
                print(insertSql, file=f)

def postCodeLoop(postCode, rowPostCode):
    for item in postCode:
        if rowPostCode == item:
            return 1
    return 0

#read_excel(r'C:\Users\HP\Desktop\provinceArea.xlsx')
fileN = r'C:\Users\HP\Desktop\provinceArea111.xlsx'
sheetN = ['Jawa Tengah'] #, 'Jawa Tengah'
#sheetN = ['BALI', 'Banten', 'Bengkulu', 'D.I Yogyakarta', ]
area = 'Kelurahan' # Kecamatan == sub-district; Kelurahan == urban/village
#read_excel_cityDistrict(fileN, sheetN, area)
#read_excel_urbanVilage(fileN, sheetN, area)
read_excel_import_ggblock(fileN, sheetN)
