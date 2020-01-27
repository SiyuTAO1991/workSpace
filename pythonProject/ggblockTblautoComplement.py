import xlrd
import xlwt

def ggblock_tbl_auto_comp(fileName, sheetsName):
    workbook = xlrd.open_workbook(fileName)
    # 获得provinceMap = {'province name': province code(float)}
    # 获得cityMap = = {'city name': city code(float)}
    provinceMap, cityMap = get_provice_city_map(workbook)
    missCodeFile = open(r"C:\Users\HP\Desktop\ggblock_autoComplement\missCode.txt", "w")
    postCodeSet = set()
    #print(provinceMap)
    #print(cityMap)
    for sheets in sheetsName:
        sqlFileName = r"C:\Users\HP\Desktop\ggblock_autoComplement\\" + str(sheets) + "_ggblock_tbl_auto_comp_SQL.sql"
        sqlFile = open(sqlFileName, "w")
        # 该sheet的索引Map
        sheetIndex = {}
        # 循环取sheets
        sheet = workbook.sheet_by_name(sheets)
        # 该sheet的行数和列数
        rowNum = sheet.nrows
        colNum = sheet.ncols
        # 取第一行（标题行）
        sheetHead = sheet.row_values(0)
        KecamatanCodeIndex = 1
        ifCityChange = ""
        for i in range(1, len(sheetHead)):
            sheetIndex[sheetHead[i]] = i
        # print(sheetIndex)
        for j in range(2, rowNum): # 取非标题内容行
            rowValue = sheet.row_values(j)
            flag = postCodeLoop(postCodeSet, rowValue[4])

            if flag == 1:
                continue
            else:
                postCodeSet.add(rowValue[4])
                provinceCode = provinceMap.get(rowValue[0].upper().strip())
                cityCode = cityMap.get(rowValue[1].upper().strip())

                if not provinceCode or not cityCode:
                    if not provinceCode:
                        print("表：" + sheets + " 中的省：" + rowValue[0] + "无法找到对应的code。", file=missCodeFile)
                    if not cityCode:
                        print("表：" + sheets + " 中的市：" + rowValue[1] + "无法找到对应的code。", file=missCodeFile)
                else:
                    if ifCityChange == rowValue[1]:
                        KecamatanCodeIndex = KecamatanCodeIndex + 1
                        kecamatanCode = cityCode * 100 + KecamatanCodeIndex

                        insertSql = "insert into ggblock (postalcode, countrycode, countryname, provincecode, provincename, " \
                                    "districtcode, districtname, creatorcode, createtime, updatercode, updatetime, validdate, validind" \
                                    ", kecamatancode, kecamatan) values " \
                                    "('" + str(rowValue[4])[0:-2] + "', 'IDN', 'Indonesia', '" + str(provinceCode)[0:-2] + "', '" + \
                                    rowValue[0].upper() + "', '" + str(cityCode)[0:-2] + "', '" + rowValue[1].upper() + "', 'taosy', localtimestamp(0), " \
                                    "'taosy', localtimestamp(0), localtimestamp(0), 't', '" + str(kecamatanCode)[0:-2] + "', '" + rowValue[2] + "');"
                        print(insertSql, file=sqlFile)
                    else:
                        ifCityChange = rowValue[1]
                        KecamatanCodeIndex = 1
                        kecamatanCode = cityCode * 100 + KecamatanCodeIndex
                        insertSql = "insert into ggblock (postalcode, countrycode, countryname, provincecode, provincename, " \
                                    "districtcode, districtname, creatorcode, createtime, updatercode, updatetime, validdate, validind" \
                                    ", kecamatancode, kecamatan) values " \
                                    "('" + str(rowValue[4])[0:-2] + "', 'IDN', 'Indonesia', '" + str(provinceCode)[
                                                                                                 0:-2] + "', '" + \
                                    rowValue[0].upper() + "', '" + str(cityCode)[0:-2] + "', '" + rowValue[
                                        1].upper() + "', 'taosy', localtimestamp(0), " \
                                                     "'taosy', localtimestamp(0), localtimestamp(0), 't', '" + str(
                            kecamatanCode)[0:-2] + "', '" + rowValue[2] + "');"
                        print(insertSql, file=sqlFile)


def get_provice_city_map(workBook):
    sheetName = 'Sheet2'
    sheet = workBook.sheet_by_name(sheetName)
    provinceMap = {}
    cityMap = {}
    cityNameMap={}
    rowNum = sheet.nrows
    for i in range(1, rowNum):
        rowValue = sheet.row_values(i)
        provinceMap[rowValue[1]] = rowValue[0]
        cityMap[rowValue[3]] = rowValue[2]
        cityNameMap[rowValue[2]] = rowValue[3]

    return provinceMap, cityMap, cityNameMap

def postCodeLoop(postCodeSet, rowPostCode):
    for item in postCodeSet:
        if rowPostCode == item:
            return 1
    return 0

def kecamatanCode_get_map(workbook, sheetName):
    sheet = workbook.sheet_by_name(sheetName)
    rowNum = sheet.nrows
    kecamatanCodeMap = {}
    for i in range(1, rowNum):
        rowValue = sheet.row_values(i)
        #print("rowValue is:" + str(rowValue))
        kecamatanCodeMap[(rowValue[2], rowValue[1])] = rowValue[0]
        #print(kecamatanCodeMap)
    #print(kecamatanCodeMap)
    return kecamatanCodeMap

def get_kelurahanIndex(kelurahanIndexMap, kecamatanName):
    if kecamatanName in kelurahanIndexMap:
        kelurahanIndex = kelurahanIndexMap.get(kecamatanName)
        kelurahanIndex = kelurahanIndex + 1
        kelurahanIndexMap[kecamatanName] = kelurahanIndex
        return kelurahanIndexMap, kelurahanIndex
    else:
        kelurahanIndexMap[kecamatanName] = 1
        return kelurahanIndexMap, 1

def ggtreecode_subDistrict_auto_compl(fileName, sheetName):
    area = "Kecamatan"
    displayNo = 3000
    kelurahanDisplayNo = 12750
    workbook = xlrd.open_workbook(fileName)
    provinceMap, cityMap, cityNameMap = get_provice_city_map(workbook)
    kecMissFile = open(r"C:\Users\HP\Desktop\KecamatanCode_autocomp\kecamatanMissFile.txt", "w")
    keluMissFile = open(r"C:\Users\HP\Desktop\KelurahanCode_autocomp\kecamatanCodeMissFile.txt", "w")
    kecamatanMap = get_kecamatanMap(workbook)
    kecamatanCodeMap = kecamatanCode_get_map(workbook, 'Sheet3')
    for sheets in sheetName:
        sqlFileName = r"C:\Users\HP\Desktop\KecamatanCode_autocomp\\" + str(sheets) + "_ggblock_tbl_auto_comp_SQL.sql"
        kelurahanSqlFileName = r"C:\Users\HP\Desktop\KelurahanCode_autocomp\\" + str(sheets) + "_ggblock_tbl_auto_comp_SQL.sql"
        sqlFile = open(sqlFileName, "w")
        kelurahanSqlFile = open(kelurahanSqlFileName, "w")
        # 循环取sheets
        sheet = workbook.sheet_by_name(sheets)
        # 该sheet的行数和列数
        rowNum = sheet.nrows
        subDistrictSet = set()
        kecamatanName = ""
        cityName = ""
        kelurahanCityName = ""
        kecamatanIndex = 1
        kelurahanIndexMap = {}

        for i in range(2, rowNum):
            rowValue = sheet.row_values(i)
            flag = subDistrictLoop(subDistrictSet, rowValue[2])
            if flag == 1:
                continue
            else:
                subDistrictSet.add(rowValue[2])
                #provinceCode = provinceMap.get(rowValue[0].upper().strip())
                cityCode = cityMap.get(rowValue[1].upper().strip())
                #kecamatanCode = kecamatanMap.get(rowValue[4]) # 根据post Code去取KecamatanCode

                if not cityCode:
                    if not cityCode:
                        print("表：" + sheets + " 中的cityCode：" + rowValue[1] + "无法找到对应的code。", file=kecMissFile)

                else:
                    if cityName == rowValue[1]:
                        displayNo = displayNo + 1
                        kecamatanIndex = kecamatanIndex + 1
                        kecamatanCode = str(cityCode * 100 + kecamatanIndex)[0:-2]
                        SQL = "insert into ggtreecode (codetreetype, codetreecode, codetreename, uppercode, " \
                                "displayno, creatorcode, createtime, updatercode, updatetime, validind) " \
                                "values (" + "'" + area + "'" + ", '" + kecamatanCode + \
                                "', '" + rowValue[2] + "', '" + str(cityCode)[0:-2] + "', '" + str(displayNo) + \
                                "', 'taosy', localtimestamp(0), 'taosy', localtimestamp(0), '1');"
                        #print(SQL, file=sqlFile)
                    else:
                        displayNo = displayNo + 1
                        cityName = rowValue[1]
                        kecamatanIndex = 1
                        kecamatanCode = str(cityCode * 100 + kecamatanIndex)[0:-2]
                        SQL = "insert into ggtreecode (codetreetype, codetreecode, codetreename, uppercode, " \
                              "displayno, creatorcode, createtime, updatercode, updatetime, validind) " \
                              "values (" + "'" + area + "'" + ", '" + kecamatanCode + \
                              "', '" + rowValue[2] + "', '" + str(cityCode)[0:-2] + "', '" + str(displayNo) + \
                              "', 'taosy', localtimestamp(0), 'taosy', localtimestamp(0), '1');"
                        #print(SQL, file=sqlFile)


        for j in range(2, rowNum):
            rowValue = sheet.row_values(j)
            kelurahanArea = "Kelurahan"
            #kecamatanCode = kecamatanMap.get(rowValue[4]) # 根据postalCode去取KecamatanCode
            cityCode = cityMap.get(rowValue[1].upper().strip())
            #print("cityCode is:" + str(cityCode) + " rowValue[2] is: " + rowValue[2])
            #print("KecamatanCodeMap is:" + str(kecamatanCodeMap))
            kecamatanCode = kecamatanCodeMap.get((cityCode, rowValue[2]))
            #print(kecamatanCode)
            if kecamatanName == rowValue[2]:
                if not kecamatanCode:
                    kecamatanName = rowValue[2]
                    print("表：" + sheets + " 中的kecamatanCode：" + rowValue[2] + "无法找到对应的code。", file=keluMissFile)
                else:
                    kelurahanIndex = kelurahanIndex + 1
                    #kelurahanCode = str(kecamatanCode * 100 + kelurahanIndex)[0:-2]
                    kelurahanIndexMap,kelurahanIndex = get_kelurahanIndex(kelurahanIndexMap, rowValue[2])
                    kelurahanCode = str(kecamatanCode * 1000 + kelurahanIndex)[0:-2]
                    SQL = "insert into ggtreecode (codetreetype, codetreecode, codetreename, uppercode, " \
                            "displayno, creatorcode, createtime, updatercode, updatetime, validind) " \
                            "values (" + "'" + kelurahanArea + "'" + ", '" + kelurahanCode + \
                            "', '" + rowValue[3] + "', '" + str(kecamatanCode)[0:-2] + "', '" + str(kelurahanDisplayNo) + \
                            "', 'admin', localtimestamp(0), 'admin', localtimestamp(0), '1');"
                    print(SQL, file=kelurahanSqlFile)
            else:
                if not kecamatanCode:
                    kecamatanName = rowValue[2]
                    #print("表：" + sheets + " 中的kecamatanCode：" + rowValue[2] + "无法找到对应的code。", file=keluMissFile)
                else:
                    #kecamatanName = rowValue[2]
                    #kelurahanIndex = 1
                    #kelurahanCode = str(kecamatanCode * 100 + kelurahanIndex)
                    kelurahanIndexMap, kelurahanIndex = get_kelurahanIndex(kelurahanIndexMap, rowValue[2])
                    #print(type(kecamatanCode), type(kelurahanIndex), kelurahanIndex)
                    #print(type(kecamatanCode),kecamatanCode, type(kelurahanIndex), kelurahanIndex)
                    kelurahanCode = str(kecamatanCode * 1000 + kelurahanIndex)[0:-2]
                    SQL = "insert into ggtreecode (codetreetype, codetreecode, codetreename, uppercode, " \
                          "displayno, creatorcode, createtime, updatercode, updatetime, validind) " \
                          "values (" + "'" + kelurahanArea + "'" + ", '" + kelurahanCode + \
                          "', '" + rowValue[3] + "', '" + str(kecamatanCode)[0:-2] + "', '" + str(kelurahanDisplayNo) + \
                          "', 'admin', localtimestamp(0), 'admin', localtimestamp(0), '1');"
                    print(SQL, file=kelurahanSqlFile)

def subDistrictLoop(subDistrict, rowSubDistrict):
    for item in subDistrict:
        if rowSubDistrict == item:
            return 1
    return 0
def get_kecamatanMap(workbook):
    kecamatanMap = {}
    sheet = workbook.sheet_by_name('kecamatan autocomplement map')
    rowNum = sheet.nrows
    for i in range(1, rowNum):
        rowValue = sheet.row_values(i)
        kecamatanMap[rowValue[0]] = rowValue[2]
    print(kecamatanMap)
    return kecamatanMap





fileName = r'C:\Users\HP\Desktop\provinceArea111.xlsx'
# sheetName = ['Jawa Timur', 'Kalimantan Barat', 'Kalimantan Selatan',
#              'Kalimantan Tengah', 'Kalimantan Timur', 'Kalimantan Utara',
#              'Kepulauan Bangka Belitung', 'Kepulauan Riau', 'Lampung',
#              'Maluku', 'Maluku Utara', 'Nusa Tenggara Barat (NTB)',
#              'Nusa Tenggara Timur', 'Papua', 'Papua Barat', 'Riau',
#              'Sulawesi Barat', 'Sulawesi Selatan', 'Sulawesi Tengah',
#              'Sulawesi Tenggara', 'Sulawesi Utara', 'Sumatera Barat',
#              'Sumatera Selatan', 'Sumatera Utara ']
#ggblock_tbl_auto_comp(fileName, sheetName)
sheetName = ['Jawa Tengah']

ggtreecode_subDistrict_auto_compl(fileName, sheetName)

