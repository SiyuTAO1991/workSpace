import xlrd

def read_excel_file(fileName):
    #https://www.cnblogs.com/zhoujie/p/python18.html
    #https://www.cnblogs.com/lhj588/archive/2012/01/06/2314181.html
    workbook = xlrd.open_workbook(fileName)
    sqlFile = open(r"C:\Users\HP\Desktop\postCodeFile.txt", "w")
    sheet = workbook.sheet_by_name("Jawa Tengah")
    rowNo = sheet.nrows
    postcode_set = set()
    for i in range(2, rowNo):
        postcode = str(sheet.row_values(i)[7])[0:-2]
        if postcode not in postcode_set:
            postcode_set.add(postcode)
            sql = "INSERT INTO \"ggblock\" (\"postalcode\", \"countrycode\", \"countryname\", \"provincecode\", " \
                  "\"provincename\", \"districtcode\", \"districtname\", \"streetname\", \"blockcode\", \"blockdesc\", \"buildingname\", " \
                  "\"source\", \"unitno\", \"unitno1\", \"construction\", \"districtclass\", \"occupation\", \"occupancy\", \"heightofbuilding\", " \
                  "\"creatorcode\", \"createtime\", \"updatercode\", \"updatetime\", \"validdate\", \"invaliddate\", \"validind\", \"remark\", \"flag\", " \
                  "\"postalcodeautoind\", \"kecamatancode\", \"kecamatan\") " \
                  "VALUES (" + str(sheet.row_values(i)[7])[0:-2] + ", 'IDN', 'INDONESIA', '" + str(sheet.row_values(i)[0])[0:-2] + "', '" + sheet.row_values(i)[1] + "', '" + str(sheet.row_values(i)[2])[0:-2] + "', '" + sheet.row_values(i)[3] + "', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 'taosy', localtimestamp(0), 'taosy', localtimestamp(0), localtimestamp(0), NULL, 't', NULL, NULL, NULL, '" + str(sheet.row_values(i)[4])[0:-2] + "', '" + sheet.row_values(i)[5] + "');"
            print(sql, file=sqlFile)
        else:
            continue


read_excel_file(r'C:\Users\HP\Desktop\provinceArea111.xlsx')