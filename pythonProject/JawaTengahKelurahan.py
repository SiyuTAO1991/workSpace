import xlrd

def read_excel_file(fileName):
    #https://www.cnblogs.com/zhoujie/p/python18.html
    #https://www.cnblogs.com/lhj588/archive/2012/01/06/2314181.html
    workbook = xlrd.open_workbook(fileName)
    sqlFile = open(r"C:\Users\HP\Desktop\sqlFile.txt", "w")
    sheet = workbook.sheet_by_name("Jawa Tengah")
    rowNo = sheet.nrows

    kecamatanCode = set()
    code_in_tbl = set()
    miss_kelu_code = set()
    kecamatan_dict = {}


    for i in range(2, rowNo):
        # print(sheet.row_values(i)[4])
        kecamatanCode.add(str(sheet.row_values(i)[4])[0:-2])

    for j in range(2, rowNo):
        if len(str(sheet.row_values(j)[8])[0:-2]) > 0:
            code_in_tbl.add(str(sheet.row_values(j)[8])[0:-2])
    for m in kecamatanCode:
        if m not in code_in_tbl:
            miss_kelu_code.add(m)
    print(miss_kelu_code)

    for kelu_code in miss_kelu_code:
        for j in range(2, rowNo):
            if kelu_code == str(sheet.row_values(j)[4])[0:-2]:
                # print()
                kecamatan_dict, kelurahan_code = kelu_code_generate(kecamatan_dict, kelu_code)
                sql = "INSERT INTO ggtreecode (\"codetreetype\", \"codetreecode\", \"codetreename\", \"uppercode\", " \
                      "\"displayno\", \"creatorcode\", \"createtime\", \"updatercode\", \"updatetime\", \"validdate\", " \
                      "\"validind\", \"remark\", \"flag\", \"language\", \"invaliddate\") " \
                      "VALUES ('Kelurahan', '" + kelurahan_code + "', '" + sheet.row_values(j)[6] + "', '" + \
                      str(sheet.row_values(j)[4])[0:-2] + "', " + "1270" + ", 'taosy', localtimestamp(0), 'admin', " \
                      "localtimestamp(0), NULL, '1', NULL, NULL, NULL, NULL);"
                print(sql)
                print(sql, file=sqlFile)

def kelu_code_generate(kecamatan_dict, kecamatan_code):
    if kecamatan_code not in kecamatan_dict:
        kelurahan_index = 10
        kelurahan_code = kecamatan_code + str(kelurahan_index)
        kecamatan_dict[kecamatan_code] = kelurahan_index
        return kecamatan_dict, kelurahan_code
    else:
        kelurahan_index = kecamatan_dict.get(kecamatan_code) + 1
        kelurahan_code = kecamatan_code + str(kelurahan_index)
        kecamatan_dict[kecamatan_code] = kelurahan_index
        return kecamatan_dict, kelurahan_code






read_excel_file(r'C:\Users\HP\Desktop\provinceArea111.xlsx')