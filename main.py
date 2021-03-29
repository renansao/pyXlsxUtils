from xlsxReadWrite import xlsxUtils

reader = xlsxUtils("test.xlsx", "Dados Pessoais")


data = reader.returnData()

dataDict = {
    "cpf":"9999",
    "exec":"exec1",
    "data": "2022",
    "utils": "231"
}

dataDict2 = {
    "cpf":"8888",
    "exec":"exec1",
    "data": "2022"
}

reader.insertNewRow(dataDict)
reader.insertNewRow(dataDict2)

print(data)