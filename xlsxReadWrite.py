import openpyxl

class xlsxUtils():

    def __init__(self, path, sheetName):
        self.path = path
        self.sheetName = sheetName
        self.workBook = object
        self.sheet = object
        self.sheetMaxRow = 0
        self.sheetMaxColumn = 0
        self.headerDict = {}
        self.headerDataIndex = {}
        self.initializeSheet()
    
    def initializeSheet(self):

        #Initialize WorkBook
        self.workBook = openpyxl.load_workbook(self.path)
        
        #Check if sheetname exists in WorkBook
        if (self.sheetName not in self.workBook.sheetnames):
            errorMessage = "SheetName '%s' not found in WorkBook" % self.sheetName
            raise Exception(errorMessage)
        
        #Initialize Sheet based on SheetName
        for sheet in self.workBook.worksheets:
            if sheet.title == self.sheetName:
                self.sheet = sheet
                break
        else:
            errorMessage = "SheetName '%s' could not be initialized" % self.sheetName
            raise Exception(errorMessage)
        
        #Define max Column
        self.__defineMaxColumn()

        #Define max Row
        self.__defineMaxRow()

        #Create Dict with header data
        #Create Dict with index of header data
        for headerColumn in range(1, self.sheetMaxColumn + 1):
            self.headerDict[self.sheet.cell(1, headerColumn).value] = ""
            self.headerDataIndex[self.sheet.cell(1, headerColumn).value] = headerColumn

    def returnData(self):

        #Instantiate list with Dicts
        sheetData = []
        
        #For each row, add data to a rowDict then append to a list of rows
        for row in range(2, self.sheetMaxRow + 1):

            #Copy the dict of header data
            rowDict = self.headerDict.copy()

            for column in range(1, self.sheetMaxColumn + 1):

                #If cell value is not None, add it to the row data dict
                if (self.sheet.cell(row, column).value != None):
                    rowDict[self.sheet.cell(1, column).value] = self.sheet.cell(row, column).value

            sheetData.append(rowDict)
        
        return sheetData

    def insertNewRow(self, dataDict):

        #Get the number of the last row and add 1
        newRow = self.sheetMaxRow + 1

        #Iterate through dataDict
        for key, value in dataDict.items():

            if self.headerDataIndex.get(key) == None:
                errorMessage = "Couldnt find Index for Header '%s'" % key
                raise Exception(errorMessage)

            #Assign the value given in the data dict to the correspondent cell
            self.sheet.cell(newRow, self.headerDataIndex.get(key), value)

            #Save WorkBook
            self.workBook.save(self.path)

        #Initialize the sheet again
        self.initializeSheet()
    
    def __defineMaxColumn(self):

        value = ""
        column = 1
        while value != None:
            value = self.sheet.cell(1,column).value
            column += 1
        
        self.sheetMaxColumn = column - 2
    
    def __defineMaxRow(self):

        value = ""
        row = 1
        while value != None:
            value = self.sheet.cell(row,1).value
            row += 1
        
        self.sheetMaxRow = row - 2