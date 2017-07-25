from openpyxl import load_workbook

class WriteExcel():
    data = {}

    def writeAllData(self, data):
        wb = Workbook()
        saveSheet = self.setSaveSheetTemplate(data)
        saveSheet = self.writeSheetName(saveSheet, data)
        saveSheet = self.writeSheetHeader(saveSheet, data)
        saveSheet = self.writeSheetContext(savaSheet, data)
        wb.save('document.xlsx')

    def setSaveSheetTemplate(self, data):
        listSaveTemplate = []

        intAllSheetLength = self.setAllSheetLength(data)
        for i in range(intAllSheetLength):
            listSaveTemplate.append([])
        return listSaveTemplate

    def setAllSheetLength(self, data):
        allSheet = data['sheet']
        return len(allSheet)

    def writeSheetName(self, saveSheet, data):
        for i in range(len(saveSheet)):
            dataSheetName = data['sheet'][i]
            if i == 0:
                saveSheet[i] = wb.active
                saveSheet[i].title = dataSheetName
            else:
                saveSheet[i] = wb.create_sheet(title = dataSheetName)
        return saveSheet

    def writeSheetHeader(self, saveSheet, data):
        for sheet in range(len(saveSheet)):
            dataHeader = data['header'][sheet]
            for column in range(len(dataHeader)):
                saveSheet[sheet]['%s1' % chr(ord('A') + column)] = dataHeader[column]
        return saveSheet

    def writeSheetContext(self, saveSheet, data):
        for sheet in range(len(saveSheet)):
            dataContext = data['context'][sheet]
            for row in range(len(dataContext)):
                for column in range(len(dataContext[row])):
                saveSheet[sheet]['%s%d' % (chr(ord('A') + column , row+2)]
        return saveSheet
