from openpyxl import Workbook

class WriteExcel():
    def __init__(self, file):
        fileName = self.setFileName()
        self.writeAllData(file, fileName)

    def setFileName(self):
        print('Please enter save file name you want:')
        fileName = input()
        fileName = self.correctFileName(fileName)
        return fileName
		
    def correctFileName(self, fileName):
        if fileName.find('.xlsx') != -1:
            fileName += '.xlsx'
        return fileName
	
    def writeAllData(self, data, fileName):
        wb = Workbook()
        saveSheet = self.setSaveSheetTemplate(data)
        saveSheet = self.writeSheetName(wb, saveSheet, data)
        saveSheet = self.writeSheetHeader(saveSheet, data)
        saveSheet = self.writeSheetContext(saveSheet, data)
        wb.save('fileName')

    def setSaveSheetTemplate(self, data):
        listSaveTemplate = []

        intAllSheetLength = self.setAllSheetLength(data)
        for i in range(intAllSheetLength):
            listSaveTemplate.append([])
        return listSaveTemplate

    def setAllSheetLength(self, data):
        allSheet = data['sheet']
        return len(allSheet)

    def writeSheetName(self, workbook, saveSheet, data):
        for i in range(len(saveSheet)):
            dataSheetName = data['sheet'][i]
            if i == 0:
                saveSheet[i] = workbook.active
                saveSheet[i].title = dataSheetName
            else:
                saveSheet[i] = workbook.create_sheet(title = dataSheetName)
        return saveSheet

    def writeSheetHeader(self, saveSheet, data):
        for sheet in range(len(saveSheet)):
            dataHeader = data['header'][sheet]
            for column in range(len(dataHeader)):
                saveSheet[sheet]['%s1' % chr(ord('A')+column)] = dataHeader[column]
        return saveSheet

    def writeSheetContext(self, saveSheet, data):
        for sheet in range(len(saveSheet)):
            dataContext = data['context'][sheet]
            for row in range(len(dataContext)):
                for column in range(len(dataContext[row])):
                    saveSheet[sheet]['%s%d' % (chr(ord('A')+column), row+2)] = dataContext[row][column]
        return saveSheet

# Demo
if __name__ == '__main__':
    data = {'sheet': ['A Song of Ice and Fire', 'Harry Potter'],
            'header': [
                ['No.', 'Name', 'Author', 'Year', 'ISBN'],
                ['No.', 'Name', 'Author', 'Year', 'Rate']
                ],
            'context': [
                [
                    [1, 'A Game of Thrones', 'George R. R. Martin', '1996', '0-553-10354-7'],
                    [2, 'A Clash of Kings', 'George R. R. Martin', '1998', '0-553-10803-4'],
                    [3, 'A Storm of Swords', 'George R. R. Martin', '2000', '0-553-10663-5'],
                    [4, 'A Feast for Crows', 'George R. R. Martin', '2005', '0-553-80150-3'],
                    [5, 'A Dance with Dragons', 'George R. R. Martin', '2011', '978-0553801477'],
                    [6, 'The Winds of Winter', 'George R. R. Martin', 'forthcoming', None],
                    [7, 'A Dream of Spring', 'George R. R. Martin', 'forthcoming', None]
                    ],
                [
                    [10, "The Sorcerer's Stone", 'J.K. Rowling', '1999', '4.7'],
                    [11, 'The Chamber of Secrets', 'J.K. Rowling', '2000', '4.7'],
                    [12, 'The Prisoner of Azkaban', 'J.K. Rowling', '2001', '4.7'],
                    [13, 'The Goblet of Fire', 'J.K. Rowling', '2002', '4.7'],
                    [14, 'The Order of The Phoenix', 'J.K. Rowling', '2004', '4.7'],
                    [15, 'The Half-Blood Prince', 'J.K. Rowling', '2006', '4.5'],
                    [16, 'The Deathly Hallows', 'J.K. Rowling', '2009', '4.7']
                    ]
                ]
            }
    result = WriteExcel(data)
