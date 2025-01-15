import openpyxl as opx

class ExcelDataModel:
    def __init__(self, _path) -> None:
        """
        
        """
        self._path = _path
        self.workbook = opx.load_workbook(self._path,read_only=True,data_only=True)
    
    def getdata2dict(self, area):
        """
        params: 
        area: dict
        """
        if not isinstance(area, dict): raise Exception("area类型错误")
        data = {}
        for sheet in self.workbook.worksheets:
            sheetData = {}
            sheetname = sheet.title
            for map in area:
                cellData = []
                cells = sheet[area[map]]
                for celltuple in cells:
                    cell = celltuple[0]
                    cellData.append(cell.value)
                sheetData[map] = cellData
            data[sheetname] = sheetData
        return data
        
    def save2Excel(self, data):
        """
        params:
        data: dict
        """
        wb = opx.Workbook()
        ws = wb.active
        ws.append([])
        datakeys = list(data.keys())
        for v1count in range(len(datakeys)):
            # v1值
            datakey = datakeys[v1count]
            v1 = data[datakey]
            v2keys = list(v1.keys())
            for v2count in range(len(v2keys)):
                # v2值
                v2key = v2keys[v2count]
                v2 = v1[v2key]
                for v3count in range(len(v2)):
                    # v3值
                    v3 = v2[v3count]
                    celldata = [datakey, v2key, v3]
                    ws.append(celldata)
                
        wb.save(r".\newfile.xlsx")
        
        
def test1():
    path = r"C:\Users\d'm'l\Desktop\PythonTree\数据源.xlsx"
    area = {
        "项":"d5:d7",
        "数量":"e5:e7",
        "单价":"f5:f7",
        "总价":"g5:g7",
        }
    ex = ExcelDataModel(path)
    data = ex.getdata2dict(area)
    ex.save2Excel(data)

if __name__ == '__main__':
    test1()