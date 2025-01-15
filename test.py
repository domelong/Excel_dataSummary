import openpyxl as opx

class ExcelDataModel:
    def __init__(self, _path) -> None:
        """
        ExcelDataModel 数据模型对象 数据获取 数据分析
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
    path = r"C:\Users\d'm'l\Desktop\PythonTree\项目\25.1.12关于单工作簿多工作表同一区域零散数据汇总的解决方式\res\到达白班24年1至6月数据汇总.xlsx"
    area = {
        "片区1":"c5:c8",
        "负责人1":"d5:d8",
        "出勤人数1":"f5:f8",
        "休息人数1":"f5:f8",
        }
    ex = ExcelDataModel(path)
    data = ex.getdata2dict(area)
    print(data)
    ex.save2Excel(data)

def test2():
    data = [i for i in range(10)]
    

if __name__ == '__main__':
    test1()