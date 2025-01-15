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
                cellindex = 1
                cellData = {}
                cells = sheet[area[map]]
                for celltuple in cells:
                    cell = celltuple[0]
                    celldatakey = map + "_" + str(cellindex)
                    cellData[celldatakey] = cell.value
                    cellindex += 1
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
        datalist = []
        def iter_data(dictionary, path=[]):
            """
            深度优先遍历字典的函数
            :param dictionary: 要遍历的字典
            :param path: 存储当前路径，默认为空列表
            """
            for key, value in dictionary.items():
                # 将当前键添加到路径中
                new_path = path + [key]
                if isinstance(value, dict):
                    # 如果值是字典，递归调用 dfs_dict 函数
                    iter_data(value, new_path)
                else:
                    # 如果值不是字典，打印当前路径和对应的值
                    # print(f"Path: {new_path}, Value: {value}")
                    row = new_path + [value]
                    datalist.append(row)
                    ws.append(row)
        iter_data(data)
                        
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
    # print(data)
    ex.save2Excel(data)

def test2():
    data = [i for i in range(10)]
    

if __name__ == '__main__':
    test1()