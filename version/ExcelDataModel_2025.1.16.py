import openpyxl as opx
import os

class ExcelDataModel:
    def __init__(self, path) -> None:
        """
        ExcelDataModel 数据模型对象 数据获取 数据分析
        """
        self._path = path
        self.workbook = opx.load_workbook(self._path,read_only=True,data_only=True)

    def getdata(self, area:dict, sheets=[]) -> dict:
        """
        params: 
        area: dict 一维字典
        sheets: list 工作表名称列表
        """
        if not isinstance(area, dict): raise Exception("area参数类型错误")
        if not isinstance(sheets, list): raise Exception("sheets参数类型错误")
        if sheets: thissheets = [self.workbook[sheetname] for sheetname in sheets]
        else: thissheets = self.workbook.worksheets 
        data = {}
        for sheet in thissheets:
            sheetData = {}
            sheetname = sheet.title
            for map in area:
                cellindex = 1
                cellData = {}
                cells = sheet[area[map]]
                for celltuple in cells:
                    for cell in celltuple:
                        if not isinstance(cell, opx.cell.read_only.ReadOnlyCell): continue
                        celldatakey = map + "_" + str(cellindex)
                        cellData[celldatakey] = cell.value
                        cellindex += 1
                sheetData[map] = cellData
            data[sheetname] = sheetData
        return data

    def save_to_Excel(self, data:dict) -> list:
        """
        params:
        data: dict
        """
        if not isinstance(data, dict): raise Exception("data数据类型错误")
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
                    # 如果值是字典，递归调用 iter_data 函数
                    iter_data(value, new_path)
                else:
                    # 如果值不是字典，打印当前路径和对应的值
                    # print(f"Path: {new_path}, Value: {value}")
                    row = new_path + [value]
                    datalist.append(row)
                    ws.append(row)
        iter_data(data)
        wb.save(r".\newfile.xlsx")
        return datalist
    
    def getdata_for_workbooks(self, path:str):
        """
        params:
        path: str 多工作簿的目标文件夹
        """

        pass

def test1():
    path = r".\res"
    # 以单维字典 { 字段名:区域 } 的格式声明区域
    area = {
        "项":"d5:g7",
    }
    edm = ExcelDataModel(path)
    data = edm.getdata(area, ["2025.1.1", "2025.1.2"])
    edm.save_to_Excel(data)

if __name__ == '__main__':
    test1()