import openpyxl as opx
import os

class ExcelDataModel:
    def __init__(self, workbook=opx.Workbook()) -> None:
        """
        ExcelDataModel 数据模型对象 数据获取 数据分析
        """
        self.workbook = workbook

    # 获取单工作簿多工作表数据
    def getdata(self, workbook:opx.Workbook, area:dict, sheets=[]) -> dict:
        """
        params: 
        workbook: opx.Workbook workbook对象
        area: dict 一维字典
        sheets: list 工作表名称列表
        """
        if not isinstance(workbook, opx.Workbook): raise Exception("workbook参数类型应为workbook对象")
        if not isinstance(area, dict): raise Exception("area参数类型应为一维dict对象")
        if not isinstance(sheets, list): raise Exception("sheets参数类型应为一维list对象")
        if sheets: thissheets = [workbook[sheetname] for sheetname in sheets]
        else: thissheets = workbook.worksheets 
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
                        if not (isinstance(cell, opx.cell.read_only.ReadOnlyCell) or isinstance(cell, opx.cell.cell.Cell)): continue
                        celldatakey = map + "_" + str(cellindex)
                        cellData[celldatakey] = cell.value
                        cellindex += 1
                sheetData[map] = cellData
            data[sheetname] = sheetData
        return data

    # 获取多工作簿多工作表数据 后续调整异步
    def getdata_for_workbooks(self, path:str, area:dict, filenames=[], sheets=[]) -> list: 
        """
        params:
        path: str 文件夹路径
        area: dict 一维字典
        filename: list 工作簿名称列表
        sheets: list 工作表名称列表
        """
        isfile = os.path.isfile(path)
        if not isinstance(area, dict): raise Exception("area参数类型应为一维dict对象")
        if not isinstance(filenames, list): raise Exception("filenames参数类型应为一维list对象")
        if not isinstance(sheets, list): raise Exception("sheets参数类型应为一维list对象")
        if isfile: raise Exception("目标路径应为文件夹而不是文件")
        filepaths = []
        datalist = []
        listdir = filenames or os.listdir(path)
        for filename in listdir:
            filepath = path + "\\" + filename 
            if os.path.exists(filepath) and os.path.isfile(filepath):
                filepaths.append(filepath)
                wb = opx.load_workbook(filepath, read_only=True, data_only=True)
                data = self.getdata(wb, area, sheets)
                datalist.append(data)
            else: continue
        return datalist

    def save_to_Excel(self, datalist) -> None:
        """
        params:
        data: list
        """
        if not isinstance(datalist, list): datalist = [datalist]
        wb = self.workbook
        wb.create_sheet("newSheet")
        ws = wb.active
        ws.append([])
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
                    # datalist.append(row)
                    ws.append(row)
        for data in datalist: iter_data(data)
        wb.save(r".\newfile.xlsx")

def test1():
    path = r".\res\数据源_test1.xlsx"
    wb = opx.load_workbook(path)
    # path = r".\res"
    # 以单维字典 { 字段名:区域 } 的格式声明区域
    area = {
        "项":"d5:g7",
    }
    edm = ExcelDataModel()
    data = edm.getdata(wb, area)
    edm.save_to_Excel(data)

def test2():
    ws = opx.Workbook().active
    cell = ws["a1"]
    print(isinstance(cell, opx.cell.cell.Cell))
if __name__ == '__main__':
    test1()