from version.ExcelDataModel import ExcelDataModel

def test1():
    path = r".\res"
    # 以单维字典 { 字段名:区域 } 的格式声明区域
    area = {
        "项":"d5:g7",
    }
    edm = ExcelDataModel()
    data = edm.getdata_for_workbooks(path, area)
    edm.save_to_Excel(data)

if __name__ == '__main__':
    test1()