# excel汇总 
### ExcelDataModel(Excel数据模型对象)
### 属性:
1. workbook: openpyxl.Workbook对象
### 方法:
1. getdata(path, area, sheets=[]) -> dict
    *path: str 文件路径*
    *area: dict 一维字典 区域映射*
    *sheets: list 一维列表 工作表名称列表*
    </br>
2. getdata_for_workbooks(path, area, filenames=[], sheets=[]) -> list
    *path: str 文件路径*
    *area: dict 一维字典*
    *filenames: list 一维列表 文件名列表*
    *sheets: list 一维列表 工作表名称列表*
    </br>
3. save_to_Excel(datalist)
    *datalist: list 列表 元素为由getdata创建的dict*
    </br>
4. split_sheets_to_workbooks(path, startindex=1, endindex=None)
    *path: str 文件路径*
    *startindex: int 起始索引*
    *endindex: int 结束索引*
### 目标:
1. 
### 日志:
1. 2025.1.15新增功能 同一工作簿下多工作表同一区域数据汇总 
2. 2025.1.16 15:21新增功能 多工作簿数据汇总成单工作簿单工作表 
*ps: 本来是想汇总成多工作簿多工作表 但是后来想想没必要 既然是数据汇总再拆分成多工作表意义不大 哈哈*
3.
---
# excel拆分
### 目标:
1. 以目标工作簿为模板创建多工作表
### 日志:
1. 2025.1.17 13:18新增功能 工作簿内多工作表数据拆分成多工作簿
### 依赖: openpyxl 只支持.xlsx .xlrd
