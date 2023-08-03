# -*- coding: utf-8 -*-

from dataclasses import dataclass, field
import xlrd
import xlwt
from tqdm import tqdm

@dataclass
class ExcelList:
    page : int = field(init = False)
    x_position : float = field(init = False)
    y_position : float = field(init = False)
    height : float = field(init = False)
    width : float = field(init = False)
    color : list[float] = field(init = False)
    fontcolor : str = field(init = False)
    fontsize : float = field(init = False)
    text : str = field(init = False)

def LoadExcelListFromExcel(excel_path):
    file = xlrd.open_workbook(excel_path)
    sheet = file.sheet_by_index(0)
    nrows = sheet.nrows
    data_list = list()
    with tqdm(total = nrows) as pbar:
        pbar.set_description('\tHave loaded：')
        for i in range(1, nrows):
            ans_data = ExcelList()
            ans_data.page = int(sheet.cell(i, 0).value)
            ans_data.x_position = float(sheet.cell(i, 1).value)
            ans_data.y_position = float(sheet.cell(i, 2).value)
            ans_data.height = float(sheet.cell(i, 3).value)
            ans_data.width = float(sheet.cell(i, 4).value)
            color = str(sheet.cell(i, 5).value).split(',')
            ans_data.color = [float(item) for item in color]
            ans_data.fontcolor = str(sheet.cell(i, 6).value)
            ans_data.fontsize = str(sheet.cell(i, 7).value)
            ans_data.text = str(sheet.cell(i, 8).value)
            data_list.append(ans_data)
            pbar.update(1)
    return data_list

def WriteExcelListToExcel(data_list, excel_path):
    file = xlwt.Workbook()
    sheet = file.add_sheet("Sheet 0")
    names = ["PAGE", "Rectangle X Position (IN)", "Rectangle  Y Position (IN)", "Rectangle High (IN)", "Rectangle  Weight (IN)", "Background Color", "Caption (Text)"]
    for (i, name) in enumerate(names):
        sheet.write(0, i, label = name)
    with tqdm(total = len(data_list)) as pbar:
        pbar.set_description('\tHave stored：')
        for (i, ans_data) in enumerate(data_list):
            sheet.write(i+1, 0, label = ans_data.page)
            sheet.write(i+1, 1, label = ans_data.x_position)
            sheet.write(i+1, 2, label = ans_data.y_position)
            sheet.write(i+1, 3, label = ans_data.height)
            sheet.write(i+1, 4, label = ans_data.width)
            color = ans_data.color
            sheet.write(i+1, 5, label = "{},{},{}".format(color[0], color[1], color[2]))
            sheet.write(i+1, 6, label = ans_data.fontcolor)
            sheet.write(i+1, 7, label = ans_data.fontsize)
            sheet.write(i+1, 8, label = ans_data.text)
            pbar.update(1)
    file.save(excel_path)
    
def TransformExcelListToDict(data_list):
    data_dict = dict()
    for ans_data in data_list:
        page = ans_data.page - 1
        ans_data_dict = dict()
        ans_data_dict['x_position'] = float(ans_data.x_position)
        ans_data_dict['y_position'] = float(ans_data.y_position)
        ans_data_dict['height'] = float(ans_data.height)
        ans_data_dict['width'] = float(ans_data.width)
        ans_data_dict['color'] = ans_data.color
        ans_data_dict['fontcolor'] = str(ans_data.fontcolor)
        ans_data_dict['fontsize'] = float(ans_data.fontsize)
        ans_data_dict['text'] = ans_data.text
        if(page not in data_dict):
            data_dict[page] = [ans_data_dict, ]
        else:
            data_dict[page].append(ans_data_dict)
    return data_dict