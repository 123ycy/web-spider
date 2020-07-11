# -*- coding: utf-8 -*-
"""
Created on Sat May  9 19:54:58 2020

@author: 91579
"""

import wordcloud
import easygui as g
# import imageio
# import jieba
import xlrd            
            
class ReadExcel(object):
    def __init__(self, filename, sheetname):
        self.wb = xlrd.open_workbook(filename)
        self.sheet = self.wb.sheet_by_name(sheetname)
        
    def read_data_rows(self, name):
        table = self.sheet
        n_rows = table.nrows
        
        cases = []
        for i in range(n_rows):
            c_name = table.row_values(i, 2, 3)
            if c_name[0] == name:
                cases = table.row_values(i, 2)
        return cases
    
    def read_data_cols(self, col):
        table = self.sheet
        return table.col_values(col,1)
    
data_set = ReadExcel("豆瓣电影TOP250.xls","豆瓣电影TOP250")
name_set = data_set.read_data_cols(2)
get = False

g.msgbox("欢迎来到电影词云制作小程序！",ok_button="开始制作词云！", title="电影词云制作小程序")
while get == False:
    name = g.choicebox(msg="请选择你想制作的电影", title="电影词云制作小程序", choices = name_set)
    data_name = data_set.read_data_rows(name)
    data_name = str(data_name).strip('[]').replace('\'', '').replace(',', '').replace('主演:', '').replace('导演:','').replace('\\xa0','')
    # move = dict.fromkeys((ord(c) for c in u"\xa0\n\t"))
    # data_name = data_name.translate(move)
    
    wordc = wordcloud.WordCloud(background_color='white', font_path='msyh.ttc', width=800, height=600)
    wordc.generate(data_name)
    wordc.to_file("temp.png")
    a = g.buttonbox("生成的词云如下：",image="temp.png",choices=("另存图片为","重新制作词云", "结束"), title="词云图片成功生成")
    if a == "另存图片为":
        path = g.filesavebox(default=".png")
        wordcloud.to_file(path)
        get = True
    if a == "结束":
        get = False

















