# -*- coding: utf-8 -*-
"""
多個關鍵字搜尋，方法是if a in item and b in item do something
a, b為不同關鍵字，可藉由輸入兩個以上的較短關鍵字，達到精確搜尋又不需要記得完整關鍵字
降低搜尋難度

程式邏輯:
先定義一個function，列出目標資料夾，並以for loop iteration
for loop內是開啟PTT讀取內文的function
由於Win32com的PPT開啟需輸入絕對路徑，開啟檔案的function要微調
"""
import win32com
from genericpath import exists
from win32com.client import Dispatch, constants
import sys,os 
import io

class PowerPoint_keyword_search():
    def __init__(self):
        self.target_path = ("E:\\7PCT\\Taiwanese\\TEST")
        self.items = os.listdir(self.target_path)
        search_mode = int(raw_input('若關鍵字為1組，請輸入1，若關鍵字為2組，請輸入2。'))
        if search_mode == 1:
            #self.keyword1 = u'鐘聲大鳴'
            keyword1 = raw_input('請輸入關鍵字:') #要想辦法轉換成unicode
            keyword1_big5 = keyword1.decode('big5')# work!
            #keyword1_big5 = keyword1.decode('big5').encode('utf-8')
            # not work，big5不能轉換成utf-8? error code = UnicodeDecodeError: 'ascii' codec can't decode byte 0xe9 in position 0: ordinal not in range(128)
            print 'keyword: ', keyword1_big5
            file_list = self.ForSpecificFileInRootDir(self.items)
            self.ExtractWords(file_list, search_mode, keyword1_big5)
        elif search_mode == 2:
            keyword1 = raw_input('請輸入關鍵字1:')
            keyword2 = raw_input('請輸入關鍵字2:')
            keyword1_big5 = keyword1.decode('big5') #decode = 已知編碼是為何，告訴程式正確編碼；encode = 強迫程式編碼為指定編碼
            keyword2_big5 = keyword2.decode('big5')
            """
            keyword1_big5 = '你'
            keyword1_big5 = keyword1_big5.decode('utf-8')
            keyword2_big5 = u'hello'
            """
            print 'keyword1: ', keyword1_big5, 'keyword2:', keyword2_big5
            file_list = self.ForSpecificFileInRootDir(self.items)
            self.ExtractWords(file_list, search_mode, keyword1_big5, keyword2_big5)

    def ForSpecificFileInRootDir(self, items):
        file_list = []
        for names in items:
            if names[0]!='~':
               file_list.append(names)
        return file_list

    def ExtractWords(self, file_list, search_mode, keyword1 = None, keyword2 = None):
        cunt = 1
        if search_mode == 1:
            for file in file_list:
                #print file
                folder = self.target_path + '\\'
                file_name_complete = folder + file
                print file_name_complete, '\n'
                ppt = win32com.client.Dispatch('PowerPoint.Application')
                #ppt.Visible = 1
                pptSel = ppt.Presentations.Open(file_name_complete)
                slide_count = pptSel.Slides.Count
                for i in range(1,slide_count + 1):
                    shape_count = pptSel.Slides(i).Shapes.Count
                    for j in range(1,shape_count + 1):
                        if pptSel.Slides(i).Shapes(j).HasTextFrame:
                            s=pptSel.Slides(i).Shapes(j).TextFrame.TextRange.Text
                            if s!='' and keyword1 in s:
                               print s, '======='
                print '==========complete', cunt, '/', len(file_list), 'cycle=========='
                cunt+=1
                pptSel.Close()
        elif search_mode == 2:
            for file in file_list:
                folder = self.target_path + '\\'
                file_name_complete = folder + file
                print file_name_complete, '\n'
                ppt = win32com.client.Dispatch('PowerPoint.Application')
                #ppt.Visible = 1
                pptSel = ppt.Presentations.Open(file_name_complete)
                slide_count = pptSel.Slides.Count
                for i in range(1,slide_count + 1):
                    shape_count = pptSel.Slides(i).Shapes.Count
                    for j in range(1,shape_count + 1):
                        if pptSel.Slides(i).Shapes(j).HasTextFrame:
                            s=pptSel.Slides(i).Shapes(j).TextFrame.TextRange.Text
                            if s!='' and keyword1 in s and keyword2 in s:
                                print s, '======='
                print '==========complete', cunt, '/', len(file_list), 'cycle=========='
                cunt+=1
                pptSel.Close()

if __name__ == "__main__":
    PowerPoint_keyword_search()#起始點



