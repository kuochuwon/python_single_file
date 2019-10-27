# /usr/bin/env python
# -*- coding: UTF-8 -*-

"""
progress:
1. 準備設定def: get url json
2. 整理資料格式
"""
import requests
import sys
import numpy as np
from matplotlib import pyplot as plt
# import matplotlib as mpl


class get_NBA_score_data():
	# elf.gamemode = {'Playoffs': '4', 'Season':'2'}
	def get_json_raw_data(self, gamemode_code):
		each_url = 'https://tw.global.nba.com/stats2/league/teamstats.json?conference=All&division=All&locale=zh_TW&season=2018&seasonType='+gamemode_code
		data_response = requests.request(method="GET", url=each_url)
		raw_data = data_response.json()
		return raw_data

	def get_statistical_data(self, raw_data):
		NBA_Teams_datalist = raw_data['payload']['teams']
		tppct = {}  #三分球
		turnoversPg = {}  #失誤
		pointsPg = {}  #得分
		for each_team_dict in NBA_Teams_datalist:
			team_name = each_team_dict['profile']['name']  #隊名: code = name_us; name = name_zh
			#print team_name
			profile_dict = each_team_dict['profile']
			statAverage_dict = each_team_dict['statAverage']
			statTotal_dict = each_team_dict['statTotal']
			#standings_dict = each_team_dict['standings'] #似乎只有None
			#for statAverage_dict_key, statAverage_dict_value in statAverage_dict.items():	
			#for profile_dict_key, profile_dict_value in statTotal_dict.items():
			#print team_name
			tppct.update({team_name:statAverage_dict['tppct']})#將每一隊的tppct(三分球命中率)數據新增到tppct
			turnoversPg.update({team_name:statAverage_dict['turnoversPg']})
			pointsPg.update({team_name:statAverage_dict['pointsPg']})
		return tppct, turnoversPg, pointsPg

	def show_bar_plot(self, tppct, ylabel, title, file_name):
		tppct_label = tppct.keys()
		tppct_data = tppct.values()

		plt.figure(figsize=(14,5))#要先宣告figsize，若之後才宣告會拆成兩張圖
		plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei'] #顯示中文
		plt.rcParams['axes.unicode_minus'] = False #還不知確切作用
		plt.ylabel(ylabel)
		plt.title(title)

		c = np.arange(len(tppct_label))# type: numpy.ndarray
		#####為每個條形圖新增數值標籤#####
		for a,b in zip(c,tppct_data):#a = number of key, b = each data
			#print b
			plt.text(a, b+0.5, b, ha='center', va= 'center',fontsize=10)#1st and 2nd 參數用來設定字體位置的x, y座標(非輸出資料本身)
		#####為每個條形圖新增數值標籤#####
		plt.bar(range(len(tppct_data)), tppct_data, tick_label = tppct_label)#參數分別為X, Y and X軸標籤
		plt.savefig(file_name, dpi=200)#指定分辨率，檔案類型預設 = png
		plt.show()#show圖片以後，plt的圖片會沒辦法另外存檔
		plt.close()

	def deviation_data(self, pla_data, sea_data):
		Pla_Sea_deviation = {}
		for each_team, each_tppct in pla_data.items():
			deviation = abs(each_tppct - sea_data[each_team]) #計算季後賽與例行賽的數據差值
			#print each_team, deviation
			Pla_Sea_deviation.update({each_team : deviation})
		return Pla_Sea_deviation
	
class execution():
	print 'parse binging...'
	my_obj = get_NBA_score_data()
	Playoffs_raw_data = my_obj.get_json_raw_data('4') #2 = Seasons = Sea, 4 = Playoffs = Pla
	Seasons_raw_data = my_obj.get_json_raw_data('2') #2 = Seasons = Sea, 4 = Playoffs = Pla
	#print raw_data
	pla_tppct, pla_turnoversPg, pla_pointsPg = my_obj.get_statistical_data(Playoffs_raw_data)
	sea_tppct, sea_turnoversPg, sea_pointsPg = my_obj.get_statistical_data(Seasons_raw_data)
	Pla_Sea_tppct_deviation = my_obj.deviation_data(pla_tppct, sea_tppct)
	#print tppct
	my_obj.show_bar_plot(pla_tppct, u'三分球命中率', u'各球隊三分球命中率數據 季後賽', 'pla_tppct.png')
	my_obj.show_bar_plot(sea_tppct, u'三分球命中率', u'各球隊三分球命中率數據 例行賽', 'sea_tppct.png')
	my_obj.show_bar_plot(Pla_Sea_tppct_deviation, u'三分球命中率差值', u'各球隊三分球命中率差值數據 季後 VS 例行', 'pla_sea_tppct.png')

	#my_obj.show_bar_plot(turnoversPg, u'平均失誤次數', u'各球隊失誤數據', 'turnoversPg.png')
	#my_obj.show_bar_plot(pointsPg, u'平均得分', u'各球隊平均得分數據', 'pointsPg.png')



if __name__ == "__main__":
    execution()#起始點