import wbpy
import pandas as pd
from pprint import pprint
import openpyxl

wb = openpyxl.Workbook()

api = wbpy.IndicatorAPI()

# api:"SP.POP.TOTL","SI.POV.GINI","NY.GDP.PCAP.CD"
country_name = ["不丹", "保加利亚", "立陶宛", "西奈半岛（埃及）", "波兰", "黎巴嫩",\
                "老挝", "也门", "乌兹别克斯坦", "乌克兰", "白俄罗斯", "巴基斯坦",\
                "亚美尼亚", "伊拉克", "印度", "俄罗斯", "阿富汗", "蒙古国", "巴林",\
                "格鲁吉亚", "克罗地亚", "马来西亚", "约旦", "巴勒斯坦", "吉尔吉斯斯坦",\
                "匈牙利", "拉脱维亚", "文莱", "马尔代夫", "波黑", "塞尔维亚", "缅甸",\
                "阿尔巴尼亚", "沙特阿拉伯", "罗马尼亚", "马其顿", "塔吉克斯坦", "泰国",\
                "卡塔尔", "孟加拉", "尼泊尔", "捷克", "阿曼", "黑山", "土耳其", "斯洛伐克",\
                "新加坡", "阿拉伯联合酋长国", "阿塞拜疆", "摩尔多瓦", "叙利亚", "土库曼斯坦",\
                "哈萨克斯坦", "斯洛文尼亚", "伊朗", "印度尼西亚", "越南", "科威特", "希腊",\
                "柬埔寨", "塞浦路斯", "斯里兰卡", "以色列", "菲律宾", "爱沙尼亚"]

year = ["2010", "2011", "2012", "2013", "2014", "2015", "2016", "2017", "2018", "2019"]

iso_country_codes = ["BT", "BG", "LT", "EG", "PL", "LB", "LA", "YE", "UZ", "UA", "BY", "PK",\
                     "AM", "IQ", "IN", "RU", "AF", "MN", "BH", "GE", "HR", "MY", "JO", "PS",\
                     "KG", "HU", "LV", "BN", "MV", "BA", "RS", "MM", "AL", "SA", "RO", "MK",\
                     "TJ", "TH", "QA", "BD", "NP", "CZ", "OM", "ME", "TR", "SK", "SG", "AE",\
                     "AZ", "MD", "SY", "TM", "KZ", "SI", "IR", "ID", "VN", "KW", "GR", "KH",\
                     "CY", "LK", "IL", "PH", "EE"]

total_population = "SP.POP.TOTL"
dataset1 = api.get_dataset(total_population, iso_country_codes, date="2010:2019")
data1 = dataset1.as_dict()

gini_index = "SI.POV.GINI"
dataset2 = api.get_dataset(gini_index, iso_country_codes, date="2010:2019")
data2 = dataset2.as_dict()

avg_gdp = "NY.GDP.PCAP.CD"
dataset3 = api.get_dataset(avg_gdp, iso_country_codes, date="2010:2019")
data3 = dataset3.as_dict()

sheet = wb['Sheet'] 

sheet['A1'] = "国家"
sheet['B1'] = "年份"
sheet['C1'] = "人口规模"
sheet['D1'] = "GINI系数"
sheet['E1'] = "人均 GDP(现价美元)"

for i in range(len(country_name)):
  for j in range(len(year)):
    sheet['A' + str(i * len(year) + j + 2)] = country_name[i]
    sheet['B' + str(i * len(year) + j + 2)] = year[j]
    sheet['C' + str(i * len(year) + j + 2)] = data1[iso_country_codes[i]][year[j]]
    sheet['D' + str(i * len(year) + j + 2)] = data2[iso_country_codes[i]][year[j]]
    sheet['E' + str(i * len(year) + j + 2)] = data3[iso_country_codes[i]][year[j]]

wb.save('./output.xlsx')