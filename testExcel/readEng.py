import xlrd
from bs4 import BeautifulSoup  # HTML＆XML解析库，爬虫相关
import requests
import time
import xlsxwriter as xw
import win32com.client  # 用到其自带的朗读功能

engWordsWorkbook = xlrd.open_workbook('C:/Users/aby/Desktop/engWordsList.xlsx')
engWordsSheet = engWordsWorkbook.sheet_by_index(0)

transWorkbook = xw.Workbook('C:/Users/aby/Desktop/engWordsTrans.xlsx')
transSheet = transWorkbook.add_worksheet()

speaker = win32com.client.Dispatch('SAPI.SpVoice')

for row in range(0, engWordsSheet.nrows):
    # time.sleep(1)  # 避免因查询速度过快而导致加载错误
    word = engWordsSheet.cell(row, 0).value
    url = 'http://www.youdao.com/w/eng/' + word  # 利用爬虫在有道词典中查找选中的word

    webData = requests.get(url).text  # 获取查找结果，开始解析
    soup = BeautifulSoup(webData, 'html.parser')  # 之前使用lxml解析器报错，因为需要c语言环境
    meaning = str(soup.select('#phrsListTab > div.trans-container > ul > li')).replace('<li>', '').replace('</li>', '')
    trans = meaning[1:-1]
    print(word)
    transSheet.write(row, 0, word)
    transSheet.write(row, 1, trans)

    wordSegment = []  # 空变量
    spellingWord = []
    for i in word:
        wordSegment.append(i)
        wordSegment.append('-')
        spellingWord = ''.join(wordSegment)  # 单词拆开后拼接在一起

    speaker.Speak(str(word))  # 播放顺序：单词→拼写→单词→翻译
    speaker.Speak(str(spellingWord))
    speaker.Speak(str(word))
    speaker.Speak(str(trans))

transWorkbook.close()  # 有这一句才能生成文件
