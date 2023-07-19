
import requests
import smtplib
import re
import openai
import os
import json

from datetime import datetime

import pandas as pd
from bs4 import BeautifulSoup

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import matplotlib.pyplot as plt
from wordcloud import WordCloud, STOPWORDS, ImageColorGenerator
from operator import itemgetter
from konlpy.tag import Twitter
from collections import Counter

import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5 import uic
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QMovie

import FinanceDataReader as fdr

import re


# 전역변수

form_class = uic.loadUiType("untitled.ui")[0]

fontpath = 'C:/Windows/Fonts/batang.ttc'
start_page=1
end_page=10

target_email = [] # 대상 이메일
target_keyword = [] # 대상 기업 keyword 

access_token_path= '\kakao_code.json'

gif_path = r"C:\Users\User_\Desktop\haru\img\emot_001_x3-1.gif"


class MyListModel(QAbstractListModel):
    def __init__(self, data=[], parent=None):
        super().__init__(parent)
        self.data = data
    
    def rowCount(self, parent=None):
        return len(self.data)
    
    def data(self, index, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            return str(self.data[index.row()])
        
        return None

def get_StockList():
    df_krx= fdr.StockListing('KRX')
    return df_krx

class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)
        
        global target_keyword
        global target_email

        # 기존 코드에서 사용하던 버튼 연결 코드
        self.pushButton.clicked.connect(self.execute)
        self.pushButton_2.clicked.connect(self.button2Function)
        self.pushButton_3.clicked.connect(self.button3Function)
        self.pushButton_4.clicked.connect(self.button4Function)
        self.pushButton_5.clicked.connect(self.button5Function)

        #버튼에 기능을 연결하는 코드
        # self.setWindowTitle("Window Title")

       
        word_list= get_StockList().Name
        completer = QCompleter(word_list)
        self.lineEdit.setCompleter(completer)
        box = QVBoxLayout()        
        box.addWidget(self.lineEdit)
        self.setLayout(box)


        
        
        self.gif_path = r"C:\Users\User_\Desktop\haru\img\emot_001_x3-1.gif"

        self.label = QLabel(self)
        self.label.setScaledContents(True)
        self.label.setGeometry(10, 250, 250, 250)
        
        self.movie = QMovie(self.gif_path, QByteArray(), self)
        self.movie.setCacheMode(QMovie.CacheAll)
        
        self.label.setMovie(self.movie)
        self.movie.start()

        target_keyword=self.read_lines_from_file('.\data\keyword.txt')
        target_email=self.read_lines_from_file('.\data\email.txt')

        # 기존 코드에서 사용하던 리스트뷰 갱신 코드
        self.update_list_view(target_keyword)
        self.update_email_list_view(target_email)

        self.setWindowTitle('하루요약')
        #self.setCentralWidget(self.lineEdit)
    
        self.show()

    def read_lines_from_file(self,file_path):
        lines = []
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                lines.append(line.strip())
        return lines



    def execute(self):
        file_path_keyword = '.\data\keyword.txt'
        file_path_email = '.\data\email.txt'        
        
        try : 
            write_list_to_txt(target_keyword, file_path_keyword)        
            write_list_to_txt(target_email, file_path_email) 

            self.show_alert("알림","텍스트파일에 저장되었습니다.")

        except : self.show_alert("경고","텍스트파일에 저장되지 않았습니다.")
             

        # if not target_keyword and not target_email : 
        #     body = ""    
        #     for li in target_keyword:
        #         a=search_naver(li)
        #         body += write_email(li, a)
        #         make_wordCloud(a)

        #     #send_email(body)
        # else : print("keyword , email ERROR") 

    # 알림창 표시 함수
    def show_alert(self,title, message):

        QMessageBox.information(self,title,message)
       
    


    def update_list_view(self, my_list):
        self.model = MyListModel(my_list)
        self.company_list.setModel(self.model)
    
    def update_email_list_view(self, my_list):
        self.model = MyListModel(my_list)
        self.email_list.setModel(self.model)

        #btn_1이 눌리면 작동할 함수
    def button2Function(self) :
        self.add_to_keyword_list()

    #btn_2가 눌리면 작동할 함수
    def button3Function(self) :
        self.add_to_email_list()
      

    def button4Function(self) :
        self.delete_selected_item()

    def button5Function(self) :
        self.delete_selected_email()


    def delete_selected_item(self):
        selected_indexes = self.company_list.selectedIndexes()
        if selected_indexes:
            selected_row = selected_indexes[0].row()
            del target_keyword[selected_row]
            
            self.update_list_view(target_keyword)
    
    def delete_selected_email(self):
        selected_indexes = self.email_list.selectedIndexes()
        if selected_indexes:
            selected_row = selected_indexes[0].row()
            del target_email[selected_row]
            
            self.update_email_list_view(target_email)

    def add_to_keyword_list(self):
        new_items = self.lineEdit.text()  # QPlainTextEdit의 텍스트 가져오기
        #new_items = text.split('\n')  # 엔터 기준으로 리스트화
        #print("new_items : ",type(new_items), new_items)
        #print("target_keyword : ",target_keyword)
        #print("new_items not in target_keyword : ",new_items in target_keyword)
        if new_items!='' and new_items not in target_keyword :
            target_keyword.append(new_items)  # 기존 리스트에 새로운 아이템 추가

            print(target_keyword)  # 추가된 리스트 출력
            self.update_list_view(target_keyword)
            self.lineEdit.clear()

    def add_to_email_list(self):
        #이메일 정규식
        a  = re.compile('^[a-zA-Z0-9+-\_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$')

        text= self.plainTextEdit_2.toPlainText()
        new_items = text.split('\n')  # 엔터 기준으로 리스트화
        
        for li in new_items:
            email_test = a.match(li)
            if email_test is not None:
                target_email.append(li)  # append 일 경우 str을 고대로 넣음
                self.update_email_list_view(target_email)

        self.plainTextEdit_2.clear()

    def setupCompleter(self):
        word_list = ["apple", "banana", "cherry", "grape", "kiwi"]
        self.model.setStringList(word_list)
        self.completer.setModel(self.model)
        self.completer.setCompletionMode(QCompleter.PopupCompletion)
        self.completer.setCaseSensitivity(False)
        self.completer.setWidget(self)
        self.completer.setWrapAround(False)


def write_list_to_txt(lst, file_path):
    with open(file_path, 'w', encoding='utf-8') as file:
        for item in lst:
            file.write(str(item) + '\n')



# OpenAI 인터페이스를 감싸는 함수, 매개변수는 프롬프트, 반환 값은 해당 결과
def get_completion(prompt, model="gpt-3.5-turbo"):
    '''
    prompt: 해당 프롬프트
    model: 사용되는 모델, 기본 값은 gpt-3.5-turbo(ChatGPT), 베타 테스트에 참여한 사용자는 gpt-4를 선택할 수 있습니다.
    '''
    messages = [{"role": "user", "content": prompt}]
    response = openai.ChatCompletion.create(
        model=model,
        messages=messages,
        temperature=0, # 모델 출력의 온도 계수, 출력의 랜덤성을 제어합니다
    )
    # OpenAI의 ChatCompletion 인터페이스를 호출합니다
    return response.choices[0].message["content"]



# 페이지 url 형식에 맞게 바꾸어 주는 함수 만들기
  #입력된 수를 1, 11, 21, 31 ...만들어 주는 함수tka
def makePgNum(num):
    if num == 1:
        return num
    elif num == 0:
        return num+1
    else:
        return num+9*(num-1)


# 크롤링할 url 생성하는 함수 만들기(검색어, 크롤링 시작 페이지, 크롤링 종료 페이지)
def makeUrl(search,start_pg,end_pg):
    if start_pg == end_pg:
        start_page = makePgNum(start_pg)
        url = "https://search.naver.com/search.naver?where=news&sm=tab_pge&query=" + search + "&start=" + str(start_page)
        
        return url
    else:
        urls= []
        for i in range(start_pg,end_pg+1):
            page = makePgNum(i)
            url = "https://search.naver.com/search.naver?where=news&sm=tab_pge&query=" + search + "&start=" + str(page)
            urls.append(url)

        return urls



def analysis_news() : 
    #openapi 키 
    openai.api_key = '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'
    
    

# 네이버에서 키워드로 검색하여 검색 결과를 가져오는 함수
# return dataframe 
def search_naver(keyword):

    #변수 초기화 
    title=''
    link=''
    data=''

    headers = {'User-Agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.86 Safari/537.36'}
    search_pages=makeUrl(keyword, start_page, end_page)
    
    result_list = []
    
    # 페이지별로 검색 
    for page in search_pages : 

        data = requests.get(page,headers=headers)

        soup = BeautifulSoup(data.text, 'html.parser')

        list_news = soup.select('#main_pack > section > div > div.group_news > ul > li')

        #검색된 페이지에서 태그별로
        for li in list_news:
            a = li.select_one('a.news_tit')
            info_group = li.select_one('.info_group')
            infos = info_group.select('a')

            try : 
                press_info = infos[0]
                naver_info = infos[1].attrs['href']

            except : 
                pass

            press_name = press_info.get_text(strip=True) if press_info else ''

            press_name = re.sub(r'(언론사|선정)', '', press_name)

            title = a.text
            link = naver_info
            press=press_name

            print(title, "  ", link, "  ",press)

            if title!='' and link != '' and press !='' :
                result_list.append([title, link, press])  # 딕셔너리 대신 리스트로 추가
            else : pass
    
    df = pd.DataFrame(result_list, columns=['title', 'link','press'])  # result_list를 2차원 배열로 변환

    df.drop_duplicates(['link'])     #중복 제거
    
    print(df['press'].value_counts())
    return df

#검색결과 이메일 전송
def write_email(keyword, article):
    data = f'<a >{keyword}</a><br>'
    print(data)
    for index, row in article.iterrows():
        title = row['title']
        link = row['link']
        data += f'<a href="{link}" target="_blank" style="text-decoration: underline;">{title}</a><br>'
    return data


# 카카오톡 json 토큰 파일 저장 
def generate_token():
    url = 'https://kauth.kakao.com/oauth/token'
    client_id = '자신의 REST 키값'
    redirect_uri = 'https://example.com/oauth'
    code = '자신의 CODE 값'

    data = {
        'grant_type':'authorization_code',
        'client_id':client_id,
        'redirect_uri':redirect_uri,
        'code': code,
        }

    response = requests.post(url, data=data)
    tokens = response.json()

    #발행된 토큰 저장
    with open("token.json","w") as kakao:
        json.dump(tokens, kakao)




# 액세스 토큰 재발급 함수
def refresh_token(refresh_token):
    client_id='730f60d67658e98f9cd2c04f8571901a'
    url = "https://kauth.kakao.com/oauth/token"
    
    data = {
        "grant_type" : "refresh_token",
        "client_id" : client_id,
        "refresh_token" : refresh_token
    }
    response = requests.post(url, data=data)
    tokens = response.json()

    with open("kakao_code.json", "w") as fp: # 재발급받은 토큰을 파일에 저장
        json.dump(tokens, fp)

    return tokens




#검색결과 카카오톡 전송
def write_Kakao_message(article):
    #발행한 토큰 불러오기
    with open("kakao_code.json","r") as kakao:
        tokens = json.load(kakao)

    url="https://kapi.kakao.com/v2/api/talk/memo/default/send"

    headers={
        "Authorization" : "Bearer " + tokens["access_token"]
    }

    data = {
       "object_type" : "text",
                                     "text" : "Test "+ str(datetime.now().strftime("%Y-%m-%d %H:%M"))+ "\n" +"https://news.naver.com/",
                                     "link" : {
                                                 "web_url" : "https://www.google.co.kr/search?q=drone&source=lnms&tbm=nws"
                                              },
                                    "button_title" : "확인하기"
    }
    
    data = {'template_object': json.dumps(data)}
    response = requests.post(url, headers=headers, data=data)
    response.status_code


    
    
# 이메일로 검색 결과를 보내는 함수
def send_email(search_result):
    print("send_email : ", search_result)
    if search_result:
        # 이메일 설정 및 인증 정보
        smtp_server = 'smtp.naver.com'  # SMTP 서버 주소
        smtp_port = 587  # SMTP 포트 번호
        sender_email = '4_night@naver.com'  # 발신자 이메일 주소
        sender_password = 'gksmfEKd1!'  # 발신자 이메일 비밀번호
        
        for li in target_email:
            receiver_email = li  # 수신자 이메일 주소

            # 이메일 생성
            


            msg = MIMEMultipart()
            msg.attach(MIMEText(search_result, 'html'))
            msg['Subject'] = '네이버 검색 결과'
            msg['From'] = sender_email
            msg['To'] = receiver_email

            # 이메일 발송
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, receiver_email, msg.as_string())
    else:
        print("검색 결과가 없습니다.")

        
        
def parse_string_to_list(input_string):
    # 입력받은 문자열을 쉼표를 기준으로 분리하여 리스트로 저장
    items = input_string.split(',')
    
    # 분리된 리스트의 각 아이템을 공백을 제거하여 리스트로 변환
    result_list = [item.strip() for item in items]
    
    return result_list
        
# 워드클라우드만들기
def make_wordCloud(a):
    # 워드클라우드 만들

    day_text = " ".join(li for li in a['title'].astype(str))
    
    nlpy = Twitter()
    nouns = nlpy.nouns(day_text)

    count = Counter(nouns)

    tag_count = []
    tags = []

    for n, c in count.most_common(100):
        dics = {'tag': n, 'count': c}
        if len(dics['tag']) >= 2 and len(tags) <= 49:

            tag_count.append(dics)
            tags.append(dics['tag'])

    for tag in tag_count:
        print(" {:<14}".format(tag['tag']), end='\t')
        print("{}".format(tag['count']))

    plt.subplots(figsize=(15,15))
    wordcloud = WordCloud(background_color='white', width = 1500,height=1500,
                          max_words=100,stopwords=None, 
                          font_path=fontpath).generate(day_text)
    plt.axis('off')
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.show()


# 메인 코드

if __name__ == '__main__':
    app = QApplication(sys.argv)
    widgetSample = WindowClass()
    widgetSample.show()
    sys.exit(app.exec_())

