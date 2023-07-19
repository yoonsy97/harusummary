import requests
import smtplib
import pandas as pd

import re
from bs4 import BeautifulSoup


import matplotlib.pyplot as plt
from wordcloud import WordCloud, STOPWORDS, ImageColorGenerator
from operator import itemgetter
from konlpy.tag import Twitter
from collections import Counter


from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import copy

from datetime import datetime

start_page=1
end_page=10

fontpath = 'C:/Windows/Fonts/batang.ttc'
file_path_email = r'C:\Users\User_\Desktop\haru\data\email.txt'
file_path_keyword = r'C:\Users\User_\Desktop\haru\data\keyword.txt'





def read_lines_from_file(file_path):
    lines = []
    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
            lines.append(line.strip())
    return lines

    

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



# 이메일로 검색 결과를 보내는 함수
def send_email(search_result,target_email):
    print("send_email : ", search_result)
    if search_result:
        # 이메일 설정 및 인증 정보
        smtp_server = 'smtp.naver.com'  # SMTP 서버 주소
        smtp_port = 587  # SMTP 포트 번호
        sender_email = '4_night@naver.com'  # 발신자 이메일 주소
        sender_password = 'gksmfEKd1!'  # 발신자 이메일 비밀번호
        
        now = datetime.now()
        mail_title = now.strftime('%Y%m%d') + '일자 - 네이버 검색결과'

        for li in target_email:
            receiver_email = li  # 수신자 이메일 주소

            # 이메일 생성


            msg = MIMEMultipart()
            msg.attach(MIMEText(search_result, 'html'))
            msg['Subject'] = mail_title
            msg['From'] = sender_email
            msg['To'] = receiver_email

            # 이메일 발송
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, receiver_email, msg.as_string())
    else:
        print("검색 결과가 없습니다.")


# 네이버에서 키워드로 검색하여 검색 결과를 가져오는 함수
# return dataframe 
def search_naver(keyword):
    title = ''
    link = ''
    data = ''

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.86 Safari/537.36'}
    search_pages = makeUrl(keyword, start_page, end_page)

    result_list = []

    for page in search_pages:
        data = requests.get(page, headers=headers)

        soup = BeautifulSoup(data.text, 'html.parser')

        list_news = soup.select('#main_pack > section > div > div.group_news > ul > li')

        for li in list_news:
            a = li.select_one('a.news_tit')
            info_group = li.select_one('.info_group')
            infos = info_group.select('a')

            press_info = infos[0] if len(infos) > 0 else None
            naver_info = infos[1].attrs['href'] if len(infos) > 1 else ''

            press_name = press_info.get_text(strip=True) if press_info else ''
            press_name = re.sub(r'(언론사|선정)', '', press_name)

            title = a.text
            link = naver_info
            press = press_name

            print(title, "  ", link, "  ", press)

            if title != '' and link != '' and press != '':
                result_list.append([title, link, press])
            else:
                pass

    df = pd.DataFrame(result_list, columns=['title', 'link', 'press'])
    df.drop_duplicates(['link'])
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




def execute():

    target_email = [""] # 대상 이메일
    target_keyword = [""] # 대상 기업 keyword 
    target_keyword= read_lines_from_file(file_path_keyword)
    target_email= read_lines_from_file(file_path_email)
    
    
    #엑셀파일 작성
    now = datetime.now()
    file_name = now.strftime('%Y%m%d')
    file_path = './excel/'+file_name+'.xlsx'

    writer = pd.ExcelWriter(file_path)

    total_result=pd.DataFrame()
    if target_keyword and target_email : 
        body = ""    
        for li in target_keyword:
            print("li : ", li)
            a=search_naver(li)
            body += write_email(li, a)
            print("a :  ",a)
            a.to_excel(writer,sheet_name=li,index=False)
            
            total_result = pd.concat([total_result, a], ignore_index=True)
        
            if not a.empty :
                make_wordCloud(a)

        send_email(body,target_email)
        #save_excel(total_result)
        writer.close()
    else : print("keyword , email ERROR") 


def save_excel(df):
    now = datetime.now()
    file_name = now.strftime('%Y%m%d')
    file_path = './excel/'+file_name+'.xlsx'
    df.to_excel(file_path,index=False)


if __name__ == '__main__':
    execute()
       
