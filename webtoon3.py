#from selenium.webdriver.common.keys import Keys
from selenium import webdriver
#from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import time
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
#글꼴설정
matplotlib.rcParams['font.family'] ='Malgun Gothic' # 폰트설정
matplotlib.rcParams['font.size'] = 15 # 글자크기
matplotlib.rcParams['axes.unicode_minus'] = False 

browser = webdriver.Chrome()

web_url = ['https://comic.naver.com/webtoon/detail?titleId=81482&no=719']

num_Data = []
page_num = []
last_num = []
try:
    # URL리스트 입력시 해당 웹툰 '페이지 번호' 및 '마지막화' Data를 반환하는 함수
    def Webtoon_URL(object): # object에 다가는 수집한 'web_url'리스트 대입
        global page_num
        global last_num
        for i in object:
            url_split = i.split('&')
            for j in url_split:
                url_Data = j.split('=')
                num_Data.append(int(url_Data[1]))
        for i in range(len(num_Data)):
            if i % 2 == 0:
                page_num.append(num_Data[i])
            else:
                last_num.append(num_Data[i])
        return page_num, last_num

    Webtoon_URL(web_url)

    # 크롤링 함수
    def crawlling(PN, No): 
        url="https://comic.naver.com/webtoon/detail?titleId=" + str(page_num[PN]) + "&no=" + str(No)
        browser.get(url)
        Name = browser.find_element(By.XPATH,"/html/body/div/div[2]/div/div[1]/div[1]/div[2]/h2/span[1]").text
        ge = browser.find_element(By.XPATH,"/html/body/div/div[2]/div/div[1]/div[1]/div[2]/p[2]/span[1]").text
        star = browser.find_element(By.XPATH,"/html/body/div/div[2]/div/div[1]/div[4]/dl/dd[1]/div/span[2]/strong").text
        star_ap = browser.find_element(By.XPATH,"/html/body/div/div[2]/div/div[1]/div[4]/dl/dd[1]/div/span[3]/em").text
        time.sleep(1)
        browser.switch_to.frame("commentIframe")
        reply_total=browser.find_element(By.XPATH,"/html/body/div/div/div[1]/span").text
        reply_total = reply_total.replace(',','')
        webtoon_data.append([No, Name, ge, float(star), int(star_ap), int(reply_total)])
        return webtoon_data

    # webtoon_data 리스트를 이용해 데이터프레임 만들고 엑셀 파일로 변경하는 함수
    def df_to_excel(data_list): # 인자를 리스트명으로 하기
        global pd_data
        columns = ['화','제목', '장르', '별점','별점 참여자 수', '댓글 수']
        pd_data = pd.DataFrame(data_list, columns = columns)
        pd_data_episode = pd_data.set_index('화') # '화'를 인덱스로 설정한 데이터프레임
        special_Chars = ['\ ', '| ','/ ','? ','" ','* ',': ','< ','> ']
        for i in special_Chars:
            data_list[0][1] = data_list[0][1].replace(i,'')
        name = data_list[0][1]
        pd_data_episode.to_excel('./'+name+'_web_toon.xlsx', index = True)
        return pd_data

    # 시각화 함수
    def Web_crawlling_plot(data_F):
        #sublpot설정
        fig, axs = plt.subplots(1, 2, figsize=(15, 5))
        #별점 참여자수 , 댓글 수 하나의 그래프로 표현
        axs[0].plot(data_F['화'], data_F['별점 참여자 수'],'b', label='별점 참여자 수')
        axs[0].plot(data_F['화'], data_F['댓글 수'],'r', label='댓글 수')
        axs[0].set_xlabel('화')
        axs[0].set_ylabel('별점 참여자 수')
        axs[0].set_title('별점 참여자수 와 댓글 수')
        axs[0].legend()
        #별점 그래프로 표현
        axs[1].plot(data_F['화'], data_F['별점'],'g',label='별점')
        axs[1].set_xlabel('화')
        axs[1].set_ylabel('별점')
        axs[1].set_title('별점')
        axs[1].legend()
        plt.suptitle(data_F['제목'][0], fontsize=30, fontweight ="bold")
        plt.show()

    for i in range(len(page_num)):
        webtoon_data = []
        if last_num[i] <= 100: # 100화 이하 작품
            for num in range(1, last_num[i]+1):
                crawlling(i, num)
            df_to_excel(webtoon_data)
            Web_crawlling_plot(pd_data)

        elif (last_num[i] > 100) & (last_num[i] <= 150): # 101화 ~ 150화까지 작품
            # 작품 앞 부분 크롤링
            for num in range(1,51):
                crawlling(i, num)
            time.sleep(2)
            # 작품 뒷 부분 크롤링
            for num in range(last_num[i]-49, last_num[i]+1):
                crawlling(i, num)
            df_to_excel(webtoon_data)
            Web_crawlling_plot(pd_data)

        elif (last_num[i] > 150) & (last_num[i] <= 200): # 101화 ~ 150화까지 작품
            # 작품 앞 부분 크롤링
            for num in range(1,76):
                crawlling(i, num)
            time.sleep(2)
            # 작품 뒷 부분 크롤링
            for num in range(last_num[i]-74, last_num[i]+1):
                crawlling(i, num)
            df_to_excel(webtoon_data)
            Web_crawlling_plot(pd_data)

        else: # 201화 이상 작품
            # 작품 앞 부분 크롤링
            for num in range(1,101):
                crawlling(i, num)
            time.sleep(2)
            # 작품 뒷 부분 크롤링
            for num in range(last_num[i]-99, last_num[i]+1):
                crawlling(i, num)
            df_to_excel(webtoon_data)
            Web_crawlling_plot(pd_data)
except KeyboardInterrupt:
    print('KeyboardInterrupt')
except NoSuchElementException:
    print('NoSuchElementException')
#except SystemExit:
    #print('SystemExit')
#except ZeroDivisionError:
    #print('ZeroDivisionError')
#except RemoteDisconnected:
    #print('RemoteDisconnected')
except ProtocolError:
    print('ProtocolError')
except ConnectionRefusedError:
    print('ConnectionRefusedError')
except MaxRetryError:
    print('MaxRetryError')
#예외처리 중간에 끊겼을때 뜨는 에러, 차단당했을때 뜨는 에러, 읽어드릴 데이터가 없을때 뜨는 에러 등 발생할 수 있는 다양한 에러에 대비한 
#예외처리입니다.