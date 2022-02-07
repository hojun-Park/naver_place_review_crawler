import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QDesktopWidget, QLabel, QLineEdit
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QThread, pyqtSignal, pyqtSlot
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from bs4 import BeautifulSoup
from openpyxl import Workbook
import re

global query
query=''

def rmEmoji(inputString):
    text = u''+inputString

    emoji_pattern = re.compile("["
            u"\U0001F600-\U0001F64F"  # emoticons
            u"\U0001F300-\U0001F5FF"  # symbols & pictographs
            u"\U0001F680-\U0001F6FF"  # transport & map symbols
            u"\U0001F1E0-\U0001F1FF"  # flags (iOS)
                            "]+", flags=re.UNICODE)
    new_text = re.sub(r"[-=+,#/\?:^$.@*\"※~&%ㆍ!』\\‘|\(\)\[\]\<\>`\'…》]","",emoji_pattern.sub(r'', text))
    return  new_text # no emoji

class naverCrawling(QThread):
    state = pyqtSignal(str)

    def __init__(self): 
        super().__init__()
        

    def run(self):
        self.query=query 
        chrome_driver_path="C:/Users/hooju/OneDrive/바탕 화면/종설/크롤링 코드/gui/chromedriver.exe"
        url = "https://map.naver.com/"

        browser = webdriver.Chrome(chrome_driver_path)
        browser.get(url) # 네이버 맵 url로 들어가기

        time.sleep(5)
        
        browser.find_element_by_xpath('/html/body/app/layout/div[3]/div[2]/shrinkable-layout/div/app-base/search-input-box/div/div[1]/div/input').send_keys(self.query) # 네이버 맵 검색창에 식당명 입력
        self.state.emit(self.query+"을(를) 입력합니다.")

        time.sleep(2)

        browser.find_element_by_xpath('/html/body/app/layout/div[3]/div[2]/shrinkable-layout/div/app-base/search-input-box/div/div[1]/div/input').send_keys(Keys.ENTER) # 검색창에 입력한 식당 검색
        self.state.emit(self.query+"을(를) 검색합니다.")

        time.sleep(5)

        try:
            self.state.emit("크롤링 중...")

            browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[1]/div[3]/div[2]/div/salt-marker/div/button').click() # 지도에 뜬 식당 버튼 클릭

            time.sleep(5)

            place_num = browser.current_url.split('/place/')[1].split('?')[0] # 식당 고유 번호 받아오기

            url = "https://pcmap.place.naver.com/restaurant/"+str(place_num)+"/review/visitor" # 식당 고유 번호에 해당하는 네이버 플레이스 리뷰 목록 창 url 

            browser.get(url) # 네이버 플레이스 리뷰 목록 url 창으로 들어가기

            time.sleep(5)

            last_height = browser.execute_script("return document.body.scrollHeight") # 최근 script 높이 

            print(last_height)

            time.sleep(5)

            while True:
                
                browser.execute_script("window.scrollTo(0, document.body.scrollHeight);") # 스크롤 끝까지 내리기

                time.sleep(2)

                try:
                    browser.find_element_by_xpath('/html/body/div[3]/div/div/div[2]/div[5]/div[4]/div[3]/div[2]/a').click() # 더보기 버튼 클릭
                except:
                    break

                time.sleep(2)

                new_height = browser.execute_script("return document.body.scrollHeight") # 스크롤 내린 후 스크롤 높이 다시 가져옴

                print(new_height)

                if new_height == last_height:
                    break
                last_height = new_height

            wb = Workbook()

            sheet = wb.active

            sheet.append([self.query]) # sheet에 식당명 입력
            sheet.append(["리뷰"]) # sheet에 리뷰 입력

            res = browser.page_source

            soup =BeautifulSoup(res,"lxml")

            review = soup.find_all(name='span',attrs={'class':'WoYOw'})
            self.state.emit("총 리뷰 수: "+str(len(review)))

            rm_review = 0
            filter_review = 0

            for i in range(0,len(review)):
                r = rmEmoji(review[i].text)
                if r=='':
                    rm_review+=1
                else:
                    filter_review+=1
                    sheet.append([r]) # 댓글 목록 모두 sheet에 입력

            self.state.emit("삭제된 리뷰 수: "+str(rm_review))

            time.sleep(2)

            self.state.emit("크롤링 된 리뷰 수: "+str(filter_review))

            file_name = str(self.query) + ".xlsx" # sheet명 선언

            wb.save(file_name) # sheet 저장
            self.state.emit("sheet 생성 완료!")
            self.state.emit("크롤링이 성공적으로 수행되었습니다.")
            
        except:
            self.state.emit("식당이 존재하지 않습니다.")


class MyApp(QWidget):

    def __init__(self):
        super().__init__()
        
        self.initUI()
        self.query=''
        

    def initUI(self):
        

        label1 = QLabel("리뷰를 검색 할 음식점명을 입력하세요.",self)
        label1.move(120,50)

        qle = QLineEdit(self)
        qle.resize(300,30)
        qle.move(100,100)
        qle.textChanged[str].connect(self.res_name)

        self.label2 = QLabel(self)
        self.label2.resize(300,30)
        self.label2.move(100,150)

        btn1 = QPushButton('리뷰 탐색 시작', self)
        btn1.move(190,200)
        btn1.toggle()
        btn1.clicked.connect(self.crawling)


        self.setWindowTitle('네이버 리뷰 크롤링')
        self.setWindowIcon(QIcon('naver.jfif'))
        self.resize(500, 300)
        self.center() 
        self.show()
    
    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def res_name(self, text):
        self.query=text
        global query
        query=self.query

    def crawling(self):
        self.naverCrawling = naverCrawling()
        self.naverCrawling.state.connect(self.state)
        self.naverCrawling.start()

    @pyqtSlot(str)
    def state(self, text):
        self.label2.setText(text)

if __name__ == '__main__':
   app = QApplication(sys.argv)
   ex = MyApp()
   sys.exit(app.exec_())