from openpyxl import load_workbook
from selenium import webdriver
import timeit
import datetime

print('''코드 제작 : @ccamang / 무단 수정, 배포 금지 / 문의 : curomame@naver.com
ver.21.01.28
--------------------------------------------------------------------------------

               * 실행 전 반드시 아래 <주의 사항>을 확인 후 진행해주세요 *

                                             <주의사항>
                                                
      1. 실행하려는 xlsx파일 내 A행에 키워드, B행에 유효 url이 입력 되어 있는지 확인해주세요.

      2. 실행하려는 폴더 내에 입력한 xlsx파일이 있는지 확인해주세요.

      3. xlsx파일 내 Sheet를 다음과 같이 배치해주세요.
          Sheet 1 : NAVER-pc // Sheet 2 : NAVER-mo
          Sheet 3 : DAUM-pc // Sheet 4 : DAUM-mo

      4. 실행 도중 프로그램이 종료 될 시, 구해진 데이터는 저장되지 않습니다.
          (가능한 종료시까지 다른 작업을 하지 않는 것을 추천합니다.)

      5. 파일명은 검색완료.[today,time].xlsx로 현재 폴더에 저장됩니다.

      + 검색 엔진의 HTML값이 상시 변경되어 파일이 제대로 작동하지 않을 수 있습니다. +

      주의 사항을 다 확인하셨다면 파일명을 입력 후 계속 진행해주세요. ''')

wb_name = input('파일명을 입력하세요.(확장자 포함) : ')
wb_name = wb_name.strip()

start_num = input('검색을 시작할 Sheet의 번호를 입력하세요 : ')
end_num = input('검색을 끝낼 Sheet의 번호를 입력하세요.(이 값은 시작 번호보다 같거나 커야합니다.) : ')

s = int(start_num) - 1

wb = load_workbook(wb_name)
start_time = timeit.default_timer()

num_ad = input('''
주의사항을 전부 확인하셨다면 Enter를 입력하여 프로그램을 실행합니다.

''')

while True :
   s += 1

   if s == int(end_num) + 1 : break
   
###### 1번 Naver PC 실행
   elif s == 1 :
      k = 1
      data = wb['Sheet'+str(s)]
      url_ad = str('B'+str(k))
      
      while True :
         wb_num = str('A'+str(k))
         if data[wb_num].value ==  None : break

#html
         driver = webdriver.Chrome('./chromedriver')
         url = 'https://ad.search.naver.com/search.naver?where=ad&query='+data[wb_num].value
         driver.get(url)

         from bs4 import BeautifulSoup
         html = driver.page_source

         soup = BeautifulSoup(html,'html.parser')

         driver.close()

# rank 출력 완료

         global i
         i = 0
   
         while True :
            ad_ranks = soup.select('div.url_area > a')[i].text
            ad_ranks = ad_ranks.strip()

            if ad_ranks == data[url_ad].value :
               break

            else :
     
               i += 1
         k += 1

#xlsx
   
         data = wb['Sheet1']  
         data['C'+str(k-1)] = int(i+1)

###### 2번 Naver Mo 실행
   elif s == 2 :
      k = 1
      data = wb['Sheet'+str(s)]
      url_ad = str('B'+str(k))
      
      while True :
         wb_num = str('A'+str(k))
         if data[wb_num].value ==  None : break

#html
         driver = webdriver.Chrome('./chromedriver')
         url = 'https://m.ad.search.naver.com/search.naver?where=m_expd&query='+data[wb_num].value
         driver.get(url)

         from bs4 import BeautifulSoup
         html = driver.page_source

         soup = BeautifulSoup(html,'html.parser')

         driver.close()

# rank 출력 완료


         i = 0
   
         while True :
            ad_ranks = soup.select('cite.url > span.url_link')[i].text
            ad_ranks = ad_ranks.strip()

            if ad_ranks == data[url_ad].value :
               break

            else :
     
               i += 1
         k += 1

#xlsx
   
         data = wb['Sheet2']  
         data['C'+str(k-1)] = int(i+1)

###### 3번 Daum Pc 실행
   elif s == 3 :
      k = 1
      data = wb['Sheet'+str(s)]
      url_ad = str('B'+str(k))
      
      while True :
         wb_num = str('A'+str(k))
         if data[wb_num].value ==  None : break

#html
         driver = webdriver.Chrome('./chromedriver')
         url = 'https://search.daum.net/search?w=ad&DA=YZR&q='+data[wb_num].value
         driver.get(url)

         from bs4 import BeautifulSoup
         html = driver.page_source

         soup = BeautifulSoup(html,'html.parser')

         driver.close()

# rank 출력 완료


         i = 0
   
         while True :
            ad_ranks = soup.select('div.info_main > a')[i].text
            ad_ranks = ad_ranks.strip()

            if ad_ranks == data[url_ad].value :
               break

            else :
     
               i += 1
         k += 1

#xlsx
   
         data = wb['Sheet3']  
         data['C'+str(k-1)] = int(i+1)

###### 4번 Daum Pc 실행
   elif s == 4 :
      k = 1
      data = wb['Sheet'+str(s)]
      url_ad = str('B'+str(k))
      
      while True :
         wb_num = str('A'+str(k))
         if data[wb_num].value ==  None : break

#html
         driver = webdriver.Chrome('./chromedriver')
         url = 'https://m.search.daum.net/search?w=ad&q='+data[wb_num].value
         driver.get(url)

         from bs4 import BeautifulSoup
         html = driver.page_source

         soup = BeautifulSoup(html,'html.parser')

         driver.close()

# rank 출력 완료
         i = 0
   
         while True :
            ad_ranks = soup.select('div.wrap_ad > a')[i].text
            ad_ranks = ad_ranks.replace(" 광고 ","")
            ad_ranks = ad_ranks.strip()
            
            if ad_ranks == data[url_ad].value :
               break

            else :
     
               i += 1
         k += 1

#xlsx
   
         data = wb['Sheet4']  
         data['C'+str(k-1)] = int(i+1)


today = datetime.datetime.today()
today = str(today)
today = today[:19]
today = today.replace(":","_")

wb_name = str(wb_name)
wb_name = wb_name.replace(".xlsx","")

today = wb_name + '검색완료 '+ today + '.xlsx'

wb.save(today)

terminate_time = timeit.default_timer()
print('파워랭킹 순위 작성이 완료되었습니다.')
print("순위 검색에 총%f초가 걸렸습니다." % (terminate_time - start_time))
print('''--------------------------------------------------------------------------------
오늘하루도 파이팅입니닷! :>

코드 제작 : @ccamang / 무단 수정, 배포 금지 / 문의 : curomame@naver.com
ver.21.01.28''')

         
