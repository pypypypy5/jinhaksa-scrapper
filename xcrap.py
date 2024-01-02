import webbrowser
import pyautogui
import time
import openpyxl
from PIL import Image
import pytesseract,re
import os,subprocess


webbrowser.open('https://hijinhak.jinhak.com/SAT/J1Apply/J1MyApplyList.aspx?LeftTab=1') #진학사 들어가고 로딩시간 10초 대기
time.sleep(5)

gottoget = [['-','-','-','-','-','-','-','-','-','-','-','-','-','-'],['-','-','-','-','-','-','-','-','-','-','-','-','-','-'],['-','-','-','-','-','-','-','-','-','-','-','-','-','-'],['-','-','-','-','-','-','-','-','-','-','-','-','-','-'],['-','-','-','-','-','-','-','-','-','-','-','-','-','-'],['-','-','-','-','-','-','-','-','-','-','-','-','-','-'],['-','-','-','-','-','-','-','-','-','-','-','-','-','-'],['-','-','-','-','-','-','-','-','-','-','-','-','-','-'],['-','-','-','-','-','-','-','-','-','-','-','-','-','-'],['-','-','-','-','-','-','-','-','-','-','-','-','-','-'],['-','-','-','-','-','-','-','-','-','-','-','-','-','-'],['-','-','-','-','-','-','-','-','-','-','-','-','-','-'],['-','-','-','-','-','-','-','-','-','-','-','-','-','-'],['-','-','-','-','-','-','-','-','-','-','-','-','-','-']] #[['설산공'],['설기계'],['설공지균'],['설화공'],['설재료'],['설항공'],['설컴'],['설전정'],['설첨융'],['고기계'],['고컴교과'],['고데이터'],['연컴'],['연인공지능'],['연전전'],['성반도']]
#리스트내용 순서 - '모의지원자수','모의지원등수','분석대상자수','합격예측등수','모집인원','모의지원합격자수','예상추합률'/'예상경쟁률','한계경쟁률','칸수보정','예상합격률'/'진학사 칸수','칸수내 인원','칸수내 등수'



#모지 페이지 들어가기
def intopage(scroll_x,scroll_y,into_x,into_y):
    if scroll_x == 0 and scroll_y == 0:
        pass
    else:
        pyautogui.moveTo(scroll_x,scroll_y)
        pyautogui.mouseDown()
        time.sleep(3)
        pyautogui.mouseUp()
    pyautogui.click(into_x,into_y)




#스크롤 내리기, 스크린샷(할 것들 저장)
def scrap():
    pyautogui.screenshot()        #맨 위칸

    pyautogui.click(599,388)      #왼쪽 위칸
    time.sleep(3)
    pyautogui.screenshot()
    pyautogui.click(888,387)      
    time.sleep(3)

    pyautogui.mouseDown(1800,378)  #칸수 있는 쪽
    time.sleep(3)
    pyautogui.mouseUp()
    pyautogui.screenshot()

    pyautogui.mouseDown(1803,729)  #3번
    time.sleep(3)
    pyautogui.mouseUp()
    pyautogui.click(1003,840)
    pyautogui.screenshot()

    pyautogui.mouseDown(1800,850)  #4번
    time.sleep(3)
    pyautogui.mouseUp()
    pyautogui.screenshot()

    pyautogui.click(1779,16)      #페이지 나가기





#스크린샷 값만 크게 보이도록 오리기
#테서렉트로 읽고 
#gottoget(함수가 받은 값을 인덱스로)에 저장
#스크린샷 지우기
def ocr(idx):

    nowimg = Image.open(r"C:\Users\ShinHyoJae\OneDrive - 연세대학교 (Yonsei University)\그림\스크린샷\스크린샷(232).png") #모의지원자수, 모의지원등수
    sml = nowimg.crop((1159,918,1243,952))
    text = pytesseract.image_to_string(sml)
    t = re.compile(r'([0-9]+)등/([0-9]+)')
    f = t.search(text)
    gottoget[idx][0] = f.group(2)
    gottoget[idx][1] = f.group(1)



    nowimg = Image.open(r"C:\Users\ShinHyoJae\OneDrive - 연세대학교 (Yonsei University)\그림\스크린샷\스크린샷(233).png") #분석대상자수, 합격예측등수
    sml = nowimg.crop((1363,247,1487,284))
    text = pytesseract.image_to_string(sml)
    t = re.compile(r'([0-9]+)등.*/.*([0-9]+)')
    f = t.search(text)
    gottoget[idx][2] = f.group(2)
    gottoget[idx][3] = f.group(1)



    nowimg = Image.open(r"C:\Users\ShinHyoJae\OneDrive - 연세대학교 (Yonsei University)\그림\스크린샷\스크린샷(231).png") #모집인원
    sml = nowimg.crop((1363,247,1487,284))
    text = pytesseract.image_to_string(sml)
    t = re.compile(r'([0-9]+)')
    f = t.search(text)
    gottoget[idx][4] = f.group(1)




    nowimg = Image.open(r"C:\Users\ShinHyoJae\OneDrive - 연세대학교 (Yonsei University)\그림\스크린샷\스크린샷(231).png") #모의지원합격자수
    sml = nowimg.crop((1223,976,1295,1012))
    text = pytesseract.image_to_string(sml)
    t = re.compile(r'([0-9]+)')
    f = t.search(text)
    gottoget[idx][5] = f.group(1)



    nowimg = Image.open(r"C:\Users\ShinHyoJae\OneDrive - 연세대학교 (Yonsei University)\그림\스크린샷\스크린샷(235).png") #예상추합률
    sml = nowimg.crop((1428,852,1530,897))
    text = pytesseract.image_to_string(sml)
    t = re.compile(r'([0-9]+)')
    f = t.search(text)
    gottoget[idx][6] = f.group(1)



    nowimg = Image.open(r"C:\Users\ShinHyoJae\OneDrive - 연세대학교 (Yonsei University)\그림\스크린샷\스크린샷(233).png") #진학사 칸수
    sml = nowimg.crop((760,124,793,165))
    text = pytesseract.image_to_string(sml)
    t = re.compile(r'([0-9])')
    f = t.search(text)
    gottoget[idx][11] = f.group(1)
    
    
    #TODO 칸수내 인원#칸수내 등수


    os.remove(r"C:\Users\ShinHyoJae\OneDrive - 연세대학교 (Yonsei University)\그림\스크린샷\스크린샷(231).png")
    os.remove(r"C:\Users\ShinHyoJae\OneDrive - 연세대학교 (Yonsei University)\그림\스크린샷\스크린샷(232).png")
    os.remove(r"C:\Users\ShinHyoJae\OneDrive - 연세대학교 (Yonsei University)\그림\스크린샷\스크린샷(233).png")
    os.remove(r"C:\Users\ShinHyoJae\OneDrive - 연세대학교 (Yonsei University)\그림\스크린샷\스크린샷(234).png")
    os.remove(r"C:\Users\ShinHyoJae\OneDrive - 연세대학교 (Yonsei University)\그림\스크린샷\스크린샷(235).png")


#진학사 메인 페이지에서 설산공 클릭
intopage(1907,673,1363,730) 
scrap()
ocr(0)

#진학사 메인 페이지에서 설기계 클릭
intopage(0,0,1362,673) 
scrap()
ocr(1)

#진학사 메인 페이지에서 설공지균 클릭
intopage(0,0,1364,331) 
scrap()
ocr(2)

#진학사 메인 페이지에서 설항공 클릭 
intopage(0,0,1365,846) 
scrap()
ocr(3)

#진학사 메인 페이지에서 설재료 클릭
intopage(0,0,1365,788)  
scrap()
ocr(4)

#진학사 메인 페이지에서 설첨융 클릭
intopage(0,0,1364,559)  
scrap()
ocr(7)

#진학사 메인 페이지에서 고컴교과 클릭
intopage(0,0,1385,1019)  
scrap()
ocr(8)

#진학사 메인 페이지에서 고데이터 클릭
intopage(0,0,1361,501)  
scrap()
ocr(9)

#진학사 메인 페이지에서 연전전 클릭
intopage(0,0,1364,613)  
scrap()
ocr(12)

#진학사 메인 페이지에서 성반도 클릭 
intopage(0,0,1364,385) 
scrap()
ocr(13)

#진학사 메인 페이지에서 설컴 클릭 
intopage(1908,842,1368,602) 
scrap()
ocr(5)

#진학사 메인 페이지에서 설전정 클릭
intopage(0,0,1365,426)  
scrap()
ocr(6)

#진학사 메인 페이지에서 연컴 클릭
intopage(0,0,1362,316)  
scrap()
ocr(10)

#진학사 메인 페이지에서 연인공지능 클릭
pyautogui.move(1908,174)
pyautogui.mouseDown()
time.sleep(5)
pyautogui.mouseUp()
intopage(1905,428,1366,779)  
scrap()
ocr(11)

pyautogui.click(1895,20)         #진학사 페이지 나가기



#for문 이용, gottoget을 Limit Acc Rate에 입력, 화면 캡처, 변환, gottoget에 입력

for num_count in range(14):
    rate = openpyxl.load_workbook('Limit_Acc_Rate.xlsx')
    calcul = rate['산공기계광역']
    
    for i in range(7):
        calcul[f'C{5+i}'] = gottoget[num_count][i]

    for i in range(3):           #진학사 칸수, 칸수내 인원, 칸수내 등수 입력
        calcul[f'C{17+i}'] = gottoget[num_count][12+i]
    
    calcul.save('Limit_Acc_Rate.xlsx')
    
    subprocess.Popen(['start', 'excel', r"C:\Users\ShinHyoJae\Desktop\itall\Limit_Acc_Rate.xlsx"], shell=True) #창 전체 크기, 크기 66%
    time.sleep(5)
    pyautogui.click(1658,86)
    time.sleep(1)
    pyautogui.click(1763,1005)
    time.sleep(1)
    pyautogui.screenshot()
    pyautogui.click(1399,29)

    nowimg = Image.open(r"C:\Users\ShinHyoJae\OneDrive - 연세대학교 (Yonsei University)\그림\스크린샷\스크린샷(231).png") #예상경쟁률
    sml = nowimg.crop((219,548,287,561))
    text = pytesseract.image_to_string(sml)
    t = re.compile('[0-9]+\.[0-9]+')
    f = t.search(text)
    gottoget[num_count][7] = str(f)
    
    nowimg = Image.open(r"C:\Users\ShinHyoJae\OneDrive - 연세대학교 (Yonsei University)\그림\스크린샷\스크린샷(231).png") #한계경쟁률
    sml = nowimg.crop((215,568,290,588))
    text = pytesseract.image_to_string(sml)
    t = re.compile('[0-9]+\.[0-9]+')
    f = t.search(text)
    gottoget[num_count][8] = str(f)

    nowimg = Image.open(r"C:\Users\ShinHyoJae\OneDrive - 연세대학교 (Yonsei University)\그림\스크린샷\스크린샷(231).png") #칸수보정
    sml = nowimg.crop((216,752,290,771))
    text = pytesseract.image_to_string(sml)
    t = re.compile('[0-9]+\.[0-9]+')
    f = t.search(text)
    gottoget[num_count][9] = str(f)

    nowimg = Image.open(r"C:\Users\ShinHyoJae\OneDrive - 연세대학교 (Yonsei University)\그림\스크린샷\스크린샷(231).png") #예상합격률
    sml = nowimg.crop((210,878,307,906))
    text = pytesseract.image_to_string(sml)
    t = re.compile('[0-9]+\.[0-9]+')
    f = t.search(text)
    gottoget[num_count][10] = str(f)

    os.remove(r"C:\Users\ShinHyoJae\OneDrive - 연세대학교 (Yonsei University)\그림\스크린샷\스크린샷(231).png")




#TODO for문 이용, gottoget을 ILCL에 입력, 화면 캡처, 변환, gottoget에 입력


# for문 이용, gottoget을 기록지에 입력
wb = openpyxl.load_workbook('Record_sheet.xlsx')
sheet = wb['5칸 이상 학과(한계경쟁률, 칸수계산기)']

for i in range(14):
    for j in range(11):
        sheet[f'G{4+(14*i)+j}'] = gottoget[i][j]

wb.save('Record_sheet.xlsx')




###########################
#pyautogui.move()는 그 좌표로 마우스를 옮기는 것이 아니라 좌표만큼 옮기는 것.