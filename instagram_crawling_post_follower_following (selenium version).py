import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--incognito")
chrome_options.add_argument("--headless")
driver = webdriver.Chrome('C:/Users/Paul/Downloads/chromedriver/chromedriver.exe',
                          options=chrome_options)
## url에 접근
driver.get('https://www.instagram.com/')
print("○ 로그인 페이지 접속")

username = 'aaaa'
password = 'aaaa'
driver.implicitly_wait(2)
print("○ 로그인 시도 중")

# id, pw 입력할 곳을 찾습니다.
driver.find_element_by_name("username").send_keys(username)
driver.find_element_by_name("password").send_keys(password + Keys.ENTER)
time.sleep(2)
print("○ 로그인 완료")
time.sleep(2)
print("○ 작업 시작")
# 엑셀 불러오기
print("○ 엑셀 파일 로딩")
where_file=('list.xlsx')
df=pd.read_excel(where_file, 'Sheet1', index_col=None, na_values=['NA'])

save_count = 1 # 100명 마다 한 번씩 저장

#만약 중간에 끊어졌다면, 끝어진 타이밍의 숫자를 입력하여 진행 (5라고 치면 6부터 시작)
previous_num = 129

# 엑셀의 2번째 줄부터(1번째 줄은 col로 설정) 불러와서 정보 parsing
for i in range(previous_num,len(df)):
    url=df.iloc[i,0]         # 첫 열의 첫 줄부터 순서대로 처리
    instaid=url.split('/')[3]         # 인스타 아이디 분리
    # text = requests.get(url).text        # html 코드 받기
    driver.get(url)
    text = driver.page_source

    # 원하는 값의 앞 뒤를 제거하여 값 parsing
    start = '"edge_owner_to_timeline_media":{"count":'
    end = ',"page_info":{"has_next_page"'
    post_num = text[text.find(start) + len(start):text.rfind(end,text.find(start),text.find(start)+80)]

    start = '"edge_followed_by":{"count":'
    end = '},"followed_by_viewer"'
    followers = text[text.find(start) + len(start):text.rfind(end)]

    start = '"edge_follow":{"count":'
    end = '},"follows_viewer"'
    following = text[text.find(start) + len(start):text.rfind(end)]

    # 각 행에 저장
    df.loc[i,'게시물수'] = int(post_num)
    df.loc[i,'팔로워'] = int(followers)
    df.loc[i,'팔로잉'] = int(following)
    df.loc[i, '팔로워 대비 팔로잉'] = round(float(followers)/float(following),2)
    follperc=float(followers) / float(following)
    print("%d - 인스타 ID : %s, 포스팅 : %s개, 팔로워 : %s명, 팔로잉 : %s명, 팔로워 대비 팔로워 : %.2f배" % (i+1,instaid,post_num,followers,following,follperc))

    if i == ((save_count * 100) - 1):
        df.to_excel(where_file, 'Sheet1', index=False, encoding='utf-8')
        print('\n중간 저장 - 총 %d개 계정의 정보를 저장했습니다.\n' % (i + 1))
        save_count = save_count+1


print('\n총 %d명 계정의 정보를 가져왔습니다.' % (i+1))

# 엑셀에 저장. 파일 열려있으면 권한 에러 발생
df.to_excel(where_file,'Sheet1', index=False, encoding='utf-8')

print('\n프로그램이 종료됩니다.')