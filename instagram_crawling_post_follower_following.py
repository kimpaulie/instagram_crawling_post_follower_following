import requests
import pandas as pd

# 엑셀 불러오기
where_file=('list1.xlsx')
df=pd.read_excel(where_file, 'Sheet1', index_col=None, na_values=['NA'])

# 엑셀의 2번째 줄부터(1번째 줄은 col로 설정) 불러와서 정보 parsing
for i in range(len(df)):
    url=df.iloc[i, 0         # 첫 열의 첫 줄부터 순서대로 처리
    instaid=url.split('/')[3]         # 인스타 아이디 분리
    text = requests.get(url).text        # html 코드 받기

    # 원하는 값의 앞 뒤를 제거하여 값 parsing
    start = '"edge_owner_to_timeline_media":{"count":'
    end = ',"page_info":{"has_next_page"'
    post_num = text[text.find(start) + len(start):text.rfind(end,text.find(start),text.find(start)+100)]

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


print('\n총 %d명 계정의 정보를 가져왔습니다.' % (i+1))

# 엑셀에 저장. 파일 열려있으면 권한 에러 발생
df.to_excel(where_file,'Sheet1', index=False, encoding='utf-8')

print('\n프로그램이 종료됩니다.')


