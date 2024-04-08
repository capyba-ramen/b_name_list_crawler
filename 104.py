import time
import random
import requests


def search_job(keyword, max_mun=10, filter_params=None, sort_type='符合度', is_sort_asc=False):
    """搜尋職缺"""
    jobs = []
    total_count = 0

    url = 'https://www.104.com.tw/jobs/search/list'
    query = f'ro=0&kwop=7&keyword={keyword}&expansionType=area,spec,com,job,wf,wktm&mode=s&jobsource=2018indexpoc'
    if filter_params:
        # 加上篩選參數，要先轉換為 URL 參數字串格式
        query += ''.join([f'&{key}={value}' for key, value, in filter_params.items()])

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.92 Safari/537.36',
        'Referer': 'https://www.104.com.tw/jobs/search/',
    }

    # 加上排序條件
    sort_dict = {
        '符合度': '1',
        '日期': '2',
        '經歷': '3',
        '學歷': '4',
        '應徵人數': '7',
        '待遇': '13',
    }
    sort_params = f"&order={sort_dict.get(sort_type, '1')}"
    sort_params += '&asc=1' if is_sort_asc else '&asc=0'
    query += sort_params

    page = 1
    while len(jobs) < max_mun:
        params = f'{query}&page={page}'
        r = requests.get(url, params=params, headers=headers)
        if r.status_code != requests.codes.ok:
            print('請求失敗', r.status_code)
            data = r.json()
            print(data['status'], data['statusMsg'], data['errorMsg'])
            break

        data = r.json()
        total_count = data['data']['totalCount']
        jobs.extend(data['data']['list'])

        if (page == data['data']['totalPage']) or (data['data']['totalPage'] == 0):
            break
        page += 1
        time.sleep(random.uniform(3, 5))

    return total_count, jobs[:max_mun]


filter_params = {
    'area': '6001001000',  # (地區) 台北市
    's9': '1',  # (上班時段) 日班
    'isnew': '0',  # (更新日期) 本日最新
    'jobexp': '1,3',  # (經歷要求) 1年以下,1-3年
    'zone': '16',  # (公司類型) 上市上櫃
}
total_count, jobs = search_job('python', max_mun=10, filter_params=filter_params)
print('搜尋結果職缺總數：', total_count)
