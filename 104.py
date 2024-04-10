import requests
from bs4 import BeautifulSoup
import os
import openpyxl
from datetime import datetime
from operate_excel import XlsxExcel
import requests
from urllib.parse import urlencode

def scrape_job_detail(apiUrl):
    result = None

    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.92 Safari/537.36",
            "Referer": "https://www.104.com.tw/jobs/search/",
        }

        response = requests.get(apiUrl, headers=headers)

        if response.status_code == 200:
            data = response.json()
            email = data['data']['contact']['email']
            hrName = data['data']['contact']['hrName']
            companyName = data['data']['header']['custName']
            companyUrl = data['data']['header']['custUrl']
            industry = data['data']['industry']
            custNo = data['data']['custNo']

            if email:
                result = { custNo: [companyName, industry, email, hrName, companyUrl] }

        else:
            print(f"Failed to fetch webpage. Status code: {response.status_code}")

    except KeyError:
        result = None

    except requests.exceptions.RequestException as e:
        print(f"Error occurred: {e}")

    return result


def scrape_company_joblist(companyJobListLink):
    try:
        response = requests.get(companyJobListLink)
        print(response.status_code)

        if response.status_code == 200:
            soup = BeautifulSoup(response.text, "html.parser")
            link = soup.find("a", class_="info-job__text").get('href') if soup.find("a", class_="info-job__text") else None
            if not link:
                return
            
            job_no_encoded = link.split('/')[-1]
            companyRowData = scrape_job_detail(f"https://www.104.com.tw/job/ajax/content/{job_no_encoded}")

        else:
            print(f"Failed to fetch webpage. Status code: {response.status_code}")
    
    except requests.exceptions.RequestException as e:
        print(f"Error occurred: {e}")

    return companyRowData
    


def search_companies(query):
    try:
        url = "https://www.104.com.tw/company/search"
        page_number = 1
        dict = {}

        while True:
            query['page'] = page_number
            qstr = urlencode(query)
            page_url = f"{url}?{qstr}"
            print(f"Fetching data from page {page_number}...")

            response = requests.get(page_url)

            if response.status_code == 200:
                soup = BeautifulSoup(response.text, "html.parser")
                joblistLinks = soup.find_all("div", class_="job-count-link")

                if not joblistLinks:
                    print("No more companies found. Stopping crawling.")
                    break

                for jobListLink in joblistLinks:
                    companyJobListLink = jobListLink.find('a').get('href')
                    companyRowData = None

                    if companyJobListLink:
                        companyRowData = scrape_company_joblist(companyJobListLink)

                        if companyRowData:
                            custNo = list(companyRowData.keys())[0] if companyRowData else None
                            if custNo:
                                if custNo not in dict:
                                    dict[custNo] = companyRowData[custNo]
                                    print(dict)

                if page_number >= 100:
                    print("104 only supports up to 100 pages. Stopping crawling.")
                    break

                page_number += 1
            
            else:
                if response.status_code == 404:
                    print("No more companies found. Stopping crawling.")
                    break

                else:
                    print(f"Failed to fetch webpage. Status code: {response.status_code}")
                    continue

    except requests.exceptions.RequestException as e:
        print(f"Error occurred: {e}")

    
    return dict.values()

    


def save_to_excel():
    # data = search_companies({'zone': 5 } ) # 外商公司
    data = search_companies({'zone': '4,5,16', 'order': 4 } ) # 外商上市上櫃 + 資本額高到低

    project_dir = os.path.dirname(os.path.abspath(__file__))
    excel_folder_name = 'doc'
    doc_dir = os.path.join(project_dir, excel_folder_name)

    if not os.path.exists(doc_dir):
      os.makedirs(doc_dir)
    now = datetime.now()
    file_created_time = now.strftime('%Y-%m-%d_%H%M')
    file_name = f"104_{file_created_time}.xlsx"
    file_path = os.path.join(doc_dir, file_name)

    # create a new Excel file
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = '104'
    workbook.save(file_path)

    # add data to worksheet
    import_result_file = XlsxExcel(file_path, 0)

    column_titles = ['單位名稱', '產業類別', 'Email', 'HR 名稱', '公司網站']

    # import_result_file.writeCols(0, 0, ["Fruits"] + test_data_fruits)
    import_result_file.writeRows(0, 0, column_titles)

    rows = 0
    for i, row in enumerate(data):
        import_result_file.writeRows(i + 1, 0, row)
        rows += 1

    print(f"Total rows: {rows}")

    import_result_file.save()



if __name__ == "__main__":
    save_to_excel()
