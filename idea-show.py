import requests
from bs4 import BeautifulSoup
import os
import openpyxl
from datetime import datetime
from operate_excel import XlsxExcel
from time import sleep

def scrape_page(url, data=[]):
    try:
        response = requests.get(url)
        
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, "html.parser")
            title = soup.find("h1", class_="post-title").text if soup.find("h1", class_="post-title") else None
            email = soup.find("span", class_="organizer-email").find("a").text if soup.find("span", class_="organizer-email") else None
            phone = soup.find("span", class_="organizer-tel").text if soup.find("span", class_="organizer-tel") else None
            orgName = soup.find("span", class_="organizer-name").find("a").text if soup.find("span", class_="organizer-name") else None
            dateRangeList = soup.find("p", class_="event-date-range").find_all("span")
            dateRange = f"{dateRangeList[0].text} ~ {dateRangeList[1].text}" if dateRangeList else None
            data.append([title, email, phone, orgName, dateRange, url])

        else:
            print(f"Failed to fetch webpage. Status code: {response.status_code}")


    
    except requests.exceptions.RequestException as e:
        print(f"Error occurred: {e}")

    return data

def crawl_webpage(url):
    try:
        page_number = 1
        data = []

        while True:
            page_url = f"{url}?pid={page_number}"
            print(f"Fetching data from page {page_number}...")
            response = requests.get(page_url)

            if response.status_code == 200:
                soup = BeautifulSoup(response.text, "html.parser")
                
                events = soup.find_all("article", class_="grid-item")
                
                if len(events) == 0:
                    print("No more events found. Stopping crawling.")
                    break

                # Iterate over each event and find the links within their post-thumbnails
                for event in events:
                    href = event.find("div", class_="post-thumbnail").find('a').get('href')
                    if href and href.startswith("http"):
                        scrape_page(href, data)

                page_number += 1
                sleep(1)
                

            else:
                print(f"Failed to fetch webpage. Status code: {response.status_code}")
                break

    except requests.exceptions.RequestException as e:
        print(f"Error occurred: {e}")

    return data

url = "https://news.idea-show.com/active"


def save_to_excel():
    data = crawl_webpage(url)

    project_dir = os.path.dirname(os.path.abspath(__file__))
    excel_folder_name = 'doc'
    doc_dir = os.path.join(project_dir, excel_folder_name)

    if not os.path.exists(doc_dir):
      os.makedirs(doc_dir)
    now = datetime.now()
    file_created_time = now.strftime('%Y-%m-%d_%H%M')
    file_name = f"點子秀_{file_created_time}.xlsx"
    file_path = os.path.join(doc_dir, file_name)

    # create a new Excel file
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = '點子秀'
    workbook.save(file_path)

    # add data to worksheet
    import_result_file = XlsxExcel(file_path, 0)

    column_titles = ['單位名稱', 'Email', '電話', '活動名稱', '活動日期', '網站']

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