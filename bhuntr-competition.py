import os
import openpyxl
import requests
from datetime import datetime
from operate_excel import XlsxExcel

def make_url(target: str, page: int, size: int):
    return f"https://api.bhuntr.com/tw/cms/bhuntr/contest?language=tw&target={target}&limit={size}&page={page}&sort=mixed&timeline=notEnded"


def crawl_webpage(url):
    try:
        response = requests.get(url)
        
        if response.status_code == 200:
            return response.json()["payload"]
        else:
            print(f"Failed to fetch webpage. Status code: {response.status_code}")
            return None
    except requests.exceptions.RequestException as e:
        print(f"Error occurred: {e}")
        return None


# url = "https://api.bhuntr.com/tw/cms/bhuntr/contest?language=tw&target=event&limit=2&page=1&sort=mixed&timeline=notEnded"

def craw_all_and_prepare_data():
    count = 0
    start_from = 1
    size = 300
    url = make_url("competition", start_from, size)
    data = [] # ['單位名稱', 'Email', '電話', '活動名稱', '活動日期', '網站']
    while True:
        print(f"Fetching data from page {start_from}...")
        result = crawl_webpage(url)

        if not result:
            break

        for item in result["list"]:
            count += 1
            data.append([
                item["organizerTitle"],
                item["contactEmail"],
                item["contactPhone"],
                item["title"],
                f"{datetime.fromtimestamp(item['startTime'])} 至 {datetime.fromtimestamp(item['endTime'])}",
                f"https://bhuntr.com/tw/competitions/{item['alias']}"
            ])

        if result["page"]["next"] == 1:
            break

        start_from += 1
        url = make_url("competitions", start_from, size)
    
    print(f"Total records: {count}")


    # Save data to Excel
    project_dir = os.path.dirname(os.path.abspath(__file__))
    excel_folder_name = 'doc'
    doc_dir = os.path.join(project_dir, excel_folder_name)

    if not os.path.exists(doc_dir):
     os.makedirs(doc_dir)
    now = datetime.now()
    file_created_time = now.strftime('%Y-%m-%d_%H%M')
    file_name = f"獎金獵人_competition_{file_created_time}.xlsx"
    file_path = os.path.join(doc_dir, file_name)

    # create a new Excel file
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = '獎金獵人'
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
    craw_all_and_prepare_data()