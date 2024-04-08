import requests


def crawl_webpage(url):
    try:
        response = requests.get(url)
        
        if response.status_code == 200:
            print(response.json()["payload"]["list"][0])
        else:
            print(f"Failed to fetch webpage. Status code: {response.status_code}")
    except requests.exceptions.RequestException as e:
        print(f"Error occurred: {e}")

url = "https://api.bhuntr.com/tw/cms/bhuntr/contest?language=tw&target=competition&limit=24&page=1&sort=mixed&timeline=notEnded&identify=none"

crawl_webpage(url)