import requests
from bs4 import BeautifulSoup

url = "https://quotes.toscrape.com/"

try:
    response = requests.get(url, timeout=10)

    if response.status_code == 200:
        print("Website finish join. Data reading...")

        soup = BeautifulSoup(response.text, "html.parser")

        quotes = soup.find_all("div", class_="quote")

        print(f"\n seen writer: {len(quotes)} chu\n")
        print("_" * 50)

        for index, item in enumerate(quotes, 1):
            text = item.find("span", class_="text").text
            author = item.find("small", class_="author").text
            print(f"{index}. {text}")
            print(f"   - writer: {author}\n")

    else:
        print(f" no join website. Status Code: {response.status_code}")

except Exception as e:
    print(f" false: {e}")                