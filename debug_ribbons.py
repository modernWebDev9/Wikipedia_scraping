import requests, re
from bs4 import BeautifulSoup

API = "https://en.wikipedia.org/w/api.php"
H   = {"User-Agent": "MedalCatalogBot/1.0"}

def fetch_infobox_imgs(page):
    r = requests.get(API, params={
        "action": "parse", "page": page, "prop": "text",
        "disableeditsection": "true", "format": "json", "formatversion": "2"
    }, headers=H, timeout=20).json()
    html    = r["parse"]["text"]
    soup    = BeautifulSoup(html, "html.parser")
    infobox = soup.find("table", class_=re.compile(r"infobox", re.I))

    print(f"\n=== {page} ===")
    print("-- infobox-image td spans --")
    td = infobox.find("td", class_=re.compile(r"infobox-image", re.I)) if infobox else None
    if td:
        for span in td.find_all("span", attrs={"typeof": "mw:File"}):
            img = span.find("img")
            if img:
                src = img.get("src","")
                m   = re.search(r"/thumb/[0-9a-f]/[0-9a-f]{2}/([^/]+\.(svg|png|jpg|gif))/", src, re.I)
                fname = requests.utils.unquote(m.group(1)) if m else src.split("/")[-1]
                fw = img.get("data-file-width","?")
                fh = img.get("data-file-height","?")
                cap_div = td.find("div", class_=re.compile(r"infobox-caption", re.I))
                cap = cap_div.get_text(strip=True) if cap_div else ""
                print(f"  span: {fname}  [{fw}x{fh}]  caption='{cap}'")

    print("-- all infobox tr imgs --")
    if infobox:
        for tr in infobox.find_all("tr"):
            label = tr.find(class_=re.compile(r"infobox-label", re.I))
            ltext = label.get_text(strip=True) if label else ""
            for img in tr.find_all("img"):
                src = img.get("src","")
                m   = re.search(r"/thumb/[0-9a-f]/[0-9a-f]{2}/([^/]+\.(svg|png|jpg|gif))/", src, re.I)
                fname = requests.utils.unquote(m.group(1)) if m else src.split("/")[-1]
                fw = img.get("data-file-width","?")
                fh = img.get("data-file-height","?")
                print(f"  label='{ltext}'  {fname}  [{fw}x{fh}]")

fetch_infobox_imgs("Air Medal")
fetch_infobox_imgs("General Service Medal (Canada)")
