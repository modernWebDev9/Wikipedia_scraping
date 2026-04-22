import requests
API = "https://en.wikipedia.org/w/api.php"
H = {"User-Agent": "MedalCatalogBot/1.0"}

cats = [
    "Category:New Zealand campaign medals",
    "Category:New Zealand Royal Honours System",
    "Category:Recipients of orders, decorations, and medals of New Zealand",
    "Category:Orders, decorations, and medals of New Zealand",
    "Category:Australian campaign medals",
    "Category:Orders, decorations, and medals of Australia",
    "Category:Military awards and decorations of Australia",
    "Category:Military awards and decorations of New Zealand",
]
for cat in cats:
    r = requests.get(API, params={
        "action": "query", "list": "categorymembers",
        "cmtitle": cat, "cmlimit": "5", "cmtype": "page",
        "cmnamespace": "0", "format": "json"
    }, headers=H, timeout=10).json()
    members = r.get("query", {}).get("categorymembers", [])
    print(f"{len(members):>2}  {cat}")
    for m in members[:3]:
        print(f"      → {m['title']}")
