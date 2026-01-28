# scripts/update_data.py
import json
import time
import pandas as pd
from datetime import datetime
import requests

def get_robust_session():
    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0"
    })
    return s

def main():
    url = "https://www.soilmove.tw/soilmove/dumpsiteGisQueryList"
    base_url = "https://www.soilmove.tw/soilmove/dumpsiteGisQuery"

    session = get_robust_session()

    # 先拿 session cookie
    session.get(base_url, timeout=20)
    time.sleep(1)

    r = session.post(
        url,
        data={"city": ""},
        headers={
            "Referer": base_url,
            "X-Requested-With": "XMLHttpRequest"
        },
        timeout=30
    )
    r.raise_for_status()

    text = r.text.strip().lstrip("\ufeff")
    payload = json.loads(text)

    df = pd.DataFrame(payload)

    def fix_coords(row):
        try:
            x = float(row.get("x", 0) or 0)
            y = float(row.get("y", 0) or 0)
            if 118 < x < 125: return [x, y]
            if 118 < y < 125: return [y, x]
        except:
            pass
        return [0, 0]

    df[["lng", "lat"]] = df.apply(lambda r: pd.Series(fix_coords(r)), axis=1)
    df = df[(df["lng"] > 0) & (df["lat"] > 0)]

    result = {
        "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "data": df.fillna("").to_dict(orient="records")
    }

    with open("data.json", "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"✅ data.json updated, records={len(result['data'])}")

if __name__ == "__main__":
    main()