import json
import time
import pandas as pd
from datetime import datetime
import requests
from urllib3.util.retry import Retry
from requests.adapters import HTTPAdapter

DATA_FILE = "data.json"  # ✅ 寫回 repo，Actions 才能 commit

def get_robust_session():
    session = requests.Session()

    retries = Retry(
        total=5,
        connect=5,
        read=5,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=frozenset(["GET", "POST"]),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("https://", adapter)
    session.mount("http://", adapter)

    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "*/*",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": "https://www.soilmove.tw",
    })
    return session

def main():
    url = "https://www.soilmove.tw/soilmove/dumpsiteGisQueryList"
    base_url = "https://www.soilmove.tw/soilmove/dumpsiteGisQuery"

    session = get_robust_session()

    # 1) 先 GET 取得 session cookie
    g = session.get(base_url, timeout=20, allow_redirects=True)
    g.raise_for_status()
    time.sleep(0.5)

    # 2) POST 取資料（注意 headers）
    headers_post = {
        "Referer": base_url,
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    }
    r = session.post(
        url,
        data={"city": ""},
        headers=headers_post,
        timeout=30,
        allow_redirects=True
    )
    r.raise_for_status()

    text = (r.text or "").lstrip("\ufeff").strip()

    # ✅ 不看 Content-Type，直接 parse（因為它會亂標 text/html）
    if not (text.startswith("[") or text.startswith("{")):
        raise RuntimeError(f"Upstream not JSON. sample={text[:200]}")

    payload = json.loads(text)
    df = pd.DataFrame(payload)

    def fix_coords(row):
        try:
            x = float(row.get("x", 0) or 0)
            y = float(row.get("y", 0) or 0)
            if 118 < x < 125:
                return [x, y]
            if 118 < y < 125:
                return [y, x]
        except:
            pass
        return [0.0, 0.0]

    if "x" not in df.columns: df["x"] = 0
    if "y" not in df.columns: df["y"] = 0

    df[["lng", "lat"]] = df.apply(lambda rr: pd.Series(fix_coords(rr)), axis=1)
    df = df[(df["lng"] > 0) & (df["lat"] > 0)].copy()

    result = {
        "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "data": df.fillna("").to_dict(orient="records")
    }

    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"✅ Updated {DATA_FILE}, records={len(result['data'])}")

if __name__ == "__main__":
    main()