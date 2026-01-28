import os
import io
import time
import json
import urllib3
import requests
import pandas as pd
import simplekml
from datetime import datetime
from flask import Flask, render_template, jsonify, request, send_file
from urllib3.exceptions import InsecureRequestWarning
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# 忽略 SSL 安全警告
urllib3.disable_warnings(InsecureRequestWarning)

app = Flask(__name__)

# 預設存檔路徑
DATA_FILE = os.path.join("/tmp", "data.json")

def get_robust_session():
    """設定具備重試機制的連線階段"""
    session = requests.Session()
    retries = Retry(total=3, backoff_factor=1, status_forcelist=[500, 502, 503, 504])
    session.mount('https://', HTTPAdapter(max_retries=retries))
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "X-Requested-With": "XMLHttpRequest"
    })
    return session

@app.route('/')
def index():
    """主頁面渲染"""
    return render_template('index.html')

@app.route('/api/local_data')
def local_data():
    try:
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                return jsonify(json.load(f))
        else:
            return jsonify({
                "updated": "尚未更新",
                "data": []
            })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/update', methods=['GET'])
def api_update():
    """從目標網站抓取最新土資場資料，並寫入 /tmp/data.json（serverless cache）"""
    try:
        url = "https://www.soilmove.tw/soilmove/dumpsiteGisQueryList"
        base_url = "https://www.soilmove.tw/soilmove/dumpsiteGisQuery"

        session = get_robust_session()

        # 如果你要保護這支 API，解除註解（並在 Vercel 設 UPDATE_KEY 環境變數）
        # if request.args.get("key") != os.environ.get("UPDATE_KEY"):
        #     return jsonify({"error": "unauthorized"}), 403

        # ---- 1) 模擬瀏覽行為以取得 Session（拿 cookie）----
        headers_get = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        }
        g = session.get(base_url, headers=headers_get, timeout=15, allow_redirects=True)

        # 若被導到奇怪頁面，先把前 200 字回傳（超關鍵）
        if g.status_code >= 400:
            return jsonify({
                "error": "BasePageFailed",
                "status_code": g.status_code,
                "sample": (g.text or "")[:200]
            }), 502

        time.sleep(0.5)

        # ---- 2) 用更像 XHR 的 headers 去 POST ----
        headers_post = {
            "User-Agent": headers_get["User-Agent"],
            "Accept": "application/json, text/plain, */*",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "X-Requested-With": "XMLHttpRequest",
            "Origin": "https://www.soilmove.tw",
            "Referer": base_url,
        }

        r = session.post(
            url,
            data={"city": ""},
            headers=headers_post,
            timeout=25,
            allow_redirects=True
        )

        text = (r.text or "").strip()

        # 狀態碼不對才擋
        if r.status_code != 200:
            return jsonify({
                "error": "UpstreamBadStatus",
                "status_code": r.status_code,
                "sample": text[:200]
            }), 502

        # ✅ Content-Type 不管它，改看內容是不是 JSON 開頭
        if not (text.startswith("{") or text.startswith("[")):
            return jsonify({
                "error": "UpstreamNotJSON",
                "status_code": r.status_code,
                "content_type": r.headers.get("Content-Type", ""),
                "sample": text[:200]
            }), 502

        # ✅ 用 json.loads 解析（比 r.json() 更可控）
        payload = json.loads(text)

        # ---- 4) 轉 DataFrame ----
        df = pd.DataFrame(payload)
        if df.empty:
            result_content = {
                "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "data": []
            }
            # 寫入 /tmp（可選）
            os.makedirs(os.path.dirname(DATA_FILE), exist_ok=True)
            with open(DATA_FILE, 'w', encoding='utf-8') as f:
                json.dump(result_content, f, ensure_ascii=False, indent=4)
            return jsonify(result_content)

        # ---- 5) 座標校正：判斷經緯度是否反置 ----
        def fix_coords(row):
            try:
                x = float(row.get("x", 0) or 0)
                y = float(row.get("y", 0) or 0)

                # 台灣經度約 118~125
                if 118 < x < 125:
                    return pd.Series([x, y])
                if 118 < y < 125:
                    return pd.Series([y, x])
                return pd.Series([0.0, 0.0])
            except Exception:
                return pd.Series([0.0, 0.0])

        # 若 x/y 欄位不存在也不會炸
        if "x" not in df.columns:
            df["x"] = 0
        if "y" not in df.columns:
            df["y"] = 0

        df[["lng", "lat"]] = df.apply(fix_coords, axis=1)
        df = df[(df["lng"] > 0) & (df["lat"] > 0)].copy()

        # ---- 6) 組回傳內容 ----
        result_content = {
            "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "data": df.fillna("").to_dict(orient="records")
        }

        # ---- 7) 寫入 /tmp/data.json（serverless cache）----
        os.makedirs(os.path.dirname(DATA_FILE), exist_ok=True)
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(result_content, f, ensure_ascii=False, indent=4)

        return jsonify(result_content)

    except Exception as e:
        # 回傳更可診斷的錯誤（不要太長）
        return jsonify({
            "error": str(e),
            "hint": "常見原因：上游回應非 JSON、timeout、或 session 抓不到資料"
        }), 500
    
@app.route('/api/upload_default_json', methods=['POST'])
def upload_default_json():
    """接收前端傳來的 JSON 檔案並覆寫伺服器上的預設存檔"""
    try:
        if 'file' not in request.files:
            return jsonify({"error": "未選取檔案"}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "檔案名稱不可為空"}), 400

        # 直接覆蓋儲存 data.json
        file.save(DATA_FILE)

        # 重新讀取新存好的資料內容回傳前端
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            new_data = json.load(f)

        return jsonify({
            "updated": new_data.get("updated", "手動更新"),
            "data": new_data.get("data", [])
        })
    except Exception as e:
        return jsonify({"error": f"伺服器儲存失敗: {str(e)}"}), 500
    
@app.route('/api/export/excel', methods=['POST'])
def export_excel():
    """匯出篩選後的資料為 Excel (包含 10 個完整欄位)"""
    try:
        data = request.json.get('data', [])
        if not data: return jsonify({"error": "無資料可匯出"}), 400
        
        df = pd.DataFrame(data)
        
        # 欄位對應表
        mapping = {
            "dumpname": "場名",
            "city": "縣市",
            "typename": "類型",
            "controlId": "流向編號",
            "applydate": "申報日期",
            "remain": "剩餘量(㎥)",
            "maxbury": "核准量(㎥)",
            "area": "面積(ha)",
            "lng": "經度",
            "lat": "緯度"
        }
        
        existing_cols = [c for c in mapping.keys() if c in df.columns]
        df_export = df[existing_cols].rename(columns=mapping)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False)
        
        output.seek(0)
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"土資場資料_{datetime.now().strftime('%Y%m%d')}.xlsx"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/export/kml', methods=['POST'])
def export_kml():
    """匯出 KML 檔案，並在 description 中加入詳細的 Tip 資訊"""
    try:
        data = request.json.get('data', [])
        if not data: return jsonify({"error": "沒有資料"}), 400
            
        kml = simplekml.Kml()
        for d in data:
            try:
                lng = float(d.get('lng', 0))
                lat = float(d.get('lat', 0))
                if lng == 0 or lat == 0: continue
                
                pnt = kml.newpoint(
                    name=str(d.get('dumpname', '未命名')), 
                    coords=[(lng, lat)]
                )
                
                # 設定 Google Earth 彈窗詳細描述 (Tip)
                pnt.description = (
                    f"場名：{d.get('dumpname', '－')}\n"
                    f"縣市：{d.get('city', '－')}\n"
                    f"類型：{d.get('typename', '－')}\n"
                    f"流向編號：{d.get('controlId', '－')}\n"
                    f"申報日期：{d.get('applydate', '－')}\n"
                    f"--------------------------\n"
                    f"剩餘填埋量：{d.get('remain', '0')} ㎥\n"
                    f"核准填埋量：{d.get('maxbury', '0')} ㎥\n"
                    f"面積：{d.get('area', '0')} 公頃\n"
                    f"座標：{lat:.5f}, {lng:.5f}"
                )
            except: continue 
                
        output = io.BytesIO()
        output.write(kml.kml().encode('utf-8'))
        output.seek(0)
        
        return send_file(
            output, 
            mimetype='application/vnd.google-earth.kml+xml', 
            as_attachment=True,
            download_name=f"土資場匯出_{datetime.now().strftime('%Y%m%d')}.kml"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)