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

urllib3.disable_warnings(InsecureRequestWarning)
app = Flask(__name__)

DATA_FILE = 'data.json'

def get_robust_session():
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
    return render_template('index.html')

@app.route('/api/local_data')
def get_local_data():
    """讀取預設的 data.json 檔案"""
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                content = json.load(f)
                return jsonify(content)
        except Exception as e:
            return jsonify({"error": str(e)}), 500
    return jsonify({"error": "No default data found"}), 404

@app.route('/api/update')
def api_update():
    try:
        url = "https://www.soilmove.tw/soilmove/dumpsiteGisQueryList"
        base_url = "https://www.soilmove.tw/soilmove/dumpsiteGisQuery"
        session = get_robust_session()
        session.get(base_url, timeout=15, verify=False)
        time.sleep(1)
        r = session.post(url, data={"city": ""}, headers={"Referer": base_url}, timeout=25, verify=False)
        r.raise_for_status()
        
        df = pd.DataFrame(r.json())
        def fix_coords(row):
            try:
                x, y = float(row.get("x", 0)), float(row.get("y", 0))
                if 118 < x < 125: return pd.Series([x, y])
                if 118 < y < 125: return pd.Series([y, x])
                return pd.Series([0, 0])
            except: return pd.Series([0, 0])
            
        df[["lng", "lat"]] = df.apply(fix_coords, axis=1)
        df = df[(df["lng"] > 0) & (df["lat"] > 0)].copy()
        
        result_content = {
            "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 
            "data": df.fillna("").to_dict(orient="records")
        }

        # 更新完畢後自動存入 data.json
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(result_content, f, ensure_ascii=False, indent=4)

        return jsonify(result_content)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/export/excel', methods=['POST'])
def export_excel():
    try:
        data = request.json.get('data', [])
        if not data: return jsonify({"error": "無資料"}), 400
        
        df = pd.DataFrame(data)
        
        # 這裡的 Key 必須跟 JS 裡 d.xxxx 的名稱一模一樣
        mapping = {
            "dumpname": "場名",
            "city": "縣市",
            "typename": "類型",
            "controlId": "流向編號",
            "applydate": "申報日期",
            "remain": "B1~B7 剩餘量(㎥)",
            "maxbury": "B1~B7 核准量(㎥)",
            "area": "面積(公頃)",
            "lng": "經度",
            "lat": "緯度"
        }
        
        # 篩選出資料中有的欄位
        existing_cols = [c for c in mapping.keys() if c in df.columns]
        df_export = df[existing_cols].rename(columns=mapping)

        output = io.BytesIO()
        # 強制使用 openpyxl 引擎
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False)
        
        output.seek(0)
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"土資場匯出_{datetime.now().strftime('%Y%m%d')}.xlsx"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/export/kml', methods=['POST'])
def export_kml():
    try:
        data = request.json.get('data', [])
        kml = simplekml.Kml()
        for d in data:
            try:
                # 座標必須轉換為 float，否則 Google Earth 無法讀取
                lng = float(d.get('lng', 0))
                lat = float(d.get('lat', 0))
                if lng == 0: continue
                kml.newpoint(name=str(d.get('dumpname')), coords=[(lng, lat)])
            except: continue
            
        output = io.BytesIO()
        output.write(kml.kml().encode('utf-8'))
        output.seek(0)
        return send_file(
            output,
            mimetype='application/vnd.google-earth.kml+xml',
            as_attachment=True,
            download_name='export.kml'
        )
    except Exception as e:
        return jsonify({"error": f"KML 匯出失敗: {str(e)}"}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)