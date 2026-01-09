import os
import io
import time
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
        return jsonify({"updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "data": df.fillna("").to_dict(orient="records")})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/export/excel', methods=['POST'])
def export_excel():
    try:
        data = request.json.get('data', [])
        df = pd.DataFrame(data)
        mapping = {"dumpname": "場名", "city": "縣市", "typename": "類型", "controlId": "流向編號", "applydate": "申報日期", "remain": "B1~B7 剩餘量", "maxbury": "B1~B7 核准量", "area": "面積", "lng": "經度", "lat": "緯度"}
        df_export = df[mapping.keys()].rename(columns=mapping)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False)
        output.seek(0)
        return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/export/kml', methods=['POST'])
def export_kml():
    try:
        data = request.json.get('data', [])
        kml = simplekml.Kml()
        for d in data:
            pnt = kml.newpoint(name=d.get('dumpname'), coords=[(d.get('lng'), d.get('lat'))])
            pnt.description = f"縣市：{d.get('city')}\n剩餘量：{d.get('remain')} ㎥"
        output = io.BytesIO()
        output.write(kml.kml().encode('utf-8'))
        output.seek(0)
        return send_file(output, mimetype='application/vnd.google-earth.kml+xml', as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    # 使用環境變數中的 PORT，若無則預設 5000
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)