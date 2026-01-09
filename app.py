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
DATA_FILE = 'data.json'

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
def get_local_data():
    """讀取伺服器上預設的 data.json 檔案"""
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
    """從目標網站抓取最新土資場資料"""
    try:
        url = "https://www.soilmove.tw/soilmove/dumpsiteGisQueryList"
        base_url = "https://www.soilmove.tw/soilmove/dumpsiteGisQuery"
        session = get_robust_session()
        
        # 模擬瀏覽行為以取得 Session
        session.get(base_url, timeout=15, verify=False)
        time.sleep(1)
        
        # 發送請求獲取資料
        r = session.post(url, data={"city": ""}, headers={"Referer": base_url}, timeout=25, verify=False)
        r.raise_for_status()
        
        df = pd.DataFrame(r.json())
        
        # 座標校正邏輯：判斷經緯度是否反置
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

        # 更新完畢後自動覆蓋 data.json
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(result_content, f, ensure_ascii=False, indent=4)

        return jsonify(result_content)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

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