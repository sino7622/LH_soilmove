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


app = Flask(__name__)

# repo 內固定的 data.json（GitHub Actions 會更新它）
REPO_DATA_FILE = os.path.join(os.path.dirname(__file__), "data.json")

# Vercel serverless 可寫的暫存區（非持久化）
TMP_DATA_FILE = os.path.join("/tmp", "data.json")


def load_json_safely(path: str):
    """安全讀 JSON，不存在就回 None"""
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


@app.route("/", methods=["GET"])
def home():
    # 如果你原本有 templates/index.html，這裡就能正常 render
    # 沒有模板的話，你也可以改成回傳簡單訊息
    try:
        return render_template("index.html")
    except Exception:
        return "Soilmove Viewer is running.", 200


@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({
        "ok": True,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "repo_data_exists": os.path.exists(REPO_DATA_FILE),
        "tmp_data_exists": os.path.exists(TMP_DATA_FILE),
    }), 200


@app.route("/api/local_data", methods=["GET"])
def api_local_data():
    """
    讀取資料（建議前端只打這支）：
    1) 若 /tmp/data.json 存在（例如你有上傳暫存），就優先回傳它
    2) 否則回傳 repo 內 data.json（由 GitHub Actions 定期更新）
    """
    try:
        data = load_json_safely(TMP_DATA_FILE)
        if data is not None:
            data["source"] = "tmp"
            return jsonify(data), 200

        data = load_json_safely(REPO_DATA_FILE)
        if data is not None:
            data["source"] = "repo"
            return jsonify(data), 200

        return jsonify({
            "updated": "尚未有資料",
            "data": [],
            "source": "none"
        }), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/update", methods=["GET"])
def api_update():
    """
    不再由 Vercel/Flask 去抓 soilmove.tw（會被 WAF 擋）。
    更新由 GitHub Actions 定期跑 scripts/update_data.py 覆寫 repo 的 data.json。
    """
    return jsonify({
        "message": "Updates are handled by GitHub Actions. Please check data.json commits in the repo.",
        "hint": "Call /api/local_data to get the latest cached/repo data."
    }), 200


@app.route("/api/upload_default_json", methods=["POST"])
def upload_default_json():
    """
    接收前端上傳 JSON，僅暫存到 /tmp/data.json（serverless cache）
    注意：/tmp 不保證持久，冷啟動或換 instance 會消失
    """
    try:
        if "file" not in request.files:
            return jsonify({"error": "未選取檔案"}), 400

        file = request.files["file"]
        if not file or file.filename == "":
            return jsonify({"error": "檔案名稱不可為空"}), 400

        # 先解析驗證（避免寫入壞檔）
        try:
            payload = json.load(file.stream)
        except Exception as e:
            return jsonify({"error": f"JSON 解析失敗: {str(e)}"}), 400

        if not isinstance(payload, dict):
            return jsonify({"error": "JSON 格式不正確：頂層必須是物件(dict)，需包含 updated/data"}), 400

        updated = payload.get("updated", "手動更新")
        data = payload.get("data", [])
        if not isinstance(data, list):
            return jsonify({"error": "JSON 格式不正確：data 必須是陣列(list)"}), 400

        # 寫入 /tmp
        os.makedirs(os.path.dirname(TMP_DATA_FILE), exist_ok=True)
        with open(TMP_DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

        return jsonify({
            "updated": updated,
            "count": len(data),
            "stored": "tmp",
            "note": "This is temporary storage on serverless (/tmp). It may disappear after cold start."
        }), 200

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