from flask import Flask, send_file, after_this_request
import os
import time
import threading  # 用來延遲刪除檔案
import urllib.parse
import subfunction as sub  # 假設 subfunction.py 中有 outputexcel 函數

app = Flask(__name__)

@app.route('/')
def index():
    return send_file('./templates/index.html')

@app.route('/generate')
def generate_file_route():
    # 產生 Excel 檔案
    file_path = sub.outputexcel()

    if file_path.startswith("連線失敗"):
        return file_path  # 返回錯誤訊息

    # 設定下載的檔案名稱
    filename = "王座青田明細.xlsx"
    encoded_filename = urllib.parse.quote(filename)

    # **確保在回應結束後延遲刪除檔案**
    @after_this_request
    def cleanup(response):
        def delayed_delete():
            time.sleep(10)  # **等 10 秒後再刪除**
            try:
                os.remove(file_path)
                print(f"✅ 已刪除檔案: {file_path}")
            except Exception as e:
                print(f"⚠️ 刪除檔案失敗: {e}")

        threading.Thread(target=delayed_delete, daemon=True).start()  # **開啟背景執行緒**

        return response  # **確保回應繼續執行**

    # 建立回應物件
    response = send_file(
        file_path,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # 設定 Content-Disposition 標頭，支援 UTF-8 檔案名稱
    response.headers['Content-Disposition'] = f"attachment; filename*=utf-8''{encoded_filename}"

    return response  # **回應完成後才會執行 cleanup**
    
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=70, debug=True)