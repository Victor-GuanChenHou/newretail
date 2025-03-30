from flask import Flask, send_file, request, after_this_request, jsonify
import os
import time
import threading  # 用來延遲刪除檔案
import urllib.parse
import subfunction as sub  # 假設 subfunction.py 中有 outputexcel 函數

app = Flask(__name__)

# ✅ 設定上傳和處理後的檔案夾
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER

# ✅ 確保目錄存在
for folder in [UPLOAD_FOLDER, PROCESSED_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

@app.route('/')
def index():
    return send_file('./templates/index.html')

# ✅ 下載報表（原有功能）
@app.route('/generate')
def generate_file_route():
    file_path = sub.outputexcel()

    if file_path.startswith("連線失敗"):
        return file_path

    filename = "王座青田明細.xlsx"
    encoded_filename = urllib.parse.quote(filename)

    @after_this_request
    def cleanup(response):
        def delayed_delete():
            time.sleep(10)
            try:
                os.remove(file_path)
                print(f"✅ 已刪除檔案: {file_path}")
            except Exception as e:
                print(f"⚠️ 刪除檔案失敗: {e}")

        threading.Thread(target=delayed_delete, daemon=True).start()
        return response

    response = send_file(
        file_path,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response.headers['Content-Disposition'] = f"attachment; filename*=utf-8''{encoded_filename}"

    return response

# ✅ 上傳並處理報表（新增功能）
@app.route('/upload_and_process', methods=['POST'])
def upload_and_process():
    if 'file' not in request.files:
        return jsonify({"error": "沒有檔案部分！"}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({"error": "沒有選擇檔案！"}), 400

    if file:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)
        
        # ✅ 執行處理函數，產生 Excel
        processed_file_path = sub.fullday(file_path)

        if processed_file_path.startswith("連線失敗"):
           return jsonify({"error": processed_file_path}), 500

        filename = file_path
        encoded_filename = urllib.parse.quote(filename)

        @after_this_request
        def cleanup(response):
            def delayed_delete():
                time.sleep(10)
                try:
                    os.remove(processed_file_path)
                    os.remove(file_path)
                    print(f"✅ 已刪除檔案: {processed_file_path}")
                except Exception as e:
                    print(f"⚠️ 刪除檔案失敗: {e}")

            threading.Thread(target=delayed_delete, daemon=True).start()
            return response

        response = send_file(
            processed_file_path,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        response.headers['Content-Disposition'] = f"attachment; filename*=utf-8''{encoded_filename}"

        return response
@app.route('/upload_ALL', methods=['POST'])
def upload_ALL():
    if 'file' not in request.files:
        return jsonify({"error": "沒有檔案部分！"}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({"error": "沒有選擇檔案！"}), 400

    if file:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)
        
        # ✅ 執行處理函數，產生 Excel
        processed_file_path = sub.ALL(file_path)

        if processed_file_path.startswith("連線失敗"):
           return jsonify({"error": processed_file_path}), 500

        filename = file_path
        encoded_filename = urllib.parse.quote(filename)

        @after_this_request
        def cleanup(response):
            def delayed_delete():
                time.sleep(10)
                try:
                    os.remove(processed_file_path)
                    os.remove(file_path)
                    print(f"✅ 已刪除檔案: {processed_file_path}")
                except Exception as e:
                    print(f"⚠️ 刪除檔案失敗: {e}")

            threading.Thread(target=delayed_delete, daemon=True).start()
            return response

        response = send_file(
            processed_file_path,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        response.headers['Content-Disposition'] = f"attachment; filename*=utf-8''{encoded_filename}"

        return response

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=70, debug=True)
