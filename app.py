from flask import Flask, send_file, request, after_this_request,jsonify,render_template
import os
import time
import threading  # 用來延遲刪除檔案
import urllib.parse
import subfunction as sub  
import pandas as pd
from openpyxl import Workbook
import pandas as pd
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
def home():
    return render_template("home.html")
@app.route('/index')
def index():
    return render_template("index.html")
@app.route('/productno')
def productno():
    file_path = '進價.xlsx'  # 請根據實際文件位置修改

    # 使用pandas的read_excel函數讀取Excel文件
    inputdata = pd.read_excel(file_path, engine='openpyxl')
    inputdata.columns = ['品號', '商品品名', '保存期限(天)', '成本']

    return render_template("productno.html",product_list=inputdata)
@app.route('/adddata', methods=["POST"])
def adddata():
    try:
        data = request.get_json()
        nu = data.get('nu')
        name = data.get('name')
        day = data.get('day')
        money = data.get('money')
        file_path = '進價.xlsx'  # 請根據實際文件位置修改

        # 使用pandas的read_excel函數讀取Excel文件
        inputdata = pd.read_excel(file_path, engine='openpyxl')
        inputdata.columns = ['品號', '商品品名', '保存期限(天)', '成本']
        new_row = pd.DataFrame([{
        '品號': nu,
        '商品品名': name,
        '保存期限(天)': day,
        '成本': money
        }])
        updated_data = pd.concat([inputdata, new_row], ignore_index=True)
        updated_data.to_excel(file_path, index=False, engine='openpyxl')
        return jsonify({'success': True, 'message': '資料已成功新增'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})
@app.route('/editdata', methods=["POST"])
def editdata():
    try:
        data = request.get_json()

        # 新值
        nu = data.get('enu')
        name = data.get('ename')
        day = data.get('eday')
        money = data.get('emoney')

        # 舊值，用於找要修改的列
        ori_nu = data.get('ori_nu')
        ori_name = data.get('ori_name')
        ori_day = data.get('ori_day')
        ori_money = data.get('ori_money')

        file_path = '進價.xlsx'  # Excel 檔案位置

        # 讀取資料
        df = pd.read_excel(file_path, engine='openpyxl')
        df.columns = ['品號', '商品品名', '保存期限(天)', '成本']

        # 尋找要修改的列條件（全欄比對避免只憑品號可能錯誤）
        match = (
            (df['品號'] == ori_nu) &
            (df['商品品名'] == ori_name) &
            (df['保存期限(天)'].astype(str) == str(ori_day)) &
            (df['成本'].astype(str) == str(ori_money))
        )

        if match.sum() == 0:
            return jsonify({'success': False, 'error': '找不到要修改的資料'})

        # 更新符合條件的列
        df.loc[match, '品號'] = nu
        df.loc[match, '商品品名'] = name
        df.loc[match, '保存期限(天)'] = day
        df.loc[match, '成本'] = money

        # 儲存回 Excel
        df.to_excel(file_path, index=False, engine='openpyxl')

        return jsonify({'success': True, 'message': '資料已成功更新'})
    
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})
@app.route('/deletdata', methods=["POST"])
def deletdata():
    try:
        data = request.get_json()

       
        dnu = data.get('dnu')
      
        file_path = '進價.xlsx'  # Excel 檔案位置

        # 讀取資料
        df = pd.read_excel(file_path, engine='openpyxl')
        df.columns = ['品號', '商品品名', '保存期限(天)', '成本']

        # 尋找要修改的列條件（全欄比對避免只憑品號可能錯誤）
        match = (
            (df['品號'] == dnu) )

        if match.sum() == 0:
            return jsonify({'success': False, 'error': '找不到要修改的資料'})

        # 更新符合條件的列
        df = df[~match]
        # 儲存回 Excel
        df.to_excel(file_path, index=False, engine='openpyxl')

        return jsonify({'success': True, 'message': '資料已成功更新'})
    
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/icon')
def icon():
    return send_file('./templates/kingza.ico')
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
    app.run(host='0.0.0.0', port=9001, debug=True)
