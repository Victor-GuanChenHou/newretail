def outputexcel():
    import pyodbc
    import pandas as pd
    from openpyxl import Workbook
    import pandas as pd
    from datetime import datetime
    from openpyxl.styles import PatternFill
    import os
    from openpyxl import Workbook
    
    # 讀取進價.xlsx文件
    file_path = '進價.xlsx'  # 請根據實際文件位置修改

    # 使用pandas的read_excel函數讀取Excel文件
    inputdata = pd.read_excel(file_path, engine='openpyxl')
    try:
        # 建立與 SQL Server 的連線
        conn = pyodbc.connect(
            "DRIVER={ODBC Driver 17 for SQL Server};"
            "SERVER=10.140.0.5;"  # 替換為完整的伺服器名稱
            "DATABASE=Lakaffa004;"
            "UID=sa;"  # 使用 sa 作為使用者名稱
            "PWD=dsc@42756204;"  # 替換為 sa 的密碼
            "Trusted_Connection=no;" # 明確使用 SQL Server 認證
              
        )

        # 用戶選擇的品牌（假設這來自某個介面） 
        CombMO002_ItemIndex = 'S02012'

        # SQL1 查詢語句
        Sql1 = """
            SELECT PLU_ID, PLU_NAME, PLU_SPEC, BIN_NAME, BIN_DESC, LABEL_F, UNIT, EXP_DATE, QTY, BOOKING, AVAILABLE
            FROM View_WMS02
            WHERE 1=1
        """

        # 根據品牌選擇拼接 Sql2
        if CombMO002_ItemIndex == "0":
            Sql2 = ""
        else:
            MA076 = CombMO002_ItemIndex
            Sql2 = """
                AND RTRIM(PLU_ID) COLLATE Chinese_Taiwan_Stroke_BIN IN (
                    SELECT RTRIM(MI002)
                    FROM COPMA
                    INNER JOIN LKFMJ ON MJ001 = MA001
                    LEFT JOIN LKFMH ON MH001 = MJ002
                    LEFT JOIN LKFMI ON MI001 = MH001
                    WHERE 1=1
                    AND MA076 = ?
                    UNION
                    SELECT RTRIM(MD002) AS MI002
                    FROM COPMA
                    INNER JOIN LKFMG ON MG001 = MA001
                    LEFT JOIN LKFMC ON MC001 = MG002
                    LEFT JOIN LKFMD ON MD001 = MC001
                    WHERE 1=1
                    AND MA076 = ?
                )
            """


        # 完整的查詢語句
        Sql3 =   " ORDER BY PLU_ID, BIN_NAME, EXP_DATE, LABEL_F"

        # 最終的 SQL 查詢語句
        Sql = Sql1 + Sql2 + Sql3

        # 使用 pandas 執行查詢並傳遞參數以防止 SQL 注入
        df = pd.read_sql(Sql, conn, params=(MA076, MA076))
    except:
        message='連線失敗'
        
        return message
        
    df.columns =['品號','品名','規格','標的代號','標的名稱','貨品標籤','單位','有效日期','數量','鎖定數量','可用數量']

    data={}
    shelflifeday=[]
    day_1=[]
    day_2=[]
    day_3=[]
    day_4=[]
    shelflifeday_2=[]
    end=[]
    data['品號']=df['品號']
    data['品名']=df['品名']

    for z in range(len(data['品號'])):
        filtered_df = inputdata[inputdata['品號'] == data['品號'][z]]
        if filtered_df.empty:
            shelflifeday.append(0)
        else:
            shelflifeday.append(filtered_df['保存期限\n(天)'].iloc[0])
    data['保存期限天數']=shelflifeday
    data['允收期限2/3'] = [2/3] * len(df['品名'])
    data['允收期限1/2'] = [1/2] * len(df['品名'])
    for z in range(len(data['品號'])):
        if data['保存期限天數'][z] != 0:
            day_1.append(int(data['保存期限天數'][z])*2/3)
            day_2.append(int(data['保存期限天數'][z])*1/2)
        else:
            day_1.append(0)
            day_2.append(0)
        str=df['有效日期'][z]
        str_2=f"{str[:4]}/{str[4:6]}/{str[6:]}"
        shelflifeday_2.append(str_2)
        target_date = datetime.strptime(str_2, '%Y/%m/%d')
        end.append( (target_date - datetime.today()).days)
    
    data['計算天數2/3']=day_1
    data['計算天數1/2']=day_2
    data['可用數量']=df['可用數量']
    data['有效日期']=shelflifeday_2
    data['結果']=end
    for z in range(len(data['品號'])):
        if data['計算天數2/3'][z]==0:
            day_3.append(None)
            day_4.append(None)
        else:
            day_3.append(int(data['結果'][z]-data['計算天數2/3'][z]))
            day_4.append(int(data['結果'][z]-data['計算天數1/2'][z]))
    data['1/3允收餘天數']=day_3
    data['1/2允收餘天數']=day_4
    data=pd.DataFrame(data)
    data.columns=['品號','品名','保存期限天數','允收期限2/3','允收期限1/2','計算天數2/3','計算天數1/2','可用數量','有效日期','結果','1/3允收餘天數','1/2允收餘天數']

    # 根據 "品號" 進行分組，並對 "可用數量" 求和
    df_grouped = df.groupby(['品號','品名'],    as_index=False)['可用數量'].sum()
    df_grouped.columns=['品號','品名','總和']
    # 重新設置索引，將總和欄位插入最後

    # 創建 Excel 工作簿
    wb = Workbook()

    # 添加工作表
    ws = wb.active
    ws.title = "王座青田明細"
    header_fill = PatternFill(start_color="9999FF", end_color="9999FF", fill_type="solid")

    # 在 A 到 K 欄顯示 df 的欄位名稱和數據
    for col_idx, col_name in enumerate(df.columns, start=1):
        cell=ws.cell(row=1, column=col_idx, value=col_name)  # 放置欄位名稱
        cell.fill = header_fill  
    # 將 df 輸出到 A 到 K 欄
    for r_idx, row in df.iterrows():
        for c_idx, value in enumerate(row):
            ws.cell(row=r_idx + 2, column=c_idx + 1, value=value)
    ##################################
    for col_idx, col_name in enumerate(df_grouped.columns, start=13):  # M 欄開始
        cell=ws.cell(row=1, column=col_idx, value=col_name)  # 放置欄位名稱
        cell.fill = header_fill  
    # 將 df_grouped 輸出到 M 到 O 欄
    for r_idx, row in df_grouped.iterrows():
        for c_idx, value in enumerate(row):
            ws.cell(row=r_idx + 2, column=c_idx + 13, value=value)
    #############################
    for col_idx, col_name in enumerate(data.columns, start=18):  
        cell=ws.cell(row=1, column=col_idx, value=col_name)  # 放置欄位名稱
        cell.fill = header_fill  
    for r_idx, row in data.iterrows():
        for c_idx, value in enumerate(row):
            ws.cell(row=r_idx + 2, column=c_idx + 18, value=value)

    # 儲存 Excel 文件
    filepath='./'

    # 要儲存的檔案名稱
    filename = "王座青田明細.xlsx"
    full_path = os.path.join(filepath, filename)

        # 檢查檔案是否存在，如果存在則新增編號
    counter = 1
    while os.path.exists(full_path):
        filename = f"王座青田明細({counter}).xlsx"
        full_path = os.path.join(filepath, filename)
        counter += 1
    wb.save(full_path)

    # output_file = "test.xlsx"
    # df.to_excel(output_file, index=False, engine='openpyxl')

    # 關閉連線
    conn.close()
    
    message=full_path
    return message
