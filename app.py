import streamlit as st
import docx
from docx.shared import Cm, Pt
import openai
import os

# 網頁標題設定
st.set_page_config(page_title="慈文國中數資班word閱卷系統", page_icon="📝")
st.title("📝 慈文國中數資班word閱卷系統")
st.subheader("請上傳您的作業檔案 (.docx)，系統將進行嚴格的格式審查。")

# 讓老師可以在網頁後台或側邊欄輸入 API Key (也可以在 Streamlit 設定中隱藏)
api_key = st.sidebar.text_input("請輸入 OpenAI API Key", type="password")

uploaded_file = st.file_uploader("選擇上傳 Word 檔案 (.docx)", type=["docx"])

# 解析 Word 格式的函數
def analyze_docx(file_path):
    doc = docx.Document(file_path)
    report = []
    
    # 1. 檢查邊界 (轉為公分)
    sections = doc.sections
    for i, section in enumerate(sections):
        top = round(section.top_margin.cm, 2) if section.top_margin else "未知"
        bottom = round(section.bottom_margin.cm, 2) if section.bottom_margin else "未知"
        left = round(section.left_margin.cm, 2) if section.left_margin else "未知"
        right = round(section.right_margin.cm, 2) if section.right_margin else "未知"
        report.append(f"[版面邊界] 上:{top}cm, 下:{bottom}cm, 左:{left}cm, 右:{right}cm")

    # 2. 檢查段落格式（行距、縮排、字體、字級、粗體）
    report.append("\n[段落內文與格式明細]")
    for i, p in enumerate(doc.paragraphs):
        if not p.text.strip():
            continue
        
        # 行距與縮排
        line_spacing = p.paragraph_format.line_spacing if p.paragraph_format.line_spacing else "預設(1.0)"
        indent = p.paragraph_format.left_indent.paragraphs if p.paragraph_format.left_indent else 0
        # 轉成字元大約值 (1字元約等於 0.423 公分或特定點數，這裡直接抓原始設定或交給 AI 輔助判斷)
        
        # 讀取文字區塊的細部格式 (Run)
        run_info = []
        for run in p.runs:
            font_name = run.font.name if run.font.name else "隨系統預設"
            font_size = run.font.size.pt if run.font.size else "隨系統預設"
            is_bold = run.bold if run.bold is not None else False
            color = run.font.color.rgb if run.font.color else "預設黑色"
            run_info.append(f"'{run.text}'(字體:{font_name}, 大小:{font_size}, 粗體:{is_bold}, 顏色:{color})")
            
        report.append(f"段落 {i+1}: {p.text} \n   -> 格式: 行距:{line_spacing}行, 詳細設定:[{', '.join(run_info)}]")

    # 3. 檢查表格結構
    report.append("\n[文件內表格結構]")
    for t_idx, table in enumerate(doc.tables):
        rows = len(table.rows)
        cols = len(table.columns)
        report.append(f"表格 {t_idx+1}: 共有 {rows} 列(Row) x {cols} 欄(Column)")
        # 檢查框線 (docx 的 table.style.name 可以抓到是否為無框線或基本表格)
        report.append(f"   -> 表格樣式名稱: {table.style.name}")
        for r_idx, row in enumerate(table.rows):
            row_text = [cell.text.strip() for cell in row.cells]
            report.append(f"      第 {r_idx+1} 列文字: {row_text}")

    return "\n".join(report)

if uploaded_file is not None:
    filename = uploaded_file.name
    st.info(f"📁 已偵測到上傳檔名：{filename}")
    
    # 執行檔案格式解析
    with st.spinner("正在抽取 Word 底層格式數據..."):
        try:
            docx_details = analyze_docx(uploaded_file)
        except Exception as e:
            st.error(f"檔案解析失敗，請確保是標準的 docx 檔案。錯誤訊息: {e}")
            docx_details = None

    if docx_details and api_key:
        st.success("格式數據抽取成功！正在啟動 AI 閱卷老師評分...")
        
        # 建立結合了「標準」與「學生檔案實際格式」的 AI 提示詞
        prompt = f"""
你是一位嚴格的資訊課閱卷老師。請根據以下標準，針對學生上傳的 Word 文件「底層格式數據」與「檔名」進行格式審查。
初始分數為 100 分，每發現一處未完成或不符規規定（包含文字打錯、不全、格式不對），請明確指出並「扣 5 分」，扣至 0 分為止。

【嚴格格式檢核標準】
1. 檔名：必須符合「7X-XXX-文件三寶測驗」格式（X為任意英文或中文字，例如 7A-102-文件三寶測驗）。
2. 基本資訊：標題旁邊必須包含「實驗者姓名」與「日期」。
3. 邊界與行距：上下左右邊界必須皆為 2.5cm（容許正負0.05cm微差），行距必須為 1.0 行。
4. 字體與顏色：中文字全為「標楷體」，英文與數字全為「半形 Times New Roman」，文字與數字顏色皆為黑色（或預設黑）。
5. 字級大小：只有標題是「14號字」，其餘所有內容（包含一二三、器材、步驟）皆為「12號字」。
6. 粗體限制：只有「不同溶液下，草履蟲的移動速度」、「一、目的」、「二、器材」、「三、步驟」這四個部分為粗體（True），其餘所有內文必須為正常字體（False）。
7. 器材呈現：八個器材必須使用 4*2 (4列2欄) 表格呈現，文字靠左對齊，且表格樣式必須是無框線（例如 Table Grid 且無外顯框線，或樣式名稱含 No Border/Normal）。
8. 步驟縮排：「三、步驟」底下的 (一) 到 (五) 內文，每行開頭都必須縮排 2 個字元。

-----------------------------------------
【學生的檔案實際數據如下】
檔案名稱: {filename}

解析出來的底層 XML 格式報告:
{docx_details}
-----------------------------------------

請根據上述實際數據，嚴格比對 1~8 點標準。
請用以下格式輸出閱卷結果：

### 👨‍🏫 閱卷老師評分報告
- **最終得分**：[分數] 分

#### 🟢 合格項目
- （列出完全符合標準的項目與簡短稱讚）

#### 🔴 不合格項目與扣分明細
- （明確指出哪一條不符、實際數據是什麼、扣 5 分。例如：表格非4*2而是2*4，扣5分）
"""

        # 呼叫 OpenAI API (使用 gpt-4o 確保強大的推理解析能力)
        try:
            client = openai.OpenAI(api_key=api_key)
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "你是一位精準、挑剔且絕對公正的學校資訊課閱卷評分機器人。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.0
            )
            
            # 顯示結果
            st.write("---")
            st.markdown(response.choices[0].message.content)
            
        except Exception as e:
            st.error(f"AI 連線失敗，請檢查 API Key 是否正確。錯誤報告: {e}")
            
    elif not api_key:
        st.warning("⚠️ 請在左側欄位輸入您的 OpenAI API Key 才能啟動 AI 老師評分功能。")
