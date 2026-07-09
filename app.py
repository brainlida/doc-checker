import streamlit as st
import docx
from docx.shared import Pt, Cm
import re

st.set_page_config(page_title="實驗報告格式檢查器", page_icon="📝")

st.title("📝 實驗報告格式檢查系統")
st.subheader("請上傳您的 Word 檔案 (.docx) 進行格式檢查")

# 老師的規格常數設定
TARGET_FONT_CHI = "標楷體"
TARGET_FONT_ENG = "Times New Roman"
TARGET_LINE_SPACING = 1.0

# 檢查邏輯函式
def check_report_format(file):
    try:
        doc = docx.Document(file)
    except Exception as e:
        return ["❌ 檔案讀取失敗，請確保上傳的是標準的 .docx 格式檔案。"], False

    report = []
    score = 100
    
    # --- 1. 檢查頁面邊界 ---
    section = doc.sections[0]
    # 2.5 cm 約等於 0.984 英吋
    margin_top = section.top_margin.cm
    margin_bottom = section.bottom_margin.cm
    margin_left = section.left_margin.cm
    margin_right = section.right_margin.cm
    
    if all(abs(m - 2.5) < 0.1 for m in [margin_top, margin_bottom, margin_left, margin_right]):
        report.append("✅ **頁面邊界**：符合規定 (上下左右皆為 2.5cm)")
    else:
        report.append(f"❌ **頁面邊界**：不符規定！目前為 上:{margin_top:.1f}cm, 下:{margin_bottom:.1f}cm, 左:{margin_left:.1f}cm, 右:{margin_right:.1f}cm (規定應為 2.5cm)")
        score -= 10

    # --- 2. 檢查標題、姓名與日期 ---
    first_para = doc.paragraphs[0] if doc.paragraphs else None
    if first_para:
        text = first_para.text
        if "不同溶液中草履蟲的移動速度" in text:
            # 檢查是否有姓名和日期（簡單判斷長度或關鍵字，因學生名字不固定，這裡檢查是否有其他文字與數字）
            has_date = any(char.isdigit() for char in text)
            if len(text) > 15 and has_date:
                report.append("✅ **標題與基本資料**：標題正確，且已包含實驗者與日期。")
            else:
                report.append("❌ **標題與基本資料**：有看到標題，但**別忘記標題旁邊要有實驗者(你)和日期(今天)**！")
                score -= 10
        else:
            report.append("❌ **標題**：第一行找不到正確的標題「不同溶液中草履蟲的移動速度」")
            score -= 15

    # --- 3. 檢查內文段落、字體、粗體、行距 ---
    bold_keywords = ["不同溶液中草履蟲的移動速度", "一、目的", "二、器材", "三、步驟"]
    font_error = False
    color_error = False
    spacing_error = False
    bold_error = False
    indent_error = False

    for para in doc.paragraphs:
        if not para.text.strip():
            continue
        
        # 檢查行距
        if para.paragraph_format.line_spacing and abs(para.paragraph_format.line_spacing - TARGET_LINE_SPACING) > 0.1:
            spacing_error = True

        # 檢查 步驟(一)~(五) 的縮排
        if re.match(r"^\s*\(六*肌肉|一|二|三|四|五\)", para.text.strip()) and "三、步驟" not in para.text:
            # 判斷是否有縮排（首行縮排或左縮排）
            if not (para.paragraph_format.first_line_indent or para.paragraph_format.left_indent):
                # 也有可能學生是用空格縮排
                if not para.text.startswith("  ") and not para.text.startswith("　"):
                    indent_error = True

        # 檢查粗體與字體
        is_keyword_para = any(kw in para.text for kw in bold_keywords)
        
        for run in para.runs:
            if not run.text.strip():
                continue
            
            # 粗體檢查
            if is_keyword_para:
                if any(kw in run.text for kw in bold_keywords) and not run.bold:
                    bold_error = True
            else:
                if run.bold:
                    bold_error = True
            
            # 顏色檢查 (預設或黑色皆可)
            if run.font.color and run.font.color.rgb and run.font.color.rgb != docx.shared.RGBColor(0,0,0):
                color_error = True

    if spacing_error:
        report.append("❌ **行距**：發現有段落不是 1.0 行行距。")
        score -= 10
    else:
        report.append("✅ **行距**：全符合 1.0 行規定。")

    if bold_error:
        report.append("❌ **粗體規定**：格式不符。規定只有標題、「一、目的」、「二、器材」、「三、步驟」為粗體，其他應為正常字體。")
        score -= 10
    else:
        report.append("✅ **粗體規定**：符合規定。")

    if color_error:
        report.append("❌ **文字顏色**：偵測到非黑色的文字或數字！")
        score -= 10
    else:
        report.append("✅ **文字顏色**：全為黑色。")

    if indent_error:
        report.append("❌ **步驟縮排**：『三、步驟』下面的 (一)~(五) 必須縮排 2 個字元。")
        score -= 10
    else:
        report.append("✅ **步驟縮排**：步驟縮排符合規定。")

    # --- 4. 檢查表格 (器材 4x2 表格) ---
    if len(doc.tables) > 0:
        table = doc.tables[0]
        rows = len(table.rows)
        cols = len(table.columns)
        
        if rows == 4 and cols == 2:
            report.append(f"✅ **器材表格**：成功使用 {rows}x{cols} 的表格呈現。")
        else:
            report.append(f"❌ **器材表格**：規格不符！規定要 4x2 表格，目前偵測到的是 {rows}x{cols}。")
            score -= 15
    else:
        report.append("❌ **器材表格**：沒偵測到表格！八個器材請用 4*2 表格呈現。")
        score -= 20

    return report, score

# 網頁上傳介面
uploaded_file = st.file_uploader("請選擇您的實驗報告 Word 檔案 (.docx)", type=["docx"])

if uploaded_file is not None:
    with st.spinner('正在檢查格式中，請稍候...'):
        results, final_score = check_report_format(uploaded_file)
        
        st.write("---")
        st.header("📊 檢查結果報告")
        
        if final_score == 100:
            st.balloons()
            st.success(f"💯 太棒了！您的格式完全符合規定！得分：{final_score} 分")
        elif final_score >= 60:
            st.warning(f"⚠️ 檢查完成，部分格式有誤。得分：{final_score} 分")
        else:
            st.error(f"❌ 格式錯誤較多，請修正後重新上傳。得分：{final_score} 分")
            
        for item in results:
            st.write(item)
