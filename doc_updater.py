import pandas as pd
from docx import Document
import os
import openpyxl

# 定義核心報廢 ID 列表 (依據 Excel 灰色背景與使用者截圖)
OBSOLETE_IDS = [
    'HQ-PD.497', 'HQ-PD.498', 'HQ-PD.499', 'HQ-PD.500', 
    'HQ-PD.501', 'HQ-PD.502', 'HQ-PD.514',
    'HQ-PD.748', 'HQ-PD.749', 'HQ-PD.750', 'HQ-PD.751'
]

def create_mapping(excel_path):
    xl = pd.ExcelFile(excel_path)
    sheet_name = xl.sheet_names[4]
    df = pd.read_excel(xl, sheet_name=sheet_name, header=None).fillna('')
    
    mapping = {}
    for i in range(1, len(df)):
        new_val = str(df.iloc[i, 1]).strip()
        old_val = str(df.iloc[i, 5]).strip()
        
        if old_val and new_val and old_val != new_val:
            old_clean = old_val.replace('\n', ' ')
            mapping[old_clean] = new_val
            
            if '_' in old_clean:
                old_id = old_clean.split('_')[0].strip()
                new_id = new_val.split('_')[0].strip()
                if old_id and new_id and old_id != new_id:
                    mapping[old_id] = new_id
            
            old_name_only = old_clean.split('_', 1)[-1].strip() if '_' in old_clean else old_clean
            new_name_only = new_val.split('_', 1)[-1].strip() if '_' in new_val else new_val
            if old_name_only and new_name_only and old_name_only != new_name_only:
                mapping[old_name_only] = new_name_only

    mapping["操作指導書"] = "作業指導書"
    return mapping

def process_docx(doc_path, mapping, output_path):
    doc = Document(doc_path)
    stats = {"replaced": 0, "obsolete_removed": 0}

    # 1. 處理段落 (使用徹底移除邏輯)
    # 我們需要倒序處理段落列表，以免移除元素時索引出錯
    paras = list(doc.paragraphs)
    for p in paras:
        text = p.text
        if any(oid in text for oid in OBSOLETE_IDS):
            p._element.getparent().remove(p._element)
            stats["obsolete_removed"] += 1
            continue
            
        # 正常替換
        new_text = text
        for old_str, new_str in mapping.items():
            if old_str in new_text:
                new_text = new_text.replace(old_str, new_str)
        if new_text != text:
            p.text = new_text
            stats["replaced"] += 1

    # 2. 處理表格 (維持清空邏輯，不破壞排版)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    text = p.text
                    if any(oid in text for oid in OBSOLETE_IDS):
                        p.text = "" # 清空文字，格子還在
                        stats["obsolete_removed"] += 1
                    else:
                        new_text = text
                        for old_str, new_str in mapping.items():
                            if old_str in new_text:
                                new_text = new_text.replace(old_str, new_str)
                        if new_text != text:
                            p.text = new_text
                            stats["replaced"] += 1

    # 3. 標頭標尾 (替換邏輯)
    for section in doc.sections:
        for header in [section.header, section.first_page_header, section.even_page_header]:
            if header:
                for p in header.paragraphs:
                    for old_str, new_str in mapping.items():
                        if old_str in p.text:
                            p.text = p.text.replace(old_str, new_str)
        for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
            if footer:
                for p in footer.paragraphs:
                    for old_str, new_str in mapping.items():
                        if old_str in p.text:
                            p.text = p.text.replace(old_str, new_str)

    doc.save(output_path)
    return stats

if __name__ == "__main__":
    excel_file = r"temp/ISO17025(2017)實驗室管理系統文件總覽表.xlsx"
    docx_file = r"temp/template.docx"
    output_file = r"HQ-PD-P-038_v1.1_實驗室測試方法管理程序_更新版.docx"
    
    print("正在提取更新映射表...")
    mapping_dict = create_mapping(excel_file)
    
    print("\n正在執行深度清理與更新...")
    result_stats = process_docx(docx_file, mapping_dict, output_file)
    
    print("\n完成！")
    print(f"一般資訊替換: {result_stats['replaced']} 處")
    print(f"報廢資訊徹底移除 (含段落刪除): {result_stats['obsolete_removed']} 處")
    print(f"檔案已儲存至: {output_file}")
