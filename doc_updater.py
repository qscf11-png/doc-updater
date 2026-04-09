import pandas as pd
from docx import Document
import os

def create_mapping(excel_path):
    xl = pd.ExcelFile(excel_path)
    # 我們鎖定 Index 4 (三階文件)
    sheet_name = xl.sheet_names[4]
    df = pd.read_excel(xl, sheet_name=sheet_name, header=None).fillna('')
    
    mapping = {}
    # 根據研究：F 欄 (index 5) 是舊的，B 欄 (index 1) 是新的
    # 從第 1 列開始（跳過標題）
    for i in range(1, len(df)):
        new_val = str(df.iloc[i, 1]).strip()
        old_val = str(df.iloc[i, 5]).strip()
        
        if old_val and new_val and old_val != new_val:
            # 去除可能干擾匹配的換行符或多餘空格
            old_clean = old_val.replace('\n', ' ')
            mapping[old_clean] = new_val
            
            # 有時候編號跟名稱是分開匹配的，我們也加入單純編號的匹配 (如果 old 包含編號)
            # 例如 "HQ-PD.516_名稱" -> 我們也匹配 "HQ-PD.516"
            if '_' in old_clean:
                old_id = old_clean.split('_')[0].strip()
                new_id = new_val.split('_')[0].strip()
                if old_id and new_id and old_id != new_id:
                    mapping[old_id] = new_id
            
            # 新增：純名稱匹配 (移除編號後的剩餘部分)
            old_name_only = old_clean.split('_', 1)[-1].strip() if '_' in old_clean else old_clean
            new_name_only = new_val.split('_', 1)[-1].strip() if '_' in new_val else new_val
            if old_name_only and new_name_only and old_name_only != new_name_only:
                mapping[old_name_only] = new_name_only

    # 手動加入通用名稱變更
    mapping["操作指導書"] = "作業指導書"

    return mapping

def replace_text_in_docx(doc_path, mapping, output_path):
    doc = Document(doc_path)
    
    # 建立替換統計
    stats = {"paragraphs": 0, "tables": 0, "headers": 0}

    def search_replace(text):
        if not text: return text
        new_text = text
        for old_str, new_str in mapping.items():
            if old_str in new_text:
                new_text = new_text.replace(old_str, new_str)
        return new_text

    # 1. 段落
    for p in doc.paragraphs:
        original = p.text
        updated = search_replace(original)
        if original != updated:
            p.text = updated
            stats["paragraphs"] += 1

    # 2. 表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    original = p.text
                    updated = search_replace(original)
                    if original != updated:
                        p.text = updated
                        stats["tables"] += 1

    # 3. 標頭標尾
    for section in doc.sections:
        for header in [section.header, section.first_page_header, section.even_page_header]:
            if header:
                for p in header.paragraphs:
                    original = p.text
                    updated = search_replace(original)
                    if original != updated:
                        p.text = updated
                        stats["headers"] += 1
        for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
            if footer:
                for p in footer.paragraphs:
                    original = p.text
                    updated = search_replace(original)
                    if original != updated:
                        p.text = updated
                        stats["headers"] += 1

    doc.save(output_path)
    return stats

if __name__ == "__main__":
    excel_file = r"temp/ISO17025(2017)實驗室管理系統文件總覽表.xlsx"
    docx_file = r"temp/template.docx"
    output_file = r"HQ-PD-P-038_v1.1_實驗室測試方法管理程序_更新版.docx"
    
    print("正在提取 Excel 映射表...")
    mapping_dict = create_mapping(excel_file)
    print(f"成功提取 {len(mapping_dict)} 組映射關係。")
    
    # 預覽部分映射
    print("映射範例 (前 5 組):")
    for i, (k, v) in enumerate(list(mapping_dict.items())[:5]):
        print(f"  {k} -> {v}")

    print("\n正在更新 Word 文件...")
    result_stats = replace_text_in_docx(docx_file, mapping_dict, output_file)
    
    print("\n更新完成！")
    print(f"段落更新次數: {result_stats['paragraphs']}")
    print(f"表格更新次數: {result_stats['tables']}")
    print(f"標頭標尾更新次數: {result_stats['headers']}")
    print(f"檔案已儲存至: {output_file}")
