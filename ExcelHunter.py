import pandas as pd
from pathlib import Path
import warnings

# 忽略 openpyxl 的警告信息
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def get_column_letter(col_idx):
    """将数字列索引转换为Excel列字母"""
    result = ""
    while col_idx >= 0:
        col_idx, remainder = divmod(col_idx, 26)
        result = chr(65 + remainder) + result
        col_idx -= 1
    return result

def search_excel_files(search_text):
    """在Sheets目录下的所有Excel文件中搜索指定文本"""
    # 获取Sheets目录及其所有子目录下的Excel文件
    sheets_dir = Path('./Sheets')
    excel_files = list(sheets_dir.rglob('*.xlsx')) + list(sheets_dir.rglob('*.xls'))
    
    if not excel_files:
        print("Sheets目录下没有找到Excel文件")
        return
    
    found = False
    found_files = 0
    
    for excel_file in excel_files:
        try:
            excel = pd.ExcelFile(excel_file)
            file_match_count = 0
            
            for sheet_name in excel.sheet_names:
                df = pd.read_excel(excel_file, 
                                 sheet_name=sheet_name, 
                                 nrows=None)
                
                for row_idx, row in df.iterrows():
                    for col_idx, value in enumerate(row):
                        if isinstance(value, str) and search_text in value:
                            # 如果不是第一个有结果的文件，添加分隔行
                            if found_files > 0:
                                print("\n" + "="*50 + "\n")
                                found_files = 0
                            
                            found = True
                            file_match_count += 1
                            excel_row = row_idx + 2
                            excel_col = get_column_letter(col_idx)
                            # 使用 relative_to 去除 Sheets 目录前缀
                            relative_path = excel_file.relative_to(sheets_dir)
                            print(f"\n在文件 '{relative_path}' 中找到匹配：")
                            print(f"工作表: {sheet_name}")
                            print(f"位置: {excel_col}{excel_row}")
                            print(f"内容: {value}")
            
            if file_match_count > 0:
                relative_path = excel_file.relative_to(sheets_dir)
                print(f"\n在文件 '{relative_path}' 中共找到 {file_match_count} 个匹配项")
                found_files += 1
    
        except Exception as e:
            print(f"处理文件 '{excel_file}' 时出错: {str(e)}")

    if not found:
        print(f"\n在所有Excel文件中都没有找到 '{search_text}'")

if __name__ == "__main__":
    search_text = input("请输入要搜索的文字: ")
    search_excel_files(search_text)
    input("按回车键退出...")
