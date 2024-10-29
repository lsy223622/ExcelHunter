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
    # 获取当前目录下所有的Excel文件
    excel_files = list(Path('.').glob('*.xlsx')) + list(Path('.').glob('*.xls'))
    
    if not excel_files:
        print("当前目录下没有找到Excel文件")
        return
    
    found = False
    
    for i, excel_file in enumerate(excel_files):
        try:
            # 如果不是第一个文件，在文件结果之前添加分隔行
            if i > 0:
                print("\n" + "="*50 + "\n")
                
            excel = pd.ExcelFile(excel_file)
            file_match_count = 0
            
            for sheet_name in excel.sheet_names:
                df = pd.read_excel(excel_file, 
                                 sheet_name=sheet_name, 
                                 nrows=None)
                
                for row_idx, row in df.iterrows():
                    for col_idx, value in enumerate(row):
                        if isinstance(value, str) and search_text in value:
                            found = True
                            file_match_count += 1
                            excel_row = row_idx + 2
                            excel_col = get_column_letter(col_idx)
                            print(f"\n在文件 '{excel_file}' 中找到匹配：")
                            print(f"工作表: {sheet_name}")
                            print(f"位置: {excel_col}{excel_row}")
                            print(f"内容: {value}")
            
            if file_match_count > 0:
                print(f"\n在文件 '{excel_file}' 中共找到 {file_match_count} 个匹配项")
    
        except Exception as e:
            print(f"处理文件 '{excel_file}' 时出错: {str(e)}")

    if not found:
        print(f"\n在所有Excel文件中都没有找到 '{search_text}'")

if __name__ == "__main__":
    search_text = input("请输入要搜索的文字: ")
    search_excel_files(search_text)
