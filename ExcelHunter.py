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
    
    found_files = 0  # 记录找到结果的文件数
    
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
                                found_files = 0  # 重置标记，避免重复打印分隔行
                            
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
                found_files += 1
    
        except Exception as e:
            print(f"处理文件 '{excel_file}' 时出错: {str(e)}")

    if not found:
        print(f"\n在所有Excel文件中都没有找到 '{search_text}'")

if __name__ == "__main__":
    search_text = input("请输入要搜索的文字: ")
    search_excel_files(search_text)
