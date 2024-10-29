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
    
    for excel_file in excel_files:
        try:
            # 读取Excel文件中的所有工作表
            excel = pd.ExcelFile(excel_file)
            
            for sheet_name in excel.sheet_names:
                # 设置 nrows=None 确保读取所有行
                df = pd.read_excel(excel_file, 
                                 sheet_name=sheet_name, 
                                 nrows=None)  # 添加此参数
                
                # 在数据框中搜索文本
                match_count = 0  # 添加计数器用于调试
                for row_idx, row in df.iterrows():
                    for col_idx, value in enumerate(row):
                        if isinstance(value, str) and search_text in value:
                            found = True
                            match_count += 1  # 计数匹配项
                            excel_row = row_idx + 2
                            excel_col = get_column_letter(col_idx)
                            print(f"\n在文件 '{excel_file}' 中找到匹配：")
                            print(f"工作表: {sheet_name}")
                            print(f"位置: {excel_col}{excel_row}")
                            print(f"内容: {value}")
                
                print(f"\n在工作表 {sheet_name} 中共找到 {match_count} 个匹配项")  # 显示匹配总数
    
        except Exception as e:
            print(f"处理文件 '{excel_file}' 时出错: {str(e)}")

    if not found:
        print(f"\n在所有Excel文件中都没有找到 '{search_text}'")

if __name__ == "__main__":
    search_text = input("请输入要搜索的文字: ")
    search_excel_files(search_text)
