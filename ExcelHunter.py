import pandas as pd
from pathlib import Path
import warnings

# 忽略 openpyxl 的警告信息
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

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
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                
                # 在数据框中搜索文本
                for row_idx, row in df.iterrows():
                    for col_idx, value in enumerate(row):
                        if isinstance(value, str) and search_text in value:
                            found = True
                            print(f"\n在文件 '{excel_file}' 中找到匹配：")
                            print(f"工作表: {sheet_name}")
                            print(f"位置: 第{row_idx + 1}行, 第{col_idx + 1}列")
                            print(f"内容: {value}")
                            
        except Exception as e:
            print(f"处理文件 '{excel_file}' 时出错: {str(e)}")
    
    if not found:
        print(f"\n在所有Excel文件中都没有找到 '{search_text}'")

if __name__ == "__main__":
    search_text = input("请输入要搜索的文字: ")
    search_excel_files(search_text)
