import pandas as pd

def modify_salaries(file_path):
    # 讀取 Excel 文件
    df = pd.read_excel(file_path)

    # 查找 "薪水" 列並將其數值乘以五
    if '薪水' in df.columns:
        df['薪水'] = df['薪水'] * 5

    # 保存修改後的數據到新文件
    new_file_path = 'modified_' + file_path
    df.to_excel(new_file_path, index=False)

    print(f'修改後的文件已保存為: {new_file_path}')

# 修改薪水
modify_salaries('demo.xlsx')
