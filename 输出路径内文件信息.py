import os
import pandas as pd
import time
from pathlib import Path
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

def get_file_info(path, root_path=''):
    file_info = []

    for root, dirs, files in os.walk(path):
        for file in files:
            file_path = os.path.join(root, file)
            size = os.path.getsize(file_path) / 1024 / 1024  # 将文件大小转换为MB单位
            _, ext = os.path.splitext(file)
            creation_time = os.path.getctime(file_path)
            modification_time = os.path.getmtime(file_path)

            file_info.append({
                '文件名': file,
                '文件大小(MB)': size,
                '文件类型': ext,
                '文件夹路径': os.path.relpath(root, root_path),  # 记录子文件夹路径
                '创建时间': time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(creation_time)),
                '最后修改时间': time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(modification_time)),
                '最后修改人': get_user_name()  # 获取最后修改人
            })

    return file_info

def save_to_excel(file_info, output_path):
    df = pd.DataFrame(file_info)
    writer = pd.ExcelWriter(output_path, engine='openpyxl')
    df.to_excel(writer, index=False)
    wb = writer.book
    ws = wb.active

    # 设置字体
    for row in ws.iter_rows():
        for cell in row:
            cell.font = Font(name='微软雅黑')

    # 自动调整列宽
    for column_cells in ws.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        ws.column_dimensions[get_column_letter(column_cells[0].column)].width = length

    # 保存文件
    writer.close()

def get_user_name():
    # 获取最后修改人的用户名
    import getpass
    return getpass.getuser()

if __name__ == '__main__':
    # 获取用户输入的路径
    source_path = input("请输入文件夹路径：")

    start_time = time.time()  # 记录开始时间

    # 获取文件夹名称和当前日期时间
    folder_name = os.path.basename(source_path)
    date_time = time.strftime('%Y%m%d-%H%M%S', time.localtime())

    # 构造输出路径
    output_folder = '请将此文字替换为你的实际csv导出路径'
    output_path = os.path.join(output_folder, f'{folder_name}_{date_time}.xlsx')

    # 获取文件信息并保存到 Excel 文件
    file_info = get_file_info(source_path, source_path)
    save_to_excel(file_info, output_path)

    end_time = time.time()  # 记录结束时间
    duration = end_time - start_time  # 计算耗时
    size = os.path.getsize(output_path) / 1024 / 1024  # 计算文件大小，单位为MB

    print(f"文件信息已保存到：{output_path}")
    print(f"本次运行耗时：{duration:.2f}秒")
    print(f"输出文件大小：{size:.2f}MB")
