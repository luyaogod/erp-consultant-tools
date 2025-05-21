import os
import pandas as pd


def csv_to_excel(input_folder, suffix="_excel", encoding="gbk"):
    """
    将指定文件夹内所有CSV文件转换为Excel文件

    参数:
        input_folder (str): 包含CSV文件的文件夹路径
        suffix (str): 输出文件夹的后缀，默认为'_excel'
        encoding (str): CSV文件的编码格式，默认为'gb2312'
    """
    # 创建输出文件夹路径
    output_folder = os.path.join(
        os.path.dirname(input_folder), os.path.basename(input_folder) + suffix
    )

    # 如果输出文件夹不存在则创建
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 遍历输入文件夹中的所有文件
    for filename in os.listdir(input_folder):
        if filename.endswith(".csv"):
            # 构造完整的输入和输出文件路径
            csv_path = os.path.join(input_folder, filename)
            excel_filename = os.path.splitext(filename)[0] + ".xlsx"
            excel_path = os.path.join(output_folder, excel_filename)

            # 读取CSV并写入Excel
            try:
                # 显式指定GB2312编码
                df = pd.read_csv(csv_path, encoding=encoding)
                df.to_excel(excel_path, index=False)
                print(f"转换成功: {filename} -> {excel_filename}")
            except Exception as e:
                print(f"转换失败 {filename}: {str(e)}")


if __name__ == "__main__":
    input_folder = r"src/test/csv2xlsx/应付单据"
    csv_to_excel(input_folder, "_xlsx")
    print("所有文件转换完成!")

    input_folder = r"src/test/csv2xlsx/应收单据"
    csv_to_excel(input_folder, "_xlsx")
    print("所有文件转换完成!")
