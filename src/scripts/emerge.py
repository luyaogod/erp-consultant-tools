import os
import pandas as pd


def merge_excel_sheets(input_folder, output_suffix=""):
    # 获取文件夹名称
    folder_name = os.path.basename(input_folder)

    # 设置输出文件路径和名称
    output_file = os.path.join(
        os.path.dirname(input_folder), f"{folder_name}{output_suffix}.xlsx"
    )

    # 创建一个新的Excel writer对象
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        # 遍历文件夹中的所有Excel文件
        for file in os.listdir(input_folder):
            if file.endswith((".xlsx", ".xls")):
                file_path = os.path.join(input_folder, file)
                file_name = os.path.splitext(file)[0]  # 获取文件名（不带扩展名）

                # 读取Excel文件中的所有工作表
                excel_file = pd.ExcelFile(file_path)

                # 遍历每个工作表
                for sheet_name in excel_file.sheet_names:
                    # 读取工作表数据
                    sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)

                    # 新工作表名称格式：原文件名_原工作表名（避免重复）
                    new_sheet_name = f"{file_name}_{sheet_name}"[
                        :31
                    ]  # Excel工作表名最多31个字符

                    # 将数据写入新Excel文件中的新工作表
                    sheet_data.to_excel(writer, sheet_name=new_sheet_name, index=False)

    print(f"所有Excel文件的工作表已合并到: {output_file}")


if __name__ == "__main__":

    # 调用合并函数
    merge_excel_sheets(r"src/test/emerge/采购订单")
    merge_excel_sheets(r"src/test/emerge/委外工单")
    merge_excel_sheets(r"src/test/emerge/应付单据")
    merge_excel_sheets(r"src/test/emerge/应收单据")
