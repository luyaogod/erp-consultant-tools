import os
import re
import json
import pandas as pd
from openpyxl.utils import column_index_from_string


def excel_to_json(
    input_file_path,
    selector,
    suffix="_json",
    output_path=None,
    sheet_name="Sheet1",  # 修改默认值为"Sheet1"
    skip_rows=0,  # 新增跳过行数参数
):
    """
    将Excel文件转换为JSON列表或键值对

    参数:
        input_file_path (str): 输入Excel文件路径
        selector (str):
            - "A" 表示提取该列为JSON列表
            - "A1:B2567" 表示提取两列为JSON键值对
        suffix (str): 输出文件后缀，默认为"_json"
        output_path (str): 输出目录，默认同输入文件目录
        sheet_name (str/int): 工作表名称/索引，默认为"Sheet1"
        skip_rows (int): 跳过的行数（表头等），默认为0

    返回:
        str: 生成的JSON文件路径
    """
    # 读取Excel文件（跳过指定行数）
    df = pd.read_excel(
        input_file_path, sheet_name=sheet_name, header=None, skiprows=skip_rows
    )

    # 处理列选择器
    if ":" in selector:
        # 解析范围选择器（如"A1:B2567"）
        pattern = re.compile(r"^([A-Za-z]+)\d*:([A-Za-z]+)\d*$")
        match = pattern.match(selector)
        if not match:
            raise ValueError("Invalid range format. Expected like 'A1:B2567'")

        col1 = column_index_from_string(match.group(1)) - 1
        col2 = column_index_from_string(match.group(2)) - 1

        # 提取键值对数据
        result = {}
        for index, row in df.iterrows():
            key = row[col1]
            if pd.notna(key):
                value = row[col2] if pd.notna(row[col2]) else None
                result[str(key)] = value
    else:
        # 解析单列选择器（如"A"）
        col = column_index_from_string(selector) - 1
        result = [item for item in df[col] if pd.notna(item)]

    # 生成输出路径
    if output_path is None:
        output_path = os.path.dirname(input_file_path)

    input_name = os.path.splitext(os.path.basename(input_file_path))[0]
    output_filename = f"{input_name}{suffix}.json"
    output_file_path = os.path.join(output_path, output_filename)

    # 写入JSON文件
    with open(output_file_path, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)

    return output_file_path


# 示例用法
if __name__ == "__main__":
    # 提取Sheet1的A列（跳过1行表头）
    # excel_to_json("src/test/e2j/维护入库单_CUSTOM_001.xlsx", "A", sheet_name="Sheet1")

    # 提取名为"Products"的Sheet中A:B列
    json_path = excel_to_json(
        "src/test/e2j/维护入库单_CUSTOM_001.xlsx",
        "A:B",
        sheet_name="Sheet1",
    )
    # print(f"JSON文件已生成：{json_path}")
