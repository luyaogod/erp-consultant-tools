import pandas as pd
import json
import os
import re
from openpyxl.utils import column_index_from_string


def coldiff(
    excel_path, range_str, output_type="json", output_path=None, suffix="_comparison"
):
    """
    分析Excel文件中指定两列的公共部分和独有部分，并输出结果文件。

    参数：
    - excel_path: 输入Excel文件路径
    - range_str: 列范围（如'A1:B306'）
    - output_type: 输出类型，可选'json'或'excel'（默认json）
    - output_path: 输出目录路径（默认同输入文件目录）
    - suffix: 输出文件名后缀（默认'_comparison'）
    """

    def parse_range(range_str):
        """解析范围字符串为行和列的索引"""
        start_end = range_str.split(":")
        if len(start_end) != 2:
            raise ValueError("范围格式应为类似'A1:B306'")
        start, end = start_end

        # 解析列和行
        pattern = r"^([A-Za-z]+)(\d+)$"
        start_match = re.match(pattern, start)
        end_match = re.match(pattern, end)
        if not start_match or not end_match:
            raise ValueError("单元格格式错误")

        col_start = start_match.group(1)
        row_start = int(start_match.group(2))
        col_end = end_match.group(1)
        row_end = int(end_match.group(2))

        # 转换为0-based索引
        col_start_idx = column_index_from_string(col_start) - 1
        col_end_idx = column_index_from_string(col_end) - 1
        row_start_idx = row_start - 1
        row_end_idx = row_end - 1

        # 验证范围有效性
        if row_start_idx > row_end_idx or col_start_idx > col_end_idx:
            raise ValueError("无效的范围")

        return row_start_idx, row_end_idx, col_start_idx, col_end_idx

    # 解析范围
    try:
        row_start, row_end, col_start, col_end = parse_range(range_str)
    except Exception as e:
        raise ValueError(f"范围解析失败: {e}")

    # 读取Excel数据
    try:
        df = pd.read_excel(excel_path, header=None, engine="openpyxl")
    except Exception as e:
        raise IOError(f"读取Excel失败: {e}")

    # 提取指定范围数据
    selected_data = df.iloc[row_start : row_end + 1, col_start : col_end + 1]

    # 验证必须为两列
    if selected_data.shape[1] != 2:
        raise ValueError("必须选择两列数据")

    # 处理数据（去重、去空、转字符串）
    col_a = selected_data.iloc[:, 0].dropna().astype(str).unique()
    col_b = selected_data.iloc[:, 1].dropna().astype(str).unique()

    set_a = set(col_a)
    set_b = set(col_b)

    # 计算结果
    common = sorted(list(set_a & set_b))
    a_unique = sorted(list(set_a - set_b))
    b_unique = sorted(list(set_b - set_a))

    result = {"common": common, "a_unique": a_unique, "b_unique": b_unique}

    # 构建输出路径
    dir_path, filename = os.path.split(excel_path)
    base_name = os.path.splitext(filename)[0]

    output_dir = dir_path if output_path is None else output_path
    os.makedirs(output_dir, exist_ok=True)

    # 生成文件名
    ext = "json" if output_type == "json" else "xlsx"
    output_file = f"{base_name}{suffix}.{ext}"
    output_fullpath = os.path.join(output_dir, output_file)

    # 输出结果
    if output_type == "json":
        with open(output_fullpath, "w", encoding="utf-8") as f:
            json.dump(result, f, indent=2, ensure_ascii=False)
    else:
        # 创建三列等长的DataFrame
        max_len = max(len(common), len(a_unique), len(b_unique))
        df_output = pd.DataFrame(
            {
                "Common": common + [None] * (max_len - len(common)),
                "A_Unique": a_unique + [None] * (max_len - len(a_unique)),
                "B_Unique": b_unique + [None] * (max_len - len(b_unique)),
            }
        )
        df_output.to_excel(output_fullpath, index=False, engine="openpyxl")

    print(f"结果已输出至：{output_fullpath}")

if __name__ == "__main__":
    # 示例用法
    coldiff('src/test/coldiff/锡圆群组对比.xlsx', 'A2:B17', output_type='excel', suffix='_diff')
