import pandas as pd
import os


def colmerge(
    file_paths: list,
    columns_to_merge: int,
    skip_rows: int = 0,
    merge_mode: str = "tail",
    output_path: str = None,
) -> None:
    """
    合并多个Excel文件中所有Sheet的指定列，并添加来源信息

    :param file_paths: 要处理的Excel文件路径列表
    :param columns_to_merge: 需要合并的前N列
    :param skip_rows: 每个Sheet跳过的行数（默认0）
    :param merge_mode: 合并模式，'head'头插或'tail'尾插（默认'head'）
    :param output_path: 输出文件路径（默认同首文件目录/merged.xlsx）
    """
    # 参数验证
    if not file_paths:
        raise ValueError("文件路径列表不能为空")
    if merge_mode not in ("head", "tail"):
        raise ValueError("合并模式必须为'head'或'tail'")
    if columns_to_merge <= 0:
        raise ValueError("合并列数必须大于0")
    if skip_rows < 0:
        raise ValueError("跳过的行数不能为负数")

    # 设置默认输出路径
    if output_path is None:
        first_file_dir = os.path.dirname(os.path.abspath(file_paths[0]))
        output_path = os.path.join(first_file_dir, "merged.xlsx")
    else:
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)

    all_data = []

    try:
        for file_idx, file_path in enumerate(file_paths):
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"文件不存在: {file_path}")

            with pd.ExcelFile(file_path) as excel:
                for sheet_idx, sheet_name in enumerate(excel.sheet_names):
                    # 读取数据并跳过指定行数
                    df = pd.read_excel(
                        excel,
                        sheet_name=sheet_name,
                        skiprows=skip_rows,
                        header=None if skip_rows else 0,
                    )

                    # 列数验证
                    if df.shape[1] < columns_to_merge:
                        raise ValueError(
                            f"文件: {os.path.basename(file_path)}\n"
                            f"Sheet: {sheet_name}\n"
                            f"有效列数不足: 需要{columns_to_merge}列，实际{df.shape[1]}列"
                        )

                    # 提取指定列并添加来源信息
                    selected_cols = df.iloc[:, :columns_to_merge].copy()
                    selected_cols["来源Sheet"] = sheet_name
                    selected_cols["来源文件"] = os.path.basename(file_path)

                    all_data.append(selected_cols)

        # 处理合并顺序
        if merge_mode == "head":
            all_data.reverse()

        # 合并所有数据
        merged_df = pd.concat(all_data, ignore_index=True)

        # 设置列名（处理跳过表头的情况）
        if skip_rows > 0:
            original_columns = [f"列{i + 1}" for i in range(columns_to_merge)]
            merged_df.columns = original_columns + ["来源Sheet", "来源文件"]

        # 保存结果
        merged_df.to_excel(output_path, index=False, engine="openpyxl")

    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")
        raise


# 使用示例
if __name__ == "__main__":
    colmerge(
        file_paths=[r"src/test/colmerge/对照-FT苏州1-middle.xlsx", r"src/test/colmerge/对照-FT苏州2-middle.xlsx"],
        columns_to_merge=2,
        skip_rows=1,
        merge_mode="tail",
        output_path=r"src/test/colmerge/merged_data.xlsx",
    )
