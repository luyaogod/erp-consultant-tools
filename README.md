# SetUP
```
pip install -r requirements.txt
```

# smap
该模块实现了 Excel 文件之间的 VLOOKUP 功能，通过查找另一个 Excel 文件中的数据范围，替换目标 Excel 文件中指定列的数据。

支持以下功能：

从 JSON 配置文件或字典加载字段映射关系。
多 Sheet 处理。
自定义匹配与未匹配处理策略（如保留原值、设置为空）。
跳过表头行和指定数量的数据行。
保存处理后的文件为新名称。

```python
    # 使用自定义处理程序示例
    class CustomHandler(MatchHandler):
        def on_match(self, cell, lookup_value):
            cell.value = f"[MATCHED] {lookup_value}"

        def on_no_match(self, cell):
            cell.value = "[UNMATCHED]"

    processor = SmapProcessor(
        target_file_path="16smap/维护BOM信息all250427.xlsx",
        lookup_file_path="16smap/昆山1对照表_no_duplicates.xlsx",
        config_path="16smap/config.json",
        header_row=2,
        skip_rows=0,
        sheet_names=["BOM_D"],
        suffix="_processed",
        match_handler=CustomHandler(),
    )

    result_path = processor.process()
    print(f"处理完成，文件已保存至：{result_path}")
```

# coldiff
该脚本用于分析 Excel 文件中指定两列数据之间的公共部分与独有部分，并支持将结果输出为 JSON 或 Excel 文件。

# colmerge
该脚本用于合并多个 Excel 文件中所有 Sheet 的指定列，并在合并结果中添加来源信息（包括文件名和 Sheet 名）。

# e2j
该脚本用于将 Excel 文件中的数据转换为 JSON 格式，支持以下两种输出模式：

- 单列列表：提取一列生成 JSON 数组；
- 键值对对象：提取两列生成 JSON 对象（第一列为 key，第二列为 value）。

# tcopy
该脚本用于复制 Excel 文件中所有 Sheet 的结构（仅保留 Sheet 名称），生成一个空模板文件，常用于创建标准化模板或清空数据后的结构复用。











