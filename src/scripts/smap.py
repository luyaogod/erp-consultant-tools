import openpyxl
from openpyxl.utils import column_index_from_string
import json
import os


class MatchHandler:
    """基类：定义匹配处理接口"""

    def on_match(self, cell, lookup_value):
        """匹配成功时的处理"""
        cell.value = lookup_value

    def on_no_match(self, cell):
        """匹配失败时的处理"""
        pass


class EmptyOrKeep(MatchHandler):
    """未找到时保留原值（默认处理）"""

    def __init__(self, on_no_match_action="empty"):
        self.not_found_action = on_no_match_action

    def on_no_match(self, cell):
        # if keep do nothing
        if self.not_found_action == "empty":
            cell.value = None


class Smap:
    """Excel跨文件VLOOKUP处理器（面向对象重构版）

    核心功能：
    1. 根据配置文件加载查找表数据
    2. 在目标文件中执行列替换
    3. 支持自定义匹配/未匹配处理策略

    设计特点：
    - 采用策略模式分离匹配处理逻辑
    - 支持多Sheet处理
    - 自动跳过指定行数
    - 可扩展的保存命名规则
    """
    def __init__(
        self,
        target_path: str,
        lookup_path: str,
        header_row: int,
        skip_rows: int = 0,
        config_path: str = "",
        config: dict = None,
        sheet_names: list = None,
        suffix: str = "_processed",
        match_handler: MatchHandler = None,
    ):
        """
        初始化处理器

        :param target_file_path: 待处理Excel文件路径
        :param lookup_file_path: 查找表Excel文件路径
        :param config_path: JSON配置文件路径（定义字段和查找范围）
        :param config: 定义字段和查找范围（优先级高于JSON配置）
        :param header_row: 列头所在行号（1-based）
        :param skip_rows: 跳过行数（从列头行之后开始计算）
        :param sheet_names: 指定处理的Sheet名称列表，None表示处理所有Sheet
        :param suffix: 输出文件后缀（默认添加'_processed'）
        :param match_handler: 自定义匹配处理器实例，None则使用默认处理器
        """
        if (not config_path) and (not config):
            raise ValueError("请提供配置文件路径或配置字典")

        # 初始化参数
        self.target_path = target_path
        self.lookup_path = lookup_path
        self.config_path = config_path
        self.config = config or None
        self.header_row = header_row
        self.skip_rows = skip_rows
        self.sheet_names = sheet_names if sheet_names else []
        self.suffix = suffix

        # 初始化处理程序
        self.match_handler = match_handler or EmptyOrKeep()
       
        # 运行时数据
        self.lookup_data = {}
        self.target_wb = None

    def process(self):
        """执行完整处理流程"""
        self._load_config()
        self._load_lookup_data()
        self._process_target_file()
        return self._save_processed_file()

    def _load_config(self):
        """加载配置文件"""
        if self.config:
            pass
        else:
            with open(self.config_path, "r", encoding="utf-8") as f:
                self.config = json.load(f)

    def _load_lookup_data(self):
        """加载查找表数据"""
        lookup_wb = openpyxl.load_workbook(
            self.lookup_path, read_only=True, data_only=True
        )

        for field, range_str in self.config.items():
            # 解析范围字符串
            start_end = range_str.split(":")
            start_col = column_index_from_string(start_end[0][0])
            start_row = int(start_end[0][1:])
            end_col = column_index_from_string(start_end[1][0])
            end_row = int(start_end[1][1:])

            # 读取数据并保持首次出现的键值
            lookup_sheet = lookup_wb.active
            field_dict = {}
            for row in lookup_sheet.iter_rows(
                min_row=start_row,
                max_row=end_row,
                min_col=start_col,
                max_col=end_col,
                values_only=True,
            ):
                key, value = row[0], row[1]
                if key not in field_dict:
                    field_dict[key] = value

            self.lookup_data[field] = field_dict

        lookup_wb.close()

    def _process_target_file(self):
        """处理目标文件"""
        self.target_wb = openpyxl.load_workbook(self.target_path)
        sheets = self.sheet_names if self.sheet_names else self.target_wb.sheetnames

        for sheet_name in sheets:
            if sheet_name not in self.target_wb.sheetnames:
                continue

            sheet = self.target_wb[sheet_name]
            for field in self.config.keys():
                target_col = self._find_target_column(sheet, field)
                if target_col is None:
                    continue

                self._process_column(sheet, field, target_col)

    def _find_target_column(self, sheet, field):
        """定位目标列"""
        for cell in sheet[self.header_row]:
            if cell.value == field:
                return cell.column
        return None

    def _process_column(self, sheet, field, target_col):
        """处理单个列"""
        start_row = self.header_row + self.skip_rows + 1
        lookup_map = self.lookup_data.get(field, {})

        for row in sheet.iter_rows(
            min_row=start_row, min_col=target_col, max_col=target_col
        ):
            cell = row[0]
            lookup_value = lookup_map.get(cell.value)

            if lookup_value is not None:
                self.match_handler.on_match(cell, lookup_value)
            else:
                self.match_handler.on_no_match(cell)

    def _save_processed_file(self):
        """保存处理结果"""
        original_dir = os.path.dirname(self.target_path)
        original_name = os.path.basename(self.target_path)
        base_name, ext = os.path.splitext(original_name)
        new_path = os.path.join(original_dir, f"{base_name}{self.suffix}{ext}")

        self.target_wb.save(new_path)
        self.target_wb.close()
        return new_path


def smap(
    target_path: str,
    lookup_path:  str,
    config_path: str = "",
    config: dict = None,
    header_row: int = 2,
    skip_rows: int = 0,
    sheet_names: list = None,
    suffix: str = "_processed",
    not_math: str = "empty",
):
    smap = Smap(
        target_path=target_path,
        lookup_path=lookup_path,
        config_path=config_path,
        config=config,
        header_row=header_row,
        skip_rows=skip_rows,
        sheet_names=sheet_names,
        suffix=suffix,
        match_handler=EmptyOrKeep(not_math),
    )
    return smap.process()


# 使用示例
if __name__ == "__main__":
    result_path = smap(
        target_path="src/test/smap/维护BOM信息all250427.xlsx",
        lookup_path="src/test/smap/昆山1对照表_no_duplicates.xlsx",
        config={"元件品号": "A2:B2853", "主芯": "E2:F2"},
    )

    print(f"处理完成，文件已保存至：{result_path}")
