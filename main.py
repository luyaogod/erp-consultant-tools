from openpyxl.styles import PatternFill
from src.scripts.smap import Smap, MatchHandler

class FillYellow(MatchHandler):
    """匹配标黄"""
    def on_match(self, cell, lookup_value):
        """匹配成功时的处理"""
        cell.fill = PatternFill(
            start_color="FFFF00", end_color="FFFF00", fill_type="solid"
        )  # 黄色背景

    def on_no_match(self, cell):
        """匹配失败时的处理（可按需扩展）"""
        pass

smap_instance = Smap(
    target_path="src/test/smap/麦歌恩微单据/委外工单.xlsx",
    lookup_path="src/test/smap/麦歌恩微单据/唯一表.xlsx",
    header_row=1,
    config={"字段描述": "A2:B31"},
    match_handler=FillYellow(),
)

# 执行处理
result_path = smap_instance.process()
print(f"处理完成，文件已保存至：{result_path}")