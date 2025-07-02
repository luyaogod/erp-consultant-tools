from nicegui import ui
from scripts import run_tool_encoder, ToolConfigEncoder

class TabPanel:
    def __init__(self):
        self.content = ""
        self.config = ToolConfigEncoder(
        excel_file=r"src/test/encoder/测试.xlsx",
        range="A1:B159",
        function_indices=[0, 1],
        separators=["",""],
        database_path="src/test/encoder/processing_records.db",
        sheet_name="Sheet1",
        begin_serial=None,
        num_zill=3,
    )

    def create_panel(self):
        @ui.refreshable
        def create_code_wigets():
            ui.code(self.content).classes("w-full h-full")

        def run_tool_encoder_wrapper(self):
            try:
                config = self.config
                btn.props("loading")
                # 尝试转换输入类型
                if config.function_indices:
                    if isinstance(config.function_indices, str):
                        config.function_indices = [
                            int(i) for i in config.function_indices.split(",")
                        ]
                if config.separators:
                    if isinstance(config.separators, str):
                        config.separators = [
                            s.strip() for s in config.separators.split(",")
                        ]
                # 执行工具函数
                # content = config.model_dump_json(indent=4)
                self.content = str(run_tool_encoder(config))
                btn.props(remove="loading")
            except Exception as e:
                self.content = e.__str__()
                btn.props(remove="loading")
            finally:
                create_code_wigets.refresh()

        # 配置网格布局：两列，
        with ui.grid(columns=2).classes("w-full gap-4"):
            # 左侧表单容器
            with ui.column().classes("space-y-4"):
                config = self.config
                btn = ui.button("", icon="play_arrow", on_click=run_tool_encoder_wrapper).props(
                    "unelevated"
                )

                ui.input(label="Excel文件路径", placeholder="输入.xlsx文件路径").classes(
                    "w-full"
                ).bind_value_to(config, "excel_file").props(
                    "clearable"
                ).value = config.excel_file

                ui.input(label="单元格范围", placeholder="例如: A1:B10").classes(
                    "w-full"
                ).bind_value_to(config, "range").value = config.range

                ui.input(label="函数索引", placeholder="例如: 0,1,2").classes(
                    "w-full"
                ).bind_value_to(config, "function_indices").tooltip(
                    "0不变;1格式日期为无分隔符"
                ).value = ",".join(map(str, config.function_indices))

                ui.input(label="分隔符", placeholder="例如: -,-,-").classes(
                    "w-full"
                ).bind_value_to(config, "separators").value = ",".join(config.separators)

                ui.input(label="工作表名称", placeholder="默认为'生成单号'").classes(
                    "w-full"
                ).bind_value_to(config, "sheet_name").value = config.sheet_name

                ui.input(label="数据库路径", placeholder="输入数据库文件路径").classes(
                    "w-full"
                ).props("clearable").bind_value_to(
                    config, "database_path"
                ).value = config.database_path

                nz = ui.number(
                    label="补位数", placeholder="3表示补到三位如001", min=1
                ).classes("w-full")
                nz.value = config.num_zill
                nz.bind_value_to(config, "num_zill", lambda x: int(x))

                nb = ui.number(
                    label="起始序号", placeholder="大于等于0的整数", min=1
                ).classes("w-full")
                nb.value = config.begin_serial
                nb.bind_value_to(
                    config, "begin_serial", lambda x: int(x) if x is not None else None
                )

            # 右侧结果区域
            with ui.column().classes("copyable h-full"):
                create_code_wigets()