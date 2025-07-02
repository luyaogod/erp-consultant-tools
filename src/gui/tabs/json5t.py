from nicegui import ui
from scripts import run_tool_json5t, ToolConfigJson5t


class TabPanel:
    def __init__(self):
        self.output_str = ""
        self.config = ToolConfigJson5t(json5_str="", indent=2)

    def create_panel(self):
        @ui.refreshable
        def create_ret_wigets():
            ui.codemirror(self.output_str, language="JSON", theme="vscodeLight").classes("w-full h-full")

        def change_handler():
            try:
                self.config.json5_str = self.editor.value  # 直接使用editor的值
                self.output_str = run_tool_json5t(self.config)
            except Exception as e:
                self.output_str = e.__str__()
            finally:
                create_ret_wigets.refresh()
                

        # 配置网格布局：两列，
        with ui.grid(columns=2).classes("w-full gap-4"):
            # 左侧表单容器
            with ui.column().classes("space-y-4 "):
                # 使用monospace字体保持格式
                self.editor = ui.codemirror(
                    value=self.config.json5_str,
                    on_change=change_handler,
                    language="JSON",
                    theme="vscodeLight",
                ).classes("w-full h-full")
                # 添加一些样式改进

            # 右侧结果区域
            with ui.column().classes("copyable h-full"):
                create_ret_wigets()
