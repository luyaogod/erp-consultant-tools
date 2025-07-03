from scripts import run_tool_visio2, ToolConfigVisio2, SUPPORT_FORMAT_VISIO2
from threading import Thread
from nicegui import ui


class TabPanel:
    def __init__(self):
        self.output_str = ""
        self.config = ToolConfigVisio2(visio_dir="", format="PDF")
        self.ret = ""
        self.format: SUPPORT_FORMAT_VISIO2 = ["PDF"] 
        self.progress = 0

    def create_panel(self):

        @ui.refreshable
        def create_code_wigets():
            ui.code(content=self.ret).classes("w-full")

        def update_progress(filename, idx, total_files):
            self.ret = f"正在处理: ({idx}/{total_files} {filename} )"
            self.progress = idx / total_files * 100
            create_code_wigets.refresh()

        def run_tool_w(config: ToolConfigVisio2, update_progress=None):
            run_tool_visio2(config, update_progress)
            self.ret = "处理完成!"
            create_code_wigets.refresh()

        def run_tool():
            try:
                btn.props("loading")
                thread = Thread(target=run_tool_w, args=(self.config, update_progress))
                thread.start()
            except Exception as e:
                self.ret = e.__str__()
            finally:
                btn.props(remove="loading")
                create_code_wigets.refresh()

        inp = ui.input("输入待处理的Visio文件路径")
        inp.classes("w-full")
        inp.bind_value(self.config, "visio_dir")
        inp.value = self.config.visio_dir
        with ui.row():
            s = ui.select(self.format, ).bind_value(self.config, "format")
            s.value = self.config.format
            btn = ui.button(
                "", on_click=run_tool, icon="play_arrow"
            ).props("unelevated")

        create_code_wigets()
        
