from nicegui import ui
from . import TabPanelEncoder, TabPanelJson5t, TabPanelVisio2

def gui_run():
    css = """
    .copyable {
            user-select: text !important;
            -webkit-user-select: text !important;
            cursor: text !important;
        }
    """

    ui.add_head_html(f"<style>{css}</style>")


    with ui.splitter(value=10).classes("w-full h-full") as splitter:
        with splitter.before:
            with ui.tabs().props("vertical").classes("w-full") as tabs:
                json5 = ui.tab("JSON5")
                visio2 = ui.tab("VISIO2")
                encoder = ui.tab("编码器")
        with splitter.after:
            with (
                ui.tab_panels(tabs, value=json5)
                .props("vertical")
                .classes("w-full h-full")
            ):
                with ui.tab_panel(json5):
                    TabPanelJson5t().create_panel()
                with ui.tab_panel(visio2):
                    TabPanelVisio2().create_panel()
                with ui.tab_panel(encoder):
                    TabPanelEncoder().create_panel()

    ui.run(native=True)
