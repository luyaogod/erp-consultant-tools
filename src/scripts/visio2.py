import os
import pythoncom
import win32com.client
import subprocess
from typing import Literal
from pydantic import BaseModel

WORD_APP_VISIBLE = "v3.0"
SUPPORT_FORMAT = Literal[
    "PDF", 
]

class Utils:
    @staticmethod
    def kill_visio_processes():
        """
        终止所有正在运行的 Visio 进程以防止文件被占用。

        使用Windows的taskkill命令强制终止所有visio.exe进程。
        如果终止失败会捕获异常并打印错误信息。
        """
        try:
            subprocess.run(["taskkill", "/F", "/IM", "visio.exe"], check=True)
            print("所有Visio进程已终止。")
        except subprocess.CalledProcessError as e:
            print(f"终止Visio进程时出错: {e}")

    @staticmethod
    def kill_word_processes(word_processor="Word"):
        """
        根据指定的应用程序类型终止 Word 或 WPS 进程。

        参数:
            word_processor (str): 指定要终止的办公软件类型，可选"Word"或"WPS"

        根据参数决定终止winword.exe(Word)或wps.exe(WPS)进程。
        使用非严格模式(taskkill的check=False)，即使进程不存在也不会报错。
        """
        targets = ["winword.exe"] if word_processor == "Word" else ["wps.exe"]
        try:
            for proc in targets:
                subprocess.run(
                    ["taskkill", "/F", "/IM", proc], check=False
                )  # 非严格模式
            print(f"已终止{word_processor}相关进程")
        except Exception as e:
            print(f"进程终止异常: {e}")


class FileLoader:
    """
    获取指定目录下所有Visio文件

        参数:
            visio_dir (str): 要扫描的目录路径
            extensions (list, 可选): 要匹配的文件扩展名列表，默认包含 .vsdx/.vsd
    """

    def __init__(self, visio_dir, extensions=None):
        self.visio_dir = visio_dir
        self.extensions = extensions or [".vsdx", ".vdx"]
        self.files = []

    def get_visio_files(self) -> list[str]:
        """
        返回:
            list: 匹配到的Visio文件名列表(相对路径)
        """

        try:
            all_files = [
                f
                for f in os.listdir(visio_dir)
                if os.path.isfile(os.path.join(visio_dir, f))
            ]

            file_list = [
                f
                for f in all_files
                if os.path.splitext(f.lower())[1] in self.extensions
            ]
            self.files = sorted(file_list)
            return self.files  # 按名称排序保证处理顺序一致

        except Exception as e:
            print(f"扫描目录失败: {e}")
            return []

class Convertor:
    def __init__(self, visio_dir:str, file_list:list[str], update_progress=None):
        self.visio_dir = visio_dir
        self.file_list = file_list
        self.update_progress = update_progress

    def converte(self, format: SUPPORT_FORMAT):
        if format == "PNG" or format == "JPEG" or format == "SVG":
            return self.convertor_img(format)
        if format == "PDF":
            return self.convertor_pdf()
        else:
            raise ValueError("不支持的格式")



    def convertor_pdf(self):
        """
        增强版Visio转PDF函数（解决授权和路径问题）

        改进点：
        1. 增加Visio许可证检查
        2. 处理中文路径兼容性
        3. 更完善的错误处理
        """
        pythoncom.CoInitialize()
        generated_files = []
        visio_app = None

        try:
            # 尝试启动Visio并检查许可证
            visio_app = win32com.client.Dispatch("Visio.Application")
            visio_app.Visible = False

            # 验证Visio是否已授权
            try:
                _ = visio_app.Version  # 触发许可证检查
            except Exception as lic_ex:
                raise RuntimeError("Visio未正确授权或试用版已过期") from lic_ex

            # 创建输出目录（兼容中文路径）
            pdf_output_dir = os.path.join(visio_dir, "PDF")
            os.makedirs(pdf_output_dir, exist_ok=True)

            total_files = len(self.file_list)
            for idx, filename in enumerate(self.file_list):
                try:
                    if self.update_progress:
                        self.update_progress(filename, idx + 1, total_files)

                    # 处理中文文件名（短路径转换）
                    safe_filename = filename.encode("gbk", errors="ignore").decode("gbk")
                    pdf_filename = os.path.splitext(safe_filename)[0] + ".pdf"
                    pdf_path = os.path.join(pdf_output_dir, pdf_filename)

                    # 使用原始API参数确保兼容性
                    visio_file_path = os.path.normpath(os.path.join(visio_dir, filename))
                    visio_doc = visio_app.Documents.Open(visio_file_path)

                    # 最简参数调用（兼容大多数版本）
                    visio_doc.ExportAsFixedFormat(
                        1,  # visFixedFormatPDF
                        pdf_path,
                        1,  # IncludeDocumentProperties
                        0,  # IgnoreDocumentStructure
                        600,  # BitmapResolution
                        1,  # OptimizeForPrint
                    )

                    generated_files.append(pdf_path)
                    visio_doc.Close()

                except Exception as file_ex:
                    print(f"[警告] 文件 {filename} 转换失败: {str(file_ex)}")
                    continue

            return generated_files

        except Exception as e:
            print(f"[严重错误] PDF导出中断: {str(e)}")
            print("请检查：1. Visio是否已激活 2. 文件路径是否合法 3. 管理员权限")
            return []
        finally:
            if visio_app is not None:
                try:
                    visio_app.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()


class ToolConfig(BaseModel):
    visio_dir:str
    format:SUPPORT_FORMAT


def run_tool(config: ToolConfig, update_progress=None):
    f = FileLoader(visio_dir)
    fiels = f.get_visio_files()
    c = Convertor(visio_dir=visio_dir, file_list=fiels, update_progress=update_progress)
    c.converte(config.format)


if __name__ == "__main__":
    visio_dir = r"D:\鼎捷项目\历史项目\进迭时空\进迭时空SOP-v2\制造SOP\采购管理"
    def update_progress(filename, idx, total_files):
        print(f"正在处理文件：({idx}/{total_files} {filename} )")

    conf = ToolConfig(
        visio_dir=visio_dir,
        format="PDF",
    )
    run_tool(conf, update_progress)


