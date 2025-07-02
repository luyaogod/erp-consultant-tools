import os
import pythoncom
import win32com.client
import subprocess
WORD_APP_VISIBLE = "v3.1"
visFixedFormatPDF = 1


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


def convertor_img(
    visio_dir,
    file_list,
    update_progress=None,
    image_format="PNG",
):
    """
    将Visio文件导出为图片到Converted_Files目录下

    参数:
        visio_dir (str): Visio文件所在目录路径
        file_list (list): 要转换的Visio文件名列表
        update_progress (function, 可选): 进度回调函数，格式为func(文件名, 当前序号, 总数)
        image_format (str): 导出的图片格式，支持"PNG"、"JPG"、"GIF"等Visio支持的格式

    返回:
        list: 生成的图片文件路径列表

    流程:
    1. 初始化COM环境
    2. 启动Visio应用程序
    3. 为每个Visio文件创建单独的子目录
    4. 将每页导出为指定格式的图片
    5. 返回所有生成的图片路径

    注意:
    - 图片会保存在visio_dir/Converted_Files/原文件名/目录下
    - 图片命名为"Page_1.png"、"Page_2.png"等形式
    - 使用前确保没有Visio进程运行(可调用kill_visio_processes)
    """
    pythoncom.CoInitialize()
    generated_files = []
    try:
        visio_app = win32com.client.Dispatch("Visio.Application")
        visio_app.Visible = False

        total_files = len(file_list)
        for idx, filename in enumerate(file_list):
            if update_progress:
                update_progress(filename, idx + 1, total_files)

            # 创建文件输出目录: visio_dir/Converted_Files/原文件名/
            output_dir = os.path.join(
                visio_dir, "Converted_Files", os.path.splitext(filename)[0]
            )
            os.makedirs(output_dir, exist_ok=True)

            # 打开Visio文件
            visio_file_path = os.path.normpath(os.path.join(visio_dir, filename))
            visio_doc = visio_app.Documents.Open(visio_file_path)

            # 导出每一页
            for i, page in enumerate(visio_doc.Pages):
                page_number = i + 1
                image_name = f"Page_{page_number}.{image_format.lower()}"
                image_path = os.path.join(output_dir, image_name)

                # 导出图片 (使用完整的导出方法确保质量)
                page.Export(image_path)
                generated_files.append(image_path)

            visio_doc.Close()

        return generated_files

    except Exception as e:
        print(f"导出图片时出错: {e}")
        return []
    finally:
        try:
            visio_app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def convertor_pdf(
    visio_dir,
    file_list,
    update_progress=None,
):
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
        pdf_output_dir = os.path.join(visio_dir, "Converted_Files", "PDF")
        os.makedirs(pdf_output_dir, exist_ok=True)

        total_files = len(file_list)
        for idx, filename in enumerate(file_list):
            try:
                if update_progress:
                    update_progress(filename, idx + 1, total_files)

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

def run_visio_task(visio_dir, func, *args, **kwargs):
    """
    自动获取 file_list 并执行指定的 Visio 处理任务函数。

    参数:
        visio_dir (str): Visio 文件所在目录
        func (callable): 要执行的处理函数（如 visio_to_word_export_png）
        *args, **kwargs: 会透传给 func 的额外参数

    返回:
        func 返回值，或 None（如果无文件或出错）
    """
    file_list = FileLoader(visio_dir).get_visio_files()
    if not file_list:
        print(f"在目录 {visio_dir} 中未找到任何 Visio 文件。")
        return None

    return func(visio_dir, file_list, *args, **kwargs)


if __name__ == "__main__":
    visio_dir = r"D:\鼎捷项目\历史项目\进迭时空\进迭时空SOP-v2\制造SOP\采购管理"
    run_visio_task(visio_dir, convertor_pdf)

    # visio_dir = r"D:\Lenovo\Desktop\进迭时空SOP-v2\财务SOP\应收管理"
    # run_visio_task(visio_dir, visio_to_word_export_png, separate_files=True)

    # visio_dir = r"D:\Lenovo\Desktop\进迭时空SOP-v2\财务SOP\资产管理"
    # run_visio_task(visio_dir, visio_to_word_export_png, separate_files=True)

    # visio_dir = r"D:\Lenovo\Desktop\进迭时空SOP-v2\财务SOP\总账管理"
    # run_visio_task(visio_dir, visio_to_word_export_png, separate_files=True)

    # visio_dir = r"D:\Lenovo\Desktop\进迭时空SOP-v2\制造SOP\采购管理"
    # run_visio_task(visio_dir, visio_to_word_export_png, separate_files=True)

    # visio_dir = r"D:\Lenovo\Desktop\进迭时空SOP-v2\制造SOP\委外仓库管理"
    # run_visio_task(visio_dir, visio_to_word_export_png, separate_files=True)

    # visio_dir = r"D:\Lenovo\Desktop\进迭时空SOP-v2\制造SOP\物料BOM"
    # run_visio_task(visio_dir, visio_to_word_export_png, separate_files=True)

    # visio_dir = r"D:\Lenovo\Desktop\进迭时空SOP-v2\制造SOP\物料BOM-IC"
    # run_visio_task(visio_dir, visio_to_word_export_png, separate_files=True)
