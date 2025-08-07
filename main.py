import os
import sys
import threading
import tempfile
import shutil
import subprocess
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ---------------------------
# 辅助函数：获取资源路径（兼容 PyInstaller 打包）
# ---------------------------
def resource_path(filename: str) -> str:
    """
    返回资源文件的绝对路径。
    当程序是用 PyInstaller 打包时，--add-data 添加的文件会存放在 sys._MEIPASS 下。
    """
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)

# ---------------------------
# 主界面类
# ---------------------------
class DualFileCompressor:
    def __init__(self, root):
        self.root = root
        self.root.title("双文件压缩器")
        self.center_window(550, 480)
        self.setup_ui()

    # 窗口居中
    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    # 初始化界面
    def setup_ui(self):
        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat")
        style.configure("TLabel", font=("Arial", 10))
        style.configure("TEntry", padding=5)

        # 文件A选择
        ttk.Label(self.root, text="文件/文件夹A:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.file1_entry = ttk.Entry(self.root, width=40)
        self.file1_entry.grid(row=0, column=1, padx=10, pady=10)
        ttk.Button(self.root, text="浏览", command=lambda: self.browse_file(self.file1_entry)).grid(row=0, column=2)

        # 密码A
        ttk.Label(self.root, text="密码A:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        self.password1_entry = ttk.Entry(self.root, width=40, show="*")
        self.password1_entry.grid(row=1, column=1, padx=10, pady=10)

        # 文件B选择
        ttk.Label(self.root, text="文件/文件夹B:").grid(row=2, column=0, padx=10, pady=10, sticky="e")
        self.file2_entry = ttk.Entry(self.root, width=40)
        self.file2_entry.grid(row=2, column=1, padx=10, pady=10)
        ttk.Button(self.root, text="浏览", command=lambda: self.browse_file(self.file2_entry)).grid(row=2, column=2)

        # 密码B
        ttk.Label(self.root, text="密码B:").grid(row=3, column=0, padx=10, pady=10, sticky="e")
        self.password2_entry = ttk.Entry(self.root, width=40, show="*")
        self.password2_entry.grid(row=3, column=1, padx=10, pady=10)

        # 输出路径
        ttk.Label(self.root, text="输出EXE路径:").grid(row=4, column=0, padx=10, pady=10, sticky="e")
        self.output_entry = ttk.Entry(self.root, width=40)
        self.output_entry.grid(row=4, column=1, padx=10, pady=10)
        self.output_entry.insert(0, os.path.join(os.getcwd(), "output.exe"))
        ttk.Button(self.root, text="浏览", command=self.save_output).grid(row=4, column=2)

        # 生成按钮
        ttk.Button(self.root, text="生成自解压EXE", command=self.compress).grid(row=5, column=1, pady=10)

        # 状态显示
        self.status_var = tk.StringVar()
        self.status_label = ttk.Label(self.root, textvariable=self.status_var, foreground="blue")
        self.status_label.grid(row=6, column=0, columnspan=3, pady=10)

        # 进度条
        self.progress = ttk.Progressbar(self.root, length=400, mode="determinate")
        self.progress.grid(row=7, column=0, columnspan=3, padx=10, pady=10)

    # ---------------------------
    # 辅助UI函数
    # ---------------------------
    def browse_file(self, entry):
        choice = messagebox.askquestion("选择类型", "您要选择文件吗？点击“否”选择文件夹", icon="question", type="yesno")
        path = filedialog.askopenfilename() if choice == "yes" else filedialog.askdirectory()
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)

    def save_output(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".exe",
            filetypes=[("EXE files", "*.exe")],
            initialdir=os.getcwd(),
            initialfile="output.exe"
        )
        if path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, path)

    # ---------------------------
    # 压缩入口：开启新线程以防止界面卡死
    # ---------------------------
    def compress(self):
        t = threading.Thread(target=self._compress_worker, daemon=True)
        t.start()

    # ---------------------------
    # 查找 7z：优先用系统的，否则使用程序目录下的（支持打包）
    # ---------------------------
    def find_7z_cmd(self):
        exe_name = "7z.exe" if os.name == "nt" else "7z"
        from shutil import which
        if which(exe_name):
            return which(exe_name)
        path = resource_path("7z.exe")
        if os.path.exists(path):
            return path
        return None

    # ---------------------------
    # 压缩与打包的工作线程
    # ---------------------------
    def _compress_worker(self):
        try:
            file1 = self.file1_entry.get().strip()
            file2 = self.file2_entry.get().strip()
            pw1 = self.password1_entry.get()
            pw2 = self.password2_entry.get()
            output_exe = self.output_entry.get().strip()

            # 参数检查
            if not all([file1, file2, pw1, pw2, output_exe]):
                self.set_status("错误: 请填写所有字段", "red")
                return
            if not os.path.exists(file1) or not os.path.exists(file2):
                self.set_status("错误: 文件/文件夹不存在", "red")
                return

            # 查找 7z
            seven_zip = self.find_7z_cmd()
            if seven_zip is None:
                self.set_status("错误: 未找到 7z，可将 7z.exe 放到脚本目录或安装 7-Zip 并添加到 PATH。", "red")
                return

            # 创建临时工作目录
            work_tmp = tempfile.mkdtemp(prefix="dfc_pack_")
            pyinstaller_work = os.path.join(work_tmp, "pyinstaller_work")
            os.makedirs(pyinstaller_work, exist_ok=True)

            # 临时文件路径
            zip1 = os.path.join(work_tmp, "data1.7z")
            zip2 = os.path.join(work_tmp, "data2.7z")
            extractor_py = os.path.join(work_tmp, "extractor.py")

            # 阶段1：压缩文件A
            self.set_status("压缩文件A中...", "blue")
            self.set_progress(5)
            self.run_7z_compress(seven_zip, file1, zip1, pw1)
            self.set_progress(30)

            # 阶段2：压缩文件B
            self.set_status("压缩文件B中...", "blue")
            self.run_7z_compress(seven_zip, file2, zip2, pw2)
            self.set_progress(55)

            # 阶段3：生成解压脚本
            self.set_status("生成解压脚本...", "blue")
            with open(extractor_py, "w", encoding="utf-8") as f:
                f.write(self._make_extractor_script(pw1, pw2))
            self.set_progress(70)

            # 阶段4：用 PyInstaller 打包 EXE
            self.set_status("正在用 PyInstaller 打包 EXE（可能需要几分钟）...", "blue")
            sep = ";" if os.name == "nt" else ":"
            add_data_list = [
                f"{zip1}{sep}.",
                f"{zip2}{sep}.",
                f"{resource_path('7z.exe')}{sep}."
            ]

            pyinstaller_cmd = [
                sys.executable, "-m", "PyInstaller",
                "--noconfirm",
                "--onefile",
                "--windowed",
                "--name", os.path.splitext(os.path.basename(output_exe))[0],
                "--distpath", os.path.dirname(os.path.abspath(output_exe)),
                "--workpath", pyinstaller_work,
                "--specpath", pyinstaller_work,
            ]
            for ad in add_data_list:
                pyinstaller_cmd += ["--add-data", ad]
            pyinstaller_cmd.append(extractor_py)

            proc = subprocess.run(pyinstaller_cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            if proc.returncode != 0:
                self.set_status("打包失败，请查看控制台输出。", "red")
                print("PyInstaller 输出:\n", proc.stdout)
                return

            self.set_progress(100)
            self.set_status(f"成功生成 EXE: {output_exe}", "green")

        except Exception as e:
            traceback.print_exc()
            self.set_status(f"错误: {e}", "red")
        finally:
            try:
                if 'work_tmp' in locals() and os.path.isdir(work_tmp):
                    shutil.rmtree(work_tmp, ignore_errors=True)
            except Exception:
                pass

    # ---------------------------
    # 调用 7z 压缩文件
    # ---------------------------
    def run_7z_compress(self, seven_zip_cmd: str, source_path: str, out_archive: str, password: str):
        """
        调用 7z 压缩 source_path 到 out_archive，并设置密码。
        """
        os.makedirs(os.path.dirname(out_archive), exist_ok=True)
        cmd = [seven_zip_cmd, "a", "-t7z", out_archive, source_path, f"-p{password}", "-mhe=on"]
        subprocess.run(cmd, check=True)

    # ---------------------------
    # 生成解压脚本（用于打包成最终EXE的入口）
    # ---------------------------
    def _make_extractor_script(self, pw1: str, pw2: str) -> str:
        # The extractor script (a small GUI) will be packaged as the final exe entrypoint.
        # It reads embedded data1.7z and data2.7z and 7z.exe from sys._MEIPASS (PyInstaller), extracts to cwd.
        return f'''import os
import sys
import shutil
import tempfile
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox

def resource_path(name):
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, name)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), name)

class ExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文件解压工具")
        self.center_window(420, 160)
        self.setup_ui()

    def center_window(self, width, height):
        sw = self.root.winfo_screenwidth(); sh = self.root.winfo_screenheight()
        x = (sw - width)//2; y = (sh - height)//2
        self.root.geometry(f"{{width}}x{{height}}+{{x}}+{{y}}".format(width=width,height=height,x=x,y=y))

    def setup_ui(self):
        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill='both', expand=True)
        ttk.Label(frm, text="请输入解压密码:").grid(row=0, column=0, pady=10, sticky="e")
        self.pw_entry = ttk.Entry(frm, show="*", width=30)
        self.pw_entry.grid(row=0, column=1, pady=10, sticky="w")
        ttk.Button(frm, text="解压", command=self.extract).grid(row=1, column=1, pady=5, sticky="e")
        self.status_var = tk.StringVar(value="")
        ttk.Label(frm, textvariable=self.status_var).grid(row=2, column=0, columnspan=2, pady=5)

    def extract(self):
        pwd = self.pw_entry.get().strip()
        outdir = os.getcwd()
        tmpdir = tempfile.mkdtemp(prefix="extract_")
        try:
            seven = resource_path("7z.exe")
            # try data1 first
            archive1 = resource_path("data1.7z")
            archive2 = resource_path("data2.7z")
            # choose which archive by trying to extract (7z will return non-zero on bad password)
            cmd1 = [seven, "x", archive1, f"-p{{pwd}}", "-o" + tmpdir, "-y"]
            proc1 = subprocess.run(cmd1, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            if proc1.returncode == 0 and os.listdir(tmpdir):
                # move extracted to current dir
                for name in os.listdir(tmpdir):
                    shutil.move(os.path.join(tmpdir, name), outdir)
                self.status_var.set("解压成功（文件A）")
                return
            # else try archive2
            shutil.rmtree(tmpdir, ignore_errors=True)
            tmpdir = tempfile.mkdtemp(prefix="extract_")
            cmd2 = [seven, "x", archive2, f"-p{{pwd}}", "-o" + tmpdir, "-y"]
            proc2 = subprocess.run(cmd2, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            if proc2.returncode == 0 and os.listdir(tmpdir):
                for name in os.listdir(tmpdir):
                    shutil.move(os.path.join(tmpdir, name), outdir)
                self.status_var.set("解压成功（文件B）")
                return
            # neither succeeded
            messagebox.showerror("错误", "密码错误或文件损坏，未解压任何内容。")
        except Exception as ex:
            messagebox.showerror("错误", f"解压出错：{{ex}}")
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

if __name__ == "__main__":
    root = tk.Tk()
    # Ensure ttk style (match main UI feel)
    style = ttk.Style()
    style.configure("TButton", padding=6, relief="flat")
    style.configure("TLabel", font=("Arial", 10))
    style.configure("TEntry", padding=5)
    app = ExtractorApp(root)
    root.mainloop()
'''.replace("{", "{{").replace("}", "}}").replace("{{pwd}}", "{pwd}").replace("{{ex}}", "{ex}")


    # ---------------------------
    # 更新状态和进度条
    # ---------------------------
    def set_status(self, text: str, color: str = "blue"):
        self.root.after(0, lambda: (self.status_var.set(text), self.status_label.config(foreground=color)))

    def set_progress(self, value: int):
        self.root.after(0, lambda: self.progress.config(value=value))


# ---------------------------
# 程序入口
# ---------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = DualFileCompressor(root)
    root.mainloop()
