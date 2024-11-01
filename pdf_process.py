from os import getcwd, makedirs
from os.path import isfile, join
from traceback import format_exception
from time import localtime, strftime
from abc import ABC, abstractmethod
from tkinter import (
    Tk,
    Frame,
    Label,
    Button,
    Entry,
    StringVar,
    IntVar,
    Listbox,
    PhotoImage,
    Text,
    Toplevel,
)
from tkinter.ttk import Notebook, Combobox, Progressbar
from tkinter.filedialog import (
    askdirectory,
    askopenfilename,
    askopenfilenames,
    asksaveasfilename,
)
from tkinter.messagebox import showerror
from windnd import hook_dropfiles
from pymupdf import open as pdf_open, Pixmap

SPLIT_FUNC_1 = "指定起始页提取"
SPLIT_FUNC_2 = "按固定间隔拆分"
SPLIT_FUNC_3 = "指定页码拆分"


class FileApp(ABC, Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.input_single_pdf_path = StringVar()
        self.input_single_pdf_path.set("请点击按钮选择文件或将文件拖动至窗口内")
        self.output_path = StringVar()
        self.output_path.set(join(getcwd(), "output"))
        self.progress_info = StringVar()
        self.progress_index = StringVar()

    @abstractmethod
    def start_handle(self):
        pass

    def create_single_input_button(self, master, button_text, position="left"):
        hook_dropfiles(master, func=self.drag_single_file)
        upload_button = Button(
            master,
            text=button_text,
            command=self.upload_single_file,
            width=10,
            height=1,
            relief="groove",
        )
        upload_button.pack(side=position)

    def create_single_input_frame(self, master):
        # 上传文件或者输入文件路径
        upload_frame = Frame(master)
        upload_frame.pack(fill="both", padx=24, pady=(18, 0))
        hook_dropfiles(master, func=self.drag_single_file)
        upload_button = Button(
            upload_frame,
            text="选择文件",
            command=self.upload_single_file,
            width=10,
            height=1,
            relief="groove",
        )
        upload_label = Label(upload_frame, textvariable=self.input_single_pdf_path)
        upload_button.pack(side="left", padx=(0, 18))
        upload_label.pack(side="left", fill="x")

    def create_select_ouput_frame(self, master):
        # 选择输出文件夹
        output_frame = Frame(master)
        output_frame.pack(fill="x", padx=24, pady=(18, 12))
        output_entry = Entry(output_frame, textvariable=self.output_path)
        output_entry.pack(side="left", fill="both", expand=True)
        output_button = Button(
            output_frame,
            text="输出文件夹",
            command=self.select_output,
            width=10,
            height=1,
            relief="groove",
        )
        output_button.pack(side="left", padx=(18, 0))

    def create_progress_frame(self, master, button_text):
        # 进度条
        progress_frame = Frame(master)
        progress_frame.pack(fill="both", padx=24, pady=0)
        start_button = Button(
            progress_frame,
            text=button_text,
            command=self.start_handle,
            width=10,
            height=1,
            relief="groove",
        )
        start_button.pack(side="left", padx=(0, 18), pady=0)
        self.progress = Progressbar(
            progress_frame,
            orient="horizontal",
            mode="determinate",
        )
        self.progress.pack(side="left", fill="both", expand=True, pady=0)
        progress_label_frame = Frame(master)
        progress_label_frame.pack(fill="both", padx=18, pady=0)
        progress_label_left = Label(
            progress_label_frame, textvariable=self.progress_info
        )
        progress_label_left.pack(side="left", padx=(100, 0))
        progress_label_right = Label(
            progress_label_frame, textvariable=self.progress_index
        )
        progress_label_right.pack(side="right")

    def create_handle_button(self, master, button_text, padx=0, position="left"):
        start_button = Button(
            master,
            text=button_text,
            command=self.start_handle,
            width=10,
            height=1,
            relief="groove",
        )
        start_button.pack(side=position, padx=padx)

    def select_output(self):
        folder_path = askdirectory()
        if folder_path:
            self.output_path.set(folder_path)

    def callback(self, cur, tot, name):
        self.progress["value"] = cur
        self.progress["maximum"] = tot
        if name:
            self.progress_info.set(f"处理中: {name} ...")
        else:
            self.progress_info.set("处理完成!")
        self.progress_index.set(f"{cur}/{tot}")
        self.update()

    def drag_file_post_process(self, files):
        return files

    def drag_files(self, files):
        files = [
            f.decode("utf-8").replace("\\", "/")
            for f in files
            if isfile(f.decode("utf-8")) and f.decode("utf-8").endswith(".pdf")
        ]
        return self.drag_file_post_process(files)

    def drag_single_file_post_process(self, file):
        return file

    def drag_single_file(self, files):
        files = [
            f.decode("utf-8").replace("\\", "/")
            for f in files
            if isfile(f.decode("utf-8")) and f.decode("utf-8").endswith(".pdf")
        ]
        self.input_single_pdf_path.set(files[0])
        self.update()
        return self.drag_single_file_post_process(files[0])

    def upload_single_file_post_process(self, file):
        return file

    def upload_single_file(self):
        file = askopenfilename(filetypes=[("PDF files", "*.pdf")])
        file.replace("\\", "/")
        if isfile(file):
            self.input_single_pdf_path.set(file)
            return self.upload_single_file_post_process(file)

    def upload_files_post_process(self, files):
        return files

    def upload_files(self):
        files = askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if not files:
            return
        files = [file.replace("\\", "/") for file in files]
        return self.upload_files_post_process(files)


class AboutApp(Frame):
    def __init__(self, master, text):
        super().__init__(master)
        self.master = master
        text_box = Text(self)
        text_box.insert("1.0", text)
        text_box.config(state="disabled", wrap="word")
        text_box.pack(expand=True, fill="both")
        self.pack(expand=True, fill="both")


class DetailApp(FileApp):
    def __init__(self, master):
        super().__init__(master)
        self.create_vars()
        self.create_widgets()
        self.pack(expand=True, fill="both")

    def create_vars(self):
        self.pdf_info = StringVar()
        self.pdf_cur = 0
        self.pdf_tot = 0
        self.page_text = ""
        self.pix_image = None
        self.page_rotations = []
        self.rate = 1

    def create_widgets(self):
        # 鼠标进入自动获取焦点
        self.bind("<Enter>", lambda event: self.focus_set())
        # 键盘输入处理
        self.bind("<Key>", self.key_press)
        # 自适应窗口大小
        self.bind("<Configure>", self.window_change)
        pdf_frame = Frame(self)
        pdf_frame.pack(fill="both", side="left", expand=True)
        # 功能栏
        banner = Frame(pdf_frame)
        banner.pack(fill="both", side="top", padx=(24, 0), pady=(18, 0))
        self.create_single_input_button(banner, "选择文件")
        self.create_handle_button(banner, "保存", padx=(18, 0))
        # pdf操作
        self.switch_side_panel = Button(
            banner,
            text="<<",
            command=self.side_panel_control,
            width=2,
            height=1,
            relief="groove",
        )
        self.switch_side_panel.pack(side="right")
        pdf_extract_text = Button(
            banner,
            text="复制文本",
            command=self.copy_text,
            width=6,
            height=1,
            relief="groove",
        )
        pdf_extract_text.pack(side="right", padx=(18, 0))
        pdf_rotate_right = Button(
            banner,
            text="右转90°",
            command=self.rotate_right,
            width=6,
            height=1,
            relief="groove",
        )
        pdf_rotate_right.pack(side="right", padx=(18, 0))
        pdf_rotate_left = Button(
            banner,
            text="左转90°",
            command=self.rotate_left,
            width=6,
            height=1,
            relief="groove",
        )
        pdf_rotate_left.pack(side="right")
        # 展示pdf
        pdf_show_frame = Frame(pdf_frame)
        pdf_show_frame.pack(fill="both", expand=True, padx=24, pady=(12, 18))
        self.pdf_show = Label(pdf_show_frame, width=1, height=1)
        self.pdf_show.pack(fill="both", expand=True)
        pdf_control = Frame(pdf_show_frame)
        pdf_control.pack(fill="x", side="bottom")
        pdf_control_prev = Button(
            pdf_control,
            text="<",
            command=self.get_pdf_prev,
            width=6,
            height=1,
            relief="groove",
        )
        pdf_control_prev.pack(fill="both", side="left")
        pdf_control_next = Button(
            pdf_control,
            text=">",
            command=self.get_pdf_next,
            width=6,
            height=1,
            relief="groove",
        )
        pdf_control_next.pack(fill="both", side="right")
        pdf_control_info = Label(pdf_control, textvariable=self.pdf_info)
        pdf_control_info.pack(fill="y")
        # 侧边栏
        self.side_frame = Frame(self)
        # 禁止外层frame大小受子组件大小影响
        self.side_frame.pack_propagate(0)
        self.pdf_extract_text_area = Text(self.side_frame)
        self.pdf_extract_text_area.pack(fill="both", side="right")

    def window_change(self, event):
        self.resize_image()
        self.resize_side_frame()

    def key_press(self, event):
        if event.keysym == "Left":
            self.get_pdf_prev()
        elif event.keysym == "Right":
            self.get_pdf_next()

    def start_handle(self):
        if not isfile(self.input_single_pdf_path.get()):
            return
        file = asksaveasfilename(
            confirmoverwrite=True,
            filetypes=[("PDF", "*.pdf")],
            defaultextension=".pdf",
            initialfile=self.input_single_pdf_path.get().rsplit("/", 1)[1],
        )
        if not file:
            return
        with pdf_open(self.input_single_pdf_path.get()) as f:
            for i in range(self.pdf_tot):
                page = f.load_page(i)
                page.set_rotation(self.page_rotations[i])
            f.save(file)

    def rotate_left(self):
        with pdf_open(self.input_single_pdf_path.get()) as f:
            page = f.load_page(self.pdf_cur - 1)
            self.page_rotations[self.pdf_cur - 1] = (
                self.page_rotations[self.pdf_cur - 1] - 90
            ) % 360
            page.set_rotation(self.page_rotations[self.pdf_cur - 1])
            self.pix_image = page.get_pixmap(alpha=False)
        self.resize_image()

    def rotate_right(self):
        with pdf_open(self.input_single_pdf_path.get()) as f:
            page = f.load_page(self.pdf_cur - 1)
            self.page_rotations[self.pdf_cur - 1] = (
                self.page_rotations[self.pdf_cur - 1] + 90
            ) % 360
            page.set_rotation(self.page_rotations[self.pdf_cur - 1])
            self.pix_image = page.get_pixmap(alpha=False)
        self.resize_image()

    def copy_text(self):
        if not self.page_text:
            return
        self.clipboard_clear()
        self.clipboard_append(self.page_text)

    def get_pdf_text(self, page):
        self.page_text = page.get_text()
        self.pdf_extract_text_area.delete("1.0", "end")
        self.pdf_extract_text_area.insert("insert", self.page_text)

    def resize_side_frame(self):
        self.side_frame.configure(width=self.winfo_width() // 3)

    def side_panel_control(self):
        if not self.side_frame.winfo_ismapped():
            self.switch_side_panel.configure(text=">>")
            self.resize_side_frame()
            self.side_frame.pack(fill="both", side="right", pady=18)
        else:
            self.switch_side_panel.configure(text="<<")
            self.side_frame.pack_forget()
        self.resize_image()

    def resize_image(self, *args):
        if not self.pix_image:
            return
        w, h = self.pdf_show.winfo_width(), self.pdf_show.winfo_height()
        rate_w = w / self.pix_image.width
        rate_h = h / self.pix_image.height
        if rate_w > rate_h:
            w = int(self.pix_image.width * rate_h)
            self.rate = rate_h
        else:
            h = int(self.pix_image.height * rate_w)
            self.rate = rate_w
        self.image = PhotoImage(data=Pixmap(self.pix_image, w, h, None).tobytes("ppm"))
        self.pdf_show.configure(image=self.image)

    def get_pdf_prev(self):
        if self.pdf_cur <= 1:
            return
        self.pdf_cur -= 1
        with pdf_open(self.input_single_pdf_path.get()) as f:
            page = f.load_page(self.pdf_cur - 1)
            self.get_pdf_text(page)
            page.set_rotation(self.page_rotations[self.pdf_cur - 1])
            self.pix_image = page.get_pixmap(alpha=False)
        self.resize_image()
        self.pdf_info.set(f"{self.pdf_cur}/{self.pdf_tot}")

    def get_pdf_next(self):
        if self.pdf_cur >= self.pdf_tot:
            return
        self.pdf_cur += 1
        with pdf_open(self.input_single_pdf_path.get()) as f:
            page = f.load_page(self.pdf_cur - 1)
            self.get_pdf_text(page)
            page.set_rotation(self.page_rotations[self.pdf_cur - 1])
            self.pix_image = page.get_pixmap(alpha=False)
        self.resize_image()
        self.pdf_info.set(f"{self.pdf_cur}/{self.pdf_tot}")

    def upload_single_file_post_process(self, file):
        with pdf_open(file) as f:
            self.pdf_tot = f.page_count
            self.pix_image = f.load_page(0).get_pixmap(alpha=False)
            self.get_pdf_text(f.load_page(0))
            self.page_rotations = [f.load_page(i).rotation for i in range(self.pdf_tot)]
        self.resize_image()
        self.pdf_cur = 1
        self.pdf_info.set(f"{self.pdf_cur}/{self.pdf_tot}")

    def drag_single_file_post_process(self, file):
        with pdf_open(file) as f:
            self.pdf_tot = f.page_count
            self.pix_image = f.load_page(0).get_pixmap(alpha=False)
            self.get_pdf_text(f.load_page(0))
            self.page_rotations = [f.load_page(i).rotation for i in range(self.pdf_tot)]
        self.resize_image()
        self.pdf_cur = 1
        self.pdf_info.set(f"{self.pdf_cur}/{self.pdf_tot}")


class MergeHandler:
    def __init__(self):
        self.output_filename = "[MERGE] {timestamp}.pdf"

    def merge(self, files, callback):
        if not files:
            return
        filename = self.output_filename.format(
            timestamp=strftime("%Y-%m-%d_%H_%M_%S", localtime())
        )
        save_file = asksaveasfilename(
            confirmoverwrite=True,
            filetypes=[("PDF", "*.pdf")],
            defaultextension=".pdf",
            initialfile=filename,
        )
        if not save_file:
            return
        with pdf_open() as writer:
            tot = len(files)
            for i, file in enumerate(files):
                callback(i, tot, file)
                with pdf_open(file) as f:
                    writer.insert_pdf(f)
            callback(tot - 1, tot, "正在写入新文件...")
            writer.save(save_file)
        callback(tot, tot, None)


class MergeApp(FileApp):
    def __init__(self, master):
        super().__init__(master)
        self.create_vars()
        self.create_widgets()
        self.pack(expand=True, fill="both")

    def create_vars(self):
        self.handler = MergeHandler()
        self.files_set = set()
        self.files = list()
        self.files_variable = StringVar()

    def create_widgets(self):
        files_frame = Frame(self)
        files_frame.pack(fill="both", expand=True, padx=24, pady=18)
        files_list_frame = Frame(files_frame)
        files_list_frame.pack(fill="both", side="left", expand=True)
        self.files_list = Listbox(
            files_list_frame,
            listvariable=self.files_variable,
            selectmode="extended",
            activestyle="none",
        )
        self.files_list.pack(fill="both", expand=True, padx=0, pady=0)
        files_ops_frame = Frame(files_frame)
        files_ops_frame.pack(fill="both", side="right", padx=(18, 0), pady=0)
        hook_dropfiles(self, func=self.drag_files)
        files_button = Button(
            files_ops_frame,
            text="添加文件",
            command=self.upload_files,
            width=10,
            height=1,
            relief="groove",
        )
        files_button.pack(side="top", pady=(0, 12))
        files_button = Button(
            files_ops_frame,
            text="删除文件",
            command=self.delete_file,
            width=10,
            height=1,
            relief="groove",
        )
        files_button.pack(side="top", pady=(0, 12))
        files_button = Button(
            files_ops_frame,
            text="清空文件",
            command=self.clear_files,
            width=10,
            height=1,
            relief="groove",
        )
        files_button.pack(side="top", pady=(0, 12))
        files_button = Button(
            files_ops_frame,
            text="上移文件",
            command=self.move_up,
            width=10,
            height=1,
            relief="groove",
        )
        files_button.pack(side="top", pady=(0, 12))
        files_button = Button(
            files_ops_frame,
            text="下移文件",
            command=self.move_down,
            width=10,
            height=1,
            relief="groove",
        )
        files_button.pack(side="top", pady=(0, 12))
        files_button = Button(
            files_ops_frame,
            text="按名称升序",
            command=self.sort_by_name_asc,
            width=10,
            height=1,
            relief="groove",
        )
        files_button.pack(side="top", pady=(0, 12))
        files_button = Button(
            files_ops_frame,
            text="按名称降序",
            command=self.sort_by_name_desc,
            width=10,
            height=1,
            relief="groove",
        )
        files_button.pack(side="top")
        self.create_progress_frame(self, "开始合并")

    def start_handle(self):
        self.handler.merge(self.files, self.callback)

    def drag_file_post_process(self, files):
        for file in files:
            if file not in self.files_set:
                self.files_set.add(file)
                self.files.append(file)
        self.files_variable.set(self.files)

    def upload_files_post_process(self, files):
        for file in files:
            if file not in self.files_set:
                self.files_set.add(file)
                self.files.append(file)
        self.files_variable.set(self.files)

    def delete_file(self):
        tmp = [self.files[i] for i in self.files_list.curselection()]
        for each in tmp:
            self.files_set.discard(each)
            self.files.remove(each)
        self.files_variable.set(self.files)

    def clear_files(self):
        self.files_set.clear()
        self.files.clear()
        self.files_variable.set([])

    def move_up(self):
        if not self.files_list.curselection():
            return
        idx = self.files_list.curselection()[0]
        if idx > 0:
            self.files[idx], self.files[idx - 1] = self.files[idx - 1], self.files[idx]
            self.files_list.select_set(idx - 1)
            self.files_list.activate(idx - 1)
            self.files_list.select_clear(idx)
        self.files_variable.set(self.files)

    def move_down(self):
        if not self.files_list.curselection():
            return
        idx = self.files_list.curselection()[0]
        if idx < len(self.files) - 1:
            self.files[idx], self.files[idx + 1] = self.files[idx + 1], self.files[idx]
            self.files_list.select_set(idx + 1)
            self.files_list.activate(idx + 1)
            self.files_list.select_clear(idx)
        self.files_variable.set(self.files)

    def sort_by_name_asc(self):
        self.files.sort()
        self.files_variable.set(self.files)

    def sort_by_name_desc(self):
        self.files.sort(reverse=True)
        self.files_variable.set(self.files)


class SplitHandler:
    def __init__(self):
        self.raw_path = None
        self.pages = 0
        self.output_filename = "[SPLIT]{name} from {begin} to {end}.pdf"
        self.split_func = {
            SPLIT_FUNC_1: self.split_with_range,
            SPLIT_FUNC_2: self.split_with_stride,
            SPLIT_FUNC_3: self.split_with_idx,
        }

    def get_pages(self, path):
        self.raw_path = path
        with pdf_open(path) as f:
            self.pages = f.page_count

    def split(self, output_path, func, param, callback):
        makedirs(output_path, exist_ok=True)
        filename = join(output_path, self.output_filename)
        self.split_func[func](filename, param, callback)

    def split_with_range(self, filename, param, callback):
        begin = param["begin"].get()
        end = param["end"].get()
        if begin < 1 or end > self.pages or begin > end:
            raise Exception(
                "非法输入, 请检查输入的页码范围是否合法!\n注意: 结束页码必须大于等于开始页码"
            )
        raw_name = self.raw_path.rsplit("/", 1)[1].rsplit(".", 1)[0]
        output_filename = filename.format(name=raw_name, begin=begin, end=end)
        callback(0, 1, output_filename)
        with pdf_open() as writer:
            with pdf_open(self.raw_path) as f:
                writer.insert_pdf(f, from_page=begin - 1, to_page=end - 1)
            writer.save(output_filename)
        callback(1, 1, None)

    def split_with_stride(self, filename, param, callback):
        stride = param.get()
        if stride < 1 or stride > self.pages:
            raise Exception(
                "非法输入, 请检查输入的步长是否合法!\n注意: 步长必须大于0且小于总页数"
            )
        raw_name = self.raw_path.rsplit("/", 1)[1].rsplit(".", 1)[0]
        tot = (self.pages + stride - 1) // stride
        cur = 0
        with pdf_open(self.raw_path) as f:
            for l in range(1, self.pages + 1, stride):
                r = min(l + stride - 1, self.pages)
                output_filename = filename.format(name=raw_name, begin=l, end=r)
                callback(cur, tot, output_filename)
                with pdf_open() as writer:
                    writer.insert_pdf(f, from_page=l - 1, to_page=r - 1)
                    writer.save(output_filename)
                cur += 1
        callback(cur, tot, None)

    def split_with_idx(self, filename, param, callback):
        param = param.get()
        if not self.split_idx_check(param, self.pages):
            raise Exception(
                "非法输入, 请检查输入的页码序列是否合法!\n注意: 页码之间只能用空格分隔"
            )
        # 这里去重 + 排序
        arr = sorted(list(set(map(int, param.split()))), reverse=True)
        l = 1
        ranges = []
        for r in arr:
            ranges.append((l, r))
            l = r + 1
        if l <= self.pages:
            ranges.append((l, self.pages))
        raw_name = self.raw_path.rsplit("/", 1)[1].rsplit(".", 1)[0]
        tot = len(arr) + 1 if arr[-1] < self.pages else len(arr)
        cur = 0
        with pdf_open(self.raw_path) as f:
            for l, r in ranges:
                output_filename = filename.format(name=raw_name, begin=l, end=r)
                callback(cur, tot, output_filename)
                with pdf_open() as writer:
                    writer.insert_pdf(f, from_page=l - 1, to_page=r - 1)
                    writer.save(output_filename)
                cur += 1
        callback(cur, tot, None)

    def split_idx_check(self, input, n):
        if not input:
            return False
        if not all(i.isdigit() for i in input.split()):
            return False
        arr = list(map(int, input.split()))
        return max(arr) <= n and min(arr) >= 1


class SplitApp(FileApp):
    def __init__(self, master):
        super().__init__(master)
        self.create_vars()
        self.create_widgets()
        self.pack(expand=True, fill="both")

    def create_vars(self):
        self.handler = SplitHandler()
        self.pdf_info = StringVar()
        self.func_choice = StringVar()
        self.func_info = {}
        self.func_vars = {}
        self.func_combos = []
        self.progress_filename = StringVar()
        self.progress_index = StringVar()

    def create_widgets(self):
        self.create_single_input_frame(self)
        upload_comment_frame = Frame(self)
        upload_comment_frame.pack(fill="x", padx=24)
        # 展示pdf信息
        upload_comment_label = Label(upload_comment_frame, textvariable=self.pdf_info)
        upload_comment_label.pack(side="left")
        # pdf拆分功能
        func_frame = Frame(self)
        func_frame.pack(fill="x", padx=24)
        func_select = Combobox(
            func_frame, textvariable=self.func_choice, width=20, height=1
        )
        # 初始化选框
        self.init_func_info(func_frame)
        func_select["value"] = list(self.func_info.keys())
        func_select.current(0)
        func_select.pack(side="left", fill="both")
        self.func_info[self.func_choice.get()].pack(side="left")
        # 选框更换事件
        func_select.bind("<<ComboboxSelected>>", self.func_switch)
        label_text = "输出文件命名格式: 输出文件夹/[SPLIT]{原pdf名称} from {起始页} to {结束页}.pdf\n注意: 若有相同文件名会自动覆盖"
        output_label = Label(self, text=label_text)
        output_label.pack(expand=True)
        self.create_select_ouput_frame(self)
        self.create_progress_frame(self, "开始拆分")

    def start_handle(self):
        self.handler.split(
            self.output_path.get(),
            self.func_choice.get(),
            self.func_vars[self.func_choice.get()],
            self.callback,
        )

    def drag_single_file_post_process(self, file):
        self.handler.get_pages(file)
        # 刷新选框值域
        self.reset_func_combos()
        self.pdf_info.set(f"pdf文件上传成功, 共{self.handler.pages}页!")

    def upload_single_file_post_process(self, file):
        self.handler.get_pages(file)
        # 刷新选框值域
        self.reset_func_combos()
        self.pdf_info.set(f"pdf文件上传成功, 共{self.handler.pages}页!")

    def init_func_info(self, master) -> dict:
        self.func_vars[SPLIT_FUNC_1] = {"begin": IntVar(), "end": IntVar()}
        self.func_vars[SPLIT_FUNC_2] = IntVar()
        self.func_vars[SPLIT_FUNC_3] = StringVar()
        self.func_info[SPLIT_FUNC_1] = Frame(master)
        self.create_combobox(
            self.func_info[SPLIT_FUNC_1],
            "from: ",
            self.func_vars[SPLIT_FUNC_1]["begin"],
        )
        self.create_combobox(
            self.func_info[SPLIT_FUNC_1],
            "to: ",
            self.func_vars[SPLIT_FUNC_1]["end"],
        )

        self.func_info[SPLIT_FUNC_2] = Frame(master)
        self.create_combobox(
            self.func_info[SPLIT_FUNC_2],
            "stride: ",
            self.func_vars[SPLIT_FUNC_2],
        )

        self.func_info[SPLIT_FUNC_3] = Entry(
            master,
            textvariable=self.func_vars[SPLIT_FUNC_3],
        )

    def create_combobox(self, master, label_text, var):
        frame = Frame(master)
        frame.pack(side="left", fill="x", padx=(24, 0))
        label = Label(frame, text=label_text, width=8)
        label.pack(side="left")
        box = Combobox(frame, textvariable=var, width=8, height=1)
        box.pack(side="left", fill="both")
        self.func_combos.append(box)

    def reset_func_combos(self):
        for each in self.func_combos:
            each["values"] = list(range(1, self.handler.pages + 1))
            each.current(0)

    def func_switch(self, _):
        for frame in self.func_info.values():
            frame.pack_forget()
        self.func_info[self.func_choice.get()].pack(side="left")


class Application(Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.master.report_callback_exception = self.log_exception
        self.set_window()
        self.set_vars()
        self.create_widgets()
        self.pack(expand=True, fill="both")

    def set_window(self):
        self.master.title("PDF处理工具v0.2")
        self.master.geometry("600x400")

    def set_vars(self):
        self.about_text = (
            "                          PDF处理工具v0.2\n"
            + "                          作者: Dervey D\n"
            + "                     开源协议: AGPL-3.0 license\n"
            + "           Github: https://github.com/DerveyD/PdfProcessor"
        )
        with open("LICENSE") as f:
            self.license_text = f.read()

    def create_widgets(self):
        # 创建功能选项卡
        notebook = Notebook(self)
        merge_frame = MergeApp(notebook)
        split_frame = SplitApp(notebook)
        detail_frame = DetailApp(notebook)
        license_frame = AboutApp(notebook, self.license_text)
        about_frame = AboutApp(notebook, self.about_text)
        notebook.add(merge_frame, text="合并PDF")
        notebook.add(split_frame, text="拆分PDF")
        notebook.add(detail_frame, text="PDF浏览")
        notebook.add(about_frame, text="关于")
        notebook.add(license_frame, text="开源协议")
        notebook.pack(expand=True, fill="both")

    def log_exception(self, *args):
        exc_message = "".join(format_exception(*args))
        log_message = f"[{strftime('%Y-%m-%d %H:%M:%S', localtime())}] " + exc_message
        print(log_message)
        with open("error.log", "a") as f:
            f.write(log_message)
        showerror("Error", exc_message)


if __name__ == "__main__":
    app = Application(Tk())
    app.mainloop()
