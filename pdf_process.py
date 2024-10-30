# -*- coding:utf-8 -*-

from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import os
import time


def arr_check(key, n):
    if not key:
        return False
    if not all(i.isdigit() for i in key.split()):
        return False
    arr = list(map(int, key.split()))
    arr.sort()
    i = 1
    for each in arr:
        if each == i:
            i += 1
        else:
            return False
    return i == n + 1


def merge():
    print("注意:\n-请将需要合并的pdf全部放到本程序同目录下")
    print("-大量文件合并时建议先对文件重命名, 使按文件名升序排序结果为合并顺序")

    arr = []
    # 列出目录中的所有pdf
    files = sorted([each for each in os.listdir() if each.endswith(".pdf")])
    while not arr:
        print("当前扫描到的pdf(按合并顺序如下):")
        for i, each in enumerate(files):
            print(f"{i+1} --- {each}")
        key = input(
            "是否按当前顺序合并?\n-输入y: 开始合并\n-输入r: 重新扫描pdf文件\n-输入上述序号的排列(按空格分隔): 指定新合并顺序\n"
        ).strip()
        while key != "y" and key != "r" and not arr_check(key, len(files)):
            key = input("非法输入! 请输入y或序号的合法排列!\n").strip()
        if key == "y":
            arr = files
        elif key == "r":
            files = sorted([each for each in os.listdir() if each.endswith(".pdf")])
        else:
            idx = list(map(int, key.split()))
            files = [files[i - 1] for i in idx]
    files = arr
    merger = PdfMerger()
    for file in files:
        with open(file, "rb") as f:
            merger.append(f)
    filename = f"[MERGE]{time.strftime('%Y-%m-%d_%H%M%S', time.localtime())}.pdf"
    with open(filename, "wb") as fout:
        merger.write(fout)
    print(f"任务完成, 合并pdf ====> {filename}\n")
    key = input("输入r返回功能选择页面, 输入其他任意键结束...\n").strip()
    if key == "r":
        main()


def split_idx_check(key, n):
    if not key:
        return False
    if not all(i.isdigit() for i in key.split()):
        return False
    arr = list(map(int, key.split()))
    return max(arr) <= n and min(arr) >= 1


def write_split_file(path, filename, last, cnt, writer):
    # 创建输出文件名
    output_filename = os.path.join(path, filename.format(begin=last, end=cnt))
    with open(output_filename, "wb") as output_file:
        writer.write(output_file)
    print(f"已分割页面 {last} - {cnt} 到文件 {output_filename}")


def split():
    print("注意: 拆分的源pdf请放到本程序同目录")
    file = input("请输入需要拆分的pdf名称(可以不带.pdf后缀):\n").strip()
    while not os.path.isfile(file) and not os.path.isfile(file + ".pdf"):
        print("非法输入,请重试!\n")
        file = input("请输入需要拆分的pdf名称(可以不带.pdf后缀):\n").strip()
    if not os.path.isfile(file):
        file += ".pdf"
    with open(file, "rb") as f:
        reader = PdfReader(f)
        print(f"{file} 共有 {len(reader.pages)} 页")
        print("需要输入拆分的页码, 多页码使用空格分隔, 文件将在指定的页后分割")
        print(
            "   如:输入 1 3 5, 文件将被分割为 第一页, 第二页和第三页, 第四页和第五页 三个文件"
        )
        key = input("请输入拆分位置的页码:\n").strip()
        while not split_idx_check(key, len(reader.pages)):
            key = input("非法输入!\n请输入合法的页码:\n")

        arr = sorted(list(map(int, key.split())), reverse=True)
        writer = PdfWriter()
        path = f"output_{time.strftime('%Y-%m-%d_%H%M%S', time.localtime())}"
        os.makedirs(path, exist_ok=True)
        filename = "[SPLIT]from {begin} to {end}.pdf"
        cnt = 0
        last = 1
        for each in reader.pages:
            writer.add_page(each)
            cnt += 1
            if arr and cnt == arr[-1]:
                arr.pop()
                write_split_file(path, filename, last, cnt, writer)
                writer = PdfWriter()
                last = cnt + 1
        if cnt >= last:
            write_split_file(path, filename, last, cnt)
    print(f"任务完成, 拆分pdf ====> {path}文件夹")
    key = input("输入r返回功能选择页面, 输入其他任意键结束...\n").strip()
    if key == "r":
        main()


def main():
    key = input("输入序号选择功能:\n1.合并pdf\n2.拆分pdf\n").strip()
    if key == "1":
        merge()
    elif key == "2":
        split()
    else:
        main()


if __name__ == "__main__":
    main()
