from docx import Document
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import os


# 定义移除格式并保存到新的Word文件的函数
def remove_format_and_save_to_word(input_file, output_file):
    try:
        # 读取输入文件
        doc = Document(input_file)

        # 清除所有的格式（例如：删除字体、段落样式，保留纯文本）
        for para in doc.paragraphs:
            # 只保留文本，清除段落的样式和格式
            para.style = 'Normal'  # 将所有段落的样式设置为 'Normal'
            for run in para.runs:
                run.font.name = None  # 清除字体
                run.font.size = None  # 清除字体大小
                run.font.bold = None  # 清除加粗
                run.font.italic = None  # 清除斜体
                run.font.underline = None  # 清除下划线
                run.font.color.rgb = None  # 清除字体颜色

        # 保存新的Word文档
        doc.save(output_file)
        print(f"文件已保存到 {output_file}")
    except Exception as e:
        print(f"出现错误: {e}")


# 使用 Tkinter 打开文件选择对话框
def get_input_file():
    Tk().withdraw()  # 隐藏主窗口
    input_file = askopenfilename(title="请选择输入的Word文档", filetypes=[("Word 文档", "*.docx")])
    return input_file


def main():
    # 获取输入文件
    input_file = get_input_file()

    # 检查文件是否被正确选择
    if not input_file:
        print("未选择文件，程序退出")
        return

    # 确保文件是一个有效的 Word 文档 (.docx)
    if not input_file.lower().endswith(".docx"):
        print("请选择一个有效的 Word (.docx) 文件")
        return

    # 设置输出文件路径
    output_file = os.path.join(os.path.dirname(input_file), "output.docx")

    # 调用函数处理文档
    remove_format_and_save_to_word(input_file, output_file)


if __name__ == "__main__":
    main()
