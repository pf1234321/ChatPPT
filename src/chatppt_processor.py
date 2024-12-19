# chatppt_processor.py
import os
from input_parser import parse_input_text
from ppt_generator import generate_presentation
from template_manager import load_template, get_layout_mapping, print_layouts
from layout_manager import LayoutManager
from config import Config
from logger import LOG  # 引入 LOG 模块

def process_user_input(user_input):
    config = Config()  # 加载配置文件

    # 将用户输入保存为临时 markdown 文件
    temp_md_file = "temp_input.md"
    with open(temp_md_file, 'w', encoding='utf-8') as file:
        file.write(user_input)

    # 检查输入的 markdown 文件是否存在
    if not os.path.exists(temp_md_file):
        LOG.error(f"{temp_md_file} 不存在。")  # 如果文件不存在，记录错误日志
        return

    # 读取 markdown 文件的内容
    with open(temp_md_file, 'r', encoding='utf-8') as file:
        input_text = file.read()

    # 加载 PowerPoint 模板，并打印模板中的可用布局
    prs = load_template(config.ppt_template)  # 加载模板文件
    LOG.info("可用的幻灯片布局:")  # 记录信息日志，打印可用布局
    print_layouts(prs)  # 打印模板中的布局

    # 初始化 LayoutManager，使用配置文件中的 layout_mapping
    layout_manager = LayoutManager(config.layout_mapping)

    # 调用 parse_input_text 函数，解析输入文本，生成 PowerPoint 数据结构
    powerpoint_data, presentation_title = parse_input_text(input_text, layout_manager)

    LOG.info(f"解析转换后的 ChatPPT PowerPoint 数据结构:\n{powerpoint_data}")  # 记录调试日志，打印解析后的 PowerPoint 数据

    # 定义输出 PowerPoint 文件的路径
    output_pptx = f"outputs/{presentation_title}.pptx"

    # 调用 generate_presentation 函数生成 PowerPoint 演示文稿
    generate_presentation(powerpoint_data, config.ppt_template, output_pptx)

    return output_pptx
