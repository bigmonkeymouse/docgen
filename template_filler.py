#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word模板填充工具
根据Excel数据替换Word模板中的{{占位符}}
支持格式保留和页眉页脚脚注尾注处理
"""

import argparse
import sys
import re
import zipfile
import os
from xml.etree import ElementTree as ET
from pathlib import Path
from docx import Document
from openpyxl import load_workbook
import datetime
import json
from pypinyin import pinyin, Style


def load_excel_data(excel_path):
    """
    从Excel文件加载数据
    新格式：两列，第一列是字段名，第二列是对应的值
    返回字典列表，每行数据为一个字典
    """
    # 验证文件存在性
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel文件不存在: {excel_path}")
    
    workbook = load_workbook(excel_path, data_only=True)
    sheet = workbook.active

    # 读取两列数据
    raw_data = {}
    for row in sheet.iter_rows(min_row=1, values_only=True):
        if len(row) >= 2 and row[0] is not None:
            key = str(row[0]).strip()
            value = str(row[1]) if row[1] is not None else ""
            if key:  # 只处理有字段名的行
                raw_data[key] = value

    # 生成处理后的数据
    processed_data = {}
    
    # 1. 所在分行缩写
    if "所在分行" in raw_data:
        branch = raw_data["所在分行"]
        # 提取城市名（去掉"分行"二字）
        city_name = branch.replace("分行", "")
        # 提取拼音首字母大写
        pinyin_list = pinyin(city_name, style=Style.FIRST_LETTER)
        pinyin_abbr = "".join([p[0].upper() for p in pinyin_list])
        processed_data["所在分行缩写"] = f"{pinyin_abbr}分行"
    
    # 2. 批复金额
    if "批复金额" in raw_data:
        try:
            amount = float(raw_data["批复金额"])
            # 格式化保留两位小数（国际计数法，3位一节用逗号隔开）
            formatted_amount = "{:,.2f}".format(amount)
            # 存储格式化后的金额
            processed_data["批复金额"] = formatted_amount
            # 转换为银行标准大写
            processed_data["批复金额大写"] = num_to_Chinese(amount)
        except ValueError:
            processed_data["批复金额"] = ""
            processed_data["批复金额大写"] = ""
    
    # 3. 额度启用日期和到期日期
    if "额度启用日期" in raw_data:
        enable_date_str = raw_data["额度启用日期"]
        try:
            # 解析日期（格式如"2026年2月"）
            parts = enable_date_str.split("年")
            year = int(parts[0])
            month = int(parts[1].replace("月", ""))
            # 额度启用日期缩写
            processed_data["额度启用日期缩写"] = f"{year}.{month:02d}"
            # 额度到期日期（加一年）
            processed_data["额度到期日期"] = f"{year+1}年{month}月"
        except (IndexError, ValueError):
            processed_data["额度启用日期缩写"] = ""
            processed_data["额度到期日期"] = ""
    
    # 4. 合同制作日期
    current_date = datetime.datetime.now()
    processed_data["合同制作日期"] = current_date.strftime("%Y%m%d")
    
    # 5. 合同制作号
    processed_data["合同制作号"] = generate_contract_number()
    
    # 6. 各类合同编号
    if "所在分行缩写" in processed_data and "部门" in raw_data:
        branch_abbr = processed_data["所在分行缩写"]
        department = raw_data["部门"]
        contract_date = processed_data["合同制作日期"]
        contract_number = processed_data["合同制作号"]
        
        processed_data["综字1"] = f"平银{branch_abbr}{department}综字{contract_date}{contract_number}"
        processed_data["资池字2"] = f"平银{branch_abbr}{department}资池字{contract_date}{contract_number}"
        processed_data["资池质字3"] = f"平银{branch_abbr}{department}资池质字{contract_date}{contract_number}"
        processed_data["自由票字4"] = f"平银{branch_abbr}{department}自由票字{contract_date}{contract_number}"
        processed_data["线融字5"] = f"平银{branch_abbr}{department}线融字{contract_date}{contract_number}"
        processed_data["国内信字6"] = f"平银{branch_abbr}{department}国内信字{contract_date}{contract_number}"
        processed_data["国内信商字7"] = f"平银{branch_abbr}{department}国内信商字{contract_date}{contract_number}"
    
    # 合并原始数据和处理后的数据
    data = {**raw_data, **processed_data}
    
    print(f"处理后的数据: {list(processed_data.keys())}")
    
    # 返回包含一个字典的列表
    return [data] if data else []


def num_to_Chinese(num):
    """
    将数字转换为银行标准大写金额
    """
    digits = ["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"]
    units = ["", "拾", "佰", "仟"]
    big_units = ["", "万", "亿"]
    
    if num == 0:
        return "零元整"
    
    # 处理整数部分和小数部分
    integer_part = int(num)
    decimal_part = int(round((num - integer_part) * 100))
    
    result = ""
    
    # 处理整数部分
    if integer_part > 0:
        unit_index = 0
        big_unit_index = 0
        temp = ""
        
        while integer_part > 0:
            digit = integer_part % 10
            if digit > 0:
                temp = digits[digit] + units[unit_index] + temp
            else:
                # 避免连续的零
                if temp and not temp.startswith("零"):
                    temp = "零" + temp
            
            unit_index += 1
            if unit_index == 4:
                # 每四位处理一次
                if temp:
                    result = temp + big_units[big_unit_index] + result
                temp = ""
                unit_index = 0
                big_unit_index += 1
            
            integer_part = integer_part // 10
        
        if temp:
            result = temp + big_units[big_unit_index] + result
        
        result += "元"
    
    # 处理小数部分
    if decimal_part == 0:
        result += "整"
    else:
        jiao = decimal_part // 10
        fen = decimal_part % 10
        if jiao > 0:
            result += digits[jiao] + "角"
        if fen > 0:
            result += digits[fen] + "分"
    
    return result


def generate_contract_number():
    """
    生成合同制作号
    规则：每日早上8点开始编号，初始值为"第001号"，每1分30秒递增1，次日早上8点重置
    """
    # 存储上次生成的信息
    state_file = os.path.join(os.path.dirname(__file__), "contract_number_state.json")
    
    now = datetime.datetime.now()
    today = now.date()
    
    # 检查是否需要重置
    if now.hour < 8:
        # 凌晨到8点，使用前一天的编号
        yesterday = today - datetime.timedelta(days=1)
        reset_time = datetime.datetime.combine(yesterday, datetime.time(8, 0, 0))
    else:
        reset_time = datetime.datetime.combine(today, datetime.time(8, 0, 0))
    
    # 计算从重置时间到现在经过的秒数
    seconds_since_reset = (now - reset_time).total_seconds()
    
    # 每1分30秒（90秒）递增一个编号
    interval = int(seconds_since_reset // 90)
    
    # 读取或初始化状态
    try:
        if os.path.exists(state_file):
            with open(state_file, 'r', encoding='utf-8') as f:
                state = json.load(f)
        else:
            state = {
                "last_reset_date": reset_time.strftime("%Y-%m-%d"),
                "last_interval": -1,
                "current_number": 0
            }
    except Exception:
        state = {
            "last_reset_date": reset_time.strftime("%Y-%m-%d"),
            "last_interval": -1,
            "current_number": 0
        }
    
    # 检查是否需要重置
    if state["last_reset_date"] != reset_time.strftime("%Y-%m-%d"):
        state["last_reset_date"] = reset_time.strftime("%Y-%m-%d")
        state["last_interval"] = -1
        state["current_number"] = 0
    
    # 计算应该的编号（从1开始）
    current_number = interval + 1
    
    # 保存状态
    state["last_interval"] = interval
    state["current_number"] = current_number
    
    try:
        with open(state_file, 'w', encoding='utf-8') as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    
    # 生成编号
    return f"第{current_number:03d}号"


def find_placeholders(text):
    """
    查找文本中的所有占位符
    返回：[(占位符全文本如{{name}}, 占位符名称如name), ...]
    """
    pattern = r'\{\{([^}]+)\}\}'
    matches = re.finditer(pattern, text)
    return [(match.group(0), match.group(1).strip()) for match in matches]


def replace_in_paragraph(paragraph, replacement_dict, debug=False):
    """
    在段落中替换占位符，保留原有格式
    支持占位符跨越多个run的情况
    """
    if not paragraph.runs:
        return False

    # 合并所有run的文本用于匹配
    full_text = "".join([run.text for run in paragraph.runs])

    # 查找所有占位符
    placeholders = find_placeholders(full_text)

    if not placeholders:
        return False

    replaced = False
    if debug:
        print(f"  找到段落: {full_text[:50]}...")

    # 逐个处理每个占位符
    for placeholder_text, placeholder_name in placeholders:
        if placeholder_name in replacement_dict:
            replacement_value = str(replacement_dict[placeholder_name])

            # 在完整文本中找到占位符的位置
            placeholder_pos = full_text.find(placeholder_text)
            if placeholder_pos == -1:
                continue

            placeholder_end = placeholder_pos + len(placeholder_text)

            # 找到占位符跨越的所有run
            current_pos = 0
            runs_to_modify = []
            for run in paragraph.runs:
                run_start = current_pos
                run_end = current_pos + len(run.text)
                current_pos = run_end

                # 检查这个run是否与占位符有交集
                if run_start < placeholder_end and run_end > placeholder_pos:
                    runs_to_modify.append((run, run_start, run_end))

            if runs_to_modify:
                # 检查占位符是否完全在一个run中
                first_run, first_start, first_end = runs_to_modify[0]
                last_run, last_start, last_end = runs_to_modify[-1]

                if len(runs_to_modify) == 1:
                    # 占位符完全在一个run中，直接替换
                    first_run.text = first_run.text.replace(placeholder_text, replacement_value)
                else:
                    # 占位符跨越多个run，需要合并处理
                    # 计算占位符在第一个run中的偏移
                    offset_in_first = placeholder_pos - first_start
                    # 计算占位符在最后一个run中的结束位置
                    offset_in_last = placeholder_end - last_start

                    # 保留第一个run中占位符之前的文本
                    prefix = first_run.text[:offset_in_first] if offset_in_first > 0 else ""
                    # 保留最后一个run中占位符之后的文本
                    suffix = last_run.text[offset_in_last:] if offset_in_last < len(last_run.text) else ""

                    # 清空中间的所有run
                    for run, _, _ in runs_to_modify[1:-1]:
                        run.text = ""

                    # 在第一个run中设置完整替换值
                    first_run.text = prefix + replacement_value + suffix
                    # 清空最后一个run（内容已合并到第一个run）
                    last_run.text = ""

                replaced = True

                if debug:
                    print(f"    替换: {placeholder_text} -> {replacement_value}")

                # 更新full_text用于后续查找
                full_text = "".join([run.text for run in paragraph.runs])

    return replaced


def replace_in_table(table, replacement_dict, debug=False):
    """
    在表格中替换占位符
    """
    replaced = False
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                if replace_in_paragraph(paragraph, replacement_dict, debug):
                    replaced = True
    return replaced


def replace_in_headers(doc, replacement_dict, debug=False):
    """
    在页眉中替换占位符
    """
    replaced = False

    # 遍历所有section的header
    for i, section in enumerate(doc.sections):
        if section.header:
            if debug:
                print(f"\n处理页眉 Section {i}:")
            for paragraph in section.header.paragraphs:
                if replace_in_paragraph(paragraph, replacement_dict, debug):
                    replaced = True

            # 处理页眉中的表格
            for table in section.header.tables:
                if replace_in_table(table, replacement_dict, debug):
                    replaced = True

    return replaced


def replace_in_footers(doc, replacement_dict, debug=False):
    """
    在页脚中替换占位符
    """
    replaced = False

    # 遍历所有section的footer
    for i, section in enumerate(doc.sections):
        if section.footer:
            if debug:
                print(f"\n处理页脚 Section {i}:")
            for paragraph in section.footer.paragraphs:
                if replace_in_paragraph(paragraph, replacement_dict, debug):
                    replaced = True

            # 处理页脚中的表格
            for table in section.footer.tables:
                if replace_in_table(table, replacement_dict, debug):
                    replaced = True

    return replaced


def replace_in_footers_xml(docx_path, replacement_dict, output_path, debug=False):
    """
    直接操作XML文件，替换所有footer中的占位符
    python-docx可能无法读取某些footer，使用XML直接处理
    """
    replaced = False

    # 创建临时目录
    temp_dir = Path(output_path).parent / "temp"
    temp_dir.mkdir(exist_ok=True)

    # 解压docx文件
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    # 处理所有footer XML文件
    footer_files = list(temp_dir.glob("word/footer*.xml"))

    if debug:
        print(f"\n=== 处理Footer XML文件 ===")
        print(f"找到 {len(footer_files)} 个footer文件")

    for footer_file in footer_files:
        if debug:
            print(f"\n处理: {footer_file.name}")

        # 读取XML
        tree = ET.parse(footer_file)
        root = tree.getroot()

        # 递归替换所有文本节点中的占位符
        def replace_text_in_element(element):
            nonlocal replaced
            if element.text:
                placeholders = find_placeholders(element.text)
                for placeholder_text, placeholder_name in placeholders:
                    if placeholder_name in replacement_dict:
                        replacement_value = str(replacement_dict[placeholder_name])
                        if debug:
                            print(f"  替换: {placeholder_text} -> {replacement_value}")
                        element.text = element.text.replace(placeholder_text, replacement_value)
                        replaced = True

            for child in element:
                replace_text_in_element(child)
                if child.tail:
                    placeholders = find_placeholders(child.tail)
                    for placeholder_text, placeholder_name in placeholders:
                        if placeholder_name in replacement_dict:
                            replacement_value = str(replacement_dict[placeholder_name])
                            if debug:
                                print(f"  替换(尾文本): {placeholder_text} -> {replacement_value}")
                            child.tail = child.tail.replace(placeholder_text, replacement_value)
                            replaced = True

        replace_text_in_element(root)

        # 保存修改后的XML
        tree.write(footer_file, encoding='utf-8', xml_declaration=True)

    # 重新打包docx
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, temp_dir)
                zipf.write(file_path, arcname)

    # 清理临时目录
    import shutil
    shutil.rmtree(temp_dir)

    return replaced


def fill_template(template_path, excel_path, output_path, debug=False):
    """
    填充Word模板
    """
    # 加载Word模板
    doc = Document(template_path)

    # 加载Excel数据
    excel_data = load_excel_data(excel_path)

    if not excel_data:
        print("警告：Excel中没有数据")
        return

    # 使用第一行数据（如需多行生成，需要多次处理）
    replacement_dict = excel_data[0]

    print(f"从Excel中加载了 {len(replacement_dict)} 个替换字段")
    if debug:
        print(f"替换字段列表: {list(replacement_dict.keys())}")

    total_replaced = 0

    # 替换段落中的占位符
    if debug:
        print("\n=== 处理正文段落 ===")
    for paragraph in doc.paragraphs:
        if replace_in_paragraph(paragraph, replacement_dict, debug):
            total_replaced += 1

    # 替换表格中的占位符
    if debug:
        print("\n=== 处理正文表格 ===")
    for table in doc.tables:
        if replace_in_table(table, replacement_dict, debug):
            total_replaced += 1

    # 替换页眉中的占位符
    if replace_in_headers(doc, replacement_dict, debug):
        if debug:
            print("\n=== 页眉处理完成 ===")
        total_replaced += 1

    # 替换页脚中的占位符（python-docx方式）
    if replace_in_footers(doc, replacement_dict, debug):
        if debug:
            print("\n=== 页脚处理完成 ===")
        total_replaced += 1

    print(f"\n已完成正文替换，共处理 {total_replaced} 个段落/元素")

    # 保存中间结果
    temp_output = str(Path(output_path).parent / f"temp_{Path(output_path).name}")
    doc.save(temp_output)

    # 使用XML方式处理footer
    if replace_in_footers_xml(temp_output, replacement_dict, output_path, debug):
        print("Footer XML处理完成")
    else:
        # 如果XML处理没有替换，直接使用中间结果
        import os
        os.replace(temp_output, output_path)

    print(f"已保存到: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="Word模板填充工具")
    parser.add_argument("--template", required=True, help="Word模板文件路径")
    parser.add_argument("--excel", default="input.xlsx", help="Excel数据文件路径（默认：input.xlsx）")
    parser.add_argument("--output", required=True, help="输出文件路径")
    parser.add_argument("--debug", action="store_true", help="显示详细调试信息")

    args = parser.parse_args()

    # 检查文件存在性
    template_path = Path(args.template)
    excel_path = Path(args.excel)
    output_path = Path(args.output)

    if not template_path.exists():
        print(f"错误：模板文件不存在: {template_path}")
        sys.exit(1)

    if not excel_path.exists():
        print(f"错误：Excel文件不存在: {excel_path}")
        sys.exit(1)

    print(f"=== Word模板填充工具 ===")
    print(f"模板文件: {template_path}")
    print(f"Excel数据: {excel_path}")
    print(f"输出文件: {output_path}")

    # 创建输出目录（如果需要）
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # 执行填充
    try:
        fill_template(template_path, excel_path, output_path, debug=args.debug)
        print("\n填充成功！")
    except Exception as e:
        print(f"\n填充失败: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
