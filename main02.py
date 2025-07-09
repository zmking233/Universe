import re
import sys
import fitz  # PyMuPDF
from docx import Document

def reduce_repeated_chars(text):
    def replacer(m):
        return m.group(2)
    pattern = r'(([\u4e00-\u9fa5])(\s*\2){2,})'
    return re.sub(pattern, replacer, text)

def fuzzy_kw_regex(kw):
    return r'\s*'.join([fr'{ch}(?:\s*{ch})*' for ch in kw])

def fuzzy_find_section(text, start_keywords, end_keywords):
    start_pattern = '|'.join(fuzzy_kw_regex(kw) for kw in start_keywords)
    end_pattern = '|'.join(fuzzy_kw_regex(kw) for kw in end_keywords)
    regex = re.compile(fr'({start_pattern})(.*?)({end_pattern})', re.DOTALL)
    m = regex.search(text)
    return m.group(2).strip() if m else ""

def clean_text(text):
    text = re.sub(r'第?\s*\d+\s*页', '', text)
    text = re.sub(r'\d+\s*/\s*\d+\s*页', '', text)
    watermark_pattern = r'(人\s*民\s*法\s*院\s*案\s*例\s*库){1,}'
    text = re.sub(watermark_pattern, '', text)
    text = reduce_repeated_chars(text)
    # 不去除空行，保留分段
    lines = text.splitlines()
    lines = [line.rstrip() for line in lines]
    return '\n'.join(lines)

def process_pdf_text_lines(raw_text):
    lines = raw_text.splitlines()
    new_lines = []
    buffer = ""
    for line in lines:
        line = line.strip()
        # 跳过分页导致的空行（如上一页结尾和下一页开头之间的空行）
        if not line:
            continue
        # 跳过分页符号（如“第X页”已在clean_text去除）
        if buffer and not buffer.endswith(('。', '！', '？', '：', '.', '!', '?', ':', '；')):
            buffer += line
        else:
            if buffer:
                new_lines.append(buffer)
            buffer = line
    if buffer:
        new_lines.append(buffer)

    # 分段逻辑：如果一行少于20个汉字且以句号结尾，则在其后分段
    final_lines = []
    for i, line in enumerate(new_lines):
        final_lines.append(line)
        hanzi_count = len([ch for ch in line if '\u4e00' <= ch <= '\u9fa5'])
        if hanzi_count < 20 and line.endswith('。'):
            final_lines.append('')  # 插入空行实现分段
    # 只保留我们主动插入的空行，忽略原始PDF的空行
    return '\n\n'.join([l for l in final_lines if l or (i > 0 and final_lines[i-1])])

def parse_fields(text):
    # 修复案号提取：支持2025-12-3-001-007格式
    m_num = re.search(r'(\d{4}(?:-\d+){3,4})', text)
    case_number = m_num.group(1) if m_num else "未知编号"

    # 修复案名提取：案号之后到第一个“案”为止
    m_name = re.search(fr'{re.escape(case_number)}\s*\n?(.*?案)', text, re.DOTALL)
    case_name = m_name.group(1).strip() if m_name else "未知案件名称"

    # 简要案情：从第一个“案”字后到“关键词”前
    m_desc = re.search(fr'案(.*?){fuzzy_kw_regex("关键词")}', text, re.DOTALL)
    case_desc = m_desc.group(1).strip() if m_desc else ""

    key_word = fuzzy_find_section(text, ["关键词"], ["基本案情"])
    case_text = fuzzy_find_section(text, ["基本案情"], ["裁判理由"])
    trial_process = fuzzy_find_section(text, ["裁判理由"], ["裁判要旨"])
    trial_abbr = fuzzy_find_section(text, ["裁判要旨"], ["关联索引"])
    # 关联索引：提取“关联索引”之后的所有文本
    m_index = re.search(fr'{fuzzy_kw_regex("关联索引")}(.*)', text, re.DOTALL)
    relevant_index = m_index.group(1).strip() if m_index else ""

    return {
        "case_number": case_number,
        "case_name": case_name,
        "case_desc": case_desc,
        "key_word": key_word,
        "case_text": case_text,
        "trial_process": trial_process,
        "trial_abbr": trial_abbr,
        "relevant_index": relevant_index
    }

def styled_paragraph(text, color, size, bold=False, align='justify', line_height=2, indent_px=16, margin_top=0, margin_bottom=0, background="#ffffff"):
    style = f"""
        color:{color};
        font-size:{size}px;
        text-align:{align};
        line-height:{line_height};
        margin-top:{margin_top}px;
        margin-bottom:{margin_bottom}px;
        margin-left:{indent_px}px;
        margin-right:{indent_px}px;
        background-color:{background};
        {'font-weight:bold;' if bold else ''}
    """
    return f'<p style="{style.strip()}">{text}</p>'

def generate_wechat_html(data):
    parts = []
    parts.append(styled_paragraph(
        f'<span style="background-color:#5287b7;color:#ffffff;">今日案例播客版  干货知识轻松听</span>',
        "#5e5e5e", 16, bold=False, align='justify', line_height=1.6, indent_px=16, margin_top=0, margin_bottom=0, background="#ffffff"
    ))
    parts.append('<br/>')

    parts.append(styled_paragraph(
        f"本期推送人民法院案例库编号为{data['case_number']}的参考案例",
        "#5e5e5e", 16, bold=False, align='justify', line_height=1.6, indent_px=16,  margin_top=8, margin_bottom=8
    ))

    parts.append(styled_paragraph(
        f'<span style="background-color:#5287b7;color:#ffffff;">延伸阅读</span>',
        "#5e5e5e", 16, bold=False, align='justify', line_height=1.6, indent_px=16, margin_top=0, margin_bottom=0, background="#ffffff"
    ))
    parts.append('<br/>')

    parts.append(styled_paragraph(
        data['case_name'], "#5287b7", 18, bold=True, align='left', line_height=1.6, indent_px=16, margin_top=0, margin_bottom=0
    ))

    parts.append(styled_paragraph(
        data['case_desc'], "#424242", 18, bold=True, align='left', line_height=2, indent_px=16, margin_top=0, margin_bottom=0
    ))
    parts.append('<br/>' + styled_paragraph(
        '【关键词】', "#5287b7", 16, bold=True, align='left', line_height=2, indent_px=16, margin_top=0, margin_bottom=0
    ) + '<br/>')

    parts.append(styled_paragraph(
        data['key_word'], "#5e5e5e", 16, bold=False, align='justify', line_height=2, indent_px=16, margin_top=0, margin_bottom=0
    ))

    parts.append('<br/>' + styled_paragraph(
        '【基本案情】', "#5287b7", 16, bold=True, align='left', line_height=2, indent_px=16, margin_top=0, margin_bottom=0
    ) + '<br/>')

    # 分段渲染：将case_text按空行分段，分别渲染为段落，并在每段后插入一个<br/>实现空一行
    case_text_paragraphs = [p for p in data['case_text'].split('\n\n') if p.strip()]
    for para in case_text_paragraphs:
        parts.append(styled_paragraph(
            para, "#5e5e5e", 16, bold=False, align='justify', line_height=2, indent_px=16, margin_top=0, margin_bottom=0
        ))
        parts.append('<br/>')  # 每段后空一行

    parts.append(styled_paragraph(
        '【裁判理由】', "#5287b7", 16, bold=True, align='left', line_height=2, indent_px=16, margin_top=0, margin_bottom=0
    ) + '<br/>')

    # 分段渲染：将trial_process按空行分段，分别渲染为段落
    trial_process_paragraphs = [p for p in data['trial_process'].split('\n\n') if p.strip()]
    for para in trial_process_paragraphs:
        parts.append(styled_paragraph(
            para, "#5e5e5e", 16, bold=False, align='justify', line_height=2, indent_px=16, margin_top=0, margin_bottom=0
        ))
        parts.append('<br/>')  # 每段后空一行

    parts.append(styled_paragraph(
        '【裁判要旨】', "#5287b7", 16, bold=True, align='left', line_height=2, indent_px=16, margin_top=0, margin_bottom=0
    ) + '<br/>')

    trial_abbr_paragraphs = [p for p in data['trial_abbr'].split('\n\n') if p.strip()]
    for para in trial_abbr_paragraphs:
        parts.append(styled_paragraph(
            para, "#5e5e5e", 16, bold=False, align='justify', line_height=2, indent_px=16, margin_top=0, margin_bottom=0
        ))

    parts.append('<br/>' + styled_paragraph(
        '【关联索引】', "#5287b7", 16, bold=True, align='left', line_height=2, indent_px=16, margin_top=0, margin_bottom=0
    ) + '<br/>')
    # 在此处处理 relevant_index 的换行和格式化
    relevant_index = data['relevant_index']
    if relevant_index:
        # 1. 第二个及之后的"《"前换行
        count = [0]
        def replace_bracket(match):
            count[0] += 1
            return '<br/>《' if count[0] > 1 else '《'
        relevant_index = re.sub(r'《', replace_bracket, relevant_index)

        # 2. "一审"和"二审"前换行（用<br/>以便HTML可见）
        relevant_index = re.sub(r'(一审)', r'<br/><br/>\1', relevant_index)
        relevant_index = re.sub(r'(二审)', r'<br/>\1', relevant_index)

        # 3. "本案例文本已于"前换行并空一行
        relevant_index = re.sub(r'(本案例文本已于)', r'<br/><br/>\1', relevant_index)

    relevant_index_paragraphs = [p for p in relevant_index.split('\n\n') if p.strip()]
    for para in relevant_index_paragraphs:
        parts.append(styled_paragraph(
            para, "#5e5e5e", 16, bold=False, align='justify', line_height=2, indent_px=16, margin_top=0, margin_bottom=0
        ))

    return "\n".join(parts)

def extract_text_from_pdf(path):
    doc = fitz.open(path)
    full_text = [page.get_text() for page in doc]
    return process_pdf_text_lines("\n".join(full_text))

def extract_text_from_docx(path):
    doc = Document(path)
    return "\n\n".join([para.text.strip() for para in doc.paragraphs if para.text.strip()])

def main():
    if len(sys.argv) < 2:
        print("用法: python script.py 文件路径")
        return

    import os
    path = sys.argv[1]
    base_name = os.path.splitext(os.path.basename(path))[0]
    text_dir = os.path.join(os.path.dirname(path), 'text')
    output_dir = os.path.join(os.path.dirname(path), 'output')
    os.makedirs(text_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    if path.endswith(".pdf"):
        raw_text = extract_text_from_pdf(path)
    elif path.endswith(".docx"):
        raw_text = extract_text_from_docx(path)
    else:
        print("仅支持PDF和DOCX文件")
        return

    cleaned = clean_text(raw_text)
    parsed = parse_fields(cleaned)

    text_file = os.path.join(text_dir, f"{base_name}-文本.txt")
    with open(text_file, "w", encoding="utf-8") as f:
        f.write(cleaned)

    html_output = generate_wechat_html(parsed)
    html_file = os.path.join(output_dir, f"{base_name}-公众号格式.html")
    with open(html_file, "w", encoding="utf-8") as f:
        f.write(html_output)

    print(f"✅ 提取完成，生成文件：\n - {text_file}\n - {html_file}")

if __name__ == "__main__":
    main()
