#!/usr/bin/env python3
"""师助AI - PPT生成 API (部署到 Railway / Render)"""
import os, json, io, base64, http.server, uuid, tempfile, urllib.request, urllib.error
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

DL_DIR = tempfile.mkdtemp(prefix="seat_dl_")

THEME = {
    'exam':    ('#1A237E', '#2B579A', '#E53935', '#E8EAF6'),
    'safety':  ('#BF360C', '#E65100', '#FF6F00', '#FFF3E0'),
    'holiday': ('#880E4F', '#AD1457', '#D81B60', '#FCE4EC'),
    'health':  ('#1B5E20', '#2E7D32', '#43A047', '#E8F5E9'),
    'plan':    ('#4A148C', '#6A1B9A', '#7B1FA2', '#F3E5F5'),
    'default': ('#283593', '#3949AB', '#4A6CF7', '#E8EAF6'),
}

def rgb(h):
    return RGBColor(int(h[1:3],16), int(h[3:5],16), int(h[5:7],16))

def detect(t):
    s = (t or '').lower()
    if '考试' in s or '期中' in s or '复习' in s: return 'exam'
    if '安全' in s or '消防' in s or '防溺水' in s: return 'safety'
    if '节日' in s or '元旦' in s or '国庆' in s: return 'holiday'
    if '健康' in s or '环保' in s or '运动' in s: return 'health'
    if '总结' in s or '计划' in s or '规划' in s: return 'plan'
    return 'default'

def gen(data):
    th = THEME.get(detect(data.get('title','')), THEME['default'])
    D = rgb(th[0]); P = rgb(th[1]); A = rgb(th[2]); L = rgb(th[3])
    W = RGBColor(0xFF,0xFF,0xFF); T = RGBColor(0x33,0x33,0x33)
    prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    slides = data.get('slides',[]); n = len(slides)

    for idx, s in enumerate(slides):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        if idx == 0:
            slide.background.fill.solid(); slide.background.fill.fore_color.rgb = D
            tb = slide.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(10), Inches(2))
            tf = tb.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]; p.text = data.get('title',''); p.font.size = Pt(52)
            p.font.bold = True; p.font.color.rgb = W; p.alignment = PP_ALIGN.CENTER
            if s.get('content'):
                p2 = tf.add_paragraph(); p2.text = s['content'][0] or ''
                p2.font.size = Pt(22); p2.font.color.rgb = A; p2.alignment = PP_ALIGN.CENTER; p2.space_before = Pt(12)
            line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.5), Inches(1.9), Inches(4), Pt(3))
            line.fill.solid(); line.fill.fore_color.rgb = A; line.line.fill.background()
            continue
        if idx == n - 1:
            slide.background.fill.solid(); slide.background.fill.fore_color.rgb = L
            tb = slide.shapes.add_textbox(Inches(2), Inches(2), Inches(9), Inches(2))
            tf = tb.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]; p.text = s.get('title',''); p.font.size = Pt(40)
            p.font.bold = True; p.font.color.rgb = D; p.alignment = PP_ALIGN.CENTER
            if s.get('content'):
                p2 = tf.add_paragraph(); p2.text = s['content'][0]; p2.font.size = Pt(22)
                p2.font.color.rgb = A; p2.alignment = PP_ALIGN.CENTER; p2.space_before = Pt(16)
            continue
        lt = (idx - 1) % 3 + 1  # 0=封面已用, 1~3=内容布局(跳过深色章节)
        if lt == 0:
            slide.background.fill.solid(); slide.background.fill.fore_color.rgb = D
            tb = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11), Inches(1.5))
            tf = tb.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]; p.text = s.get('title',''); p.font.size = Pt(40)
            p.font.bold = True; p.font.color.rgb = W; p.alignment = PP_ALIGN.CENTER
        elif lt == 1:
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(1.4))
            bar.fill.solid(); bar.fill.fore_color.rgb = D; bar.line.fill.background()
            tf = bar.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]; p.text = '  ' + s.get('title',''); p.font.size = Pt(32); p.font.bold = True; p.font.color.rgb = W
            for i, item in enumerate(s.get('content',[])[:6]):
                tb2 = slide.shapes.add_textbox(Inches(1.2), Inches(2.0+i*0.6), Inches(11), Inches(0.55))
                p2 = tb2.text_frame.paragraphs[0]; p2.text = '▪  ' + item; p2.font.size = Pt(18); p2.font.color.rgb = T
        elif lt == 2:
            items = s.get('content',[]); mid = (len(items)+1)//2
            tb2 = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11), Inches(0.9))
            p2 = tb2.text_frame.paragraphs[0]; p2.text = s.get('title',''); p2.font.size = Pt(30); p2.font.bold = True; p2.font.color.rgb = P
            for i, item in enumerate(items[:mid]):
                card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.6+i*1.0), Inches(6), Inches(0.8))
                card.fill.solid(); card.fill.fore_color.rgb = L; card.line.fill.background()
                tf2 = card.text_frame; tf2.word_wrap = True; p3 = tf2.paragraphs[0]; p3.text = item
                p3.font.size = Pt(16); p3.font.color.rgb = D; p3.font.bold = True; p3.alignment = PP_ALIGN.CENTER
            for i, item in enumerate(items[mid:]):
                card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.6+i*1.0), Inches(6), Inches(0.8))
                card.fill.solid(); card.fill.fore_color.rgb = L; card.line.fill.background()
                tf2 = card.text_frame; tf2.word_wrap = True; p3 = tf2.paragraphs[0]; p3.text = item
                p3.font.size = Pt(16); p3.font.color.rgb = D; p3.font.bold = True; p3.alignment = PP_ALIGN.CENTER
        else:
            tb2 = slide.shapes.add_textbox(Inches(0.8), Inches(0.3), Inches(11), Inches(0.9))
            p2 = tb2.text_frame.paragraphs[0]; p2.text = s.get('title',''); p2.font.size = Pt(30); p2.font.bold = True; p2.font.color.rgb = P
            line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.2), Inches(11.5), Pt(4))
            line.fill.solid(); line.fill.fore_color.rgb = A; line.line.fill.background()
            for i, item in enumerate(s.get('content',[])[:6]):
                y = 1.8 + i * 1.4
                circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.8), Inches(y), Inches(0.7), Inches(0.7))
                circle.fill.solid(); circle.fill.fore_color.rgb = P; circle.line.fill.background()
                tf2 = circle.text_frame; tf2.word_wrap = False; p3 = tf2.paragraphs[0]; p3.text = str(i+1)
                p3.font.size = Pt(20); p3.font.bold = True; p3.font.color.rgb = W; p3.alignment = PP_ALIGN.CENTER
                tb3 = slide.shapes.add_textbox(Inches(2.0), Inches(y+0.05), Inches(10), Inches(0.6))
                p4 = tb3.text_frame.paragraphs[0]; p4.text = item; p4.font.size = Pt(18); p4.font.color.rgb = T

    # Process [插图：xxx] markers - replace with image placeholders
    img_key = os.environ.get('AI_API_KEY', '')
    if slides and img_key:
        for idx, s in enumerate(slides):
            for ci, item in enumerate(s.get('content', [])):
                import re
                m = re.search(r'\[\u63d2\u56fe\uff1a(.+?)\]', str(item))
                if m:
                    query = m.group(1).strip()
                    s['content'][ci] = item.replace(m.group(0), '').strip()
                    try:
                        # Search for image using 通义千问
                        import urllib.request, urllib.error
                        search_q = f'搜索一张关于{query}的图片，返回一个可以直接访问的图片URL'
                        sd = json.dumps({
                            'model': 'qwen-turbo',
                            'messages': [{'role': 'user', 'content': search_q}],
                            'enable_search': True
                        }).encode()
                        sq = urllib.request.Request(
                            'https://ws-5ol6m5p8f4hikz1a.cn-beijing.maas.aliyuncs.com/compatible-mode/v1/chat/completions',
                            data=sd,
                            headers={'Authorization': 'Bearer ' + img_key, 'Content-Type': 'application/json'}
                        )
                        sr = urllib.request.urlopen(sq, timeout=30)
                        sr_result = json.loads(sr.read().decode())
                        img_url = ''
                        for word in sr_result['choices'][0]['message']['content'].split():
                            w = word.strip('.,;:!?"\'()[]<>')
                            if w.startswith('http') and any(w.endswith(e) for e in ['.jpg','.jpeg','.png','.gif']):
                                img_url = w
                                break
                        if img_url:
                            urllib.request.urlretrieve(img_url, '/tmp/ppt_img_' + str(idx) + '.jpg')
                            slide = prs.slides[idx]
                            slide.shapes.add_picture('/tmp/ppt_img_' + str(idx) + '.jpg', Inches(8.5), Inches(1.5), Inches(4), Inches(3))
                    except:
                        # Placeholder if image search fails
                        try:
                            slide = prs.slides[idx]
                            shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.5), Inches(1.5), Inches(4), Inches(3))
                            shape.fill.solid(); shape.fill.fore_color.rgb = L; shape.line.fill.background()
                            tf = shape.text_frame; tf.word_wrap = True
                            p = tf.paragraphs[0]; p.text = query; p.font.size = Pt(14)
                            p.font.color.rgb = D; p.alignment = PP_ALIGN.CENTER
                        except: pass
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

def gen_docx(data):
    """生成Word文档(python-docx)，带【】小标题识别、字体、字号、缩进"""
    from docx import Document
    from docx.shared import Pt, Cm, Emu
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    
    doc = Document()
    
    # 默认样式：宋体 小四 1.3倍行距
    style = doc.styles['Normal']
    style.font.name = 'SimSun'
    style.font.size = Pt(12)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    pf = style.paragraph_format
    pf.line_spacing = 1.3
    
    title = data.get('title', '文档')
    doc_title = data.get('docTitle', '') or title
    doc_subtitle = data.get('docSubtitle', '')
    
    # 大标题：小二 居中 加粗（纯文本，无蓝色主题）
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = tp.add_run(doc_title)
    tr.bold = True
    tr.font.size = Pt(18)
    tr.font.name = 'SimSun'
    tr.font.color.rgb = None
    tr.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    
    if doc_subtitle:
        sp = doc.add_paragraph()
        sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sr = sp.add_run(doc_subtitle)
        sr.bold = True
        sr.font.size = Pt(14)
        sr.font.name = 'SimSun'
        sr.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    
    content = data.get('content', '')
    subtitle_counter = 0
    
    if content:
        for para in content.split('\n'):
            para = para.strip()
            if not para:
                continue
            try:
                # 【小标题】：加粗大号、自动编号、不缩进
                import re
                bracket_match = re.search(r'【(.+?)】', para)
                if bracket_match:
                    sub_text = bracket_match.group(1)
                    subtitle_counter += 1
                    if not re.match(r'^[一二三四五六七八九十\d]', sub_text):
                        sub_text = str(subtitle_counter) + '. ' + sub_text
                    # 去掉【】部分
                    para = re.sub(r'【.+?】', '', para).strip()
                    sp = doc.add_paragraph()
                    sp.paragraph_format.space_before = Pt(12)
                    sp.paragraph_format.space_after = Pt(6)
                    sp.paragraph_format.line_spacing = 1.3
                    sr = sp.add_run(sub_text)
                    sr.bold = True
                    sr.font.size = Pt(15)
                    sr.font.name = 'SimSun'
                    sr.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                    if para:
                        dp = doc.add_paragraph(para)
                        dp.paragraph_format.first_line_indent = Pt(24)
                        dp.paragraph_format.line_spacing = 1.3
                        dp.paragraph_format.space_after = Pt(6)
                    continue
                
                if para.startswith('# ') or para.startswith('## '):
                    hp = doc.add_paragraph()
                    hp.paragraph_format.line_spacing = 1.3
                    hr = hp.add_run(para[para.index(' ')+1:])
                    hr.bold = True
                    hr.font.size = Pt(15)
                    hr.font.name = 'SimSun'
                    hr.font.color.rgb = None
                    hr.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                elif para.startswith('**') and para.endswith('**'):
                    d = doc.add_paragraph()
                    d.paragraph_format.line_spacing = 1.3
                    r = d.add_run(para[2:-2])
                    r.bold = True
                    r.font.size = Pt(12)
                    r.font.name = 'SimSun'
                    r.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                elif para.startswith('- ') or para.startswith('* '):
                    li = doc.add_paragraph(para[2:], style='List Bullet')
                    li.paragraph_format.line_spacing = 1.3
                else:
                    dp = doc.add_paragraph(para)
                    dp.paragraph_format.first_line_indent = Pt(24)
                    dp.paragraph_format.line_spacing = 1.3
                    dp.paragraph_format.space_after = Pt(6)
            except Exception:
                dp = doc.add_paragraph(para)
                dp.paragraph_format.line_spacing = 1.3
                dp.paragraph_format.space_after = Pt(6)
    
    buf = __import__('io').BytesIO()
    doc.save(buf)
    return buf.getvalue()

def parse_file(data):
    """解析文档/PDF/Excel，返回文本内容"""
    import tempfile, os
    content = data.get('content', '')  # base64 encoded file
    filename = data.get('filename', '')
    if not content:
        return '（无文件内容）'
    
    import base64
    raw = base64.b64decode(content)
    ext = filename.lower().split('.')[-1] if '.' in filename else ''
    
    try:
        if ext in ('txt', 'csv', 'json', 'md', 'xml'):
            return raw.decode('utf-8')
        elif ext in ('jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp'):
            api_key = os.environ.get('AI_API_KEY', '')
            if not api_key:
                return f'\u3010\u56fe\u7247\uff1a{filename}\u3011\n\uff08\u56fe\u7247\u6587\u5b57\u8bc6\u522b\u672a\u914d\u7f6e\uff0c\u9700\u8bbe\u7f6eAI_API_KEY\uff09'
            import urllib.request, urllib.error, base64
            img_b64 = base64.b64encode(raw).decode('ascii')
            ocr_data = json.dumps({
                'model': 'Qwen3-VL-Flash',
                'messages': [{'role': 'user', 'content': [
                    {'type': 'image_url', 'image_url': {'url': f'data:image/jpeg;base64,{img_b64}'}},
                    {'type': 'text', 'text': '\u8bf7\u63d0\u53d6\u8fd9\u5f20\u56fe\u7247\u4e2d\u7684\u6240\u6709\u6587\u5b57\u5185\u5bb9\uff0c\u76f4\u63a5\u8f93\u51fa\u6587\u5b57\uff0c\u4e0d\u8981\u989d\u5916\u8bf4\u660e'}
                ]}]
            }).encode()
            req = urllib.request.Request(
                'https://ws-5ol6m5p8f4hikz1a.cn-beijing.maas.aliyuncs.com/compatible-mode/v1/chat/completions',
                data=ocr_data,
                headers={'Authorization': 'Bearer ' + api_key, 'Content-Type': 'application/json'}
            )
            try:
                r = urllib.request.urlopen(req, timeout=60)
                result = json.loads(r.read().decode())
                text = result['choices'][0]['message']['content']
                return f'\u3010\u56fe\u7247\uff1a{filename}\u3011\n{text}'
            except Exception as e:
                return f'\u3010\u56fe\u7247\uff1a{filename}\u3011\n\uff08OCR\u8bc6\u522b\u5931\u8d25\uff1a{str(e)[:200]}\uff09'
        elif ext == 'pdf':
            import pdfplumber
            import io
            text = ''
            with pdfplumber.open(io.BytesIO(raw)) as pdf:
                for page in pdf.pages:
                    text += page.extract_text() or ''
            return text[:50000]
        elif ext in ('docx', 'doc'):
            import docx
            import io
            doc = docx.Document(io.BytesIO(raw))
            return '\n'.join([p.text for p in doc.paragraphs])[:50000]
        elif ext in ('pptx', 'ppt'):
            from pptx import Presentation
            import io
            prs = Presentation(io.BytesIO(raw))
            text = ''
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, 'text') and shape.text:
                        text += shape.text + '\n'
            return text[:50000] if text else '（PPT文件无可提取文字）'
        elif ext in ('xlsx', 'xls'):
            import openpyxl
            import io
            wb = openpyxl.load_workbook(io.BytesIO(raw))
            rows = []
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                rows.append(f'【工作表：{sheet}】')
                for row in ws.iter_rows(values_only=True):
                    rows.append('\t'.join([str(c) if c is not None else '' for c in row]))
            return '\n'.join(rows)[:50000]
        else:
            return f'不支持的文件格式：.{ext}'
    except Exception as e:
        return f'文件解析失败：{str(e)}'

def edit_excel(data):
    """根据AI指令修改Excel文件"""
    import openpyxl, re
    file_b64 = data.get('file', '')
    instructions = data.get('instructions', '')
    if not file_b64:
        return None, '未提供Excel文件'
    
    try:
        raw = base64.b64decode(file_b64)
        wb = openpyxl.load_workbook(io.BytesIO(raw))
    except:
        wb = openpyxl.Workbook()
    
    ws = wb.active
    
    # 解析简单指令：修改单元格
    # 格式："设置 A1=张三" 或 "修改 B2=95"
    for line in instructions.split('\n'):
        line = line.strip()
        m = re.match(r'设置\s*([A-Z]+)(\d+)\s*=\s*(.+)', line)
        if m:
            cell = m.group(1) + m.group(2)
            ws[cell] = m.group(3).strip()
    
    buf = io.BytesIO()
    wb.save(buf)
    return base64.b64encode(buf.getvalue()).decode('ascii'), None

# === HTTP Server ===
def make_handler():
    class Handler(http.server.BaseHTTPRequestHandler):
        def _set_cors(self):
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
            self.send_header('Access-Control-Allow-Headers', 'Content-Type')

        def do_OPTIONS(self):
            self.send_response(204)
            self._set_cors()
            self.end_headers()

        def do_GET(self):
            if self.path.startswith('/dl/'):
                file_id = self.path.split('/')[-1]
                for f in os.listdir(DL_DIR):
                    if f.endswith(file_id):
                        fpath = os.path.join(DL_DIR, f)
                        try:
                            with open(fpath, 'rb') as fh:
                                data = fh.read()
                            os.remove(fpath)
                            self.send_response(200)
                            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                            self.send_header('Content-Disposition', 'attachment; filename="seat.xlsx"')
                            self._set_cors()
                            self.end_headers()
                            self.wfile.write(data)
                            return
                        except: pass
                self.send_response(404)
                self._set_cors()
                self.end_headers()
                return
            self.send_response(200)
            self.send_header('Content-Type', 'text/plain')
            self._set_cors()
            self.end_headers()
            self.wfile.write(('师助AI PPT API is running. POST JSON to /gen').encode())

        def do_POST(self):
            try:
                length = int(self.headers.get('Content-Length', 0))
                body = json.loads(self.rfile.read(length))
                doc_type = body.get('type', 'ppt')
                
                if doc_type == 'create-xlsx':
                    import openpyxl
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = '学生名单'
                    names = body.get('names', [])
                    ws.cell(1, 1, '序号')
                    ws.cell(1, 2, '姓名')
                    for i, name in enumerate(names):
                        ws.cell(i+2, 1, i+1)
                        ws.cell(i+2, 2, name)
                    buf = io.BytesIO()
                    wb.save(buf)
                    # Save to temp file and return download URL
                    fid = str(uuid.uuid4())
                    fpath = os.path.join(DL_DIR, fid)
                    with open(fpath, 'wb') as fh:
                        fh.write(buf.getvalue())
                    resp = json.dumps({'code': 0, 'url': '/dl/' + fid})
                    print(f'Created xlsx: {len(names)} names, url=/dl/{fid}')
                elif doc_type == 'seat-export':
                    import openpyxl
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = '座位表'
                    grid = body.get('grid', [])
                    col_types = body.get('cols', [])
                    ws.cell(1, 1, '排')
                    for ci, ct in enumerate(col_types):
                        if ct == 'seat':
                            ws.cell(1, ci+2, f'列{ci+1}')
                        else:
                            ws.cell(1, ci+2, '过道')
                    for ri, row in enumerate(grid):
                        ws.cell(ri+2, 1, f'第{ri+1}排')
                        for ci, cell in enumerate(row):
                            if isinstance(cell, dict) and cell.get('student'):
                                ws.cell(ri+2, ci+2, cell['student'])
                    buf = io.BytesIO()
                    wb.save(buf)
                    # Save to temp file and return download URL
                    fid = str(uuid.uuid4())
                    fpath = os.path.join(DL_DIR, fid)
                    with open(fpath, 'wb') as fh:
                        fh.write(buf.getvalue())
                    resp = json.dumps({'code': 0, 'url': '/dl/' + fid})
                    print(f'Seat export: {len(grid)} rows, url=/dl/{fid}')
                elif doc_type == 'doc':
                    print(f'Generating DOC: {body.get("title","")}')
                    docx_bytes = gen_docx(body)
                    b64 = base64.b64encode(docx_bytes).decode('ascii')
                    resp = json.dumps({'code': 0, 'data': b64, 'file': f"师助AI_{body.get('title','')}.docx"})
                    print(f'Done: {len(docx_bytes)} bytes')
                elif doc_type == 'parse':
                    text = parse_file(body)
                    resp = json.dumps({'code': 0, 'data': text})
                elif doc_type == 'edit-excel':
                    b64_data, err = edit_excel(body)
                    if err:
                        resp = json.dumps({'code': -1, 'error': err})
                    else:
                        resp = json.dumps({'code': 0, 'data': b64_data, 'file': 'edited.xlsx'})
                elif doc_type == 'chat':
                    import urllib.request, urllib.error
                    # Use custom api_key from request if provided, otherwise env var
                    api_key = body.get('api_key', '') or os.environ.get('DEEPSEEK_API_KEY', '')
                    if not api_key:
                        resp = json.dumps({'code': -1, 'error': '服务器未配置DeepSeek API Key，请在Railway环境变量中设置 DEEPSEEK_API_KEY'})
                    else:
                        messages = body.get('messages', [])
                        model = body.get('model', 'deepseek-chat')
                        temp = body.get('temperature', 0.7)
                        max_tok = body.get('max_tokens', 4096)
                        post_data = json.dumps({'model': model, 'messages': messages, 'temperature': temp, 'max_tokens': max_tok, 'stream': False}).encode()
                        req = urllib.request.Request(
                            'https://api.deepseek.com/v1/chat/completions',
                            data=post_data,
                            headers={'Authorization': 'Bearer ' + api_key, 'Content-Type': 'application/json'}
                        )
                        try:
                            r = urllib.request.urlopen(req, timeout=120)
                            result = json.loads(r.read().decode())
                            reply = result['choices'][0]['message']['content']
                            resp = json.dumps({'code': 0, 'data': reply})
                            print(f'Chat reply: {len(reply)} chars')
                        except urllib.error.HTTPError as e:
                            resp = json.dumps({'code': -1, 'error': 'API错误: ' + e.read().decode()[:200]})
                        except Exception as e:
                            resp = json.dumps({'code': -1, 'error': str(e)[:200]})
                elif doc_type == 'search':
                    api_key = os.environ.get('AI_API_KEY', '')
                    if not api_key:
                        resp = json.dumps({'code': -1, 'error': '联网搜索功能未配置（需在Railway设置 AI_API_KEY）'})
                    else:
                        import urllib.request, urllib.error
                        query = body.get('query', '')
                        search_data = json.dumps({
                            'model': 'qwen-turbo',
                            'messages': [{'role': 'user', 'content': query}],
                            'enable_search': True
                        }).encode()
                        req = urllib.request.Request(
                            'https://ws-5ol6m5p8f4hikz1a.cn-beijing.maas.aliyuncs.com/compatible-mode/v1/chat/completions',
                            data=search_data,
                            headers={'Authorization': 'Bearer ' + api_key, 'Content-Type': 'application/json'}
                        )
                        try:
                            r = urllib.request.urlopen(req, timeout=60)
                            result = json.loads(r.read().decode())
                            reply = result['choices'][0]['message']['content']
                            resp = json.dumps({'code': 0, 'data': reply})
                        except Exception as e:
                            resp = json.dumps({'code': -1, 'error': '搜索失败: ' + str(e)[:200]})
                elif doc_type == 'vision':
                    api_key = os.environ.get('AI_API_KEY', '')
                    if not api_key:
                        resp = json.dumps({'code': -1, 'error': '图片识别功能未配置（需在Railway设置 AI_API_KEY）'})
                    else:
                        import urllib.request, urllib.error
                        image_b64 = body.get('image', '')
                        prompt = body.get('prompt', '请描述这张图片')
                        vision_data = json.dumps({
                            'model': 'Qwen3-VL-Flash',
                            'messages': [{'role': 'user', 'content': [
                                {'type': 'image_url', 'image_url': {'url': f'data:image/jpeg;base64,{image_b64}'}},
                                {'type': 'text', 'text': prompt}
                            ]}]
                        }).encode()
                        req = urllib.request.Request(
                            'https://ws-5ol6m5p8f4hikz1a.cn-beijing.maas.aliyuncs.com/compatible-mode/v1/chat/completions',
                            data=vision_data,
                            headers={'Authorization': 'Bearer ' + api_key, 'Content-Type': 'application/json'}
                        )
                        try:
                            r = urllib.request.urlopen(req, timeout=60)
                            result = json.loads(r.read().decode())
                            reply = result['choices'][0]['message']['content']
                            resp = json.dumps({'code': 0, 'data': reply})
                        except Exception as e:
                            resp = json.dumps({'code': -1, 'error': '识别失败: ' + str(e)[:200]})
                elif doc_type == 'generate-image':
                    api_key = os.environ.get('AI_API_KEY', '')
                    if not api_key:
                        resp = json.dumps({'code': -1, 'error': 'AI生图功能未配置（需在Railway设置 AI_API_KEY）'})
                    else:
                        import urllib.request, urllib.error, time
                        prompt = body.get('prompt', '')
                        img_data = json.dumps({
                            'model': 'Z-Image-Turbo',
                            'input': {'prompt': prompt},
                            'parameters': {'size': '1024*1024', 'n': 1}
                        }).encode()
                        req = urllib.request.Request(
                            'https://dashscope.aliyuncs.com/api/v1/services/aigc/text2image/image-synthesis',
                            data=img_data,
                            headers={'Authorization': 'Bearer ' + api_key, 'Content-Type': 'application/json'}
                        )
                        try:
                            r = urllib.request.urlopen(req, timeout=120)
                            result = json.loads(r.read().decode())
                            task_id = result.get('output', {}).get('task_id', '')
                            if task_id:
                                # Poll for result
                                for _ in range(30):
                                    time.sleep(2)
                                    poll_req = urllib.request.Request(
                                        f'https://dashscope.aliyuncs.com/api/v1/tasks/{task_id}',
                                        headers={'Authorization': 'Bearer ' + api_key}
                                    )
                                    poll_r = urllib.request.urlopen(poll_req, timeout=30)
                                    poll_result = json.loads(poll_r.read().decode())
                                    status = poll_result.get('output', {}).get('task_status', '')
                                    if status == 'SUCCEEDED':
                                        img_url = poll_result.get('output', {}).get('results', [{}])[0].get('url', '')
                                        resp = json.dumps({'code': 0, 'data': img_url, 'format': 'url'})
                                        break
                                    elif status in ('FAILED', 'CANCELED'):
                                        resp = json.dumps({'code': -1, 'error': '生图失败'})
                                        break
                                else:
                                    resp = json.dumps({'code': -1, 'error': '生图超时'})
                            else:
                                resp = json.dumps({'code': -1, 'error': '提交生图任务失败'})
                        except Exception as e:
                            resp = json.dumps({'code': -1, 'error': '生图失败: ' + str(e)[:200]})
                elif doc_type == 'handwriting':
                    api_key = os.environ.get('AI_API_KEY', '')
                    if not api_key:
                        resp = json.dumps({'code': -1, 'error': '\u9700\u8981\u914d\u7f6eAI_API_KEY'})
                    else:
                        import urllib.request, urllib.error, base64 as b64mod
                        image_b64 = body.get('image', '')
                        # Step 1: OCR using qwen-vl-max
                        ocr_data = json.dumps({
                            'model': 'Qwen3-VL-Flash',
                            'messages': [{'role': 'user', 'content': [
                                {'type': 'image_url', 'image_url': {'url': f'data:image/jpeg;base64,{image_b64}'}},
                                {'type': 'text', 'text': '\u8bf7\u8bc6\u522b\u8fd9\u5f20\u56fe\u7247\u4e2d\u7684\u6240\u6709\u6587\u5b57\u3002\u5982\u679c\u662f\u8868\u683c\u6570\u636e\uff0c\u7528\u5236\u8868\u7b26\t\u5206\u9694\u5217\uff0c\u6bcf\u884c\u4e00\u6761\u8bb0\u5f55\u3002\u5982\u679c\u662f\u6bb5\u843d\u6587\u5b57\uff0c\u76f4\u63a5\u8f93\u51fa\u6587\u5b57\u5185\u5bb9\u3002'}
                            ]}]
                        }).encode()
                        req = urllib.request.Request(
                            'https://ws-5ol6m5p8f4hikz1a.cn-beijing.maas.aliyuncs.com/compatible-mode/v1/chat/completions',
                            data=ocr_data,
                            headers={'Authorization': 'Bearer ' + api_key, 'Content-Type': 'application/json'}
                        )
                        try:
                            r = urllib.request.urlopen(req, timeout=60)
                            result = json.loads(r.read().decode())
                            text = result['choices'][0]['message']['content']
                            # Step 2: Detect if table or text
                            detect_data = json.dumps({
                                'model': 'qwen-turbo',
                                'messages': [{'role': 'user', 'content': f'\u5224\u65ad\u4ee5\u4e0b\u5185\u5bb9\u662f\u8868\u683c\u6570\u636e\u8fd8\u662f\u6bb5\u843d\u6587\u5b57\uff0c\u53ea\u56de\u7b54"table"\u6216"text"\uff1a\n{text[:1500]}'}],
                                'temperature': 0.1
                            }).encode()
                            detect_req = urllib.request.Request(
                                'https://ws-5ol6m5p8f4hikz1a.cn-beijing.maas.aliyuncs.com/compatible-mode/v1/chat/completions',
                                data=detect_data,
                                headers={'Authorization': 'Bearer ' + api_key, 'Content-Type': 'application/json'}
                            )
                            detect_r = urllib.request.urlopen(detect_req, timeout=30)
                            detect_result = json.loads(detect_r.read().decode())
                            doc_type_detected = detect_result['choices'][0]['message']['content'].strip().lower()
                            
                            if 'table' in doc_type_detected:
                                # Generate Excel
                                import openpyxl
                                wb = openpyxl.Workbook()
                                ws = wb.active
                                ws.title = '\u6570\u636e'
                                lines = text.strip().split('\n')
                                for ri, line in enumerate(lines):
                                    cells = line.split('\t')
                                    for ci, cell in enumerate(cells):
                                        ws.cell(ri+1, ci+1, cell.strip())
                                buf = io.BytesIO()
                                wb.save(buf)
                                fid = str(uuid.uuid4())
                                fpath = os.path.join(DL_DIR, fid)
                                with open(fpath, 'wb') as fh: fh.write(buf.getvalue())
                                resp = json.dumps({'code': 0, 'url': '/dl/' + fid, 'detected': 'table', 'ocrText': text[:5000]})
                                print(f'Handwriting: detected as table, {len(lines)} rows')
                            else:
                                # Generate Word
                                w_data = {'title': '\u624b\u5199\u6587\u6863', 'content': text}
                                docx_bytes = gen_docx(w_data)
                                fid = str(uuid.uuid4())
                                fpath = os.path.join(DL_DIR, fid)
                                with open(fpath, 'wb') as fh: fh.write(docx_bytes)
                                resp = json.dumps({'code': 0, 'url': '/dl/' + fid, 'detected': 'text', 'ocrText': text[:5000]})
                                print(f'Handwriting: detected as text, {len(text)} chars')
                        except urllib.error.HTTPError as e:
                            resp = json.dumps({'code': -1, 'error': 'API\u9519\u8bef: ' + e.read().decode()[:200]})
                        except Exception as e:
                            resp = json.dumps({'code': -1, 'error': str(e)[:200]})
                else:
                    print(f'Generating PPT: {body.get("title","")}')
                    pptx_bytes = gen(body)
                    fid = str(uuid.uuid4())
                    fpath = os.path.join(DL_DIR, fid)
                    with open(fpath, 'wb') as fh:
                        fh.write(pptx_bytes)
                    resp = json.dumps({'code': 0, 'url': '/dl/' + fid})
                    print(f'Done: {len(pptx_bytes)} bytes, url=/dl/{fid}')
                
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self._set_cors()
                self.end_headers()
                self.wfile.write(resp.encode())
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self._set_cors()
                self.end_headers()
                self.wfile.write(json.dumps({'code': -1, 'error': str(e)}).encode())

        def log_message(self, fmt, *args):
            print(f'[PPT API] {args[0] if args else fmt}')
    return Handler

if __name__ == '__main__':
    PORT = int(os.environ.get('PORT', 8080))
    server = http.server.HTTPServer(('0.0.0.0', PORT), make_handler())
    print(f'师助AI PPT API running on port {PORT}')
    server.serve_forever()
