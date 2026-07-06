#!/usr/bin/env python3
"""师助AI - PPT生成 API (部署到 Railway / Render)"""
import os, json, io, base64, http.server
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

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
        lt = (idx - 1) % 4
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
            for i, item in enumerate(s.get('content',[])):
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
            for i, item in enumerate(s.get('content',[])):
                y = 1.8 + i * 1.4
                circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.8), Inches(y), Inches(0.7), Inches(0.7))
                circle.fill.solid(); circle.fill.fore_color.rgb = P; circle.line.fill.background()
                tf2 = circle.text_frame; tf2.word_wrap = False; p3 = tf2.paragraphs[0]; p3.text = str(i+1)
                p3.font.size = Pt(20); p3.font.bold = True; p3.font.color.rgb = W; p3.alignment = PP_ALIGN.CENTER
                tb3 = slide.shapes.add_textbox(Inches(2.0), Inches(y+0.05), Inches(10), Inches(0.6))
                p4 = tb3.text_frame.paragraphs[0]; p4.text = item; p4.font.size = Pt(18); p4.font.color.rgb = T

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
                        dp.paragraph_format.first_line_indent = Cm(0.74)
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
                    dp.paragraph_format.first_line_indent = Cm(0.74)
                    dp.paragraph_format.line_spacing = 1.3
                    dp.paragraph_format.space_after = Pt(6)
            except Exception:
                dp = doc.add_paragraph(para)
                dp.paragraph_format.line_spacing = 1.3
                dp.paragraph_format.space_after = Pt(6)
    
    buf = __import__('io').BytesIO()
    doc.save(buf)
    return buf.getvalue()


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
                
                if doc_type == 'doc':
                    print(f'Generating DOC: {body.get("title","")}')
                    docx_bytes = gen_docx(body)
                    b64 = base64.b64encode(docx_bytes).decode('ascii')
                    resp = json.dumps({'code': 0, 'data': b64, 'file': f"师助AI_{body.get('title','')}.docx"})
                    print(f'Done: {len(docx_bytes)} bytes')
                else:
                    print(f'Generating PPT: {body.get("title","")}')
                    pptx_bytes = gen(body)
                    b64 = base64.b64encode(pptx_bytes).decode('ascii')
                    resp = json.dumps({'code': 0, 'data': b64, 'file': f"师助AI_{body.get('title','PPT')}.pptx"})
                    print(f'Done: {len(pptx_bytes)} bytes')
                
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
