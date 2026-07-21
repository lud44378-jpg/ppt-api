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

def generate_image_url(prompt, api_key):
    """调用阿里 Z-Image 同步文生图接口，返回临时图片 URL。"""
    image_prompt = '面向中国中小学课堂的教育教学PPT插图，人物和场景符合中国校园语境，不要正面对镜头摆拍；画面绝对不要任何文字、汉字、数字、字母、标志、海报或水印；构图简洁、主体清晰：' + prompt
    payload = json.dumps({
        'model': 'z-image-turbo',
        'input': {'messages': [{'role': 'user', 'content': [{'text': image_prompt}]}]}
    }).encode()
    req = urllib.request.Request(
        'https://ws-5ol6m5p8f4hikz1a.cn-beijing.maas.aliyuncs.com/api/v1/services/aigc/multimodal-generation/generation',
        data=payload,
        headers={'Authorization': 'Bearer ' + api_key, 'Content-Type': 'application/json'}
    )
    result = json.loads(urllib.request.urlopen(req, timeout=120).read().decode('utf-8'))
    if result.get('code'):
        raise RuntimeError('阿里生图失败：' + result.get('message', result['code']))
    content = result.get('output', {}).get('choices', [{}])[0].get('message', {}).get('content', [])
    for item in content:
        if isinstance(item, dict) and item.get('image'):
            return item['image']
    raise RuntimeError('阿里生图未返回图片地址')

def search_image_url(query, api_key, timeout=120):
    """使用阿里文搜图获取真实图片地址，适用于人物、事件、器材等事实性页面。"""
    payload = json.dumps({
        'model': 'qwen3.6-flash',
        'input': '请搜索一张可用于教学PPT的真实图片：' + query,
        'tools': [{'type': 'web_search_image'}],
        'store': False
    }).encode()
    req = urllib.request.Request(
        'https://ws-5ol6m5p8f4hikz1a.cn-beijing.maas.aliyuncs.com/compatible-mode/v1/responses',
        data=payload,
        headers={'Authorization': 'Bearer ' + api_key, 'Content-Type': 'application/json'}
    )
    response = json.loads(urllib.request.urlopen(req, timeout=timeout).read().decode('utf-8'))
    for item in response.get('output', []):
        if item.get('type') != 'web_search_image_call':
            continue
        try:
            images = json.loads(item.get('output', '[]'))
            if images and images[0].get('url'):
                return images[0]['url']
        except (TypeError, ValueError):
            continue
    raise RuntimeError('联网搜图未返回可下载图片')

def build_ppt_research(topic, api_key):
    """阿里先完成事实检索、叙事建议和图片素材检索，供后续故事板使用。"""
    import re
    prompt = '''你是中国中小学教师的备课研究员。请联网研究主题：%s。
只输出 JSON，不要 Markdown，格式必须为：
{
 "facts":["3-5条可核实的事实、案例或规范，每条含简短来源名"],
 "sources":["2-4个来源名称或机构"],
 "teaching_arc":["适合本主题的4-6步课堂叙事结构"],
 "image_queries":[{"query":"一条具体且适合教学PPT的中文搜图关键词","caption":"这张图用于说明什么"}]
}
不得编造姓名、数字、事故或来源。image_queries 最多3条；只给确有必要的真实事件、物件、地点或现象，不要泛泛搜索人物摆拍、课堂合影、游泳池合影或新闻网页截图。''' % topic
    payload = json.dumps({
        'model': 'qwen-plus',
        'messages': [{'role': 'user', 'content': prompt}],
        'enable_search': True,
        'search_options': {'forced_search': True},
        'temperature': 0.2,
    }).encode()
    req = urllib.request.Request(
        'https://ws-5ol6m5p8f4hikz1a.cn-beijing.maas.aliyuncs.com/compatible-mode/v1/chat/completions',
        data=payload,
        headers={'Authorization': 'Bearer ' + api_key, 'Content-Type': 'application/json'}
    )
    result = json.loads(urllib.request.urlopen(req, timeout=120).read().decode('utf-8'))
    raw = result['choices'][0]['message']['content']
    match = re.search(r'\{[\s\S]*\}', raw)
    if not match:
        raise RuntimeError('联网研究未返回结构化资料')
    try:
        pack = json.loads(match.group(0))
    except ValueError as err:
        raise RuntimeError('联网研究资料格式错误') from err
    candidates = [
        item for item in (pack.get('image_queries') or [])[:3]
        if isinstance(item, dict) and str(item.get('query', '')).strip()
    ]
    # 三张图片并行检索，不能让素材预取的耗时累加到数分钟。
    from concurrent.futures import ThreadPoolExecutor, as_completed
    assets = []
    def fetch_asset(item):
        query = str(item['query']).strip()
        return {
            'query': query,
            'caption': str(item.get('caption', query)).strip(),
            'url': search_image_url(query, api_key, timeout=45),
        }
    with ThreadPoolExecutor(max_workers=min(3, len(candidates) or 1)) as pool:
        futures = [pool.submit(fetch_asset, item) for item in candidates]
        for future in as_completed(futures):
            try:
                asset = future.result()
                asset['id'] = 'A' + str(len(assets) + 1)
                assets.append(asset)
            except Exception as err:
                print('[PPT API] research image unavailable:', str(err)[:120])
    return {
        'facts': (pack.get('facts') or [])[:5],
        'sources': (pack.get('sources') or [])[:4],
        'teaching_arc': (pack.get('teaching_arc') or [])[:6],
        'assets': assets,
    }

def fit_slides(raw_slides):
    """把过长要点拆成续页，确保每页最多 4 条、每条最多约 42 个字符。"""
    if len(raw_slides) <= 2:
        return raw_slides
    fitted = [raw_slides[0]]
    for source in raw_slides[1:-1]:
        kind = str(source.get('type', '')).lower()
        items = []
        for item in source.get('content', []):
            item = str(item).strip()
            if not item:
                continue
            # 案例页允许一段完整叙事，不把一个事件机械切成多条短句。
            while kind != 'case' and len(item) > 42:
                cut = max(item.rfind('，', 0, 42), item.rfind('。', 0, 42), item.rfind('；', 0, 42), item.rfind('、', 0, 42))
                cut = cut + 1 if cut >= 18 else 42
                items.append(item[:cut])
                item = item[cut:].lstrip('，。；、 ')
            if item:
                items.append(item)
        # 互动和步骤页需要留白；事实、案例、做法页可以承载更多信息。
        max_items = {
            'scenario': 3, 'steps': 6, 'compare': 6,
            'fact': 6, 'case': 6, 'action': 6, 'sources': 4,
        }.get(kind, 5)
        chunks = [items[i:i + max_items] for i in range(0, len(items), max_items)] or [[]]
        for chunk_index, chunk in enumerate(chunks):
            slide = dict(source)
            slide['content'] = chunk
            if chunk_index:
                slide['title'] = str(source.get('title', '')) + '（续）'
            fitted.append(slide)
    fitted.append(raw_slides[-1])
    return fitted

def gen(data):
    th = THEME.get(data.get('theme') or detect(data.get('title','')), THEME['default'])
    D = rgb(th[0]); P = rgb(th[1]); A = rgb(th[2]); L = rgb(th[3])
    W = RGBColor(0xFF,0xFF,0xFF); T = RGBColor(0x33,0x33,0x33)
    prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    slides = [dict(s, content=list(s.get('content', []))) for s in data.get('slides', [])]
    asset_map = {
        str(asset.get('id')): asset for asset in data.get('assets', [])
        if isinstance(asset, dict) and asset.get('id') and asset.get('url')
    }
    # 先移除插图标记，避免它作为文字渲染到 PPT 上。每份 PPT 最多生成两张图，
    # 控制等待时间和费用；提示词也会要求模型最多标记两页。
    import re
    image_requests = []
    for source in slides:
        clean_content = []
        seen_items = set()
        visual = source.get('visual') if isinstance(source.get('visual'), dict) else {}
        if visual.get('mode') == 'asset' and str(visual.get('asset_id', '')) in asset_map:
            asset = asset_map[str(visual['asset_id'])]
            source['_image_mode'] = '素材'
            source['_image_url'] = asset['url']
            source['_image_prompt'] = asset.get('caption', '')
            image_requests.append(source)
        elif visual.get('mode') in ('search', 'generate') and str(visual.get('prompt', '')).strip():
            source['_image_mode'] = '搜图' if visual['mode'] == 'search' else '生图'
            source['_image_prompt'] = str(visual['prompt']).strip()
            image_requests.append(source)
        for item in source.get('content', []):
            match = re.search(r'\[(\u641c\u56fe|\u751f\u56fe|\u63d2\u56fe)\uff1a(.+?)\]', str(item))
            if match and not source.get('_image_prompt'):
                source['_image_mode'] = match.group(1)
                source['_image_prompt'] = match.group(2).strip()
                item = str(item).replace(match.group(0), '').strip()
                image_requests.append(source)
            # 模型常自行加编号，版式也会加编号；统一去除，避免“1. 1.”。
            item = re.sub(r'^\s*(?:[0-9]+[、.．]|[一二三四五六七八九十]+、|[▪•\-])\s*', '', str(item)).strip()
            normalized = re.sub(r'\s+', '', item)
            if item and normalized not in seen_items:
                clean_content.append(item)
                seen_items.add(normalized)
        source['content'] = clean_content
    # 模型偶尔会漏掉配图标记。默认使用与主题一致的无文字插画；
    # 泛化搜图很容易得到不相关的摆拍或海外素材，真实事件才由故事板明确指定搜图。
    # 封面和结束页不放图，优先选择有正文要点的中间页。
    target_images = 3 if len(slides) >= 6 else 2
    if len(image_requests) < target_images:
        for source in slides[1:-1]:
            if len(image_requests) >= target_images:
                break
            if source.get('_image_prompt') or not source.get('content'):
                continue
            source['_image_mode'] = '生图'
            source['_image_prompt'] = str(source.get('title', '')) + '，' + '；'.join(str(x) for x in source['content'][:2])
            image_requests.append(source)
    if len(image_requests) > 3:
        # 故事板可能给多页标注配图；不应因此放弃整份PPT。
        # 优先保留真实搜图和案例/事实/情境页，其他页面自然回退为纯文字版。
        priority = {'case': 0, 'fact': 1, 'scenario': 2, 'steps': 3, 'compare': 4}
        ranked = sorted(
            enumerate(image_requests),
            key=lambda pair: (
                0 if pair[1].get('_image_mode') == '搜图' else 1,
                priority.get(str(pair[1].get('type', '')).lower(), 5),
                pair[0],
            )
        )
        selected = {id(source) for _, source in ranked[:3]}
        for source in image_requests:
            if id(source) not in selected:
                source.pop('_image_mode', None)
                source.pop('_image_prompt', None)
        image_requests = [source for _, source in ranked[:3]]
    image_key = os.environ.get('AI_API_KEY', '').strip()
    for source in image_requests:
        if not image_key:
            raise RuntimeError('PPT 需要插图，但 Railway 未配置 AI_API_KEY')
        image_path = os.path.join(tempfile.mkdtemp(prefix='ppt_img_'), 'image.png')
        try:
            if source.get('_image_mode') == '素材':
                image_url = source['_image_url']
            elif source.get('_image_mode') == '搜图':
                image_url = search_image_url(source['_image_prompt'], image_key)
            else:
                image_url = generate_image_url(source['_image_prompt'], image_key)
            # 不少新闻/图片站会拒绝无 User-Agent 的服务器下载。
            image_request = urllib.request.Request(image_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(image_request, timeout=60) as image_response, open(image_path, 'wb') as image_file:
                image_file.write(image_response.read())
            source['_image_path'] = image_path
        except Exception as image_error:
            # 搜到的外站图片常有防盗链；改用无文字 AI 插图，不能因此丢掉整份课件。
            if source.get('_image_mode') in ('搜图', '素材'):
                try:
                    image_url = generate_image_url(source['_image_prompt'], image_key)
                    image_request = urllib.request.Request(image_url, headers={'User-Agent': 'Mozilla/5.0'})
                    with urllib.request.urlopen(image_request, timeout=60) as image_response, open(image_path, 'wb') as image_file:
                        image_file.write(image_response.read())
                    source['_image_path'] = image_path
                    print('[PPT API] search image unavailable; used AI illustration instead:', str(image_error)[:120])
                    continue
                except Exception as fallback_error:
                    print('[PPT API] image skipped after search and AI fallback failed:', str(fallback_error)[:120])
            else:
                print('[PPT API] AI image skipped:', str(image_error)[:120])
            source.pop('_image_path', None)
    slides = fit_slides(slides); n = len(slides)

    for idx, s in enumerate(slides):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        # 除封面外统一页码，避免内容页出现重复或错乱编号。
        if idx > 0:
            footer = slide.shapes.add_textbox(Inches(11.8), Inches(7.03), Inches(1.0), Inches(0.25))
            fp = footer.text_frame.paragraphs[0]
            fp.text = f'{idx + 1} / {len(slides)}'
            fp.font.size = Pt(9); fp.font.color.rgb = P; fp.alignment = PP_ALIGN.RIGHT
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
        kind = str(s.get('type', '')).lower()
        items = s.get('content', [])
        # 故事板页面按教学目的渲染，避免所有内容页都是同一种项目符号列表。
        if kind == 'scenario':
            slide.background.fill.solid(); slide.background.fill.fore_color.rgb = L
            q = slide.shapes.add_textbox(Inches(1.15), Inches(1.1), Inches(11.0), Inches(1.3))
            qp = q.text_frame.paragraphs[0]; qp.text = s.get('title', '')
            qp.font.size = Pt(34); qp.font.bold = True; qp.font.color.rgb = D; qp.alignment = PP_ALIGN.CENTER
            panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(2.6), Inches(10.3), Inches(2.8))
            panel.fill.solid(); panel.fill.fore_color.rgb = W; panel.line.color.rgb = A
            tf = panel.text_frame; tf.word_wrap = True
            for i, item in enumerate(items[:3]):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.text = item; p.font.size = Pt(21); p.font.color.rgb = T; p.alignment = PP_ALIGN.CENTER
                p.space_after = Pt(12)
            continue
        if kind == 'case':
            title = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.8), Inches(0.8))
            tp = title.text_frame.paragraphs[0]; tp.text = s.get('title', ''); tp.font.size = Pt(30); tp.font.bold = True; tp.font.color.rgb = D
            has_image = bool(s.get('_image_path'))
            text_x = Inches(5.25) if has_image else Inches(1.1)
            text_w = Inches(6.9) if has_image else Inches(11.1)
            panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, text_x, Inches(1.55), text_w, Inches(4.5))
            panel.fill.solid(); panel.fill.fore_color.rgb = L; panel.line.fill.background()
            tf = panel.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]; p.text = '\n'.join(items[:2]); p.font.size = Pt(18); p.font.color.rgb = T
            if has_image:
                slide.shapes.add_picture(s['_image_path'], Inches(0.85), Inches(1.7), Inches(3.85), Inches(3.85))
            continue
        if kind == 'steps':
            title = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.8), Inches(0.8))
            tp = title.text_frame.paragraphs[0]; tp.text = s.get('title', ''); tp.font.size = Pt(30); tp.font.bold = True; tp.font.color.rgb = D
            # 3×2 六宫格：例如“六不准”必须完整留在一页，而不是拆成续页。
            count = min(max(len(items), 1), 6)
            columns = 3 if count > 4 else count
            card_w = 11.4 / columns
            for i, item in enumerate(items[:6]):
                col, row = i % columns, i // columns
                x = 0.9 + col * card_w
                y = 1.55 + row * 2.35
                num = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.22), Inches(y), Inches(0.56), Inches(0.56))
                num.fill.solid(); num.fill.fore_color.rgb = A; num.line.fill.background()
                np = num.text_frame.paragraphs[0]; np.text = str(i + 1); np.font.size = Pt(16); np.font.bold = True; np.font.color.rgb = W; np.alignment = PP_ALIGN.CENTER
                card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y + 0.72), Inches(card_w - 0.32), Inches(1.25))
                card.fill.solid(); card.fill.fore_color.rgb = L; card.line.fill.background()
                cp = card.text_frame.paragraphs[0]; cp.text = item; cp.font.size = Pt(16); cp.font.color.rgb = D; cp.font.bold = True; cp.alignment = PP_ALIGN.CENTER
            continue
        if kind in ('compare', 'sources'):
            title = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.8), Inches(0.8))
            tp = title.text_frame.paragraphs[0]; tp.text = s.get('title', ''); tp.font.size = Pt(30); tp.font.bold = True; tp.font.color.rgb = D
            if kind == 'sources':
                for i, item in enumerate(items[:4]):
                    tb = slide.shapes.add_textbox(Inches(1.4), Inches(1.8 + i * 0.9), Inches(10.3), Inches(0.6))
                    p = tb.text_frame.paragraphs[0]; p.text = '来源  ' + item; p.font.size = Pt(18); p.font.color.rgb = T
            else:
                mid = (len(items) + 1) // 2
                for column, group in enumerate((items[:mid], items[mid:])):
                    x = 0.8 + column * 6.0
                    panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.55), Inches(5.6), Inches(4.4))
                    panel.fill.solid(); panel.fill.fore_color.rgb = L; panel.line.fill.background()
                    tf = panel.text_frame; tf.word_wrap = True
                    for i, item in enumerate(group):
                        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                        p.text = item; p.font.size = Pt(17); p.font.color.rgb = D; p.space_after = Pt(14)
            continue
        lt = 1 if s.get('_image_path') else (idx - 1) % 3 + 1
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
            items = s.get('content', [])
            has_image = bool(s.get('_image_path'))
            text_width = Inches(6.4) if has_image else Inches(11)
            font_size = Pt(16 if len(items) >= 4 or any(len(str(x)) > 28 for x in items) else 18)
            text_x = Inches(6.2) if has_image and idx % 2 == 0 else Inches(1.2)
            for i, item in enumerate(items):
                tb2 = slide.shapes.add_textbox(text_x, Inches(1.9+i*1.05), text_width, Inches(0.9))
                tf2 = tb2.text_frame; tf2.word_wrap = True
                p2 = tf2.paragraphs[0]; p2.text = '▪  ' + item; p2.font.size = font_size; p2.font.color.rgb = T
            if has_image:
                image_x = Inches(1.0) if idx % 2 == 0 else Inches(8.2)
                slide.shapes.add_picture(s['_image_path'], image_x, Inches(2.0), Inches(4.1), Inches(4.1))
        elif lt == 2:
            items = s.get('content',[]); mid = (len(items)+1)//2
            tb2 = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11), Inches(0.9))
            p2 = tb2.text_frame.paragraphs[0]; p2.text = s.get('title',''); p2.font.size = Pt(30); p2.font.bold = True; p2.font.color.rgb = P
            for i, item in enumerate(items[:mid]):
                card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.6+i*1.0), Inches(6), Inches(0.8))
                card.fill.solid(); card.fill.fore_color.rgb = L; card.line.fill.background()
                tf2 = card.text_frame; tf2.word_wrap = True; p3 = tf2.paragraphs[0]; p3.text = item
                p3.font.size = Pt(14); p3.font.color.rgb = D; p3.font.bold = True; p3.alignment = PP_ALIGN.CENTER
            for i, item in enumerate(items[mid:]):
                card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.6+i*1.0), Inches(6), Inches(0.8))
                card.fill.solid(); card.fill.fore_color.rgb = L; card.line.fill.background()
                tf2 = card.text_frame; tf2.word_wrap = True; p3 = tf2.paragraphs[0]; p3.text = item
                p3.font.size = Pt(14); p3.font.color.rgb = D; p3.font.bold = True; p3.alignment = PP_ALIGN.CENTER
        else:
            tb2 = slide.shapes.add_textbox(Inches(0.8), Inches(0.3), Inches(11), Inches(0.9))
            p2 = tb2.text_frame.paragraphs[0]; p2.text = s.get('title',''); p2.font.size = Pt(30); p2.font.bold = True; p2.font.color.rgb = P
            line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.2), Inches(11.5), Pt(4))
            line.fill.solid(); line.fill.fore_color.rgb = A; line.line.fill.background()
            items = s.get('content', [])
            item_step = 1.15 if len(items) >= 4 else 1.4
            item_font = Pt(16 if len(items) >= 4 else 18)
            for i, item in enumerate(items):
                y = 1.8 + i * item_step
                circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.8), Inches(y), Inches(0.7), Inches(0.7))
                circle.fill.solid(); circle.fill.fore_color.rgb = P; circle.line.fill.background()
                tf2 = circle.text_frame; tf2.word_wrap = False; p3 = tf2.paragraphs[0]; p3.text = str(i+1)
                p3.font.size = Pt(20); p3.font.bold = True; p3.font.color.rgb = W; p3.alignment = PP_ALIGN.CENTER
                tb3 = slide.shapes.add_textbox(Inches(2.0), Inches(y+0.05), Inches(10), Inches(0.6))
                p4 = tb3.text_frame.paragraphs[0]; p4.text = item; p4.font.size = item_font; p4.font.color.rgb = T
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
            api_key = os.environ.get('AI_API_KEY', '').strip()
            if not api_key:
                raise RuntimeError('图片文字识别未配置：请在 Railway Variables 中设置 AI_API_KEY')
            import urllib.request, urllib.error, base64
            img_b64 = base64.b64encode(raw).decode('ascii')
            # Auto-detect image format
            if raw[:4] == b'\x89PNG':
                mime_type = 'png'
            elif raw[:3] == b'\xff\xd8\xff':
                mime_type = 'jpeg'
            elif raw[:4] in (b'GIF8',):
                mime_type = 'gif'
            elif raw[:2] == b'BM':
                mime_type = 'bmp'
            else:
                mime_type = 'jpeg'
            ocr_data = json.dumps({
                'model': 'qwen3-vl-flash',
                'input': {
                    'messages': [{'role': 'user', 'content': [
                        {'image': 'data:image/' + mime_type + ';base64,' + img_b64},
                        {'text': '请提取这张图片中的所有文字内容，直接输出文字，不要额外说明'}
                    ]}]
                }
            }).encode()
            req = urllib.request.Request(
                'https://ws-5ol6m5p8f4hikz1a.cn-beijing.maas.aliyuncs.com/api/v1/services/aigc/multimodal-generation/generation',
                data=ocr_data,
                headers={'Authorization': 'Bearer ' + api_key, 'Content-Type': 'application/json'}
            )
            # 截图 OCR 正常应在数十秒内完成；超时要返回明确错误，不能让小程序一直转圈。
            response = urllib.request.urlopen(req, timeout=45)
            result = json.loads(response.read().decode('utf-8'))
            text = result['output']['choices'][0]['message']['content']
            if isinstance(text, list):
                text = ''.join(
                    part.get('text', '') if isinstance(part, dict) else str(part)
                    for part in text
                )
            return '【图片：' + filename + '】\n' + str(text).strip()
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
        # 让 HTTP 层返回 code=-1。不能把异常伪装成文件文字，
        # 否则前端会继续把它交给 DeepSeek 做“文件总结”。
        raise RuntimeError(f'文件解析失败：{str(e)}') from e

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
                    if f == file_id or f.startswith(file_id + '.'):
                        fpath = os.path.join(DL_DIR, f)
                        try:
                            with open(fpath, 'rb') as fh:
                                data = fh.read()
                            os.remove(fpath)
                            ext = f.rsplit('.', 1)[-1].lower() if '.' in f else 'bin'
                            content_types = {
                                'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                                'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                                'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            }
                            self.send_response(200)
                            self.send_header('Content-Type', content_types.get(ext, 'application/octet-stream'))
                            self.send_header('Content-Disposition', 'attachment; filename="szhuAI.' + ext + '"')
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
                    fpath = os.path.join(DL_DIR, fid + '.xlsx')
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
                    fpath = os.path.join(DL_DIR, fid + '.xlsx')
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
                    # 统一使用阿里百炼：研究、故事板、文档文本与视觉能力都在同一工作空间内。
                    api_key = os.environ.get('AI_API_KEY', '').strip()
                    if not api_key:
                        resp = json.dumps({'code': -1, 'error': '服务器未配置阿里 AI_API_KEY，请在Railway环境变量中设置 AI_API_KEY'})
                    else:
                        messages = body.get('messages', [])
                        model = body.get('model', 'qwen-plus')
                        if str(model).startswith('deepseek'):
                            model = 'qwen-plus'
                        temp = body.get('temperature', 0.7)
                        max_tok = body.get('max_tokens', 4096)
                        post_data = json.dumps({'model': model, 'messages': messages, 'temperature': temp, 'max_tokens': max_tok, 'stream': False}).encode()
                        req = urllib.request.Request(
                            'https://ws-5ol6m5p8f4hikz1a.cn-beijing.maas.aliyuncs.com/compatible-mode/v1/chat/completions',
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
                    api_key = os.environ.get('AI_API_KEY', '').strip()
                    if not api_key:
                        resp = json.dumps({'code': -1, 'error': '联网搜索功能未配置（需在Railway设置 AI_API_KEY）'})
                    else:
                        import urllib.request, urllib.error
                        query = body.get('query', '')
                        search_data = json.dumps({
                            'model': 'qwen-plus',
                            'messages': [{'role': 'user', 'content': query}],
                            'enable_search': True,
                            'search_options': {'forced_search': True}
                        }).encode()
                        req = urllib.request.Request(
                            'https://ws-5ol6m5p8f4hikz1a.cn-beijing.maas.aliyuncs.com/compatible-mode/v1/chat/completions',
                            data=search_data,
                            headers={'Authorization': 'Bearer ' + api_key, 'Content-Type': 'application/json'}
                        )
                        try:
                            r = urllib.request.urlopen(req, timeout=120)
                            result = json.loads(r.read().decode())
                            reply = result['choices'][0]['message']['content']
                            resp = json.dumps({'code': 0, 'data': reply})
                        except Exception as e:
                            resp = json.dumps({'code': -1, 'error': '搜索失败: ' + str(e)[:200]})
                elif doc_type == 'ppt-research':
                    api_key = os.environ.get('AI_API_KEY', '').strip()
                    if not api_key:
                        resp = json.dumps({'code': -1, 'error': 'PPT联网研究未配置（需在Railway设置 AI_API_KEY）'})
                    else:
                        try:
                            pack = build_ppt_research(body.get('topic', ''), api_key)
                            resp = json.dumps({'code': 0, 'data': pack}, ensure_ascii=False)
                        except Exception as e:
                            resp = json.dumps({'code': -1, 'error': 'PPT联网研究失败: ' + str(e)[:200]})
                elif doc_type == 'vision':
                    api_key = os.environ.get('AI_API_KEY', '').strip()
                    if not api_key:
                        resp = json.dumps({'code': -1, 'error': '图片识别功能未配置（需在Railway设置 AI_API_KEY）'})
                    else:
                        import urllib.request, urllib.error
                        image_b64 = body.get('image', '')
                        prompt = body.get('prompt', '请描述这张图片')
                        vision_data = json.dumps({
                            'model': 'qwen3-vl-flash',
                            'input': {
                                'messages': [{'role': 'user', 'content': [
                                    {'image': f'data:image/jpeg;base64,{image_b64}'},
                                    {'text': prompt}
                                ]}]
                            }
                        }).encode()
                        req = urllib.request.Request(
                            'https://dashscope.aliyuncs.com/api/v1/services/aigc/multimodal-generation/generation',
                            data=vision_data,
                            headers={'Authorization': 'Bearer ' + api_key, 'Content-Type': 'application/json'}
                        )
                        try:
                            r = urllib.request.urlopen(req, timeout=120)
                            result = json.loads(r.read().decode())
                            reply = result['output']['choices'][0]['message']['content']
                            resp = json.dumps({'code': 0, 'data': reply})
                        except Exception as e:
                            resp = json.dumps({'code': -1, 'error': '识别失败: ' + str(e)[:200]})
                elif doc_type == 'generate-image':
                    api_key = os.environ.get('AI_API_KEY', '').strip()
                    if not api_key:
                        resp = json.dumps({'code': -1, 'error': 'AI生图功能未配置（需在Railway设置 AI_API_KEY）'})
                    else:
                        import urllib.request, urllib.error, time
                        prompt = body.get('prompt', '')
                        img_data = json.dumps({
                            'model': 'z-image-turbo',
                            'input': {'prompt': prompt},
                            'parameters': {'size': '1024*1024', 'n': 1}
                        }).encode()
                        req = urllib.request.Request(
                            'https://ws-5ol6m5p8f4hikz1a.cn-beijing.maas.aliyuncs.com/api/v1/services/aigc/text2image/image-synthesis',
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
                                        f'https://ws-5ol6m5p8f4hikz1a.cn-beijing.maas.aliyuncs.com/api/v1/tasks/{task_id}',
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
                    api_key = os.environ.get('AI_API_KEY', '').strip()
                    if not api_key:
                        resp = json.dumps({'code': -1, 'error': '\u9700\u8981\u914d\u7f6eAI_API_KEY'})
                    else:
                        import urllib.request, urllib.error, base64 as b64mod
                        image_b64 = body.get('image', '')
                        # Step 1: OCR using qwen-vl-max
                        ocr_data = json.dumps({
                            'model': 'qwen3-vl-flash',
                            'input': {
                                'messages': [{'role': 'user', 'content': [
                                    {'image': f'data:image/jpeg;base64,{image_b64}'},
                                    {'text': '\u8bf7\u8bc6\u522b\u8fd9\u5f20\u56fe\u7247\u4e2d\u7684\u6240\u6709\u6587\u5b57\u3002\u5982\u679c\u662f\u8868\u683c\u6570\u636e\uff0c\u7528\u5236\u8868\u7b26\t\u5206\u9694\u5217\uff0c\u6bcf\u884c\u4e00\u6761\u8bb0\u5f55\u3002\u5982\u679c\u662f\u6bb5\u843d\u6587\u5b57\uff0c\u76f4\u63a5\u8f93\u51fa\u6587\u5b57\u5185\u5bb9\u3002'}
                                ]}]       # close content[], msg_obj{}, msg_array[]
                            }              # close input{}
                        }).encode()
                        req = urllib.request.Request(
                            'https://dashscope.aliyuncs.com/api/v1/services/aigc/multimodal-generation/generation',
                            data=ocr_data,
                            headers={'Authorization': 'Bearer ' + api_key, 'Content-Type': 'application/json'}
                        )
                        try:
                            r = urllib.request.urlopen(req, timeout=120)
                            result = json.loads(r.read().decode())
                            text = result['output']['choices'][0]['message']['content']
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
                                fpath = os.path.join(DL_DIR, fid + '.xlsx')
                                with open(fpath, 'wb') as fh: fh.write(buf.getvalue())
                                resp = json.dumps({'code': 0, 'url': '/dl/' + fid, 'detected': 'table', 'ocrText': text[:5000]})
                                print(f'Handwriting: detected as table, {len(lines)} rows')
                            else:
                                # Generate Word
                                w_data = {'title': '\u624b\u5199\u6587\u6863', 'content': text}
                                docx_bytes = gen_docx(w_data)
                                fid = str(uuid.uuid4())
                                fpath = os.path.join(DL_DIR, fid + '.docx')
                                with open(fpath, 'wb') as fh: fh.write(docx_bytes)
                                resp = json.dumps({'code': 0, 'url': '/dl/' + fid, 'detected': 'text', 'ocrText': text[:5000]})
                                print(f'Handwriting: detected as text, {len(text)} chars')
                        except urllib.error.HTTPError as e:
                            body = e.read().decode()[:300]
                            resp = json.dumps({'code': -1, 'error': '\u8bc6\u522b\u5931\u8d25 (' + str(e.code) + '): ' + body})
                        except Exception as e:
                            resp = json.dumps({'code': -1, 'error': str(e)[:200]})
                else:
                    print(f'Generating PPT: {body.get("title","")}')
                    pptx_bytes = gen(body)
                    fid = str(uuid.uuid4())
                    fpath = os.path.join(DL_DIR, fid + '.pptx')
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
