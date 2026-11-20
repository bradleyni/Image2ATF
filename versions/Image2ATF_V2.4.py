"""
ATF OCR Tool  —  PyQt6
Image text extraction → dual-column ATF Word file for translation
Supports Claude / Gemini / Mistral, 5-thread concurrency, persistent API keys
Includes Tab 4 (Image->PSD Builder) and Tab 5 (Adobe DTP XML/Word Extraction).
Version: 2.4
"""

import sys, os, json, base64, re, traceback, io, time
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Optional

import requests
from PIL import Image as PILImage
from PIL import ImageDraw, ImageFont, ImageOps

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QSplitter, QVBoxLayout, QHBoxLayout,
    QTabWidget, QListWidget, QListWidgetItem, QLabel, QPushButton, QTextEdit,
    QLineEdit, QCheckBox, QComboBox, QProgressBar, QFileDialog, QMessageBox,
    QFrame, QSizePolicy, QPlainTextEdit, QStatusBar
)
from PyQt6.QtCore import (
    Qt, pyqtSignal, QSettings, QSize, QThreadPool,
    QRunnable, QObject, pyqtSlot, QRect, QPoint
)
from PyQt6.QtGui import (
    QPixmap, QIcon, QDragEnterEvent, QDropEvent, QFont, QColor,
    QPainter, QPen, QBrush, QTransform
)

# Optional google-genai SDK
try:
    from google import genai
    from google.genai import types
    HAS_GENAI = True
except ImportError:
    HAS_GENAI = False
    print("Warning: 'google-genai' library not found. Will fallback to requests.")

# Optional pywin32 for Tab4 & Tab5 DTP
try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False
    print("Warning: 'pywin32' not found. Tab4/Tab5 requires Windows & pywin32.")

# ─────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────
APP_NAME    = "Welocalize DTP 图片准备工具 V2.4"
APP_TITLE   = "Welocalize DTP 图片准备工具 V2.4"
ORG_NAME    = "DTPTools"
MAX_THREADS = 5
THUMB_W     = 320   # px
THUMB_H     = 220   # px
IMG_EXTS    = {'.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp', '.tiff', '.tif'}

DEFAULT_PROXY = ""

# ─────────────────────────────────────────────────────────────
# LANGUAGE SYSTEM  (EN / ZH, toggled at runtime)
# ─────────────────────────────────────────────────────────────
_LANG = ['EN']   

_UI = {
    'proxy_lbl':          {'EN': '🌐  Proxy:',                          'ZH': '🌐  代理:'},
    'use_proxy':          {'EN': 'Use Proxy',                           'ZH': '使用代理'},
    'api_key_lbl':        {'EN': 'API Key:',                            'ZH': 'API 密钥:'},
    'model_lbl':          {'EN': 'Model:',                              'ZH': '模型:'},
    'tab_gemini':         {'EN': 'ATF-creator (Gemini)',                'ZH': 'ATF-creator (Gemini)'},
    'tab_claude':         {'EN': 'ATF-creator (Claude)',                'ZH': 'ATF-creator (Claude)'},
    'tab_mask':           {'EN': 'ATF-creator (Mask Mode)',             'ZH': 'ATF-creator (Mask Mode)'},
    'add_files':          {'EN': '📄 Add Files',                        'ZH': '📄 添加文件'},
    'add_folder':         {'EN': '📁 Add Folder',                       'ZH': '📁 添加文件夹'},
    'rot_l':              {'EN': '↺ Rotate L',                          'ZH': '↺ 左转'},
    'rot_r':              {'EN': '↻ Rotate R',                          'ZH': '↻ 右转'},
    'auto_rotate':        {'EN': '🔄 Auto Rotate',                      'ZH': '🔄 自动旋转'},
    'drop_hint':          {'EN': '',                                    'ZH': ''},
    'img_preview':        {'EN': 'IMAGE PREVIEW',                       'ZH': '图片预览'},
    'ocr_result':         {'EN': 'OCR RESULT',                          'ZH': 'OCR 结果'},
    're_ocr':             {'EN': '↺ Re-OCR',                            'ZH': '↺ 重新识别'},
    'ocr_placeholder':    {'EN': 'OCR result will appear here...',      'ZH': 'OCR结果将显示在这里...'},
    'start_ocr':          {'EN': '▶ Start OCR',                         'ZH': '▶ 开始 OCR'},
    'stop':               {'EN': '■ Stop',                              'ZH': '■ 停止'},
    'reset':              {'EN': '↺ Reset',                             'ZH': '↺ 重置'},
    'gen_atf':            {'EN': '⬇ Generate ATF',                      'ZH': '⬇ 生成 ATF'},
    'clear_tip':          {'EN': 'Clear list',                          'ZH': '清除列表'},
    'summary_no_files':   {'EN': '=== WORKFLOW SUMMARY ===\n\nNo files loaded.', 'ZH': '=== 工作流摘要 ===\n\n未加载文件。'},
    'summary_hdr':        {'EN': '=== WORKFLOW SUMMARY ===\n\n',        'ZH': '=== 工作流摘要 ===\n\n'},
    'summary_ocr':        {'EN': '[1] OCR EXTRACTION\n',                'ZH': '[1] OCR 提取\n'},
    'summary_stats':      {'EN': '    Total: {t} | Done: {d} | Running: {r} | Pending: {p} | Failed: {e}\n\n', 'ZH': '    合计: {t} | 完成: {d} | 处理中: {r} | 待处理: {p} | 失败: {e}\n\n'},
    'summary_export':     {'EN': '[2] ATF EXPORT\n    Status: {s}\n',   'ZH': '[2] ATF 导出\n    状态: {s}\n'},
    'summary_failed':     {'EN': '\n🚨 [FAILED FILES]:\n',              'ZH': '\n🚨 [失败文件]:\n'},
    'export_not_started': {'EN': 'Not started',                         'ZH': '未开始'},
    'export_success':     {'EN': 'Success',                             'ZH': '成功'},
    'export_failed_s':    {'EN': 'Failed',                              'ZH': '失败'},
    'no_pending':         {'EN': 'No pending files.',                   'ZH': '没有待处理文件。'},
    'missing_key':        {'EN': 'Please enter the {api} API Key.',     'ZH': '请输入 {api} API Key。'},
    'confirm_clear':      {'EN': 'Clear all files from [{api}]?',       'ZH': '清除 [{api}] 中的所有文件？'},
    'confirm_reset':      {'EN': 'Clear all in [{api}] for a new batch?','ZH': '清除 [{api}] 中的所有内容以开始新批次？'},
    'log_reset':          {'EN': 'Reset — ready for new batch',         'ZH': '重置 — 准备开始新批次'},
    'log_start':          {'EN': 'Starting {n} files ({t} threads)',    'ZH': '开始处理 {n} 个文件 ({t} 线程)'},
    'log_stopped':        {'EN': 'Stopped by user',                     'ZH': '用户停止'},
    'log_batch_done':     {'EN': 'Batch complete — ✓ {d} succeeded, ✗ {e} failed', 'ZH': '批量完成 — ✓ {d} 成功, ✗ {e} 失败'},
    'log_atf_done':       {'EN': 'ATF generated: {p}',                  'ZH': 'ATF 已生成: {p}'},
    'log_atf_fail':       {'EN': 'ATF generation failed: {e}',          'ZH': 'ATF 生成失败: {e}'},
    'no_ocr_results':     {'EN': 'No OCR results to export.',           'ZH': '没有 OCR 结果可导出。'},
    'save_atf_title':     {'EN': 'Save ATF File',                       'ZH': '保存 ATF 文件'},
    'atf_saved':          {'EN': 'ATF saved to:\n{p}',                  'ZH': 'ATF 已保存至:\n{p}'},
    'gen_failed':         {'EN': 'Generation failed:\n{e}',             'ZH': '生成失败:\n{e}'},
    'status_ready':       {'EN': 'Ready  |  Max threads: {t}',          'ZH': '就绪  |  最大线程: {t}'},
    'rot_left_log':       {'EN': '{n} rotated -90°',                    'ZH': '{n} 左转 90°'},
    'rot_right_log':      {'EN': '{n} rotated +90°',                    'ZH': '{n} 右转 90°'},
    'auto_rot_start':     {'EN': '🔄 AI analysing rotation ({n} images)...', 'ZH': '🔄 AI 正在批量分析旋转方向 ({n} 张)...'},
    'auto_rot_applied':   {'EN': '✓ {n} auto-rotated {a}°',             'ZH': '✓ {n} 自动旋转 {a}°'},
    'auto_rot_err':       {'EN': '✗ Auto-rotate failed: {n} - {e}',    'ZH': '✗ 自动旋转失败: {n} - {e}'},
    'auto_rot_key_err':   {'EN': 'Auto-rotate requires the {api} API Key.', 'ZH': '自动旋转依赖 {api} API Key。'},
    'lang_btn':           {'EN': '中文',                                 'ZH': 'English'},
    'missing_key_title':  {'EN': 'Missing API Key',                     'ZH': 'API Key 缺失'},
    'confirm_title':      {'EN': 'Confirm',                             'ZH': '确认'},
    'reset_title':        {'EN': 'Reset',                               'ZH': '重置'},
    'error_title':        {'EN': 'Error',                               'ZH': '错误'},
    'done_title':         {'EN': 'Done',                                'ZH': '完成'},
    'info_title':         {'EN': 'Info',                                'ZH': '提示'},
    'files_count':        {'EN': '{n} file(s)',                         'ZH': '{n} 个文件'},
    'mask_batch_rot':     {'EN': '🔄 Auto Detect Rotation',             'ZH': '🔄 自动检测旋转'},
    'mask_batch_box':     {'EN': '⚡ Batch AI Auto-Box',                'ZH': '⚡ 批量 AI 自动画框'},
    'mask_autobox':       {'EN': '🤖 AI Auto-Box (Single)',             'ZH': '🤖 AI 单图自动画框'},
    'mask_clearbox':      {'EN': '🗑 Clear Masks',                      'ZH': '🗑 清空遮罩'},
    'mask_engine':        {'EN': 'Engine:',                             'ZH': '引擎:'},
    
    # === Tab 4 & 5 UI ===
    'tab_img2psd':        {'EN': '🖼️ Image->PSD Builder',              'ZH': '🖼️ Image->PSD Builder'},
    'tab_adobe':          {'EN': 'PSD/AI export & import',              'ZH': 'PSD/AI export & import'},
    'img2psd_start':      {'EN': '▶ Build PSDs',                        'ZH': '▶ 自动合成 PSD'},
    'dtp_export':         {'EN': '⬆ Auto Export (PS->Word / AI->XML)',  'ZH': '⬆ 一键导出 (PS->Word / AI->XML)'},
    'dtp_import':         {'EN': '⬇ Import to Source',                  'ZH': '⬇ 注入翻译到源文件'},
    'dtp_clear':          {'EN': '🗑 Clear DTP List',                   'ZH': '🗑 清空排版列表'},
    'dtp_log_hidden':     {'EN': '⚠️ [HIDDEN] Layer: {name}',           'ZH': '⚠️ [隐藏图层] 请注意: {name}'},
    'dtp_log_locked':     {'EN': '🔒 [LOCKED] Layer: {name}',           'ZH': '🔒 [锁定图层] 请解锁: {name}'},
    'dtp_format_err':     {'EN': 'COM Error. Ensure PS/AI is installed.', 'ZH': 'COM 接口调用失败，请确保已安装相应软件。'},
}

def T(key, **kwargs):
    text = _UI.get(key, {}).get(_LANG[0], key)
    if kwargs:
        try: text = text.format(**kwargs)
        except Exception: pass
    return text

_OCR_PROMPT_COMMON = """You are a professional OCR (Optical Character Recognition) assistant for localization and translation preparation.
STRICT RULES — follow ALL without exception:
1. EXACT EXTRACTION ONLY: Copy the text EXACTLY as it appears in the image, character by character. Do NOT paraphrase, summarize, interpret, describe, or add any commentary whatsoever.
2. ABSOLUTE SILENCE RULE: If the image contains no text, or the text is completely illegible, output exactly one word: BLANK
3. NO STRUCTURAL MARKDOWN: Do NOT use #, ##, ###, or > as formatting markers at the start of lines.
4. EXTRACT EVERY LINE AS-IS: Output each visible line of text on its own output line, faithfully. A separate post-processing step handles paragraph joining — your job is only faithful extraction.
5. ILLEGIBLE TEXT: Replace unreadable characters with [illegible]. Never guess.
6. EXCLUDE NON-TEXT ELEMENTS: Do NOT extract logos, brand marks, icons, watermarks, or purely decorative graphical elements.
7. READING ORDER: Top-to-bottom, left-to-right (right-to-left for RTL scripts).
8. RICH TEXT MARKERS — mark formatting ONLY where clearly visible in the image:
   - Bold text (heavier/darker letterforms): **bold text**
   - Italic text (clearly slanted letterforms): *italic text*
   - Underline: __underlined text__
   - Superscript (text raised above baseline): ^superscript^
   - Subscript (text lowered below baseline): ~subscript~
   - Colored text (intentional, clearly non-black): [color:#RRGGBB]text[/color]
9. INCLUDE ALL TEXT: headers, body text, captions, labels, timestamps, names, numbers, URLs, reference numbers, figure captions, page numbers — everything readable.
10. NO PREAMBLE OR POSTAMBLE: Start with the first extracted word. End with the last. Nothing before or after."""

OCR_PROMPTS = {
    'claude': _OCR_PROMPT_COMMON + """
CLAUDE-SPECIFIC RULES:
- BLANK LINES — CRITICAL: Do NOT add a blank line after every individual line of text. A blank line must ONLY appear when there is a clearly visible vertical gap (paragraph break) separating two distinct text blocks in the image. If in doubt, use NO blank line.
- CHAT/MESSAGING IMAGES: Each chat message bubble and its timestamp form one unit. Do not add blank lines between consecutive messages unless there is a large visible gap in the image.
- TECHNICAL/PATENT DIAGRAMS — CRITICAL: For network diagrams, system architecture diagrams, flowcharts, block diagrams, and patent figures, you MUST extract ALL visible text labels: box labels, reference numbers, axis labels, figure captions, page numbers, and any other readable text.
- CHEMICAL/TECHNICAL DIAGRAMS: Extract all text labels exactly.
- COLOR: You MUST use [color:#RRGGBB]text[/color] for any clearly non-black text (blue links, colored dates, highlighted text). Do not omit color.""",
    'gemini': _OCR_PROMPT_COMMON + """
GEMINI-SPECIFIC RULES:
- LINE WRAPPING — MOST CRITICAL: Text displayed in a narrow column that wraps word-by-word or character-by-character across multiple visual lines MUST be joined into a single output line. Only start a new output line for a genuinely new and separate text element.
- NO DUPLICATE EXTRACTION: Extract each piece of text ONCE only.
- CHEMICAL/TECHNICAL DIAGRAMS: Extract all text labels exactly.
- BOLD DETECTION: Text that appears visually heavier/thicker/darker than surrounding body text is bold. Mark it **like this**. When text is white on a dark background AND appears with heavier strokes, mark it as both bold AND colored: **[color:#FFFFFF]text[/color]**
- SUPERSCRIPT/SUBSCRIPT — CRITICAL: Mark with ^super^ and ~sub~. Chemical subscript numbers MUST be marked.
- COLOR: Use [color:#RRGGBB]text[/color] for any clearly non-black text. White text on dark background = [color:#FFFFFF]text[/color].""",
    'mistral': _OCR_PROMPT_COMMON + """
MISTRAL-SPECIFIC RULES:
- COMPLETENESS — ABSOLUTELY CRITICAL: Extract EVERY piece of text in the image without exception.
- ITALIC — STRICTLY FORBIDDEN unless visually confirmed.
- COLOR — MANDATORY for every colored text element.
- SCIENTIFIC/TECHNICAL DIAGRAMS: Extract every text label.
- BLANK LINES: Only between clearly separated paragraph blocks.""",
}

MASK_OCR_PROMPT = _OCR_PROMPT_COMMON + """
MASK MODE SPECIFIC RULES:
The image provided is a generated "snippet board". It contains a vertical list of text snippets cropped precisely from the original document.
Each snippet is aligned next to a RED numeric tag (e.g., [1], [2]).
CRITICAL INSTRUCTIONS:
1. Extract the text EXACTLY as it appears next to each numeric tag.
2. You MUST prefix each extracted paragraph with its exact numeric tag.
   Example Format:
   [1] First extracted text here
   [2] Second extracted text here
3. Treat the content next to EACH tag as EXACTLY ONE continuous paragraph. Join wrapped lines within the same snippet with a single space.
4. NEVER merge texts from DIFFERENT numbered tags.
5. PRESERVE ALL RICH TEXT FORMATTING (**bold**, ~sub~, ^super^, [color:#RRGGBB]text[/color]). DO NOT drop any formatting markers.
6. EXTREMELY IMPORTANT: The snippets often contain original document reference numbers and brackets (e.g., "[150]", "(180)", "601"). YOU MUST EXTRACT THEM. Do NOT confuse them with the RED tags. Do NOT drop any numbers or brackets visible in the text snippet."""

SETTINGS_KEYS = {
    'claude': {'key': 'claude_api_key', 'model': 'claude_model'},
    'gemini': {'key': 'gemini_api_key', 'model': 'gemini_model'},
    'mistral': {'key': 'mistral_api_key', 'model': 'mistral_model'},
    'proxy_url': 'proxy_url',
    'proxy_enabled': 'proxy_enabled',
}

CLAUDE_MODELS  = ['claude-sonnet-4-5', 'claude-opus-4-5', 'claude-haiku-4-5-20251001']
GEMINI_MODELS  = ['gemini-2.5-flash', 'gemini-2.5-pro', 'gemini-2.0-flash', 'gemini-1.5-flash', 'gemini-1.5-pro']
MISTRAL_MODELS = ['pixtral-large-latest', 'pixtral-12b-2409', 'mistral-small-latest']

def get_settings(): return QSettings(ORG_NAME, APP_NAME)
def save_setting(key, value): s = get_settings(); s.setValue(key, value); s.sync()
def load_setting(key, default=''): return get_settings().value(key, default)

def get_proxies(proxy_url: str, enabled: bool) -> Optional[dict]:
    if not enabled or not proxy_url.strip(): return None
    url = proxy_url.strip()
    if not url.startswith('http'): url = 'http://' + url
    return {'http': url, 'https': url}

def _semantic_restructure(raw_ocr: str, api_call_fn) -> str:
    if not raw_ocr.strip() or raw_ocr.strip().upper() == 'BLANK': return raw_ocr
    lines = raw_ocr.split('\n')
    if len(lines) <= 2: return raw_ocr
    marker_pat = re.compile(r'\[color:#[0-9A-Fa-f]{6}\](.*?)\[/color\]|\*\*(.*?)\*\*|\*(.*?)\*|__(.*?)__|\^(.*?)\^|~(.*?)~')
    def strip_markers(s): return marker_pat.sub(lambda m: next(g for g in m.groups() if g is not None), s)
    plain_lines = [strip_markers(l) for l in lines]
    numbered = '\n'.join(f'{i}| {plain_lines[i]}' for i in range(len(lines)))
    restructure_prompt = f"""You are a text formatting assistant. Your ONLY job: decide which consecutive lines should be joined into one paragraph because they were split by narrow text-box wrapping.
INPUT: Numbered lines of OCR text (format: "N| text")
OUTPUT RULES — CRITICAL:
- Output a JSON array of arrays.
- Output ONLY the JSON array. No explanation. No markdown. No code fences.
JOINING RULES:
- JOIN lines that together form one grammatically complete sentence split by wrapping.
- DO NOT JOIN: separate labels, separate list items, chemical labels, chat messages.
NOW PROCESS:
{numbered}"""
    raw_json = api_call_fn(restructure_prompt)
    if not raw_json or not raw_json.strip(): return raw_ocr
    try:
        import json
        clean = re.sub(r'(?:json)?\s*|\s* ', '', raw_json).strip()
        m = re.search(r'\[\s*\[.*?\]\s*\]', clean, re.DOTALL)
        if m: clean = m.group(0)
        groups = json.loads(clean)
        if not isinstance(groups, list): return raw_ocr
    except Exception: return raw_ocr
    result_lines = []
    seen_indices = set()
    for group in groups:
        if not isinstance(group, list): continue
        valid = [i for i in group if isinstance(i, int) and 0 <= i < len(lines)]
        if not valid: continue
        valid = [i for i in valid if i not in seen_indices]
        if not valid: continue
        seen_indices.update(valid)
        joined = ' '.join(lines[i].strip() for i in sorted(valid) if lines[i].strip())
        result_lines.append(joined)
    for i in range(len(lines)):
        if i not in seen_indices: result_lines.append(lines[i])
    return '\n'.join(result_lines)

def _get_rotated_b64(image_path: str, rotation_angle: int) -> tuple:
    with PILImage.open(image_path) as src:
        img = ImageOps.exif_transpose(src)
        if img.mode not in ('RGB', 'L', 'RGBA'): img = img.convert('RGB')
        if rotation_angle % 360 != 0:
            pil_angle = {90: -90, 180: 180, 270: 90, -90: 90}.get(rotation_angle % 360, 0)
            if pil_angle: img = img.rotate(pil_angle, expand=True)
        buf = io.BytesIO()
        img.convert('RGB').save(buf, format='PNG')
    return base64.b64encode(buf.getvalue()).decode(), 'image/png'

def ocr_claude(image_path: str, api_key: str, model: str, proxies: Optional[dict], rotation_angle: int = 0) -> str:
    b64, mime = _get_rotated_b64(image_path, rotation_angle)
    resp = requests.post(
        'https://api.anthropic.com/v1/messages',
        headers={'x-api-key': api_key, 'anthropic-version': '2023-06-01', 'content-type': 'application/json'},
        json={'model': model, 'max_tokens': 4096, 'messages': [{'role': 'user', 'content': [{'type': 'image', 'source': {'type': 'base64', 'media_type': mime, 'data': b64}}, {'type': 'text', 'text': OCR_PROMPTS['claude']}]}]},
        proxies=proxies, timeout=120
    )
    resp.raise_for_status()
    raw_ocr = resp.json()['content'][0]['text']
    def claude_text_call(prompt_text):
        r = requests.post('https://api.anthropic.com/v1/messages', headers={'x-api-key': api_key, 'anthropic-version': '2023-06-01', 'content-type': 'application/json'}, json={'model': model, 'max_tokens': 4096, 'messages': [{'role': 'user', 'content': prompt_text}]}, proxies=proxies, timeout=60)
        r.raise_for_status()
        return r.json()['content'][0]['text']
    return _semantic_restructure(raw_ocr, claude_text_call)

def ocr_gemini(image_path: str, api_key: str, model: str, proxies: Optional[dict], rotation_angle: int = 0) -> str:
    b64, mime = _get_rotated_b64(image_path, rotation_angle)
    resp = requests.post(
        f'https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}',
        json={'contents': [{'parts': [{'inline_data': {'mime_type': mime, 'data': b64}}, {'text': OCR_PROMPTS['gemini']}]}], 'generationConfig': {'maxOutputTokens': 4096, 'temperature': 0}},
        proxies=proxies, timeout=120
    )
    resp.raise_for_status()
    raw_ocr = resp.json()['candidates'][0]['content']['parts'][0]['text']
    def gemini_text_call(prompt_text):
        r = requests.post(f'https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}', json={'contents': [{'parts': [{'text': prompt_text}]}], 'generationConfig': {'maxOutputTokens': 4096, 'temperature': 0}}, proxies=proxies, timeout=60)
        r.raise_for_status()
        return r.json()['candidates'][0]['content']['parts'][0]['text']
    return _semantic_restructure(raw_ocr, gemini_text_call)

def ocr_mistral(image_path: str, api_key: str, model: str, proxies: Optional[dict], rotation_angle: int = 0) -> str:
    b64, mime = _get_rotated_b64(image_path, rotation_angle)
    resp = requests.post(
        'https://api.mistral.ai/v1/chat/completions',
        headers={'Authorization': f'Bearer {api_key}', 'content-type': 'application/json'},
        json={'model': model, 'max_tokens': 4096, 'temperature': 0, 'messages': [{'role': 'user', 'content': [{'type': 'text', 'text': OCR_PROMPTS['mistral']}, {'type': 'image_url', 'image_url': {'url': f'data:{mime};base64,{b64}'}}]}]},
        proxies=proxies, timeout=120
    )
    resp.raise_for_status()
    raw_ocr = resp.json()['choices'][0]['message']['content']
    def mistral_text_call(prompt_text):
        r = requests.post('https://api.mistral.ai/v1/chat/completions', headers={'Authorization': f'Bearer {api_key}', 'content-type': 'application/json'}, json={'model': model, 'max_tokens': 4096, 'temperature': 0, 'messages': [{'role': 'user', 'content': prompt_text}]}, proxies=proxies, timeout=60)
        r.raise_for_status()
        return r.json()['choices'][0]['message']['content']
    return _semantic_restructure(raw_ocr, mistral_text_call)

def _parse_inner(text: str, color=None) -> list:
    runs = []
    inner_pat = re.compile(r'\*\*(.*?)\*\*|\*(.*?)\*|__(.*?)__|\^(.*?)\^|~(.*?)~|([^*_^~]+)')
    base = {'color': color, 'bold': False, 'italic': False, 'underline': False, 'super': False, 'sub': False}
    for m in inner_pat.finditer(text):
        bold_t, italic_t, under_t, super_t, sub_t, plain = m.groups()
        if bold_t is not None: runs.append({**base, 'text': bold_t, 'bold': True})
        elif italic_t is not None: runs.append({**base, 'text': italic_t, 'italic': True})
        elif under_t is not None: runs.append({**base, 'text': under_t, 'underline': True})
        elif super_t is not None: runs.append({**base, 'text': super_t, 'super': True})
        elif sub_t is not None: runs.append({**base, 'text': sub_t, 'sub': True})
        elif plain is not None:
            if (runs and not any(runs[-1].get(k) for k in ('bold', 'italic', 'underline', 'super', 'sub')) and runs[-1].get('color') == color):
                runs[-1]['text'] += plain
            else:
                runs.append({**base, 'text': plain})
    return runs

def parse_rich_text(text: str) -> list:
    runs = []
    color_pat = re.compile(r'\[color:(#[0-9A-Fa-f]{6})\](.*?)\[/color\]')
    last = 0
    for cm in color_pat.finditer(text):
        before = text[last:cm.start()]
        if before: runs.extend(_parse_inner(before, color=None))
        color_hex = cm.group(1).lstrip('#'); inner = cm.group(2)
        runs.extend(_parse_inner(inner, color=color_hex))
        last = cm.end()
    if last < len(text): runs.extend(_parse_inner(text[last:], color=None))
    return runs

def _force_black_on_white(color_hex: str) -> str:
    try:
        h = color_hex.lstrip('#')
        r, g, b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
        luminance = 0.2126*r + 0.7152*g + 0.0722*b
        if luminance > 180: return '000000'
        return h
    except Exception: return '000000'

def build_atf_docx(items: list, output_path: str):
    try:
        from docx import Document
        from docx.shared import Inches
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
    except ImportError:
        raise ImportError("python-docx not installed.")
    COL_W = 4608; TABLE_W = COL_W * 2 
    doc = Document(); section = doc.sections[0]
    section.page_width = Inches(8.5); section.page_height = Inches(11)
    section.left_margin = Inches(0.75); section.right_margin = Inches(0.75)
    section.top_margin = Inches(0.75); section.bottom_margin = Inches(0.75)
    table = doc.add_table(rows=1, cols=2); table.style = 'Table Grid'
    tbl = table._tbl; tblPr = tbl.tblPr
    for old in tblPr.findall(qn('w:tblW')): tblPr.remove(old)
    tblW_el = OxmlElement('w:tblW'); tblW_el.set(qn('w:w'), str(TABLE_W)); tblW_el.set(qn('w:type'), 'dxa'); tblPr.append(tblW_el)
    tblLayout = OxmlElement('w:tblLayout'); tblLayout.set(qn('w:type'), 'fixed'); tblPr.append(tblLayout)
    for old in tbl.findall(qn('w:tblGrid')): tbl.remove(old)
    tblGrid = OxmlElement('w:tblGrid')
    for _ in range(2): gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'), str(COL_W)); tblGrid.append(gc)
    tbl.insert(list(tbl).index(tblPr) + 1, tblGrid); tbl.remove(table.rows[0]._tr)
    IMG_W_INCHES = TABLE_W / 1440.0
    for item in items:
        name = item['name']; img_path = item['image_path']; raw_text = item['ocr_text'].strip()
        stripped_raw = raw_text.strip()
        if stripped_raw.upper() == 'BLANK' or not stripped_raw: ocr_lines = []
        else:
            cleaned = re.sub(r'^#{1,6}\s+|^[-*+]\s+', '', stripped_raw, flags=re.MULTILINE)
            cleaned = re.sub(r'\[color:(#[0-9A-Fa-f]{6})\](.*?)\[/color\]', lambda m: '[color:' + m.group(1) + ']' + m.group(2).replace('\n', ' ') + '[/color]', cleaned, flags=re.DOTALL)
            cleaned = re.sub(r'^\s*\[\d+\]\s+', '', cleaned, flags=re.MULTILINE)
            lines = [l for l in cleaned.split('\n') if l.strip()]
            seen_lines = set(); deduped = []
            for l in lines:
                if l.strip() not in seen_lines: seen_lines.add(l.strip()); deduped.append(l)
            ocr_lines = deduped
        tr_a = OxmlElement('w:tr'); tc_a = OxmlElement('w:tc'); tcPr_a = OxmlElement('w:tcPr')
        tcW_a = OxmlElement('w:tcW'); tcW_a.set(qn('w:w'), str(TABLE_W)); tcW_a.set(qn('w:type'), 'dxa'); gs_a = OxmlElement('w:gridSpan'); gs_a.set(qn('w:val'), '2')
        tcPr_a.append(tcW_a); tcPr_a.append(gs_a); tc_a.append(tcPr_a)
        p_a = OxmlElement('w:p'); r_a = OxmlElement('w:r'); rPr_a = OxmlElement('w:rPr'); rPr_a.append(OxmlElement('w:vanish')); t_a = OxmlElement('w:t'); t_a.text = name; r_a.append(rPr_a); r_a.append(t_a); p_a.append(r_a); tc_a.append(p_a); tr_a.append(tc_a); tbl.append(tr_a)
        tr_b = OxmlElement('w:tr'); tc_b = OxmlElement('w:tc'); tcPr_b = OxmlElement('w:tcPr')
        tcW_b = OxmlElement('w:tcW'); tcW_b.set(qn('w:w'), str(TABLE_W)); tcW_b.set(qn('w:type'), 'dxa'); gs_b = OxmlElement('w:gridSpan'); gs_b.set(qn('w:val'), '2')
        tcPr_b.append(tcW_b); tcPr_b.append(gs_b); tc_b.append(tcPr_b); p_b = OxmlElement('w:p'); tc_b.append(p_b); tr_b.append(tc_b); tbl.append(tr_b)
        if os.path.exists(img_path):
            try:
                img_pil = PILImage.open(img_path)
                ratio = img_pil.height / img_pil.width if img_pil.width else 1
                disp_w = IMG_W_INCHES; disp_h = disp_w * ratio
                if disp_h > 5.5: disp_h = 5.5; disp_w = disp_h / ratio
                tmp_p = doc.add_paragraph(); tmp_p.add_run().add_picture(img_path, width=Inches(disp_w))
                drawing = tmp_p._p.find('.//' + qn('w:drawing'))
                if drawing is not None: r_img = OxmlElement('w:r'); r_img.append(drawing); p_b.append(r_img)
                tmp_p._p.getparent().remove(tmp_p._p)
            except Exception as e:
                t_err = OxmlElement('w:t'); t_err.text = f'[image error: {e}]'; r_err = OxmlElement('w:r'); r_err.append(t_err); p_b.append(r_err)
        for line in ocr_lines:
            tr_c = OxmlElement('w:tr'); rich_runs = parse_rich_text(line)
            def build_rpr(rdata, hidden: bool):
                rPr = OxmlElement('w:rPr')
                if hidden: rPr.append(OxmlElement('w:vanish'))
                if rdata.get('bold'): rPr.append(OxmlElement('w:b'))
                if rdata.get('italic'): rPr.append(OxmlElement('w:i'))
                if rdata.get('underline'): u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rPr.append(u)
                if rdata.get('super'): va = OxmlElement('w:vertAlign'); va.set(qn('w:val'), 'superscript'); rPr.append(va)
                if rdata.get('sub'): va = OxmlElement('w:vertAlign'); va.set(qn('w:val'), 'subscript'); rPr.append(va)
                if rdata.get('color') and not hidden: clr = OxmlElement('w:color'); clr.set(qn('w:val'), _force_black_on_white(rdata['color'])); rPr.append(clr)
                return rPr
            def build_cell(twips: int, runs: list, hidden: bool):
                tc = OxmlElement('w:tc'); tcPr = OxmlElement('w:tcPr')
                tcW = OxmlElement('w:tcW'); tcW.set(qn('w:w'), str(twips)); tcW.set(qn('w:type'), 'dxa'); tcPr.append(tcW); tc.append(tcPr)
                p = OxmlElement('w:p'); pPr = OxmlElement('w:pPr'); spacing = OxmlElement('w:spacing')
                spacing.set(qn('w:before'), '0'); spacing.set(qn('w:after'), '120'); pPr.append(spacing); p.append(pPr)
                for rdata in runs:
                    r = OxmlElement('w:r'); r.append(build_rpr(rdata, hidden))
                    t = OxmlElement('w:t'); t.text = rdata['text']; t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve'); r.append(t); p.append(r)
                tc.append(p)
                return tc
            tr_c.append(build_cell(COL_W, rich_runs, hidden=True)); tr_c.append(build_cell(COL_W, rich_runs, hidden=False)); tbl.append(tr_c)
    doc.save(output_path)

class DetectSignals(QObject):
    orientation_done  = pyqtSignal(str, int)
    orientation_error = pyqtSignal(str, str)
    boxes_done        = pyqtSignal(str, list)
    boxes_error       = pyqtSignal(str, str)

class AIOrientationWorker(QRunnable):
    def __init__(self, path, api_key, proxy, signals):
        super().__init__(); self.path = path; self.api_key = api_key; self.proxy = proxy; self.signals = signals
    @pyqtSlot()
    def run(self):
        try:
            if not self.api_key: return
            if self.proxy: os.environ['HTTP_PROXY'] = self.proxy.get('http', '').replace('http://', ''); os.environ['HTTPS_PROXY'] = self.proxy.get('https', '').replace('http://', '')
            with PILImage.open(self.path) as src: 
                working_img = ImageOps.exif_transpose(src)
                if working_img.mode == 'CMYK': working_img = working_img.convert('RGB')
                working_img = working_img.copy()
            detect_prompt = "Analyze the text in this image. In engineering diagrams, vertical text reading from bottom-to-top is common. If the text is vertical (bottom-to-top), reply 'CW'. If it is vertical (top-to-bottom), reply 'CCW'. If horizontal, reply '0'. Reply EXACTLY one word: CW, CCW, 180, or 0."
            if HAS_GENAI:
                client = genai.Client(api_key=self.api_key)
                res = client.models.generate_content(model="gemini-3.1-flash-image-preview", config=types.GenerateContentConfig(temperature=0.0), contents=[detect_prompt, working_img])
                ans = res.text.strip().upper() if (res.candidates and res.candidates[0].content.parts) else ""
            else:
                buf = io.BytesIO(); working_img.save(buf, format='PNG'); b64 = base64.b64encode(buf.getvalue()).decode()
                resp = requests.post(f'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-image-preview:generateContent?key={self.api_key}', json={'contents': [{'parts': [{'inline_data': {'mime_type': 'image/png', 'data': b64}}, {'text': detect_prompt}]}], 'generationConfig': {'temperature': 0.0}}, proxies=self.proxy, timeout=60)
                resp.raise_for_status()
                ans = resp.json()['candidates'][0]['content']['parts'][0]['text'].strip().upper()
            angle = 0
            if "CCW" in ans: angle = -90
            elif "CW" in ans: angle = 90
            elif "180" in ans: angle = 180
            self.signals.orientation_done.emit(self.path, angle)
        except Exception as e: self.signals.orientation_error.emit(self.path, str(e))

class AIBoxDetectWorker(QRunnable):
    def __init__(self, path, rot_state, api_key, proxy, signals):
        super().__init__(); self.path = path; self.rot_state = rot_state; self.api_key = api_key; self.proxy = proxy; self.signals = signals
    @pyqtSlot()
    def run(self):
        try:
            if self.proxy: os.environ['HTTP_PROXY'] = self.proxy.get('http', '').replace('http://', ''); os.environ['HTTPS_PROXY'] = self.proxy.get('https', '').replace('http://', '')
            with PILImage.open(self.path) as src: 
                working_img = ImageOps.exif_transpose(src)
                if working_img.mode == 'CMYK': working_img = working_img.convert('RGB')
                if self.rot_state == 90: working_img = working_img.rotate(-90, expand=True) 
                elif self.rot_state == 180: working_img = working_img.rotate(180, expand=True)
                elif self.rot_state == 270 or self.rot_state == -90: working_img = working_img.rotate(90, expand=True)
                working_img = working_img.copy()
            prompt = (
                "You are an expert layout analyzer. Detect all distinct SEMANTIC TEXT BLOCKS in this image. "
                "A 'Semantic Text Block' is a complete paragraph, a complete flowchart node, or a standalone heading. "
                "CRITICAL RULES: "
                "- Group all text belonging to the same paragraph or the same flowchart node into ONE bounding box. "
                "- DO NOT output separate boxes for individual lines or words if they belong together. "
                "Return exactly ONE box per semantic block. Format strictly as a JSON list of arrays: [[ymin, xmin, ymax, xmax], ...]. Scale coordinates to 0-1000."
            )
            text_res = ""
            if HAS_GENAI:
                client = genai.Client(api_key=self.api_key)
                res = client.models.generate_content(model="gemini-3.1-flash-image-preview", config=types.GenerateContentConfig(temperature=0.0), contents=[prompt, working_img])
                text_res = res.text if (res.candidates and res.candidates[0].content.parts) else ""
            else:
                buf = io.BytesIO(); working_img.save(buf, format='PNG'); b64 = base64.b64encode(buf.getvalue()).decode()
                resp = requests.post(f'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-image-preview:generateContent?key={self.api_key}', json={'contents': [{'parts': [{'inline_data': {'mime_type': 'image/png', 'data': b64}}, {'text': prompt}]}], 'generationConfig': {'temperature': 0.0}}, proxies=self.proxy, timeout=60)
                resp.raise_for_status()
                text_res = resp.json()['candidates'][0]['content']['parts'][0]['text']
            matches = re.findall(r'\[\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\]', text_res)
            raw_boxes = []
            for m in matches:
                ymin, xmin, ymax, xmax = map(float, m)
                if ymax > ymin and xmax > xmin: raw_boxes.append([ymin, xmin, ymax, xmax])
            final_boxes = []
            for i, b1 in enumerate(raw_boxes):
                y1_min, x1_min, y1_max, x1_max = b1
                area1 = (y1_max - y1_min) * (x1_max - x1_min); is_contained = False
                for j, b2 in enumerate(raw_boxes):
                    if i == j: continue
                    y2_min, x2_min, y2_max, x2_max = b2
                    area2 = (y2_max - y2_min) * (x2_max - x2_min)
                    overlap_ymin = max(y1_min, y2_min); overlap_xmin = max(x1_min, x2_min)
                    overlap_ymax = min(y1_max, y2_max); overlap_xmax = min(x1_max, x2_max)
                    if overlap_ymax > overlap_ymin and overlap_xmax > overlap_xmin:
                        overlap_area = (overlap_ymax - overlap_ymin) * (overlap_xmax - overlap_xmin)
                        if overlap_area > 0.8 * area1 and area2 > area1: is_contained = True; break
                if not is_contained: final_boxes.append(b1)
            boxes = []
            for ymin, xmin, ymax, xmax in final_boxes:
                ymin = max(0, ymin - 3); xmin = max(0, xmin - 5); ymax = min(1000, ymax + 3); xmax = min(1000, xmax + 5)
                boxes.append((xmin/1000.0, ymin/1000.0, (xmax-xmin)/1000.0, (ymax-ymin)/1000.0))
            boxes.sort(key=lambda b: b[1])
            self.signals.boxes_done.emit(self.path, boxes)
        except Exception as e: self.signals.boxes_error.emit(self.path, str(e))

def create_redacted_image_base64(image_path: str, boxes: list, rotation_angle: int = 0) -> tuple[str, str]:
    with PILImage.open(image_path) as src: img = ImageOps.exif_transpose(src).convert('RGB')
    if rotation_angle == 90: img = img.transpose(PILImage.Transpose.ROTATE_270)
    elif rotation_angle == 180: img = img.transpose(PILImage.Transpose.ROTATE_180)
    elif rotation_angle == 270 or rotation_angle == -90: img = img.transpose(PILImage.Transpose.ROTATE_90)
    img_w, img_h = img.size; crops = []
    for rx, ry, rw, rh in boxes:
        x, y = int(rx * img_w), int(ry * img_h); w, h = int(rw * img_w), int(rh * img_h)
        x1, y1 = max(0, x-2), max(0, y-2); x2, y2 = min(img_w, x+w+2), min(img_h, y+h+2)
        crops.append(img.crop((x1, y1, x2, y2)))
    try: font = ImageFont.truetype("arial.ttf", 20)
    except: font = ImageFont.load_default()
    if not crops:
        buf = io.BytesIO(); img.save(buf, format='PNG')
        return base64.b64encode(buf.getvalue()).decode(), 'image/png'
    label_w = 60; padding_y = 20
    max_crop_w = max(c.width for c in crops)
    total_w = label_w + max_crop_w + 20; total_h = sum(c.height + padding_y for c in crops) + padding_y
    canvas = PILImage.new('RGB', (total_w, total_h), (255, 255, 255)); draw = ImageDraw.Draw(canvas)
    current_y = padding_y
    for i, crop in enumerate(crops, start=1):
        label = f"[{i}]"; text_y = current_y + (crop.height // 2) - 12
        draw.text((10, text_y), label, fill=(220, 38, 38), font=font)
        canvas.paste(crop, (label_w, current_y))
        draw.rectangle([label_w-1, current_y-1, label_w+crop.width, current_y+crop.height], outline=(220, 220, 220), width=1)
        current_y += crop.height + padding_y
    buf = io.BytesIO(); canvas.save(buf, format='PNG')
    return base64.b64encode(buf.getvalue()).decode(), 'image/png'

def ocr_mask_engine(image_path: str, boxes: list, rotation_angle: int, api_type: str, api_key: str, model: str, proxies: dict) -> str:
    if not boxes: return ""
    b64, mime = create_redacted_image_base64(image_path, boxes, rotation_angle)
    dynamic_prompt = MASK_OCR_PROMPT + f"\n\nATTENTION: There are exactly {len(boxes)} snippets in this image. You MUST output exactly {len(boxes)} extracted paragraphs."
    if api_type == 'claude':
        resp = requests.post('https://api.anthropic.com/v1/messages', headers={'x-api-key': api_key, 'anthropic-version': '2023-06-01', 'content-type': 'application/json'}, json={'model': model, 'max_tokens': 4096, 'messages': [{'role': 'user', 'content': [{'type': 'image', 'source': {'type': 'base64', 'media_type': mime, 'data': b64}}, {'type': 'text', 'text': dynamic_prompt}]}]}, proxies=proxies, timeout=120)
        resp.raise_for_status()
        raw_text = resp.json()['content'][0]['text']
    else:
        resp = requests.post(f'https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}', json={'contents': [{'parts': [{'inline_data': {'mime_type': mime, 'data': b64}}, {'text': dynamic_prompt}]}], 'generationConfig': {'temperature': 0}}, proxies=proxies, timeout=120)
        resp.raise_for_status()
        raw_text = resp.json()['candidates'][0]['content']['parts'][0]['text']
    extracted = {}
    parts = re.split(r'^\s*\[(\d+)\]\s*', raw_text, flags=re.MULTILINE)
    if len(parts) > 1:
        for i in range(1, len(parts) - 1, 2):
            idx = parts[i]; text = parts[i+1].strip()
            text = re.sub(r'\n+', ' ', text) 
            extracted[idx] = text
    else:
        lines = [l.strip() for l in raw_text.split('\n') if l.strip()]
        for i, l in enumerate(lines, 1):
            l = re.sub(r'^\s*\[\d+\]\s*', '', l)
            extracted[str(i)] = l
    clean_lines = []
    for i in range(1, len(boxes) + 1):
        val = extracted.get(str(i), "").strip()
        if not val or val.upper() == "BLANK": val = "[BLANK]"
        clean_lines.append(f"[{i}] {val}")
    return '\n\n'.join(clean_lines)

class WorkerSignals(QObject):
    finished  = pyqtSignal(str, str); error = pyqtSignal(str, str); progress = pyqtSignal(int, int)

class OCRRunnable(QRunnable):
    def __init__(self, api, image_path, api_key, model, proxies, signals, done_counter, total, rotation_angle=0):
        super().__init__(); self.api = api; self.image_path = image_path; self.api_key = api_key; self.model = model
        self.proxies = proxies; self.signals = signals; self.done_counter = done_counter; self.total = total; self.rotation_angle = rotation_angle
    @pyqtSlot()
    def run(self):
        try:
            if self.api == 'claude': text = ocr_claude(self.image_path, self.api_key, self.model, self.proxies, self.rotation_angle)
            elif self.api == 'gemini': text = ocr_gemini(self.image_path, self.api_key, self.model, self.proxies, self.rotation_angle)
            else: text = ocr_mistral(self.image_path, self.api_key, self.model, self.proxies, self.rotation_angle)
            self.signals.finished.emit(self.image_path, text)
        except Exception as e: self.signals.error.emit(self.image_path, str(e))
        finally:
            self.done_counter[0] += 1
            self.signals.progress.emit(self.done_counter[0], self.total)

class MaskOCRRunnable(QRunnable):
    def __init__(self, api, image_path, boxes, rotation_angle, api_key, model, proxies, signals, done_counter, total):
        super().__init__(); self.api = api; self.image_path = image_path; self.boxes = boxes; self.rotation_angle = rotation_angle
        self.api_key = api_key; self.model = model; self.proxies = proxies; self.signals = signals
        self.done_counter = done_counter; self.total = total
    @pyqtSlot()
    def run(self):
        try:
            text = ocr_mask_engine(self.image_path, self.boxes, self.rotation_angle, self.api, self.api_key, self.model, self.proxies)
            self.signals.finished.emit(self.image_path, text)
        except Exception as e: self.signals.error.emit(self.image_path, str(e))
        finally:
            self.done_counter[0] += 1
            self.signals.progress.emit(self.done_counter[0], self.total)

class FileItem:
    def __init__(self, path: str):
        self.path = path; self.name = Path(path).name
        self.status = 'pending'; self.ocr_text = ''; self.error_msg = ''
        self.pixmap = None; self.boxes = []; self.box_detected = False; self.rotation_angle = 0 
    def set_rotation(self, angle: int):
        self.rotation_angle = angle % 360; self.pixmap = None; self.boxes.clear() 
    def load_thumb(self) -> QPixmap:
        if self.pixmap: return self.pixmap
        try:
            pix = QPixmap(self.path)
            if not pix.isNull():
                if self.rotation_angle != 0:
                    transform = QTransform().rotate(self.rotation_angle)
                    pix = pix.transformed(transform, Qt.TransformationMode.SmoothTransformation)
                self.pixmap = pix.scaled(THUMB_W, THUMB_H, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            else: self.pixmap = QPixmap(THUMB_W, THUMB_H); self.pixmap.fill(QColor('#2a2a2a'))
        except: self.pixmap = QPixmap(THUMB_W, THUMB_H); self.pixmap.fill(QColor('#2a2a2a'))
        return self.pixmap

class ScalableImageLabel(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._internal_pixmap = None; self._zoom_factor = 1.0
        self.pan_y = 0 
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setStyleSheet('background: #1a1d27; border-radius: 6px; color: #64748b;')
        self.setText('← Select an image')
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
    @property
    def _source_pixmap(self): return self._internal_pixmap
    @_source_pixmap.setter
    def _source_pixmap(self, value): self._internal_pixmap = value; self._zoom_factor = 1.0; self.pan_y = 0
    def set_image(self, path: str):
        pix = QPixmap(path)
        if pix.isNull(): self._source_pixmap = None; self.setText('Cannot load image')
        else: self._source_pixmap = pix; self.setText(''); self.update()
    def clear_image(self): self._source_pixmap = None; self.clear(); self.setText('← Select an image'); self.update()
    def wheelEvent(self, event):
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            if event.angleDelta().y() > 0: self._zoom_factor = min(5.0, self._zoom_factor + 0.1)
            else: self._zoom_factor = max(0.1, self._zoom_factor - 0.1)
        else:
            self.pan_y += event.angleDelta().y()
        self.update()
    def paintEvent(self, event):
        super().paintEvent(event)
        if not self._internal_pixmap: return
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        v_w, v_h = self.width(), self.height()
        p_w, p_h = self._internal_pixmap.width(), self._internal_pixmap.height()
        base_scale = min(v_w / p_w, v_h / p_h)
        scale_factor = base_scale * self._zoom_factor
        scaled_w, scaled_h = max(1, int(p_w * scale_factor)), max(1, int(p_h * scale_factor))
        offset_x = (v_w - scaled_w) // 2
        offset_y = (v_h - scaled_h) // 2 + self.pan_y
        scaled_pixmap = self._internal_pixmap.scaled(scaled_w, scaled_h, Qt.AspectRatioMode.IgnoreAspectRatio, Qt.TransformationMode.SmoothTransformation)
        painter.drawPixmap(offset_x, offset_y, scaled_pixmap)

class InteractiveMaskLabel(QLabel):
    boxes_changed = pyqtSignal()
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setStyleSheet('background: #1a1d27; border-radius: 6px; color: #64748b;')
        self.setText('← Select an image')
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.file_item = None; self.source_pixmap = None
        self.drawing = False; self.is_erasing = False
        self.start_pos = None; self.current_pos = None
        self.scale_factor = 1.0; self.offset_x = 0; self.offset_y = 0; self._zoom_factor = 1.0
        self.pan_y = 0 
    def set_item(self, item: FileItem):
        self._zoom_factor = 1.0; self.pan_y = 0; self.file_item = item; pix = QPixmap(item.path)
        if pix.isNull(): self.source_pixmap = None; self.setText('Cannot load image')
        else:
            if item.rotation_angle != 0: pix = pix.transformed(QTransform().rotate(item.rotation_angle), Qt.TransformationMode.SmoothTransformation)
            self.source_pixmap = pix; self.setText('')
        self.update()
    def clear_item(self):
        self._zoom_factor = 1.0; self.pan_y = 0; self.file_item = None; self.source_pixmap = None; self.setText('← Select an image'); self.update()
    def wheelEvent(self, event):
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            if event.angleDelta().y() > 0: self._zoom_factor = min(5.0, self._zoom_factor + 0.1)
            else: self._zoom_factor = max(0.1, self._zoom_factor - 0.1)
        else: 
            self.pan_y += event.angleDelta().y()
        self.update()
    def paintEvent(self, event):
        super().paintEvent(event)
        if not self.source_pixmap or not self.file_item: return
        painter = QPainter(self); painter.setRenderHint(QPainter.RenderHint.Antialiasing); painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        v_w, v_h = self.width(), self.height(); p_w, p_h = self.source_pixmap.width(), self.source_pixmap.height()
        base_scale = min(v_w / p_w, v_h / p_h); self.scale_factor = base_scale * self._zoom_factor
        scaled_w = max(1, int(p_w * self.scale_factor)); scaled_h = max(1, int(p_h * self.scale_factor))
        self.offset_x = (v_w - scaled_w) // 2; self.offset_y = (v_h - scaled_h) // 2 + self.pan_y
        scaled_pixmap = self.source_pixmap.scaled(scaled_w, scaled_h, Qt.AspectRatioMode.IgnoreAspectRatio, Qt.TransformationMode.SmoothTransformation)
        painter.drawPixmap(self.offset_x, self.offset_y, scaled_pixmap)
        pen_box = QPen(QColor(79, 142, 247, 220), 2, Qt.PenStyle.SolidLine); painter.setPen(pen_box)
        for i, (rx, ry, rw, rh) in enumerate(self.file_item.boxes):
            x = int(rx * scaled_w + self.offset_x); y = int(ry * scaled_h + self.offset_y); w = int(rw * scaled_w); h = int(rh * scaled_h)
            painter.setBrush(QBrush(QColor(79, 142, 247, 40))); painter.drawRect(x, y, w, h)
            painter.setPen(QPen(QColor(255, 255, 255))); painter.drawText(x + 5, y + 15, f"[{i+1}]"); painter.setPen(pen_box) 
        if self.drawing and self.start_pos and self.current_pos:
            painter.setPen(QPen(QColor(34, 197, 94, 255), 2, Qt.PenStyle.DashLine)); painter.setBrush(QBrush(QColor(34, 197, 94, 40)))
            painter.drawRect(QRect(self.start_pos, self.current_pos).normalized())
        if self.is_erasing and self.start_pos and self.current_pos:
            painter.setPen(QPen(QColor(239, 68, 68, 255), 2, Qt.PenStyle.DashLine)); painter.setBrush(QBrush(QColor(239, 68, 68, 40)))
            painter.drawRect(QRect(self.start_pos, self.current_pos).normalized())
    def _get_box_at(self, pos):
        if not self.file_item or not self.source_pixmap: return None
        scaled_w = self.source_pixmap.width() * self.scale_factor; scaled_h = self.source_pixmap.height() * self.scale_factor
        for box in reversed(self.file_item.boxes):
            rx, ry, rw, rh = box
            x, y = int(rx * scaled_w + self.offset_x), int(ry * scaled_h + self.offset_y); w, h = int(rw * scaled_w), int(rh * scaled_h)
            if QRect(x, y, w, h).contains(pos): return box
        return None
    def mousePressEvent(self, event):
        if not self.file_item or not self.source_pixmap: return
        self.setFocus()
        if event.button() == Qt.MouseButton.LeftButton:
            clicked_box = self._get_box_at(event.pos())
            if clicked_box: self.file_item.boxes.remove(clicked_box); self.update(); self.boxes_changed.emit()
            else: self.start_pos = event.pos(); self.current_pos = event.pos(); self.drawing = True; self.update()
        elif event.button() == Qt.MouseButton.RightButton: self.start_pos = event.pos(); self.current_pos = event.pos(); self.is_erasing = True; self.update()
    def mouseMoveEvent(self, event):
        if self.drawing or self.is_erasing: self.current_pos = event.pos(); self.update()
    def mouseReleaseEvent(self, event):
        if not self.file_item or not self.source_pixmap: return
        scaled_w = self.source_pixmap.width() * self.scale_factor; scaled_h = self.source_pixmap.height() * self.scale_factor
        if event.button() == Qt.MouseButton.LeftButton and self.drawing:
            self.drawing = False; self.current_pos = event.pos()
            rect = QRect(self.start_pos, self.current_pos).normalized(); pixmap_rect = QRect(self.offset_x, self.offset_y, int(scaled_w), int(scaled_h))
            intersected = rect.intersected(pixmap_rect)
            if not intersected.isEmpty() and intersected.width() > 5 and intersected.height() > 5:
                rx = (intersected.x() - self.offset_x) / scaled_w; ry = (intersected.y() - self.offset_y) / scaled_h
                rw = intersected.width() / scaled_w; rh = intersected.height() / scaled_h
                self.file_item.boxes.append((rx, ry, rw, rh)); self.file_item.boxes.sort(key=lambda b: b[1]); self.boxes_changed.emit()
            self.update()
            
        elif event.button() == Qt.MouseButton.RightButton and self.is_erasing:
            self.is_erasing = False; self.current_pos = event.pos(); erase_rect = QRect(self.start_pos, self.current_pos).normalized(); remaining = []
            for box in self.file_item.boxes:
                rx, ry, rw, rh = box
                x, y = int(rx * scaled_w + self.offset_x), int(ry * scaled_h + self.offset_y); w, h = int(rw * scaled_w), int(rh * scaled_h)
                if not erase_rect.intersects(QRect(x, y, w, h)): remaining.append(box)
            if len(remaining) != len(self.file_item.boxes): self.file_item.boxes = remaining; self.boxes_changed.emit()
            self.update()

class ApiConfigWidget(QWidget):
    def __init__(self, api_name: str, models: list, parent=None):
        super().__init__(parent); self.api_name = api_name; self.setFixedHeight(44)
        layout = QHBoxLayout(self); layout.setContentsMargins(8, 4, 8, 4); layout.setSpacing(10)
        layout.addWidget(QLabel(T('api_key_lbl')))
        self.key_edit = QLineEdit(); self.key_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.key_edit.setPlaceholderText({'claude': 'sk-ant-...', 'gemini': 'AIza...', 'mistral': '...'}.get(api_name, '...'))
        self.key_edit.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        if api_name in SETTINGS_KEYS:
            saved_key = load_setting(SETTINGS_KEYS[api_name]['key'], '')
            if saved_key: self.key_edit.setText(saved_key)
            self.key_edit.textChanged.connect(lambda v, api=api_name: save_setting(SETTINGS_KEYS[api]['key'], v))
        layout.addWidget(self.key_edit, 1) 
        self._default_model = models[0] if models else ""
    def get_key(self): return self.key_edit.text().strip()
    def get_model(self): return self._default_model

class ProxyWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent); layout = QHBoxLayout(self); layout.setContentsMargins(8, 4, 8, 4); layout.setSpacing(8)
        self.checkbox = QCheckBox(T('use_proxy')); self.checkbox.setChecked(load_setting('proxy_enabled', 'false') == 'true')
        self.checkbox.stateChanged.connect(lambda: save_setting('proxy_enabled', 'true' if self.checkbox.isChecked() else 'false'))
        layout.addWidget(self.checkbox)
        self.url_edit = QLineEdit(); self.url_edit.setPlaceholderText('http://host:port')
        self.url_edit.setText(load_setting('proxy_url', DEFAULT_PROXY)); self.url_edit.setMaximumWidth(260)
        self.url_edit.textChanged.connect(lambda v: save_setting('proxy_url', v))
        layout.addWidget(self.url_edit); layout.addStretch()
    def get_proxies(self): return get_proxies(self.url_edit.text(), self.checkbox.isChecked())

class FileListWidget(QListWidget):
    files_dropped = pyqtSignal(list)
    file_deleted = pyqtSignal(int)
    def __init__(self, parent=None):
        super().__init__(parent); self.setAcceptDrops(True); self.setDragDropMode(QListWidget.DragDropMode.DropOnly)
        self.setIconSize(QSize(64, 48)); self.setSpacing(2); self.setUniformItemSizes(False)
    def dragEnterEvent(self, event: QDragEnterEvent): event.acceptProposedAction() if event.mimeData().hasUrls() else super().dragEnterEvent(event)
    def dragMoveEvent(self, event): event.acceptProposedAction() if event.mimeData().hasUrls() else None
    def dropEvent(self, event: QDropEvent):
        paths = []
        for url in event.mimeData().urls():
            p = url.toLocalFile()
            if os.path.isdir(p):
                for r, _, fs in os.walk(p): paths.extend([os.path.join(r, f) for f in fs if Path(f).suffix.lower() in IMG_EXTS])
            elif Path(p).suffix.lower() in IMG_EXTS: paths.append(p)
        if paths: self.files_dropped.emit(paths)
        event.acceptProposedAction()
    def keyPressEvent(self, event):
        if event.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            rows = sorted([self.row(i) for i in self.selectedItems()], reverse=True)
            for row in rows: self.file_deleted.emit(row)
        else:
            super().keyPressEvent(event)

STATUS_COLORS = {'pending': '#888888', 'processing': '#f59e0b', 'done': '#34d399', 'error': '#ef4444'}
STATUS_LABELS = {'pending': 'Pending', 'processing': 'Processing...', 'done': 'Done', 'error': 'Error'}

def create_log_panel():
    log_splitter = QSplitter(Qt.Orientation.Horizontal)
    log = QPlainTextEdit(); log.setReadOnly(True); log.setStyleSheet("background-color: #1e1e1e; color: #d4d4d4; font-family: 'Consolas', monospace; font-size: 11px; border: none;")
    summary = QPlainTextEdit(); summary.setReadOnly(True); summary.setStyleSheet("background-color: #252526; color: #ffca28; font-family: 'Consolas', monospace; font-size: 11px; border: none; border-left: 1px solid #333;")
    log_splitter.addWidget(log); log_splitter.addWidget(summary); log_splitter.setSizes([700, 300])
    return log_splitter, log, summary

class ApiTab(QWidget):
    def __init__(self, api_name: str, models: list, proxy_widget: ProxyWidget, parent=None):
        super().__init__(parent); self.api_name = api_name; self.proxy_widget = proxy_widget
        self.file_items = []; self.selected_idx = None; self._stop_flag = False; self.export_status = T('export_not_started')
        self._thread_pool = QThreadPool(); self._thread_pool.setMaxThreadCount(MAX_THREADS); self._build_ui(models)
    def _build_ui(self, models):
        root = QVBoxLayout(self); root.setContentsMargins(0, 0, 0, 0); root.setSpacing(0)
        config_frame = QFrame(); config_frame.setFixedHeight(44); config_frame.setStyleSheet('background: #1a1d27; border-bottom: 1px solid #2e3250;')
        cf_layout = QHBoxLayout(config_frame); cf_layout.setContentsMargins(0, 0, 0, 0)
        self.config_widget = ApiConfigWidget(self.api_name, models); cf_layout.addWidget(self.config_widget); root.addWidget(config_frame); self.ui_cf_layout = cf_layout
        main_splitter = QSplitter(Qt.Orientation.Vertical); main_splitter.setHandleWidth(4)
        top_splitter = QSplitter(Qt.Orientation.Horizontal); top_splitter.setHandleWidth(4)
        left = QWidget(); left.setMinimumWidth(200); left.setMaximumWidth(300)
        lv = QVBoxLayout(left); lv.setContentsMargins(0, 0, 0, 0); lv.setSpacing(0); self.ui_left_layout = lv
        left_hdr = QHBoxLayout(); left_hdr.setContentsMargins(8, 6, 6, 4); left_hdr.setSpacing(4)
        self.file_count_label = QLabel('0 files'); self.file_count_label.setStyleSheet('color: #64748b; font-size: 11px;')
        self.btn_clear = QPushButton('✕'); self.btn_clear.setFixedSize(22, 22); self.btn_clear.setToolTip(T('clear_tip')); self.btn_clear.setStyleSheet('font-size: 10px; padding: 0;')
        left_hdr.addWidget(self.file_count_label); left_hdr.addStretch(); left_hdr.addWidget(self.btn_clear); lv.addLayout(left_hdr)
        self.file_list = FileListWidget(); lv.addWidget(self.file_list); top_splitter.addWidget(left)
        right = QWidget(); rv = QVBoxLayout(right); rv.setContentsMargins(0, 0, 0, 0); rv.setSpacing(0)
        toolbar = QFrame(); toolbar.setFixedHeight(42); toolbar.setStyleSheet('background: #1a1d27; border-bottom: 1px solid #2e3250;')
        tb = QHBoxLayout(toolbar); tb.setContentsMargins(8, 6, 8, 6); tb.setSpacing(8); self.ui_right_toolbar = tb
        self.btn_add_files = QPushButton(T('add_files')); self.btn_add_folder = QPushButton(T('add_folder'))
        self.btn_rot_l = QPushButton(T('rot_l')); self.btn_rot_r = QPushButton(T('rot_r')); self.btn_auto_rot = QPushButton(T('auto_rotate'))
        for b in [self.btn_add_files, self.btn_add_folder, self.btn_rot_l, self.btn_rot_r]: b.setFixedHeight(28)
        self.btn_rot_l.setStyleSheet('font-size: 11px; color: #f59e0b;'); self.btn_rot_r.setStyleSheet('font-size: 11px; color: #f59e0b;')
        self.btn_auto_rot.setFixedHeight(24); self.btn_auto_rot.setStyleSheet('font-size: 11px; padding: 2px; color: #34d399;')
        self.api_left_toolbar = QFrame(); ltb_layout = QHBoxLayout(self.api_left_toolbar); ltb_layout.setContentsMargins(8, 2, 8, 4); ltb_layout.addWidget(self.btn_auto_rot)
        self.ui_left_layout.insertWidget(1, self.api_left_toolbar)
        tb.addWidget(self.btn_add_files); tb.addWidget(self.btn_add_folder); tb.addWidget(self.btn_rot_l); tb.addWidget(self.btn_rot_r); tb.addStretch()
        self.drop_hint_lbl = QLabel(T('drop_hint')); self.drop_hint_lbl.setStyleSheet('color: #64748b; font-size: 11px;'); tb.addWidget(self.drop_hint_lbl); rv.addWidget(toolbar)
        preview_splitter = QSplitter(Qt.Orientation.Horizontal); preview_splitter.setHandleWidth(4)
        img_frame = QFrame(); img_layout = QVBoxLayout(img_frame); img_layout.setContentsMargins(8, 6, 4, 6); self.ui_img_layout = img_layout
        img_hdr_row = QHBoxLayout()
        self.img_hdr_lbl = QLabel(T('img_preview')); self.img_hdr_lbl.setStyleSheet('font-size: 11px; font-weight: 600; color: #94a3b8; text-transform: uppercase; letter-spacing: 1px;')
        self.img_name_label = QLabel(''); self.img_name_label.setStyleSheet('font-size: 10px; color: #64748b; font-family: monospace;')
        img_hdr_row.addWidget(self.img_hdr_lbl); img_hdr_row.addStretch(); img_hdr_row.addWidget(self.img_name_label); img_layout.addLayout(img_hdr_row)
        self.img_label = ScalableImageLabel(); self.img_label.setMinimumSize(300, 200); img_layout.addWidget(self.img_label, 1); preview_splitter.addWidget(img_frame)
        ocr_frame = QFrame(); ocr_layout = QVBoxLayout(ocr_frame); ocr_layout.setContentsMargins(4, 6, 8, 6)
        ocr_hdr_row = QHBoxLayout()
        self.ocr_hdr_lbl = QLabel(T('ocr_result')); self.ocr_hdr_lbl.setStyleSheet('font-size: 11px; font-weight: 600; color: #94a3b8; text-transform: uppercase; letter-spacing: 1px;')
        self.btn_reocr = QPushButton(T('re_ocr')); self.btn_reocr.setFixedHeight(24); self.btn_reocr.setStyleSheet('font-size: 11px;'); self.btn_reocr.setEnabled(False)
        ocr_hdr_row.addWidget(self.ocr_hdr_lbl); ocr_hdr_row.addStretch(); ocr_hdr_row.addWidget(self.btn_reocr); ocr_layout.addLayout(ocr_hdr_row)
        self.ocr_edit = QPlainTextEdit(); self.ocr_edit.setPlaceholderText(T('ocr_placeholder')); self.ocr_edit.setStyleSheet('font-family: "Consolas","Monaco",monospace; font-size: 12px;')
        self.ocr_edit.textChanged.connect(self._save_ocr_edit); ocr_layout.addWidget(self.ocr_edit, 1); preview_splitter.addWidget(ocr_frame)
        rv.addWidget(preview_splitter, 1); top_splitter.addWidget(right); top_splitter.setSizes([220, 980]); main_splitter.addWidget(top_splitter)
        log_panel_container = QFrame(); log_panel_container.setStyleSheet('background: #1a1d27; border-top: 1px solid #2e3250;')
        log_layout = QVBoxLayout(log_panel_container); log_layout.setContentsMargins(0, 0, 0, 0)
        ab = QHBoxLayout(); ab.setContentsMargins(12, 8, 12, 8); ab.setSpacing(8)
        self.btn_start = QPushButton(T('start_ocr')); self.btn_stop = QPushButton(T('stop')); self.btn_reset = QPushButton(T('reset')); self.btn_gen_atf = QPushButton(T('gen_atf'))
        for b in [self.btn_start, self.btn_stop, self.btn_reset, self.btn_gen_atf]: b.setFixedHeight(32)
        self.btn_gen_atf.setEnabled(False); self.btn_start.setStyleSheet('background: #4f8ef7; color: white; font-weight: 600; border-radius: 6px; padding: 0 16px;')
        self.btn_reset.setStyleSheet('background: #f59e0b; color: #1a1000; font-weight: 600; border-radius: 6px; padding: 0 16px;'); self.btn_gen_atf.setStyleSheet('background: #34d399; color: #0a1a12; font-weight: 600; border-radius: 6px; padding: 0 16px;')
        self.progress_bar = QProgressBar(); self.progress_bar.setFixedHeight(6); self.progress_bar.setTextVisible(False); self.progress_bar.setVisible(False)
        self.progress_label = QLabel(''); self.progress_label.setStyleSheet('font-size: 11px; color: #94a3b8; font-family: monospace;')
        ab.addWidget(self.btn_start); ab.addWidget(self.btn_stop); ab.addWidget(self.btn_reset); pv = QVBoxLayout(); pv.setSpacing(2); pv.addWidget(self.progress_bar); pv.addWidget(self.progress_label); ab.addLayout(pv, 1); ab.addWidget(self.btn_gen_atf); log_layout.addLayout(ab)
        self.log_splitter, self.tab_log, self.tab_summary = create_log_panel(); log_layout.addWidget(self.log_splitter, 1)
        main_splitter.addWidget(log_panel_container); main_splitter.setSizes([600, 200]); root.addWidget(main_splitter, 1)
        self.btn_add_files.clicked.connect(self._on_add_files); self.btn_add_folder.clicked.connect(self._on_add_folder)
        self.btn_rot_l.clicked.connect(self._on_rot_left); self.btn_rot_r.clicked.connect(self._on_rot_right)
        self.btn_auto_rot.clicked.connect(self._on_batch_auto_rot); self.btn_clear.clicked.connect(self._on_clear); self.file_list.currentRowChanged.connect(self._on_select)
        self.file_list.files_dropped.connect(self._add_paths)
        self.file_list.file_deleted.connect(self._on_delete_file)
        self.btn_start.clicked.connect(self._on_start); self.btn_stop.clicked.connect(self._on_stop)
        self.btn_reset.clicked.connect(self._on_reset); self.btn_reocr.clicked.connect(self._on_reocr); self.btn_gen_atf.clicked.connect(self._on_gen_atf)
        self.ui_preview_splitter = preview_splitter
    def _on_delete_file(self, row):
        if 0 <= row < len(self.file_items):
            del self.file_items[row]
            self.file_list.takeItem(row)
            if self.selected_idx == row:
                self.selected_idx = None
                self._clear_preview()
            elif self.selected_idx is not None and self.selected_idx > row:
                self.selected_idx -= 1
            self.file_count_label.setText(T('files_count', n=len(self.file_items)))
            self._update_summary()
    def _tab_log(self, text, is_error=False, is_warning=False):
        ts = time.strftime('%H:%M:%S'); color = "#ef4444" if is_error else ("#f59e0b" if is_warning else "#d4d4d4")
        self.tab_log.appendHtml(f"<span style='color: {color};'>[{ts}] {text}</span>"); sb = self.tab_log.verticalScrollBar(); sb.setValue(sb.maximum())
    def _update_summary(self, final=False):
        total = len(self.file_items)
        if total == 0: self.tab_summary.setPlainText(T("summary_no_files")); return
        done = len([i for i in self.file_items if i.status == 'done']); errors = len([i for i in self.file_items if i.status == 'error'])
        processing = len([i for i in self.file_items if i.status == 'processing']); pending = len([i for i in self.file_items if i.status == 'pending'])
        s = T("summary_hdr") + T("summary_ocr") + T("summary_stats", t=total, d=done, r=processing, p=pending, e=errors) + T("summary_export", s=self.export_status)
        err_items = [i for i in self.file_items if i.status == 'error']
        if err_items: s += T("summary_failed") + "\n".join(f"  - {i.name}: {i.error_msg}" for i in err_items)
        self.tab_summary.setPlainText(s)
    def _on_add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, 'Select Images', '', 'Images (*.jpg *.jpeg *.png *.webp *.gif *.bmp *.tiff *.tif)')
        if paths: self._add_paths(paths)
    def _on_add_folder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Select Folder')
        if folder: self._add_paths([os.path.join(r, f) for r, _, fs in os.walk(folder) for f in fs if Path(f).suffix.lower() in IMG_EXTS])
    def _add_paths(self, paths: list):
        existing = {item.path for item in self.file_items}; existing_names = {item.name for item in self.file_items}; added = dupes = 0
        for p in paths:
            if p in existing: dupes += 1; continue
            if Path(p).name in existing_names: self._tab_log(f'Same filename: {Path(p).name}', is_error=True)
            self.file_items.append(FileItem(p)); existing.add(p); existing_names.add(Path(p).name); added += 1
        if added: self._tab_log(f'Added {added} file(s)' + (f', skipped {dupes} duplicate(s)' if dupes else ''))
        self._refresh_list(); self._update_summary()
    def _on_clear(self):
        if self.file_items and QMessageBox.question(self, T('confirm_title'), T('confirm_clear', api=self.api_name)) == QMessageBox.StandardButton.Yes:
            self.file_items.clear(); self.selected_idx = None; self._refresh_list(); self._clear_preview(); self.btn_gen_atf.setEnabled(False)
            self.export_status = 'Not started'; self._update_summary()
    def _on_reset(self):
        if self.file_items and QMessageBox.question(self, T('reset_title'), T('confirm_reset', api=self.api_name), QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes: return
        self._stop_flag = True; self._thread_pool.clear(); self.file_items.clear(); self.selected_idx = None; self._batch_done = self._batch_errors = 0
        self._refresh_list(); self._clear_preview(); self.progress_bar.setVisible(False); self.progress_label.setText(''); self.btn_gen_atf.setEnabled(False)
        self.tab_log.clear(); self.export_status = 'Not started'; self._update_summary(); self.btn_auto_rot.setEnabled(True); self._rot_pending = 0; self._tab_log(T('log_reset'))
    def _refresh_list(self):
        self.file_list.clear()
        for item in self.file_items:
            li = QListWidgetItem(); li.setIcon(QIcon(item.load_thumb())); rot_str = " [↻]" if item.rotation_angle != 0 else ""
            li.setText(f"{ {'pending': '○', 'processing': '◌', 'done': '●', 'error': '✗'}.get(item.status,'?')} {item.name}{rot_str}")
            if item.rotation_angle != 0: li.setForeground(QColor('#f59e0b'))
            else: li.setForeground(QColor(STATUS_COLORS.get(item.status, '#888')))
            li.setToolTip(f'{item.name}\n{STATUS_LABELS.get(item.status,"")}' + (f'\n{item.error_msg}' if item.error_msg else ''))
            self.file_list.addItem(li)
        self.file_count_label.setText(T('files_count', n=len(self.file_items)))
    def _on_select(self, row):
        if row < 0 or row >= len(self.file_items): return
        self.selected_idx = row; item = self.file_items[row]
        if item.rotation_angle != 0:
            pix = QPixmap(item.path)
            if not pix.isNull():
                pix = pix.transformed(QTransform().rotate(item.rotation_angle), Qt.TransformationMode.SmoothTransformation)
                self.img_label._source_pixmap = pix; self.img_label.setText(''); self.img_label._update_scaled()
            else: self.img_label.set_image(item.path)
        else: self.img_label.set_image(item.path)
        self.img_name_label.setText(item.name); self.ocr_edit.blockSignals(True); self.ocr_edit.setPlainText(item.ocr_text); self.ocr_edit.blockSignals(False)
        self.btn_reocr.setEnabled(item.status != 'processing')
    def _clear_preview(self): self.img_label.clear_image(); self.img_name_label.setText(''); self.ocr_edit.blockSignals(True); self.ocr_edit.clear(); self.ocr_edit.blockSignals(False); self.btn_reocr.setEnabled(False)
    def _save_ocr_edit(self):
        if self.selected_idx is not None and self.selected_idx < len(self.file_items): self.file_items[self.selected_idx].ocr_text = self.ocr_edit.toPlainText()
    def _get_config(self): return self.config_widget.get_key(), self.config_widget.get_model(), self.proxy_widget.get_proxies()
    def _on_start(self):
        pending = [item for item in self.file_items if item.status == 'pending']
        if not pending: QMessageBox.information(self, T('info_title'), T('no_pending')); return
        key, model, proxies = self._get_config()
        if not key: QMessageBox.warning(self, T('missing_key_title'), T('missing_key', api=self.api_name)); return
        self._stop_flag = False; self._batch_done = self._batch_errors = 0; total = len(pending); done_counter = [0]
        signals = WorkerSignals(); signals.finished.connect(self._on_ocr_done); signals.error.connect(self._on_ocr_error); signals.progress.connect(lambda d, t: self._update_progress(d, t))
        self.progress_bar.setVisible(True); self.progress_bar.setMaximum(total); self.progress_bar.setValue(0); self.progress_label.setText(f'0 / {total}'); self.btn_gen_atf.setEnabled(False)
        self._tab_log(T('log_start', n=total, t=MAX_THREADS))
        for item in pending:
            if self._stop_flag: break
            item.status = 'processing'
            worker = OCRRunnable(self.api_name, item.path, key, model, proxies, signals, done_counter, total, item.rotation_angle)
            worker.setAutoDelete(True); self._thread_pool.start(worker)
        self._refresh_list(); self._update_summary()
    def _on_stop(self):
        self._stop_flag = True; self._thread_pool.clear()
        for item in self.file_items:
            if item.status == 'processing': item.status = 'pending'
        self._refresh_list(); self._check_enable_gen_atf(); self._tab_log(T('log_stopped'), is_error=True); self._update_summary()
    def _on_ocr_done(self, path: str, text: str):
        for item in self.file_items:
            if item.path == path: item.status, item.ocr_text = 'done', text; break
        self._batch_done = getattr(self, '_batch_done', 0) + 1; self._refresh_list()
        if self.selected_idx is not None and self.file_items[self.selected_idx].path == path:
            self.ocr_edit.blockSignals(True); self.ocr_edit.setPlainText(text); self.ocr_edit.blockSignals(False); self.btn_reocr.setEnabled(True)
        self._tab_log(f'✓ {Path(path).name}'); self._check_batch_complete(); self._update_summary()
    def _on_ocr_error(self, path: str, msg: str):
        for item in self.file_items:
            if item.path == path: item.status, item.error_msg, item.ocr_text = 'error', msg, f'[OCR Error: {msg}]'; break
        self._batch_errors = getattr(self, '_batch_errors', 0) + 1; self._refresh_list()
        self._tab_log(f'✗ {Path(path).name}: {msg}', is_error=True); self._check_batch_complete(); self._update_summary()
    def _check_batch_complete(self):
        if not [i for i in self.file_items if i.status == 'processing']:
            done = len([i for i in self.file_items if i.status == 'done'])
            errors = len([i for i in self.file_items if i.status == 'error'])
            pending = len([i for i in self.file_items if i.status == 'pending'])
            self._tab_log(T('log_batch_done', d=done, e=errors) + (f', ⏸ {pending} pending' if pending else ''))
            self._update_summary(final=True); self._check_enable_gen_atf()
    def _check_enable_gen_atf(self): self.btn_gen_atf.setEnabled(any(i.ocr_text.strip() for i in self.file_items))
    def _update_progress(self, done: int, total: int): self.progress_bar.setValue(done); self.progress_label.setText(f'{done} / {total}')
    def _on_reocr(self):
        if self.selected_idx is None: return
        item = self.file_items[self.selected_idx]
        key, model, proxies = self._get_config()
        if not key: QMessageBox.warning(self, T('missing_key_title'), T('missing_key', api=self.api_name)); return
        item.status = 'processing'; self._refresh_list(); self.btn_reocr.setEnabled(False)
        signals = WorkerSignals(); signals.finished.connect(self._on_ocr_done); signals.error.connect(self._on_ocr_error); signals.progress.connect(lambda d, t: None)
        worker = OCRRunnable(self.api_name, item.path, key, model, proxies, signals, [0], 1, item.rotation_angle)
        worker.setAutoDelete(True); self._thread_pool.start(worker); self._update_summary()
    def _on_gen_atf(self):
        done_items = [item for item in self.file_items if item.ocr_text.strip()]
        if not done_items: QMessageBox.information(self, T('info_title'), T('no_ocr_results')); return
        out_path, _ = QFileDialog.getSaveFileName(self, T('save_atf_title'), os.path.join(str(Path(done_items[0].path).parent), f'ATF_{self.api_name}.docx'), 'Word Documents (*.docx)')
        if not out_path: return
        from datetime import datetime
        ts = datetime.now().strftime('%Y%m%d_%H%M%S'); base, ext = os.path.splitext(out_path)
        if not any(c.isdigit() for c in Path(out_path).stem[-8:]): out_path = f'{base}_{ts}{ext}'
        try:
            build_atf_docx([{'name': i.name, 'image_path': i.path, 'ocr_text': i.ocr_text} for i in done_items], out_path)
            self.export_status = T('export_success'); self._tab_log(T('log_atf_done', p=out_path)); self._update_summary()
            import subprocess, sys as _sys
            try:
                if _sys.platform == 'win32': os.startfile(out_path)
                elif _sys.platform == 'darwin': subprocess.Popen(['open', out_path])
                else: subprocess.Popen(['xdg-open', out_path])
            except: QMessageBox.information(self, T('done_title'), T('atf_saved', p=out_path))
        except Exception as e:
            self.export_status = T('export_failed_s'); self._tab_log(T('log_atf_fail', e=e), is_error=True); self._update_summary()
            QMessageBox.critical(self, T('error_title'), T('gen_failed', e=e) + f'\n\n{traceback.format_exc()}')
    def _retranslate_ui(self):
        self.btn_add_files.setText(T('add_files')); self.btn_add_folder.setText(T('add_folder')); self.btn_rot_l.setText(T('rot_l')); self.btn_rot_r.setText(T('rot_r'))
        self.btn_auto_rot.setText(T('auto_rotate')); self.btn_reocr.setText(T('re_ocr')); self.btn_start.setText(T('start_ocr')); self.btn_stop.setText(T('stop'))
        self.btn_reset.setText(T('reset')); self.btn_gen_atf.setText(T('gen_atf')); self.btn_clear.setToolTip(T('clear_tip'))
        self.ocr_edit.setPlaceholderText(T('ocr_placeholder')); self.file_count_label.setText(T('files_count', n=len(self.file_items)))
        self.img_hdr_lbl.setText(T('img_preview')); self.ocr_hdr_lbl.setText(T('ocr_result')); self.drop_hint_lbl.setText(T('drop_hint')); self._update_summary()
    def _on_rot_left(self):
        if self.selected_idx is None: return
        item = self.file_items[self.selected_idx]; item.set_rotation(item.rotation_angle - 90)
        pix = QPixmap(item.path)
        if not pix.isNull() and item.rotation_angle != 0:
            pix = pix.transformed(QTransform().rotate(item.rotation_angle), Qt.TransformationMode.SmoothTransformation)
            self.img_label._source_pixmap = pix; self.img_label.setText(''); self.img_label._update_scaled()
        else: self.img_label.set_image(item.path)
        self._refresh_list(); self._tab_log(T('rot_left_log', n=item.name))
    def _on_rot_right(self):
        if self.selected_idx is None: return
        item = self.file_items[self.selected_idx]; item.set_rotation(item.rotation_angle + 90)
        pix = QPixmap(item.path)
        if not pix.isNull() and item.rotation_angle != 0:
            pix = pix.transformed(QTransform().rotate(item.rotation_angle), Qt.TransformationMode.SmoothTransformation)
            self.img_label._source_pixmap = pix; self.img_label.setText(''); self.img_label._update_scaled()
        else: self.img_label.set_image(item.path)
        self._refresh_list(); self._tab_log(T('rot_right_log', n=item.name))
    def _on_batch_auto_rot(self):
        if not self.file_items: return
        key, model, proxies = self._get_config()
        if not key: QMessageBox.warning(self, T('missing_key_title'), T('auto_rot_key_err', api=self.api_name)); return
        self.btn_auto_rot.setEnabled(False); self._tab_log(T('auto_rot_start', n=len(self.file_items)))
        for item in self.file_items:
            signals = DetectSignals(); signals.orientation_done.connect(self._on_ai_rot_res); signals.orientation_error.connect(self._on_ai_rot_err)
            worker = AIOrientationWorker(item.path, key, proxies, signals)
            worker.setAutoDelete(True); self._thread_pool.start(worker)
        self._rot_pending = len(self.file_items)
    def _on_ai_rot_res(self, path, angle):
        for item in self.file_items:
            if item.path == path:
                if angle != 0: item.set_rotation(item.rotation_angle + angle); self._tab_log(T('auto_rot_applied', n=item.name, a=angle))
                break
        self._rot_pending = getattr(self, '_rot_pending', 1) - 1
        if self._rot_pending <= 0:
            self.btn_auto_rot.setEnabled(True); self._refresh_list()
            if self.selected_idx is not None: self._on_select(self.selected_idx)
    def _on_ai_rot_err(self, path, msg):
        self._tab_log(T('auto_rot_err', n=Path(path).name, e=msg), is_error=True)
        self._rot_pending = getattr(self, '_rot_pending', 1) - 1
        if self._rot_pending <= 0: self.btn_auto_rot.setEnabled(True); self._refresh_list()

class MaskApiTab(ApiTab):
    def __init__(self, api_name: str, models: list, proxy_widget: ProxyWidget, parent=None):
        super().__init__(api_name, models, proxy_widget, parent)
        self.job_stats = {'rot_total': 0, 'rot_done': 0, 'rot_err': 0, 'box_total': 0, 'box_done': 0, 'box_empty': 0, 'box_err': 0}
    def _build_ui(self, models):
        super()._build_ui(models); self.api_left_toolbar.hide() 
        self.ui_cf_layout.takeAt(0).widget().deleteLater()
        mask_config = QWidget(); mc_layout = QHBoxLayout(mask_config); mc_layout.setContentsMargins(8, 0, 8, 0)
        lbl_engine = QLabel('Engine:'); lbl_engine.setStyleSheet("font-weight: bold; color: #4f8ef7;"); mc_layout.addWidget(lbl_engine)
        self.engine_combo = QComboBox(); self.engine_combo.addItems(['gemini', 'claude']); self.engine_combo.setMinimumWidth(100); mc_layout.addWidget(self.engine_combo)
        self.gemini_cfg = ApiConfigWidget('gemini', GEMINI_MODELS); self.claude_cfg = ApiConfigWidget('claude', CLAUDE_MODELS)
        self.sep = QFrame(); self.sep.setFrameShape(QFrame.Shape.VLine); self.sep.setStyleSheet("color: #2e3250;")
        mc_layout.addWidget(self.sep); mc_layout.addWidget(self.gemini_cfg); mc_layout.addWidget(self.claude_cfg); mc_layout.addStretch()
        self.ui_cf_layout.addWidget(mask_config)
        def update_cfg_vis(): self.claude_cfg.show() if self.engine_combo.currentText() == 'claude' else self.claude_cfg.hide()
        self.engine_combo.currentTextChanged.connect(update_cfg_vis); update_cfg_vis()
        self.ui_img_layout.itemAt(1).widget().deleteLater()
        self.img_label = InteractiveMaskLabel(); self.img_label.setMinimumSize(300, 200)
        self.ui_img_layout.addWidget(self.img_label, 1); self.img_label.boxes_changed.connect(self._save_boxes)
        batch_toolbar = QFrame(); batch_tb_layout = QHBoxLayout(batch_toolbar); batch_tb_layout.setContentsMargins(8, 2, 8, 4)
        self.btn_batch_auto_rot = QPushButton(T('mask_batch_rot')); self.btn_batch_mask = QPushButton(T('mask_batch_box'))
        self.btn_batch_auto_rot.setFixedHeight(24); self.btn_batch_mask.setFixedHeight(24)
        self.btn_batch_auto_rot.setStyleSheet('font-size: 11px; padding: 2px; color: #f59e0b;'); self.btn_batch_mask.setStyleSheet('font-size: 11px; padding: 2px; background: #2e3250; color: #34d399;')
        batch_tb_layout.addWidget(self.btn_batch_auto_rot); batch_tb_layout.addWidget(self.btn_batch_mask); self.ui_left_layout.insertWidget(1, batch_toolbar)
        self.btn_autobox = QPushButton(T('mask_autobox')); self.btn_clearbox = QPushButton(T('mask_clearbox'))
        for btn in [self.btn_autobox, self.btn_clearbox]: btn.setFixedHeight(28)
        self.ui_right_toolbar.insertWidget(4, self.btn_autobox); self.ui_right_toolbar.insertWidget(5, self.btn_clearbox)
        self.btn_batch_auto_rot.clicked.connect(self._on_batch_auto_rot); self.btn_batch_mask.clicked.connect(self._on_batch_auto_mask)
        self.btn_autobox.clicked.connect(self._on_auto_box); self.btn_clearbox.clicked.connect(self._on_clear_box)
    def _retranslate_ui(self):
        super()._retranslate_ui()
        self.btn_batch_auto_rot.setText(T('mask_batch_rot')); self.btn_batch_mask.setText(T('mask_batch_box'))
        self.btn_autobox.setText(T('mask_autobox')); self.btn_clearbox.setText(T('mask_clearbox'))
    def _update_summary(self, final=False):
        total = len(self.file_items)
        if total == 0: self.tab_summary.setPlainText(T("summary_no_files")); return
        rot_s = f"[1] ROTATION DETECTION\n    Processed: {self.job_stats['rot_done']}/{self.job_stats['rot_total']} | Errors: {self.job_stats['rot_err']}\n\n"
        box_s = f"[2] AUTO BOUNDING BOX\n    Detected: {self.job_stats['box_done']}/{self.job_stats['box_total']} | Errors: {self.job_stats['box_err']}\n"
        if self.job_stats['box_empty'] > 0: box_s += f"    ⚠️ EMPTY (0 Boxes): {self.job_stats['box_empty']} files!\n"
        ocr_done = len([i for i in self.file_items if i.status == 'done']); ocr_err = len([i for i in self.file_items if i.status == 'error'])
        ocr_pend = len([i for i in self.file_items if i.status == 'pending']); ocr_proc = len([i for i in self.file_items if i.status == 'processing'])
        ocr_s = f"\n[3] OCR EXTRACTION\n    Success: {ocr_done} | Failed: {ocr_err} | Processing: {ocr_proc} | Pending: {ocr_pend}\n\n"
        exp_s = f"[4] ATF EXPORT\n    Status: {self.export_status}\n"
        s = "=== WORKFLOW SUMMARY ===\n\n" + rot_s + box_s + ocr_s + exp_s
        err_items = [i for i in self.file_items if i.status == 'error' or (hasattr(i, 'box_detected') and i.box_detected and len(i.boxes)==0)]
        if err_items:
            s += "\n🚨 [ATTENTION REQUIRED]:\n"
            for i in err_items:
                if i.status == 'error': s += f"  - {i.name}: OCR Error\n"
                elif hasattr(i, 'box_detected') and i.box_detected and len(i.boxes)==0: s += f"  - {i.name}: 0 Boxes Detected (Empty)\n"
        self.tab_summary.setPlainText(s)
    def _on_select(self, row):
        if row < 0 or row >= len(self.file_items): return
        self.selected_idx = row; item = self.file_items[row]
        self.img_label.set_item(item); self.img_name_label.setText(item.name)
        self.ocr_edit.blockSignals(True); self.ocr_edit.setPlainText(item.ocr_text); self.ocr_edit.blockSignals(False); self.btn_reocr.setEnabled(item.status != 'processing')
    def _clear_preview(self):
        self.img_label.clear_item(); self.img_name_label.setText(''); self.ocr_edit.blockSignals(True); self.ocr_edit.clear(); self.ocr_edit.blockSignals(False); self.btn_reocr.setEnabled(False)
    def _get_gemini_key(self): return self.gemini_cfg.get_key()
    def _on_rot_left(self):
        if self.selected_idx is None: return
        item = self.file_items[self.selected_idx]; item.set_rotation(item.rotation_angle - 90)
        self.img_label.set_item(item); self._refresh_list(); self._tab_log(f'[Mask] {item.name} 内存虚拟左转 90°')
    def _on_rot_right(self):
        if self.selected_idx is None: return
        item = self.file_items[self.selected_idx]; item.set_rotation(item.rotation_angle + 90)
        self.img_label.set_item(item); self._refresh_list(); self._tab_log(f'[Mask] {item.name} 内存虚拟右转 90°')
    def _on_batch_auto_rot(self):
        if not self.file_items: return
        key = self._get_gemini_key()
        if not key: QMessageBox.warning(self, "API Key Missing", "智能旋转依赖 Gemini 视觉模型，请在配置框中填写 API Key。"); return
        self.btn_batch_auto_rot.setEnabled(False); self._tab_log(f'[Mask] 🤖 AI 正在批量分析排版走向 ({len(self.file_items)} 张)...')
        self.job_stats['rot_total'] = len(self.file_items); self.job_stats['rot_done'] = 0; self.job_stats['rot_err'] = 0; self._update_summary(); proxies = self.proxy_widget.get_proxies()
        for item in self.file_items:
            signals = DetectSignals(); signals.orientation_done.connect(self._on_ai_rot_res); signals.orientation_error.connect(self._on_ai_rot_err)
            worker = AIOrientationWorker(item.path, key, proxies, signals); worker.setAutoDelete(True); self._thread_pool.start(worker)
    def _on_ai_rot_res(self, path, angle):
        self.job_stats['rot_done'] += 1
        for item in self.file_items:
            if item.path == path:
                if angle != 0: item.set_rotation(item.rotation_angle + angle); self._tab_log(f'[Mask] 智能探测: {item.name} 建议旋转并已应用')
                break
        self._update_summary(); self._finalize_ai_worker()
    def _on_ai_rot_err(self, path, msg):
        self.job_stats['rot_err'] += 1; self._tab_log(f'[Mask] 智能旋转失败: {Path(path).name} - {msg}', is_error=True); self._update_summary(); self._finalize_ai_worker()
    def _on_auto_box(self):
        if self.selected_idx is None: return
        item = self.file_items[self.selected_idx]; key = self._get_gemini_key()
        if not key: QMessageBox.warning(self, "API Key Missing", "自动画框依赖 Gemini 模型，请填写 API Key。"); return
        self.btn_autobox.setEnabled(False); self._tab_log(f'[Mask] 🤖 AI 正在识别 {item.name} 的语义文本块...')
        self.job_stats['box_total'] += 1; self._update_summary()
        signals = DetectSignals(); signals.boxes_done.connect(self._on_ai_box_res); signals.boxes_error.connect(self._on_ai_box_err)
        worker = AIBoxDetectWorker(item.path, item.rotation_angle, key, self.proxy_widget.get_proxies(), signals)
        worker.setAutoDelete(True); self._thread_pool.start(worker)
    def _on_batch_auto_mask(self):
        pending = [i for i in self.file_items if i.status == 'pending']
        if not pending: return
        key = self._get_gemini_key()
        if not key: QMessageBox.warning(self, "API Key Missing", "批量自动画框依赖 Gemini 模型，请填写 API Key。"); return
        self.btn_batch_mask.setEnabled(False); self._tab_log(f'[Mask] 🤖 AI 正在批量识别文本框 ({len(pending)} 张)...')
        self.job_stats['box_total'] = len(pending); self.job_stats['box_done'] = 0; self.job_stats['box_empty'] = 0; self.job_stats['box_err'] = 0; self._update_summary(); proxies = self.proxy_widget.get_proxies()
        for item in pending:
            signals = DetectSignals(); signals.boxes_done.connect(self._on_ai_box_res); signals.boxes_error.connect(self._on_ai_box_err)
            worker = AIBoxDetectWorker(item.path, item.rotation_angle, key, proxies, signals); worker.setAutoDelete(True); self._thread_pool.start(worker)
    def _on_ai_box_res(self, path, boxes):
        for item in self.file_items:
            if item.path == path:
                item.boxes = boxes; item.box_detected = True
                if len(boxes) == 0: self.job_stats['box_empty'] += 1; self._tab_log(f'⚠️ [Mask] 警告: {item.name} 未检测到任何文本框！', is_warning=True)
                else: self.job_stats['box_done'] += 1; self._tab_log(f'[Mask] 画框完毕: {item.name} 共找到 {len(boxes)} 个文本段落')
                if self.selected_idx is not None and self.file_items[self.selected_idx].path == path: self.img_label.update()
                break
        self.btn_autobox.setEnabled(True); self._update_summary(); self._finalize_ai_worker()
    def _on_ai_box_err(self, path, msg):
        self.job_stats['box_err'] += 1; self._tab_log(f'[Mask] AI 画框失败: {Path(path).name} - {msg}', is_error=True); self.btn_autobox.setEnabled(True); self._update_summary(); self._finalize_ai_worker()
    def _finalize_ai_worker(self):
        active_jobs = self.job_stats['rot_total'] - (self.job_stats['rot_done'] + self.job_stats['rot_err']) + self.job_stats['box_total'] - (self.job_stats['box_done'] + self.job_stats['box_empty'] + self.job_stats['box_err'])
        if active_jobs <= 0:
            self.btn_batch_mask.setEnabled(True); self.btn_batch_auto_rot.setEnabled(True); self._refresh_list()
            if self.selected_idx is not None: self.img_label.set_item(self.file_items[self.selected_idx])
    def _on_reset(self):
        super()._on_reset()
        self.job_stats = {'rot_total': 0, 'rot_done': 0, 'rot_err': 0, 'box_total': 0, 'box_done': 0, 'box_empty': 0, 'box_err': 0}
        self.btn_batch_auto_rot.setEnabled(True); self.btn_batch_mask.setEnabled(True); self.btn_autobox.setEnabled(True); self._update_summary()
    def _save_boxes(self): pass 
    def _on_clear_box(self):
        if self.selected_idx is None: return
        self.file_items[self.selected_idx].boxes.clear(); self.file_items[self.selected_idx].box_detected = False; self.img_label.update()
    def _get_config(self):
        api = self.engine_combo.currentText()
        if api == 'claude': key, model = self.claude_cfg.get_key(), self.claude_cfg.get_model()
        else: key, model = self.gemini_cfg.get_key(), self.gemini_cfg.get_model()
        return api, key, model, self.proxy_widget.get_proxies()
    def _on_start(self):
        pending = [i for i in self.file_items if i.status == 'pending' and len(i.boxes) > 0]
        if not pending: QMessageBox.information(self, 'Info', '没有画好遮罩的待处理文件。'); return
        api, key, model, proxies = self._get_config()
        if not key: QMessageBox.warning(self, 'Missing Key', f'请确保您在上方填入了 {api.upper()} 的 API Key。'); return
        self._stop_flag = False; self._batch_done = 0; self._batch_errors = 0; total = len(pending); done_counter = [0]
        signals = WorkerSignals(); signals.finished.connect(self._on_ocr_done); signals.error.connect(self._on_ocr_error); signals.progress.connect(lambda d, t: self._update_progress(d, t))
        self.progress_bar.setVisible(True); self.progress_bar.setMaximum(total); self.progress_bar.setValue(0); self.progress_label.setText(f'0 / {total}'); self.btn_gen_atf.setEnabled(False)
        self._tab_log(f'[Mask] Starting {total} files using {api}')
        for item in pending:
            if self._stop_flag: break
            item.status = 'processing'
            worker = MaskOCRRunnable(api, item.path, item.boxes, item.rotation_angle, key, model, proxies, signals, done_counter, total)
            worker.setAutoDelete(True); self._thread_pool.start(worker)
        self._refresh_list(); self._update_summary()
    def _on_reocr(self):
        if self.selected_idx is None: return
        item = self.file_items[self.selected_idx]
        if not item.boxes: QMessageBox.warning(self, 'Warning', '请先在图片上画框再提取文本。'); return
        api, key, model, proxies = self._get_config()
        if not key: QMessageBox.warning(self, 'Missing Key', '缺少 API Key。'); return
        item.status = 'processing'; self._refresh_list(); self.btn_reocr.setEnabled(False)
        signals = WorkerSignals(); signals.finished.connect(self._on_ocr_done); signals.error.connect(self._on_ocr_error); signals.progress.connect(lambda d, t: None)
        worker = MaskOCRRunnable(api, item.path, item.boxes, item.rotation_angle, key, model, proxies, signals, [0], 1)
        worker.setAutoDelete(True); self._thread_pool.start(worker)
    def _on_gen_atf(self):
        done_items = [item for item in self.file_items if item.ocr_text.strip()]
        if not done_items: return
        out_path, _ = QFileDialog.getSaveFileName(self, 'Save ATF File', os.path.join(str(Path(done_items[0].path).parent), f'ATF_{self.api_name}.docx'), 'Word Documents (*.docx)')
        if not out_path: return
        from datetime import datetime
        ts = datetime.now().strftime('%Y%m%d_%H%M%S'); base, ext = os.path.splitext(out_path)
        if not any(c.isdigit() for c in Path(out_path).stem[-8:]): out_path = f'{base}_{ts}{ext}'
        try:
            processed_items = []
            for item in done_items:
                lines = item.ocr_text.strip().split('\n\n'); new_lines = []; seen = {}
                for l in lines:
                    val = l.strip()
                    if val in seen: seen[val] += 1; new_lines.append(val + " " * seen[val])
                    else: seen[val] = 0; new_lines.append(val)
                processed_items.append({'name': item.name, 'image_path': item.path, 'ocr_text': '\n'.join(new_lines)})
            build_atf_docx(processed_items, out_path)
            self.export_status = 'Success'; self._tab_log(f'ATF generated: {out_path}'); self._update_summary()
            import subprocess, sys as _sys
            try:
                if _sys.platform == 'win32': os.startfile(out_path)
                elif _sys.platform == 'darwin': subprocess.Popen(['open', out_path])
                else: subprocess.Popen(['xdg-open', out_path])
            except Exception: QMessageBox.information(self, 'Done', f'ATF saved to:\n{out_path}')
        except Exception as e:
            self.export_status = 'Failed'; self._tab_log(f'ATF generation failed: {e}', is_error=True); self._update_summary()
            QMessageBox.critical(self, 'Error', f'Generation failed:\n{e}\n\n{traceback.format_exc()}')


# ─────────────────────────────────────────────────────────────
# TAB 4: NEW IMAGE->PSD BUILDER (v2.4 AI 微调坐标与自定义字体版)
# ─────────────────────────────────────────────────────────────

PSD_BUILDER_PROMPT = """Analyze the provided image snippet board. Each snippet is preceded by a RED numeric tag.
For EACH tag, extract the exact text and its physical visual attributes. Output a STRICT JSON array of objects.

CRITICAL TEXT RULES:
- PRESERVE ALL SPACES EXACTLY. Do not merge separate words.
- Extract exact capitalization.

Format required:
[
  {
    "id": 1,
    "text": "The exact text here. Use \\r for paragraph newlines.",
    "color_hex": "#FFFFFF",
    "is_bold": true,
    "font_style": "sans-serif",
    "visual_line_count": 2,
    "align": "left"
  }
]

RULES:
1. "id" MUST match the numeric tag.
2. "text" is the exact text. NEVER swallow spaces.
3. "color_hex" must be a 6-character hex code representing the dominant color.
4. "is_bold" is true ONLY IF the text is a thick heading/title.
5. "font_style" MUST be either "sans-serif" (clean, straight edges) or "serif" (with decorative strokes at the ends).
6. "visual_line_count" MUST be the exact integer number of physical lines this text spans in the image.
7. "align" MUST be "left", "center", or "right". Visually determine how the text is aligned WITHIN ITS OWN paragraph block.
"""

class ImageToPsdWorker(QRunnable):
    """
    Reads masked images, extracts text via Gemini-3.1-flash-image-preview.
    Uses EXACT 1:1 User Bounding Boxes for BOTH erasing and text rendering.
    Creates a hidden original reference layer.
    """
    def __init__(self, api, image_path, boxes, rotation_angle, api_key, model, proxies, custom_fonts, signals, done_counter, total):
        super().__init__()
        self.api = api
        self.image_path = image_path
        self.boxes = boxes
        self.rotation_angle = rotation_angle
        self.api_key = api_key
        self.proxies = proxies
        self.custom_fonts = custom_fonts
        self.signals = signals
        self.done_counter = done_counter
        self.total = total

    @pyqtSlot()
    def run(self):
        temp_img_path = None
        jsx_path = None
        json_path = None
        try:
            if not HAS_WIN32COM:
                raise Exception("Windows pywin32 is required to run Photoshop COM.")
            
            # 1. 提取 JSON (强制使用 gemini-3.1-flash-image-preview)
            b64, mime = create_redacted_image_base64(self.image_path, self.boxes, self.rotation_angle)
            prompt = PSD_BUILDER_PROMPT + f"\n\nATTENTION: There are exactly {len(self.boxes)} snippets in this image. You MUST output exactly {len(self.boxes)} JSON objects."
            
            # 强制锁死模型端点
            model_endpoint = "gemini-3.1-flash-image-preview"
            api_url = f'https://generativelanguage.googleapis.com/v1beta/models/{model_endpoint}:generateContent?key={self.api_key}'
            
            resp = requests.post(
                api_url, 
                json={'contents': [{'parts': [{'inline_data': {'mime_type': mime, 'data': b64}}, {'text': prompt}]}], 'generationConfig': {'temperature': 0}}, 
                proxies=self.proxies, timeout=120
            )
            resp.raise_for_status()
            raw_text = resp.json()['candidates'][0]['content']['parts'][0]['text']

            clean_text = re.sub(r'(?:json)?\n?(.*?)\n? ', r'\1', raw_text, flags=re.DOTALL).strip()
            try:
                data = json.loads(clean_text)
            except:
                match = re.search(r'\[\s*{.*?}\s*\]', raw_text, flags=re.DOTALL)
                if match:
                    try: data = json.loads(match.group(0))
                    except: data = []
                else: data = []

            if not data:
                raise Exception("Failed to parse AI response into JSON format.")

            # 2. 生成物理底图与 1:1 数学计算
            import tempfile
            import uuid
            temp_dir = os.path.abspath(tempfile.gettempdir())
            unique_id = str(uuid.uuid4())
            
            has_alpha = False
            with PILImage.open(self.image_path) as src:
                img = ImageOps.exif_transpose(src)
                if self.rotation_angle % 360 != 0:
                    pil_angle = {90: -90, 180: 180, 270: 90, -90: 90}.get(self.rotation_angle % 360, 0)
                    if pil_angle: img = img.rotate(pil_angle, expand=True)
                
                if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
                    img = img.convert('RGBA')
                    has_alpha = True
                else:
                    img = img.convert('RGB')
            
            img_w, img_h = img.size
            
            temp_img_path = os.path.join(temp_dir, f"psd_source_{unique_id}.png")
            img.save(temp_img_path, format='PNG')
            target_image_path = temp_img_path

            js_payload = []
            for d in data:
                idx = d.get('id', 0) - 1
                if 0 <= idx < len(self.boxes):
                    rx, ry, rw, rh = self.boxes[idx]
                    
                    # 绝对 1:1 画框映射
                    x1 = int(rx * img_w)
                    y1 = int(ry * img_h)
                    w = int(rw * img_w)
                    h = int(rh * img_h)
                    
                    text_str = d.get('text', '')
                    
                    # 使用 AI 判断的视觉行数，防崩算法
                    visual_lines = int(d.get('visual_line_count', 0))
                    if visual_lines < 1:
                        visual_lines = max(1, text_str.count('\r') + text_str.count('\n') + 1)
                    
                    px_leading = h / visual_lines
                    px_size = px_leading * 0.75 # 字体占行高 75%
                    px_size = max(6.0, min(px_size, 400.0))

                    final_align = d.get('align', 'left').lower()
                    if final_align not in ["left", "center", "right"]:
                        final_align = "left"
                    
                    font_style = d.get('font_style', 'sans-serif').lower()

                    is_trans_bg = False
                    if has_alpha:
                        try:
                            box_crop = img.crop((x1, y1, x1+w, y1+h))
                            min_alpha = box_crop.split()[-1].getextrema()[0]
                            if min_alpha < 128:
                                is_trans_bg = True
                        except: pass
                    
                    js_payload.append({
                        'id': d.get('id'),
                        'text': text_str.replace('\n', '\r'),
                        'rx': rx, 'ry': ry, 'rw': rw, 'rh': rh,
                        'color': d.get('color_hex', '#FFFFFF'),
                        'align': final_align,
                        'size': px_size,
                        'leading': px_leading,
                        'is_bold': d.get('is_bold', True),
                        'font_style': font_style,
                        'is_trans_bg': is_trans_bg,
                        'custom_fonts': self.custom_fonts
                    })

            json_path = os.path.join(temp_dir, f"psd_data_{unique_id}.json")
            jsx_path = os.path.join(temp_dir, f"psd_builder_{unique_id}.jsx")
            
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(js_payload, f)

            base_dir = os.path.abspath(os.path.dirname(self.image_path))
            psd_save_path = os.path.join(base_dir, f"{Path(self.image_path).stem}.psd")
            
            target_img_js = target_image_path.replace('\\', '/')
            psd_save_js = psd_save_path.replace('\\', '/')
            json_path_js = json_path.replace('\\', '/')

            # =======================================================
            # 🚨 终极 JSX：隐藏对比层 + 1:1 原位还原 + 自定义字体保护
            # =======================================================
            jsx_code = f"""
            var originalRuler = app.preferences.rulerUnits;
            var originalType = app.preferences.typeUnits;
            
            try {{
                app.preferences.rulerUnits = Units.PIXELS;
                app.preferences.typeUnits = TypeUnits.PIXELS;

                var jsonFile = new File("{json_path_js}");
                jsonFile.open("r");
                jsonFile.encoding = "UTF-8";
                var jsonString = jsonFile.read();
                jsonFile.close();
                var data = eval("(" + jsonString + ")");

                var doc = app.open(new File("{target_img_js}"));
                var docW = doc.width.value;
                var docH = doc.height.value;

                // 🚨 复制原图作为隐藏参考层，方便译员对比
                var bgLayer = doc.layers[doc.layers.length-1];
                var refLayer = bgLayer.duplicate();
                refLayer.name = "Original Reference (Hidden)";
                refLayer.visible = false;

                if (bgLayer.isBackgroundLayer) {{
                    bgLayer.isBackgroundLayer = false;
                }}
                bgLayer.name = "Clean Background";

                // 1. 严格 1:1 背景擦除
                for(var i=0; i<data.length; i++) {{
                    var b = data[i];
                    var x1 = docW * b.rx;
                    var y1 = docH * b.ry;
                    var x2 = docW * (b.rx + b.rw);
                    var y2 = docH * (b.ry + b.rh);
                    
                    var region = [[x1, y1], [x2, y1], [x2, y2], [x1, y2]];
                    doc.selection.select(region);
                    try {{
                        if (b.is_trans_bg) {{
                            doc.selection.clear();
                        }} else {{
                            var idFl = charIDToTypeID( "Fl  " );
                            var desc = new ActionDescriptor();
                            desc.putEnumerated( charIDToTypeID( "Usng" ), stringIDToTypeID( "contentAware" ), stringIDToTypeID( "contentAware" ) );
                            executeAction( idFl, desc, DialogModes.NO );
                        }}
                    }} catch(e) {{}}
                    doc.selection.deselect();
                }}

                // 2. 严格 1:1 文本框重建
                for(var i=0; i<data.length; i++) {{
                    var b = data[i];
                    var x1 = docW * b.rx;
                    var y1 = docH * b.ry;
                    var w = docW * b.rw;
                    var h = docH * b.rh;
                    
                    var layer = doc.artLayers.add();
                    layer.kind = LayerKind.TEXT;
                    
                    var cleanName = b.text.replace(/[\\r\\n]/g, " ").substring(0, 30);
                    layer.name = cleanName ? cleanName : ("Text_" + b.id);
                    
                    var textItem = layer.textItem;
                    textItem.kind = TextType.PARAGRAPHTEXT;
                    
                    try {{
                        if (b.align == "center") textItem.justification = Justification.CENTER;
                        else if (b.align == "right") textItem.justification = Justification.RIGHT;
                        else textItem.justification = Justification.LEFT;
                    }} catch(e) {{}}
                    
                    textItem.position = [new UnitValue(x1, "px"), new UnitValue(y1, "px")];
                    textItem.width = new UnitValue(w, "px");
                    textItem.height = new UnitValue(h + 20, "px");
                    
                    textItem.contents = b.text;
                    textItem.useAutoLeading = false;
                    
                    try {{ textItem.leading = new UnitValue(b.leading, "px"); }} catch(e) {{}}
                    try {{ textItem.size = new UnitValue(b.size, "px"); }} catch(e) {{}}
                    
                    // 🚨 应用用户自定义字体 (PostScript名称)，附带错误降级保护
                    try {{
                        var fName = b.custom_fonts.sans_reg; 
                        if (b.font_style == "serif") {{
                            fName = b.is_bold ? b.custom_fonts.serif_bold : b.custom_fonts.serif_reg;
                        }} else {{
                            fName = b.is_bold ? b.custom_fonts.sans_bold : b.custom_fonts.sans_reg;
                        }}
                        textItem.font = fName;
                    }} catch(e) {{
                        try {{ textItem.font = b.is_bold ? "Arial-BoldMT" : "ArialMT"; }} catch(e2) {{}}
                    }}
                    
                    try {{
                        var solidColor = new SolidColor();
                        var hex = b.color.replace("#", "");
                        if (hex.length == 6) {{
                            solidColor.rgb.hexValue = hex;
                            textItem.color = solidColor;
                        }}
                    }} catch(e) {{}}
                }}

                var saveFile = new File("{psd_save_js}");
                var psdSaveOptions = new PhotoshopSaveOptions();
                psdSaveOptions.alphaChannels = true;
                psdSaveOptions.annotations = true;
                psdSaveOptions.embedColorProfile = true;
                psdSaveOptions.layers = true;
                doc.saveAs(saveFile, psdSaveOptions, true, Extension.LOWERCASE);
                doc.close(SaveOptions.DONOTSAVECHANGES);

            }} finally {{
                app.preferences.rulerUnits = originalRuler;
                app.preferences.typeUnits = originalType;
            }}
            """

            with open(jsx_path, 'w', encoding='utf-8') as f:
                f.write(jsx_code)

            import win32com.client
            import pythoncom
            pythoncom.CoInitialize()
            ps_app = win32com.client.Dispatch("Photoshop.Application")
            ps_app.DisplayDialogs = 3
            
            self.signals.finished.emit(self.image_path, "JSON extracted. Building 1:1 Reference PSD...")
            
            ps_app.DoJavaScriptFile(jsx_path)

            try:
                if jsx_path and os.path.exists(jsx_path): os.remove(jsx_path)
                if json_path and os.path.exists(json_path): os.remove(json_path)
                if temp_img_path and os.path.exists(temp_img_path): os.remove(temp_img_path)
            except: pass

            self.signals.finished.emit(self.image_path, f"✓ PSD Rebuilt Successfully: {Path(psd_save_path).name}")

        except Exception as e:
            self.signals.error.emit(self.image_path, f"Error: {str(e)}")
        finally:
            self.done_counter[0] += 1
            self.signals.progress.emit(self.done_counter[0], self.total)
            try: pythoncom.CoUninitialize()
            except: pass


class ImageToPsdTab(MaskApiTab):
    """Tab 4: 从静态图片自动重建带排版的可编辑 PSD"""
    def __init__(self, api_name: str, models: list, proxy_widget: ProxyWidget, parent=None):
        super().__init__(api_name, models, proxy_widget, parent)
        self.btn_start.setText("▶ Build PSDs")
        self.btn_start.setStyleSheet('background: #ec4899; color: white; font-weight: 600; border-radius: 6px; padding: 0 16px;')
        self.btn_gen_atf.hide()
        self.btn_reocr.hide()
        self.ocr_hdr_lbl.setText("JSON ATTRIBUTES / PROGRESS")
        
        # 彻底隐藏并禁用旋转
        self.btn_rot_l.hide()
        self.btn_rot_r.hide()
        self.btn_auto_rot.hide()
        self.btn_batch_auto_rot.hide()

    def _build_ui(self, models):
        super()._build_ui(models)
        
        # 🚨 UI重构1：将右下方的日志框设置得更窄
        try:
            if hasattr(self, 'ui_preview_splitter'):
                self.ui_preview_splitter.setSizes([800, 200])
        except: pass
        
        # 🚨 移除扰人的提示词和占位符
        self.drop_hint_lbl.setText("")
        self.ocr_edit.setPlaceholderText("")
        self.ocr_edit.setReadOnly(True)
        
        # 🚨 UI重构2：将字体输入框嵌入到左侧工具栏 (在 Clear Masks 后面)
        font_widget = QWidget()
        font_layout = QHBoxLayout(font_widget)
        font_layout.setContentsMargins(8, 0, 0, 0)
        font_layout.setSpacing(6)
        
        lbl_font = QLabel("🎨 PS Fonts:")
        lbl_font.setStyleSheet("color: #ec4899; font-weight: bold; font-size: 11px;")
        font_layout.addWidget(lbl_font)
        
        # UI 表面显示常规字体名，底层会自动映射为 PostScript 名称
        self.f_serif = QLineEdit("Times New Roman")
        self.f_serif_b = QLineEdit("Times New Roman Bold")
        self.f_sans = QLineEdit("Arial")
        self.f_sans_b = QLineEdit("Arial Bold")
        
        self.f_serif.setToolTip("Serif")
        self.f_serif_b.setToolTip("Serif Bold")
        self.f_sans.setToolTip("Sans-Serif")
        self.f_sans_b.setToolTip("Sans-Serif Bold")
        
        for le in [self.f_serif, self.f_serif_b, self.f_sans, self.f_sans_b]:
            le.setFixedWidth(135)
            le.setFixedHeight(24)
            le.setStyleSheet("background: #0f1117; color: #94a3b8; font-size: 10px; border: 1px solid #2e3250;")
            font_layout.addWidget(le)
            
        # 插入到右侧工具栏末尾
        self.ui_right_toolbar.insertWidget(6, font_widget)

    def _get_custom_fonts(self):
        # 🚨 智能映射：把用户填写的易读名称映射为 PS 必需的精确 PostScript 名称
        def map_font(val, fallback):
            val = val.strip()
            if val.lower() == "arial": return "ArialMT"
            if val.lower() == "arial bold": return "Arial-BoldMT"
            if val.lower() == "times new roman": return "TimesNewRomanPSMT"
            if val.lower() == "times new roman bold": return "TimesNewRomanPS-BoldMT"
            return val or fallback
            
        return {
            'sans_reg': map_font(self.f_sans.text(), "ArialMT"),
            'sans_bold': map_font(self.f_sans_b.text(), "Arial-BoldMT"),
            'serif_reg': map_font(self.f_serif.text(), "TimesNewRomanPSMT"),
            'serif_bold': map_font(self.f_serif_b.text(), "TimesNewRomanPS-BoldMT")
        }

    def _retranslate_ui(self):
        super()._retranslate_ui()
        self.btn_start.setText(T('img2psd_start'))
        self.ocr_hdr_lbl.setText("JSON ATTRIBUTES / PROGRESS")
        self.ocr_edit.setPlaceholderText("")

    def _on_start(self):
        # 🚨 无限重构功能：现在只要有框的图都可以被再次执行，无需点 Reset！
        targets = [i for i in self.file_items if len(i.boxes) > 0]
        if not targets: 
            QMessageBox.information(self, 'Info', '没有画好遮罩的图片可供生成。')
            return
            
        api, key, model, proxies = self._get_config()
        if not key: 
            QMessageBox.warning(self, 'Missing Key', f'请确保您在上方填入了 {api.upper()} 的 API Key。')
            return
            
        self._stop_flag = False; self._batch_done = 0; self._batch_errors = 0; total = len(targets); done_counter = [0]
        signals = WorkerSignals()
        signals.finished.connect(self._on_ocr_done)
        signals.error.connect(self._on_ocr_error)
        signals.progress.connect(lambda d, t: self._update_progress(d, t))
        
        self.progress_bar.setVisible(True); self.progress_bar.setMaximum(total); self.progress_bar.setValue(0)
        self.progress_label.setText(f'0 / {total}')
        self._tab_log(f'[Builder] Starting {total} files (Forced Model: Gemini-3.1-Flash)')
        
        custom_fonts = self._get_custom_fonts()
        
        for item in targets:
            if self._stop_flag: break
            item.status = 'processing' # 强行把状态拨回处理中
            worker = ImageToPsdWorker(api, item.path, item.boxes, item.rotation_angle, key, model, proxies, custom_fonts, signals, done_counter, total)
            worker.setAutoDelete(True)
            self._thread_pool.start(worker)
            
        self._refresh_list()
        self._update_summary()

# ─────────────────────────────────────────────────────────────
# TAB 5: ADOBE DTP (PS/AI) 分治架构 (PS->Word, AI->XML) 稳定版
# ─────────────────────────────────────────────────────────────
class DtpFileListWidget(QListWidget):
    files_dropped = pyqtSignal(list)
    file_deleted = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QListWidget.DragDropMode.DropOnly)
        self.setStyleSheet('background: #1a1d27; border: none; border-right: 1px solid #2e3250; padding: 5px;')
        self.valid_exts = ('.psd', '.psb', '.ai', '.eps', '.ps')

    def dragEnterEvent(self, event: QDragEnterEvent): 
        event.acceptProposedAction() if event.mimeData().hasUrls() else super().dragEnterEvent(event)
        
    def dragMoveEvent(self, event): 
        event.acceptProposedAction() if event.mimeData().hasUrls() else None
        
    def dropEvent(self, event: QDropEvent):
        paths = []
        for url in event.mimeData().urls():
            p = url.toLocalFile()
            if os.path.isdir(p):
                for r, _, fs in os.walk(p): 
                    if fs: paths.extend([os.path.join(r, f) for f in fs if f.lower().endswith(self.valid_exts)])
            elif p.lower().endswith(self.valid_exts): 
                paths.append(p)
        if paths: self.files_dropped.emit(paths)
        event.acceptProposedAction()

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            for item in self.selectedItems():
                path = item.data(Qt.ItemDataRole.UserRole)
                self.file_deleted.emit(path)
                self.takeItem(self.row(item))
        else:
            super().keyPressEvent(event)


class DtpWorkerSignals(QObject):
    log = pyqtSignal(str, bool)       
    progress = pyqtSignal(int, int)   
    finished = pyqtSignal(str, str)   
    error = pyqtSignal(str)


class AdobeDtpWorker(QRunnable):
    """独立的后台 COM 处理线程：无弹窗导出 + 六维安全富文本样式保留（含段落属性）"""
    def __init__(self, mode, files, out_dir="", base_name="", trans_paths=None, signals=None):
        super().__init__()
        self.mode = mode
        self.files = files
        self.out_dir = out_dir  
        self.base_name = base_name
        self.trans_paths = trans_paths
        self.signals = signals
        
        self.stats = {'total': len(files), 'ps_processed': 0, 'ai_processed': 0, 'fail': 0}
        self.hidden_layers = []
        self.locked_layers = []
        self.failed_layers = []
        self.skipped_layers = []

    def _rgb_to_hex(self, color_obj):
        try:
            if color_obj.Typename == 'RGBColor':
                return f"#{int(color_obj.Red):02X}{int(color_obj.Green):02X}{int(color_obj.Blue):02X}"
        except: pass
        return ""

    def _hex_to_rgb(self, ai_app, hex_str):
        try:
            hex_str = hex_str.lstrip('#')
            if len(hex_str) == 6:
                color = ai_app.RGBColor()
                color.Red = int(hex_str[0:2], 16)
                color.Green = int(hex_str[2:4], 16)
                color.Blue = int(hex_str[4:6], 16)
                return color
        except: pass
        return None

    @pyqtSlot()
    def run(self):
        if not HAS_WIN32COM:
            if self.signals: self.signals.error.emit("COM Error. Ensure PS/AI is installed.")
            return

        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()

        try:
            ps_app = None
            ai_app = None
            total_files = len(self.files)
            
            if self.mode == 'export':
                ps_extracted_items = []
                ai_extracted_items = []
                
                for idx, file_path in enumerate(self.files, 1):
                    ext = Path(file_path).suffix.lower()
                    fname = Path(file_path).name
                    if self.signals:
                        self.signals.log.emit(f"[{idx}/{total_files}] Extracting {fname}...", False)
                        self.signals.progress.emit(idx, total_files)

                    try:
                        if ext in ['.psd', '.psb']:
                            if not ps_app: 
                                ps_app = win32com.client.Dispatch("Photoshop.Application")
                                ps_app.DisplayDialogs = 3
                            
                            doc = ps_app.Open(file_path)
                            text_counter = [0]
                            def traverse_ps_export(layers):
                                for i in range(layers.Count):
                                    layer = layers.Item(i + 1)
                                    if not layer.Visible: self.hidden_layers.append(f"{fname} -> {layer.Name}")
                                    if layer.AllLocked:   self.locked_layers.append(f"{fname} -> {layer.Name}")
                                    
                                    if layer.LayerType == 2:
                                        traverse_ps_export(layer.Layers)
                                    elif layer.LayerType == 1 and layer.Kind == 2:
                                        text_counter[0] += 1
                                        tid = f"TID_{text_counter[0]}"
                                        text = layer.TextItem.Contents.strip()
                                        if text: 
                                            ps_extracted_items.append({'file': fname, 'layer': tid, 'source': text})
                            traverse_ps_export(doc.Layers)
                            doc.Close(2) 
                            self.stats['ps_processed'] += 1

                        elif ext in ['.ai', '.eps', '.ps']:
                            if not ai_app: 
                                ai_app = win32com.client.Dispatch("Illustrator.Application")
                                ai_app.UserInteractionLevel = -1 
                            
                            doc = ai_app.Open(file_path)
                            
                            import tempfile
                            temp_dir = tempfile.gettempdir()
                            jsx_path = os.path.join(temp_dir, "ai_export_fast.jsx").replace('\\', '/')
                            out_txt_path = os.path.join(temp_dir, "ai_export_result.txt").replace('\\', '/')
                            
                            if self.signals: self.signals.log.emit(f"  -> Injecting 6D Rich Style extraction...", False)

                            jsx_code = 'var destFile = new File("' + out_txt_path + '");\n'
                            jsx_code += """
                            destFile.open("w");
                            destFile.encoding = "UTF-8";

                            var doc = app.activeDocument;
                            var frames = doc.textFrames;

                            for (var i = 0; i < frames.length; i++) {
                                var tf = frames[i];
                                var tid = "TID_" + (i + 1);

                                if (tf.hidden || tf.locked || (tf.layer && (tf.layer.locked || !tf.layer.visible))) {
                                    destFile.writeln("SKIP|||" + tid + "|||" + tf.name);
                                    continue;
                                }

                                try {
                                    var ranges = tf.textRanges;
                                    var frameData = [];
                                    
                                    var current_font = "", current_size = "", current_color = "";
                                    var current_base = "", current_caps = "", current_just = "", current_under = "";
                                    var current_chunk = "";
                                    
                                    for (var j = 0; j < ranges.length; j++) {
                                        var tr = ranges[j];
                                        var txt = tr.contents;
                                        if (!txt) continue;

                                        txt = txt.replace(/\\r/g, "[BR]").replace(/\\n/g, "[BR]");
                                        txt = txt.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

                                        var font = "", size = "", hex = "", base = "", caps = "", just = "", under = "";
                                        
                                        try { font = tr.characterAttributes.textFont.name; } catch(e) {}
                                        try { size = Math.round(tr.characterAttributes.size * 100) / 100; } catch(e) {}
                                        try {
                                            var color = tr.characterAttributes.fillColor;
                                            if (color.typename == 'RGBColor') {
                                                var r = Math.round(color.red).toString(16); if(r.length==1) r="0"+r;
                                                var g = Math.round(color.green).toString(16); if(g.length==1) g="0"+g;
                                                var b = Math.round(color.blue).toString(16); if(b.length==1) b="0"+b;
                                                hex = "#" + r.toUpperCase() + g.toUpperCase() + b.toUpperCase();
                                            }
                                        } catch(e) {}
                                        
                                        try {
                                            var bOpt = tr.characterAttributes.baselinePosition;
                                            if (bOpt == FontBaselineOption.SUPERSCRIPT) base = "super";
                                            else if (bOpt == FontBaselineOption.SUBSCRIPT) base = "sub";
                                        } catch(e) {}
                                        
                                        try {
                                            var cOpt = tr.characterAttributes.capitalization;
                                            if (cOpt == FontCapsOption.ALLCAPS) caps = "all";
                                            else if (cOpt == FontCapsOption.SMALLCAPS) caps = "small";
                                        } catch(e) {}
                                        
                                        try { if (tr.characterAttributes.underline) under = "yes"; } catch(e) {}
                                        
                                        try {
                                            var jOpt = tr.paragraphAttributes.justification;
                                            if (jOpt == Justification.CENTER) just = "center";
                                            else if (jOpt == Justification.RIGHT) just = "right";
                                            else if (jOpt == Justification.FULLJUSTIFY) just = "full";
                                            else just = "left"; 
                                        } catch(e) {}
                                        
                                        if (font == current_font && size == current_size && hex == current_color && 
                                            base == current_base && caps == current_caps && under == current_under && just == current_just) {
                                            current_chunk += txt;
                                        } else {
                                            if (current_chunk) {
                                                var font_attr = current_font ? ' font="' + current_font + '"' : '';
                                                var size_attr = current_size ? ' size="' + current_size + '"' : '';
                                                var col_attr  = current_color ? ' color="' + current_color + '"' : '';
                                                var base_attr = current_base ? ' base="' + current_base + '"' : '';
                                                var caps_attr = current_caps ? ' caps="' + current_caps + '"' : '';
                                                var under_attr = current_under ? ' under="yes"' : '';
                                                var just_attr = current_just ? ' just="' + current_just + '"' : '';
                                                
                                                frameData.push('<cf' + font_attr + size_attr + col_attr + base_attr + caps_attr + under_attr + just_attr + '>' + current_chunk + '</cf>');
                                            }
                                            current_font = font; current_size = size; current_color = hex;
                                            current_base = base; current_caps = caps; current_under = under; current_just = just;
                                            current_chunk = txt;
                                        }
                                    }
                                    
                                    if (current_chunk) {
                                        var font_attr = current_font ? ' font="' + current_font + '"' : '';
                                        var size_attr = current_size ? ' size="' + current_size + '"' : '';
                                        var col_attr  = current_color ? ' color="' + current_color + '"' : '';
                                        var base_attr = current_base ? ' base="' + current_base + '"' : '';
                                        var caps_attr = current_caps ? ' caps="' + current_caps + '"' : '';
                                        var under_attr = current_under ? ' under="yes"' : '';
                                        var just_attr = current_just ? ' just="' + current_just + '"' : '';
                                        
                                        frameData.push('<cf' + font_attr + size_attr + col_attr + base_attr + caps_attr + under_attr + just_attr + '>' + current_chunk + '</cf>');
                                    }
                                    
                                    if (frameData.length > 0) {
                                        destFile.writeln("OK|||" + tid + "|||" + frameData.join(""));
                                    }
                                } catch(e) {
                                    destFile.writeln("ERR|||" + tid + "|||" + e.message);
                                }
                            }
                            destFile.close();
                            """
                            
                            with open(jsx_path, 'w', encoding='utf-8') as f:
                                f.write(jsx_code)
                                
                            import time as pytime
                            start_time = pytime.time()
                            ai_app.DoJavaScriptFile(jsx_path)
                            elapsed = pytime.time() - start_time
                            if self.signals: self.signals.log.emit(f"  -> Extraction finished in {elapsed:.2f} seconds!", False)
                            
                            if os.path.exists(out_txt_path):
                                with open(out_txt_path, 'r', encoding='utf-8') as f:
                                    lines = f.readlines()
                                    for line in lines:
                                        parts = line.strip().split('|||')
                                        if len(parts) >= 3:
                                            status = parts[0]
                                            tid = parts[1]
                                            content = parts[2]
                                            
                                            if status == "SKIP":
                                                self.hidden_layers.append(f"{fname} -> {content}")
                                            elif status == "OK":
                                                final_xml = content.replace("[BR]", "\n")
                                                ai_extracted_items.append({'file': fname, 'layer': tid, 'source': final_xml})
                                
                                try:
                                    os.remove(jsx_path)
                                    os.remove(out_txt_path)
                                except: pass

                            doc.Close(2)
                            self.stats['ai_processed'] += 1
                            
                    except Exception as e:
                        self.stats['fail'] += 1
                        if self.signals: self.signals.log.emit(f"Failed to extract {fname}: {e}", True)

                prefix = f"{self.base_name}_" if self.base_name else ""
                
                if ps_extracted_items:
                    ps_out = os.path.join(self.out_dir, f"{prefix}PS.docx")
                    self._build_dtp_word(ps_extracted_items, ps_out)
                    if self.signals: self.signals.log.emit(f"✓ Saved PS Word to: {Path(ps_out).name}", False)
                if ai_extracted_items:
                    ai_out = os.path.join(self.out_dir, f"{prefix}AI.xml")
                    self._build_dtp_xml(ai_extracted_items, ai_out)
                    if self.signals: self.signals.log.emit(f"✓ Saved AI XML to: {Path(ai_out).name}", False)
                
                summary = self._generate_summary("EXPORT")
                if self.signals: self.signals.finished.emit("DONE", summary)

            elif self.mode == 'import':
                master_translations = {}
                if self.trans_paths:
                    for trans_file in self.trans_paths:
                        import_ext = Path(trans_file).suffix.lower()
                        if import_ext == '.docx':
                            tmp_dict = self._read_dtp_word(trans_file)
                        elif import_ext == '.xml':
                            tmp_dict = self._read_dtp_xml(trans_file)
                        else: continue
                        
                        for f_name, layer_dict in tmp_dict.items():
                            if f_name not in master_translations:
                                master_translations[f_name] = {}
                            master_translations[f_name].update(layer_dict)

                for idx, file_path in enumerate(self.files, 1):
                    ext = Path(file_path).suffix.lower()
                    fname = Path(file_path).name
                    
                    if self.signals: 
                        self.signals.log.emit(f"[{idx}/{total_files}] Injecting {fname}...", False)
                        self.signals.progress.emit(idx, total_files)
                    
                    try:
                        if ext in ['.psd', '.psb']:
                            if fname not in master_translations:
                                if self.signals: self.signals.log.emit(f"  -> Skipped: No translations found for {fname}.", True)
                                continue
                            file_dict = master_translations[fname]
                            
                            if not ps_app: 
                                ps_app = win32com.client.Dispatch("Photoshop.Application")
                                ps_app.DisplayDialogs = 3
                            
                            doc = ps_app.Open(file_path)
                            text_counter = [0]
                            def traverse_ps_import(layers):
                                for i in range(layers.Count):
                                    layer = layers.Item(i + 1)
                                    if layer.LayerType == 2:
                                        traverse_ps_import(layer.Layers)
                                    elif layer.LayerType == 1 and layer.Kind == 2:
                                        text_counter[0] += 1
                                        tid = f"TID_{text_counter[0]}"
                                        if tid in file_dict and file_dict[tid]:
                                            layer.TextItem.Contents = file_dict[tid]
                            traverse_ps_import(doc.Layers)
                            doc.Save()
                            doc.Close(2)
                            if self.signals: self.signals.log.emit(f"  -> {fname} injected successfully.", False)
                            self.stats['ps_processed'] += 1

                        elif ext in ['.ai', '.eps', '.ps']:
                            if fname not in master_translations:
                                if self.signals: self.signals.log.emit(f"  -> Skipped: No translations found for {fname}.", True)
                                continue
                            file_dict = master_translations[fname]
                            
                            if not ai_app: 
                                ai_app = win32com.client.Dispatch("Illustrator.Application")
                                ai_app.UserInteractionLevel = -1
                            
                            doc = ai_app.Open(file_path)
                            total_frames = doc.TextFrames.Count
                            
                            import tempfile
                            temp_dir = tempfile.gettempdir()
                            data_path = os.path.join(temp_dir, "ai_import_data.txt").replace('\\', '/')
                            jsx_path = os.path.join(temp_dir, "ai_import_fast.jsx").replace('\\', '/')
                            
                            if self.signals: self.signals.log.emit(f"  -> Injecting JSX Engine for 6D Anti-Bleeding rendering...", False)
                            
                            with open(data_path, 'w', encoding='utf-8') as f:
                                for tid, target_xml in file_dict.items():
                                    if not target_xml: continue
                                    
                                    import re
                                    parts = re.split(r'(?i)(<cf[^>]*>.*?</cf>)', target_xml, flags=re.DOTALL)
                                    
                                    for part in parts:
                                        if not part: continue
                                        part_clean = part.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')
                                        
                                        m = re.match(r'(?i)<cf([^>]*)>(.*?)</cf>', part, flags=re.DOTALL)
                                        if m:
                                            attrs = m.group(1)
                                            text = m.group(2)
                                            text = text.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')
                                            text = text.replace('\n', '[BR]').replace('\r', '[BR]')
                                            
                                            font_m = re.search(r'(?i)font=["\'](.*?)["\']', attrs)
                                            size_m = re.search(r'(?i)size=["\'](.*?)["\']', attrs)
                                            color_m = re.search(r'(?i)color=["\'](.*?)["\']', attrs)
                                            base_m = re.search(r'(?i)base=["\'](.*?)["\']', attrs)
                                            caps_m = re.search(r'(?i)caps=["\'](.*?)["\']', attrs)
                                            just_m = re.search(r'(?i)just=["\'](.*?)["\']', attrs)
                                            under_m = re.search(r'(?i)under=["\'](.*?)["\']', attrs)
                                            
                                            f_val = font_m.group(1) if font_m else ""
                                            s_val = size_m.group(1) if size_m else ""
                                            c_val = color_m.group(1).lstrip('#') if color_m else ""
                                            b_val = base_m.group(1) if base_m else ""
                                            cp_val = caps_m.group(1) if caps_m else ""
                                            j_val = just_m.group(1) if just_m else ""
                                            u_val = "yes" if under_m else ""
                                            
                                            f.write(f"{tid}|||{f_val}|||{s_val}|||{c_val}|||{b_val}|||{cp_val}|||{j_val}|||{u_val}|||{text}\n")
                                        else:
                                            text = part_clean.replace('\n', '[BR]').replace('\r', '[BR]')
                                            f.write(f"{tid}||||||||||||||||||{text}\n")
                                            
                            jsx_code = 'var dataFile = new File("' + data_path + '");\n'
                            jsx_code += """
                            dataFile.open("r");
                            dataFile.encoding = "UTF-8";
                            
                            var dict = {};
                            while (!dataFile.eof) {
                                var line = dataFile.readln();
                                if (!line) continue;
                                var parts = line.split("|||");
                                if (parts.length >= 9) {
                                    var tid = parts[0];
                                    var text = parts.slice(8).join("|||").replace(/\\[BR\\]/g, "\\r");
                                    
                                    if (!dict[tid]) dict[tid] = [];
                                    dict[tid].push({
                                        font: parts[1], size: parts[2], color: parts[3], 
                                        base: parts[4], caps: parts[5], just: parts[6], under: parts[7], text: text
                                    });
                                }
                            }
                            dataFile.close();
                            
                            var doc = app.activeDocument;
                            var frames = doc.textFrames;
                            
                            var defaultBlack = new RGBColor(); 
                            defaultBlack.red = 0; defaultBlack.green = 0; defaultBlack.blue = 0;
                            
                            for (var i = 0; i < frames.length; i++) {
                                var tf = frames[i];
                                var tid = "TID_" + (i + 1);
                                
                                if (dict[tid]) {
                                    var chunks = dict[tid];
                                    var full_text = "";
                                    for (var c=0; c<chunks.length; c++) full_text += chunks[c].text;
                                    
                                    var was_locked = false, was_hidden = false, p_locked = false;
                                    var original_pos = null;
                                    try { 
                                        was_locked = tf.locked; 
                                        was_hidden = tf.hidden; 
                                        p_locked = tf.layer.locked; 
                                        if (p_locked) tf.layer.locked = false; 
                                        if (was_locked) tf.locked = false; 
                                        if (was_hidden) tf.hidden = false; 
                                        original_pos = tf.position;
                                    } catch(e){}
                                    
                                    tf.contents = full_text;
                                    var act_len = tf.characters.length;
                                    
                                    var char_idx = 0;
                                    for (var c=0; c<chunks.length; c++) {
                                        var chk = chunks[c];
                                        var t_len = chk.text.length;
                                        if (t_len == 0) continue;
                                        
                                        var targetColor = null;
                                        if (chk.color && chk.color.length == 6) {
                                            targetColor = new RGBColor();
                                            targetColor.red = parseInt(chk.color.substring(0,2), 16);
                                            targetColor.green = parseInt(chk.color.substring(2,4), 16);
                                            targetColor.blue = parseInt(chk.color.substring(4,6), 16);
                                        }
                                        
                                        var targetFont = null;
                                        if (chk.font) { try { targetFont = app.textFonts.getByName(chk.font); } catch(e){} }
                                        
                                        for (var k=0; k<t_len; k++) {
                                            var pos = char_idx + k;
                                            if (pos >= act_len) break;
                                            var charCode = chk.text.charCodeAt(k);
                                            if (charCode === 13 || charCode === 10 || charCode === 32 || charCode === 9) continue;
                                            
                                            try {
                                                var charObj = tf.characters[pos];
                                                
                                                if (targetFont) { try { charObj.characterAttributes.textFont = targetFont; } catch(e){} }
                                                if (chk.size) { try { charObj.characterAttributes.size = parseFloat(chk.size); } catch(e){} }
                                                if (targetColor) { try { charObj.characterAttributes.fillColor = targetColor; } catch(e){} }
                                                
                                                try {
                                                    if (chk.base == "super") charObj.characterAttributes.baselinePosition = FontBaselineOption.SUPERSCRIPT;
                                                    else if (chk.base == "sub") charObj.characterAttributes.baselinePosition = FontBaselineOption.SUBSCRIPT;
                                                    else charObj.characterAttributes.baselinePosition = FontBaselineOption.NORMALBASELINE;
                                                } catch(e){}
                                                
                                                try {
                                                    if (chk.caps == "all") charObj.characterAttributes.capitalization = FontCapsOption.ALLCAPS;
                                                    else if (chk.caps == "small") charObj.characterAttributes.capitalization = FontCapsOption.SMALLCAPS;
                                                    else charObj.characterAttributes.capitalization = FontCapsOption.NORMALCAPS;
                                                } catch(e){}
                                                
                                                if (chk.under == "yes") { try { charObj.characterAttributes.underline = true; } catch(e){} }
                                                else { try { charObj.characterAttributes.underline = false; } catch(e){} }
                                                
                                                try {
                                                    if (chk.just == "center") charObj.paragraphAttributes.justification = Justification.CENTER;
                                                    else if (chk.just == "right") charObj.paragraphAttributes.justification = Justification.RIGHT;
                                                    else if (chk.just == "full") charObj.paragraphAttributes.justification = Justification.FULLJUSTIFY;
                                                    else charObj.paragraphAttributes.justification = Justification.LEFT;
                                                } catch(e){}

                                            } catch(e) {}
                                        }
                                        char_idx += t_len;
                                    }
                                    
                                    try { 
                                        if (original_pos) tf.position = original_pos;
                                        if (was_locked) tf.locked = true; 
                                        if (was_hidden) tf.hidden = true; 
                                        if (p_locked) tf.layer.locked = true; 
                                    } catch(e){}
                                }
                            }
                            """
                            
                            with open(jsx_path, 'w', encoding='utf-8') as f:
                                f.write(jsx_code)
                                
                            import time as pytime
                            start_time = pytime.time()
                            ai_app.DoJavaScriptFile(jsx_path)
                            elapsed = pytime.time() - start_time
                            if self.signals: self.signals.log.emit(f"  -> Injection finished in {elapsed:.2f} seconds!", False)
                            
                            try:
                                os.remove(jsx_path)
                                os.remove(data_path)
                            except: pass

                            doc.Save()
                            doc.Close(2)
                            self.stats['ai_processed'] += 1
                            
                    except Exception as e:
                        self.stats['fail'] += 1
                        if self.signals: self.signals.log.emit(f"Failed to inject {fname}: {e}", True)
                        
                summary = self._generate_summary("IMPORT")
                if self.signals: self.signals.finished.emit("DONE", summary)

        except Exception as e:
            if self.signals: self.signals.error.emit(f"Adobe COM Runtime Error: {str(e)}\n{traceback.format_exc()}")
        finally:
            pythoncom.CoUninitialize()

    def _generate_summary(self, mode_title):
        s = f"<br><b>{'='*15} {mode_title} SUMMARY {'='*15}</b><br>"
        s += f"▶<br>"
        s += f"Total Files: {self.stats['total']}<br>"
        s += f"PS Processed: {self.stats['ps_processed']}<br>"
        s += f"AI Processed: {self.stats['ai_processed']}<br>"
        s += f"Failed: {self.stats['fail']}<br>"
        s += f"<b>{'='*46}</b><br>"
        
        if self.hidden_layers:
            s += f"<br><b>[WARNING] Found {len(self.hidden_layers)} HIDDEN layers (Skipped):</b><br>"
            for lyr in set(self.hidden_layers): s += f"  - {lyr}<br>"
        if self.locked_layers:
            s += f"<br><b>[WARNING] Found {len(self.locked_layers)} LOCKED layers (Skipped):</b><br>"
            for lyr in set(self.locked_layers): s += f"  - {lyr}<br>"
        if self.skipped_layers:
            s += f"<br><b>[INFO] Skipped {len(self.skipped_layers)} empty translations.</b><br>"
        if self.failed_layers:
            s += f"<br><b>[ERROR] Failed to inject specific layers:</b><br>"
            for lyr in set(self.failed_layers): s += f"  - {lyr}<br>"
        return s

    def _build_dtp_word(self, items, output_path):
        try:
            from docx import Document
            from docx.shared import Inches
            from docx.oxml.ns import qn
            from docx.oxml import OxmlElement
        except ImportError: return
        
        doc = Document(); section = doc.sections[0]
        section.page_width = Inches(8.5); section.page_height = Inches(11)
        section.left_margin = Inches(0.5); section.right_margin = Inches(0.5)
        table = doc.add_table(rows=0, cols=2); table.style = 'Table Grid'
        tblPr = table._tbl.tblPr; tblLayout = OxmlElement('w:tblLayout')
        tblLayout.set(qn('w:type'), 'fixed'); tblPr.append(tblLayout)

        for item in items:
            row = table.add_row()
            cell_0 = row.cells[0]; p0 = cell_0.paragraphs[0]
            r0 = p0.add_run(f"{item['file']} || {item['layer']}")
            rPr = OxmlElement('w:rPr'); vanish = OxmlElement('w:vanish')
            rPr.append(vanish); r0._r.get_or_add_rPr().append(rPr)
            
            cell_1 = row.cells[1]
            cell_1.text = item['source']
            cell_0.width = Inches(1.5); cell_1.width = Inches(6.0)
            
        try: doc.save(output_path)
        except PermissionError: raise PermissionError(f"Please close '{Path(output_path).name}' in Microsoft Word!")

    def _read_dtp_word(self, input_path):
        from docx import Document
        doc = Document(input_path)
        translations = {}
        for table in doc.tables:
            for row in table.rows:
                if len(row.cells) >= 2:
                    meta_text = row.cells[0].text.strip()
                    target_text = row.cells[1].text.strip()
                    if not meta_text or not target_text: continue
                    parts = meta_text.split('||')
                    if len(parts) == 2:
                        fname = parts[0].strip()
                        import re
                        tid_match = re.search(r'(TID_\d+)', parts[1])
                        if tid_match:
                            if fname not in translations: translations[fname] = {}
                            translations[fname][tid_match.group(1)] = target_text
        return translations

    def _build_dtp_xml(self, items, output_path):
        import xml.etree.ElementTree as ET
        root = ET.Element("localization")
        files_dict = {}
        for item in items:
            fname = item['file']
            if fname not in files_dict: files_dict[fname] = []
            files_dict[fname].append(item)
            
        for fname, contents in files_dict.items():
            file_node = ET.SubElement(root, "file", name=fname)
            for item in contents:
                item_node = ET.SubElement(file_node, "item", layer=item['layer'])
                source_str = item['source']
                if '<cf ' in source_str:
                    dummy_xml = f"<root>{source_str}</root>"
                    try:
                        parsed_dummy = ET.fromstring(dummy_xml)
                        for child in parsed_dummy: item_node.append(child)
                    except: item_node.text = source_str
                else: item_node.text = source_str
                
        def indent(elem, level=0):
            i = "\n" + level*"    "
            if len(elem):
                if not elem.text or not elem.text.strip(): elem.text = i + "    "
                if not elem.tail or not elem.tail.strip(): elem.tail = i
                for elem in elem: indent(elem, level+1)
                if not elem.tail or not elem.tail.strip(): elem.tail = i
            else:
                if level and (not elem.tail or not elem.tail.strip()): elem.tail = i
        indent(root)
        tree = ET.ElementTree(root)
        tree.write(output_path, encoding="utf-8", xml_declaration=True)

    def _read_dtp_xml(self, input_path):
        import xml.etree.ElementTree as ET
        tree = ET.parse(input_path)
        root = tree.getroot()
        translations = {}
        for file_node in root.findall('file'):
            fname = file_node.get('name')
            if fname not in translations: translations[fname] = {}
            for item_node in file_node.findall('item'):
                layer = item_node.get('layer')
                inner_xml = (item_node.text or "") + "".join(ET.tostring(c, encoding='unicode') for c in item_node)
                if inner_xml.strip(): translations[fname][layer] = inner_xml.strip()
        return translations


class AdobeDtpTab(QWidget):
    """Tab 5: 专属的 DTP 工程文件处理界面 (Base Name 修复)"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_items = []
        self._thread_pool = QThreadPool()
        self.valid_exts = ('.psd', '.psb', '.ai', '.eps', '.ps')
        self._build_ui()

    def _build_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        toolbar = QFrame(); toolbar.setFixedHeight(50); toolbar.setStyleSheet('background: #1a1d27; border-bottom: 1px solid #2e3250;')
        tb_layout = QHBoxLayout(toolbar); tb_layout.setContentsMargins(12, 0, 12, 0)
        
        self.btn_add_files = QPushButton("📄 Add Files"); self.btn_add_files.setFixedHeight(28)
        self.btn_add_folder = QPushButton("📁 Add Folder"); self.btn_add_folder.setFixedHeight(28)
        
        self.base_name_input = QLineEdit()
        self.base_name_input.setPlaceholderText("Export Base Name (Optional)")
        self.base_name_input.setFixedWidth(180)
        self.base_name_input.setFixedHeight(28)
        self.base_name_input.setStyleSheet('background: #0f1117; border: 1px solid #2e3250; color: #e2e8f0; padding: 0 8px;')
        
        self.btn_export = QPushButton(T('dtp_export')); self.btn_export.setStyleSheet('background: #4f8ef7; color: white; padding: 6px 16px; font-weight: bold;')
        self.btn_import = QPushButton(T('dtp_import')); self.btn_import.setStyleSheet('background: #34d399; color: #0a1a12; padding: 6px 16px; font-weight: bold;')
        
        self.btn_reset = QPushButton("↻ Reset")
        self.btn_reset.setStyleSheet('background: #f59e0b; color: #1a1000; padding: 6px 16px; font-weight: bold;')
        
        tb_layout.addWidget(self.btn_add_files); tb_layout.addWidget(self.btn_add_folder)
        tb_layout.addSpacing(15)
        tb_layout.addWidget(self.base_name_input)
        tb_layout.addWidget(self.btn_export)
        tb_layout.addSpacing(15)
        tb_layout.addWidget(self.btn_import)
        tb_layout.addStretch()
        tb_layout.addWidget(self.btn_reset)
        main_layout.addWidget(toolbar)

        progress_bar_container = QFrame()
        progress_bar_container.setStyleSheet('background: #1a1d27; border-bottom: 1px solid #2e3250; padding: 4px 12px;')
        pb_layout = QHBoxLayout(progress_bar_container)
        pb_layout.setContentsMargins(0, 0, 0, 0)
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(8)
        self.progress_bar.setTextVisible(False)
        self.progress_label = QLabel("Ready")
        self.progress_label.setStyleSheet('color: #64748b; font-size: 11px; font-family: monospace;')
        pb_layout.addWidget(self.progress_bar)
        pb_layout.addWidget(self.progress_label)
        main_layout.addWidget(progress_bar_container)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        left_panel = QWidget(); left_layout = QVBoxLayout(left_panel); left_layout.setContentsMargins(0,0,0,0)
        self.list_widget = DtpFileListWidget()
        self.list_widget.files_dropped.connect(self._add_files)
        self.list_widget.file_deleted.connect(self._remove_file)
        left_layout.addWidget(self.list_widget)
        splitter.addWidget(left_panel)
        
        self.log_console = QPlainTextEdit(); self.log_console.setReadOnly(True)
        self.log_console.setStyleSheet("background-color: #1e1e1e; color: #d4d4d4; font-family: 'Consolas'; font-size: 12px; border: none;")
        splitter.addWidget(self.log_console)
        
        splitter.setSizes([300, 900])
        main_layout.addWidget(splitter, 1)

        self.btn_add_files.clicked.connect(self._on_add_files_btn)
        self.btn_add_folder.clicked.connect(self._on_add_folder_btn)
        self.btn_export.clicked.connect(self._on_export)
        self.btn_import.clicked.connect(self._on_import)
        self.btn_reset.clicked.connect(self._on_reset)

    def _on_add_files_btn(self):
        paths, _ = QFileDialog.getOpenFileNames(self, 'Select Adobe DTP Files', '', 'Adobe Files (*.psd *.psb *.ai *.eps *.ps)')
        if paths: self._add_files(paths)

    def _on_add_folder_btn(self):
        folder = QFileDialog.getExistingDirectory(self, 'Select Folder')
        if folder:
            paths = [os.path.join(r, f) for r, _, fs in os.walk(folder) for f in fs if f.lower().endswith(self.valid_exts)]
            if paths: self._add_files(paths)

    def _add_files(self, paths):
        added = 0
        for p in paths:
            if p not in self.file_items:
                self.file_items.append(p)
                item = QListWidgetItem(Path(p).name)
                item.setData(Qt.ItemDataRole.UserRole, p)
                self.list_widget.addItem(item)
                added += 1
        if added: 
            self._log(f"Added {added} Adobe file(s) ready for processing.")
            self.progress_label.setText(f"{len(self.file_items)} files loaded")

    def _remove_file(self, path):
        if path in self.file_items:
            self.file_items.remove(path)
            self._log(f"Removed: {Path(path).name}")
            self.progress_label.setText(f"{len(self.file_items)} files loaded")

    def _on_reset(self):
        self.file_items.clear()
        self.list_widget.clear()
        self.log_console.clear()
        self.base_name_input.clear()
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(100)
        self.progress_label.setText("Ready")
        self._log("DTP Engine Reset - Ready for new batch.", False)

    def _log(self, text, is_warn=False):
        color = "#ef4444" if is_warn else "#d4d4d4"
        if "SUMMARY" in text or "DONE" in text or "Export complete" in text: color = "#34d399"
        self.log_console.appendHtml(f"<span style='color: {color};'>[{time.strftime('%H:%M:%S')}] {text}</span>")
        sb = self.log_console.verticalScrollBar(); sb.setValue(sb.maximum())

    def _update_progress(self, current, total):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        self.progress_label.setText(f"Processing File: {current} / {total}")

    def _on_worker_finished(self, status, summary_text):
        self.btn_export.setEnabled(True)
        self.btn_import.setEnabled(True)
        self._log(summary_text)
        self.progress_label.setText("Job Completed.")

    def _on_export(self):
        if not self.file_items: return
        
        base_dir = str(Path(self.file_items[0]).parent)
        base_name = self.base_name_input.text().strip()
        out_base_path = os.path.join(base_dir, base_name) if base_name else os.path.join(base_dir, "")
        
        self.btn_export.setEnabled(False); self.btn_import.setEnabled(False)
        self.progress_bar.setValue(0)
        self._log(f"Starting extraction via Adobe COM API for {len(self.file_items)} files...")
        
        signals = DtpWorkerSignals()
        signals.log.connect(self._log)
        signals.progress.connect(self._update_progress)
        signals.finished.connect(self._on_worker_finished)
        signals.error.connect(lambda msg: self._log(msg, is_warn=True))
        
        worker = AdobeDtpWorker(mode='export', files=self.file_items, out_dir=base_dir, base_name=base_name, trans_paths=None, signals=signals)
        worker.setAutoDelete(True); self._thread_pool.start(worker)

    def _on_import(self):
        if not self.file_items: return
        
        default_dir = str(Path(self.file_items[0]).parent)
        file_paths, _ = QFileDialog.getOpenFileNames(self, 'Select Translated Word AND/OR XML', default_dir, 'Translation Files (*.docx *.xml)')
        if not file_paths: return

        self.btn_export.setEnabled(False); self.btn_import.setEnabled(False)
        self.progress_bar.setValue(0)
        self._log(f"Reading translations from {len(file_paths)} files and initializing Adobe injection...")
        
        signals = DtpWorkerSignals()
        signals.log.connect(self._log)
        signals.progress.connect(self._update_progress)
        signals.finished.connect(self._on_worker_finished)
        signals.error.connect(lambda msg: self._log(msg, is_warn=True))
        
        worker = AdobeDtpWorker(mode='import', files=self.file_items, out_dir="", base_name="", trans_paths=file_paths, signals=signals)
        worker.setAutoDelete(True); self._thread_pool.start(worker)

    def _retranslate_ui(self):
        self.btn_add_files.setText(T('add_files'))
        self.btn_add_folder.setText(T('add_folder'))
        self.btn_export.setText(T('dtp_export'))
        self.btn_import.setText(T('dtp_import'))
        self.btn_reset.setText(T('reset'))


# ─────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────
# MAIN WINDOW 
# ─────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.setWindowIcon(QIcon('image2ATF.jpg'))
        self.resize(1280, 820)
        self._apply_dark_theme()
        self._build_ui()

    def _apply_dark_theme(self):
        self.setStyleSheet("""
            QMainWindow, QWidget { background: #0f1117; color: #e2e8f0; }
            QSplitter::handle { background: #2e3250; }
            QTabWidget::pane { border: none; }
            QTabBar::tab {
                background: #1a1d27; color: #64748b;
                padding: 8px 20px; border: none;
                border-bottom: 3px solid transparent;
                font-size: 13px; font-weight: 500;
            }
            QTabBar::tab:selected { color: #4f8ef7; border-bottom-color: #4f8ef7; }
            QTabBar::tab:hover:!selected { color: #e2e8f0; }
            QListWidget {
                background: #1a1d27; border: none;
                border-right: 1px solid #2e3250;
            }
            QListWidget::item { padding: 6px 8px; border-radius: 4px; }
            QListWidget::item:selected { background: #222538; }
            QListWidget::item:hover { background: #1e2132; }
            QLineEdit, QComboBox, QPlainTextEdit {
                background: #0f1117; border: 1px solid #2e3250;
                color: #e2e8f0; border-radius: 5px; padding: 5px 8px;
                font-size: 13px;
            }
            QLineEdit:focus, QComboBox:focus, QPlainTextEdit:focus {
                border-color: #4f8ef7;
            }
            QComboBox::drop-down { border: none; width: 20px; }
            QComboBox::down-arrow { color: #94a3b8; }
            QComboBox QAbstractItemView {
                background: #1a1d27; border: 1px solid #2e3250;
                selection-background-color: #222538;
            }
            QPushButton {
                background: #1a1d27; color: #e2e8f0;
                border: 1px solid #2e3250; border-radius: 6px;
                padding: 5px 14px; font-size: 12px;
            }
            QPushButton:hover { border-color: #4f8ef7; color: #4f8ef7; }
            QPushButton:pressed { background: #222538; }
            QPushButton:disabled { opacity: 0.4; }
            QProgressBar {
                background: #222538; border: 1px solid #2e3250;
                border-radius: 3px; height: 6px;
            }
            QProgressBar::chunk { background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #4f8ef7, stop:1 #a78bfa); border-radius: 3px; }
            QCheckBox { color: #94a3b8; }
            QCheckBox::indicator { width: 14px; height: 14px; border: 1px solid #2e3250; border-radius: 3px; background: #0f1117; }
            QCheckBox::indicator:checked { background: #4f8ef7; border-color: #4f8ef7; }
            QFrame[frameShape="4"] { color: #2e3250; }
            QTextEdit, QPlainTextEdit { background: #0f1117; }
            QScrollBar:vertical { width: 6px; background: transparent; }
            QScrollBar::handle:vertical { background: #2e3250; border-radius: 3px; min-height: 20px; }
            QStatusBar { background: #1a1d27; border-top: 1px solid #2e3250; color: #64748b; font-size: 11px; }
            QLabel { color: #e2e8f0; }
        """)

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Proxy widget
        proxy_bar = QFrame()
        proxy_bar.setStyleSheet('background: #1a1d27; border-bottom: 1px solid #2e3250;')
        pb = QHBoxLayout(proxy_bar)
        pb.setContentsMargins(12, 4, 12, 4)
        self.proxy_lbl = QLabel(T('proxy_lbl'))
        pb.addWidget(self.proxy_lbl)
        self.proxy_widget = ProxyWidget()
        pb.addWidget(self.proxy_widget)
        pb.addStretch()
        
        # Language switcher
        self.btn_lang = QPushButton(T('lang_btn'))
        self.btn_lang.setFixedSize(72, 26)
        self.btn_lang.setStyleSheet('background: #2e3250; color: #4f8ef7; border: 1px solid #4f8ef7; border-radius: 5px; font-size: 12px; font-weight: 600;')
        self.btn_lang.clicked.connect(self._switch_lang)
        pb.addWidget(self.btn_lang)
        
        # Readme 按钮
        self.btn_readme = QPushButton("❓ Readme")
        self.btn_readme.setFixedSize(85, 26)
        self.btn_readme.setStyleSheet('background: #2e3250; color: #34d399; border: 1px solid #34d399; border-radius: 5px; font-size: 12px; font-weight: 600;')
        self.btn_readme.clicked.connect(self._show_readme)
        pb.addWidget(self.btn_readme)
        
        layout.addWidget(proxy_bar)

        # Tab widget
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        layout.addWidget(self.tabs, 1)

        self.api_tabs = {}
        configs = [
            ('gemini', 'tab_gemini', GEMINI_MODELS),
            ('claude', 'tab_claude', CLAUDE_MODELS),
        ]
        for api_name, tab_key, models in configs:
            tab = ApiTab(api_name, models, self.proxy_widget)
            self.api_tabs[api_name] = tab
            self.tabs.addTab(tab, T(tab_key))

        self.mask_tab = MaskApiTab('gemini', GEMINI_MODELS, self.proxy_widget)
        self.tabs.addTab(self.mask_tab, T('tab_mask'))

        self.image_to_psd_tab = ImageToPsdTab('gemini', GEMINI_MODELS, self.proxy_widget)
        self.tabs.addTab(self.image_to_psd_tab, T('tab_img2psd'))

        self.adobe_tab = AdobeDtpTab()
        self.tabs.addTab(self.adobe_tab, T('tab_adobe'))

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage(T('status_ready', t=MAX_THREADS))

    def _switch_lang(self):
        _LANG[0] = 'ZH' if _LANG[0] == 'EN' else 'EN'
        self.btn_lang.setText(T('lang_btn'))
        self.proxy_lbl.setText(T('proxy_lbl'))
        self.proxy_widget.checkbox.setText(T('use_proxy'))
        self.status_bar.showMessage(T('status_ready', t=MAX_THREADS))
        tab_keys = ['tab_gemini', 'tab_claude']
        for i, key in enumerate(tab_keys): self.tabs.setTabText(i, T(key))
        self.tabs.setTabText(2, T('tab_mask'))
        self.tabs.setTabText(3, T('tab_img2psd'))
        self.tabs.setTabText(4, T('tab_adobe'))
        for tab in list(self.api_tabs.values()) + [self.mask_tab, self.image_to_psd_tab, self.adobe_tab]: 
            tab._retranslate_ui()

    # 🚨 _show_readme 现在正确地缩进在这里，属于 MainWindow 的方法
    def _show_readme(self):
        import webbrowser
        lang_suffix = _LANG[0].lower() 
        filename = f"readme_{lang_suffix}.html"
        readme_path = os.path.abspath(filename)
        
        if os.path.exists(readme_path):
            webbrowser.open(f"file:///{readme_path.replace(chr(92), '/')}")
        else:
            err_msg = f"Cannot find manual file:\n{filename}\n\nPlease ensure it is in the same directory." if lang_suffix == "en" else f"找不到说明书文件：\n{filename}\n\n请确保它存放在软件同级目录下。"
            err_title = "File Not Found" if lang_suffix == "en" else "文件丢失"
            QMessageBox.warning(self, err_title, err_msg)


# ─────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────
def main():
    missing = []
    try: import requests
    except ImportError: missing.append('requests')
    try: from PIL import Image
    except ImportError: missing.append('Pillow')
    try: from docx import Document
    except ImportError: missing.append('python-docx')

    if missing:
        app = QApplication(sys.argv)
        QMessageBox.critical(None, 'Missing Dependencies', f'Please install the following:\n\npip install {" ".join(missing)}\n\npython-docx is required for Word output')
        sys.exit(1)

    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setOrganizationName(ORG_NAME)
    font = QFont('Segoe UI', 10)
    app.setFont(font)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()