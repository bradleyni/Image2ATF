"""
ATF OCR Tool  —  PyQt6
Image text extraction → dual-column ATF Word file for translation
Supports Claude / Gemini / Mistral, 5-thread concurrency, persistent API keys
Includes Tab 4 (Image->PSD Builder) and Tab 5 (Adobe DTP XML/Word Extraction).
Version: 3.0
"""

import sys, os, json, base64, re, traceback, io, time, random
import logging, logging.handlers, threading, tempfile
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
    QFrame, QSizePolicy, QPlainTextEdit, QStatusBar,
    QGroupBox, QFormLayout, QSpinBox, QFontComboBox
)

from PyQt6.QtCore import (
    Qt, pyqtSignal, QSettings, QSize, QThreadPool,
    QRunnable, QObject, pyqtSlot, QRect, QPoint, QRectF, QPointF
)
from PyQt6.QtGui import (
    QPixmap, QIcon, QDragEnterEvent, QDropEvent, QFont, QColor,
    QPainter, QPen, QBrush, QTransform, QImageReader, QPolygonF
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
# CRASH / DIAGNOSTIC LOGGING
# 目的：程序打包成 exe 后是 windowed 模式（无控制台），任何未被
# try/except 捕获的异常此前都会“静默闪退”，什么线索都留不下。
# 这里加一个落地到本地文件的日志，覆盖：未捕获异常（含跨线程信号
# 触发的槽函数）、批处理关键节点（开始/单张完成或失败/批次结束）。
# ─────────────────────────────────────────────────────────────
def _log_dir() -> str:
    base = os.environ.get('LOCALAPPDATA') or os.environ.get('APPDATA') or tempfile.gettempdir()
    d = os.path.join(base, 'Image2ATF', 'logs')
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        d = tempfile.gettempdir()
    return d

LOG_FILE = os.path.join(_log_dir(), 'image2atf.log')
logger = logging.getLogger('image2atf')

def setup_crash_logging():
    logger.setLevel(logging.DEBUG)
    if not logger.handlers:
        try:
            handler = logging.handlers.RotatingFileHandler(LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3, encoding='utf-8')
            handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] [%(threadName)s] %(message)s'))
            logger.addHandler(handler)
        except Exception:
            pass
    logger.info('=' * 70)
    logger.info(f'App start. pid={os.getpid()} platform={sys.platform} frozen={getattr(sys, "frozen", False)} exe={sys.executable}')

    def _excepthook(exc_type, exc_value, exc_tb):
        try:
            logger.critical('UNCAUGHT EXCEPTION (main thread)', exc_info=(exc_type, exc_value, exc_tb))
        except Exception:
            pass
        try:
            sys.__excepthook__(exc_type, exc_value, exc_tb)
        except Exception:
            pass

    def _thread_excepthook(args):
        try:
            tname = args.thread.name if args.thread is not None else '?'
            logger.critical(f'UNCAUGHT EXCEPTION (thread: {tname})', exc_info=(args.exc_type, args.exc_value, args.exc_traceback))
        except Exception:
            pass

    sys.excepthook = _excepthook
    try:
        threading.excepthook = _thread_excepthook
    except Exception:
        pass

# Pillow 的 ImageDraw / ImageFont(FreeType) 在多线程并发调用时并非线程安全，
# 崩溃日志已证实这是“闪退”的真实原因（崩溃时没有任何 Python 异常/退出记录，
# 说明是原生层面的崩溃，不是能被 try/except 接住的普通错误）。
# 用一把全局锁把所有会调用 ImageDraw/ImageFont 的图像合成逻辑串行化；
# 真正耗时的网络请求部分仍然保持多线程并行，速度基本不受影响。
_pil_draw_lock = threading.Lock()



# ─────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────
APP_NAME    = "Welocalize DTP 图片准备工具 V2.8"
APP_TITLE   = "Welocalize DTP 图片准备工具 V2.8"
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
    'tab_gemini':         {'EN': 'ATF-creator (Gemini)',                'ZH': 'ATF 写手 (Gemini)'},
    'tab_claude':         {'EN': 'ATF-creator (Claude)',                'ZH': 'ATF 写手 (Claude)'},
    'tab_mask':           {'EN': 'ATF-creator (Mask Mode)',             'ZH': 'ATF 写手 (Mask Mode)'},
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
    'mask_batch_box':     {'EN': '⚡ Batch AI Auto-Box',                'ZH': '⚡ AI 批量画框'},
    'mask_autobox':       {'EN': '🤖Auto-Box',             'ZH': '🤖 单图画框'},
    'mask_clearbox':      {'EN': '🗑 Clear Masks',                      'ZH': '🗑 清空遮罩'},
    'mask_engine':        {'EN': 'Engine:',                             'ZH': '引擎:'},
    
    # === Tab 4 & 5 UI ===
    'tab_img2psd':        {'EN': '🖼️ Image->PSD Builder',              'ZH': '🖼️ Image->PSD 转换'},
    'tab_adobe':          {'EN': 'PSD/AI export & import',              'ZH': 'PSD/AI 导出导入'},
    'img2psd_start':      {'EN': '▶ Build PSDs',                        'ZH': '▶ 创建 PSD'},
    'dtp_export':         {'EN': '⬆ Export (PS->Word / AI->XML)',  'ZH': '⬆ 导出文字 (PS->Word / AI->XML)'},
    'dtp_import':         {'EN': '⬇ Import Translation',                 'ZH': '⬇ 导入翻译'},
    'dtp_clear':          {'EN': '🗑 Clear Image List',                   'ZH': '🗑 清空图片列表'},
    'dtp_log_hidden':     {'EN': '⚠️ [HIDDEN] Layer: {name}',           'ZH': '⚠️ [隐藏图层] 请注意: {name}'},
    'dtp_log_locked':     {'EN': '🔒 [LOCKED] Layer: {name}',           'ZH': '🔒 [锁定图层] 请解锁: {name}'},
    'dtp_format_err':     {'EN': 'COM Error. Ensure PS/AI is installed.', 'ZH': 'COM 接口调用失败，请确保已安装相应软件。'},
    'tab4_ps_fonts':     {'EN': 'PS Fonts:',      'ZH': '默认字体:'},
    'tab4_build_single': {'EN': 'Build Single PSD',   'ZH': '单图生成 PSD'},
    'pseudo_lbl':  {'EN': 'Pseudo', 'ZH': '假翻'},
    'dedup_lines': {'EN': 'Filter duplicate sentences', 'ZH': '过滤相同句子'},
    'exclude_numeric': {'EN': 'Exclude pure-number / single-letter text', 'ZH': '排除纯数字/单字母'},
    'confirm_skip_failed': {'EN': '{n} image(s) failed OCR and will be skipped (not included in the ATF). Continue?', 'ZH': '有 {n} 张图片 OCR 失败，将被跳过（不会出现在 ATF 文档里）。是否继续？'},
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
4. EXTRACT EVERY LINE AS-IS: Output each visible line of text on its own output line, faithfully. A separate post-processing step handles paragraph joining.
5. ILLEGIBLE TEXT: Replace unreadable characters with [illegible]. Never guess.
6. EXCLUDE NON-TEXT ELEMENTS: Do NOT extract logos, brand marks, icons, watermarks, or purely decorative graphical elements.
7. READING ORDER: Top-to-bottom, left-to-right.
8. RICH TEXT MARKERS:
   - Bold text: **bold text**
   - Italic text: *italic text* (CRITICAL: Only apply if the font is CLEARLY and INTENTIONALLY slanted. Do NOT mistake normal parentheses '()', slashes '/', or condensed fonts for italics!)
   - Underline: __underlined text__ (Only if a physical line is drawn directly beneath the text)
   - Superscript: ^superscript^
   - Subscript: ~subscript~
   - Colored text: [color:#RRGGBB]text[/color]
9. SPECIAL SYMBOLS: For diameter/radius, you MUST use '∅' (U+2205). NEVER use the Latin letter 'Ø'.
10. HYPHENATION: If a word is split across two lines with a hyphen (e.g., "non- \n volatile"), you MUST KEEP the hyphen in your output.
11. NO PREAMBLE OR POSTAMBLE."""

OCR_PROMPTS = {
    'claude': _OCR_PROMPT_COMMON + """
CLAUDE-SPECIFIC RULES:
- BLANK LINES — CRITICAL: Do NOT add a blank line after every individual line of text. A blank line must ONLY appear when there is a clearly visible vertical gap (paragraph break) separating two distinct text blocks in the image. If in doubt, use NO blank line.
- CHAT/MESSAGING IMAGES: Each chat message bubble and its timestamp form one unit. Do not add blank lines between consecutive messages unless there is a large visible gap in the image.
- TECHNICAL/PATENT DIAGRAMS — CRITICAL: For network diagrams, system architecture diagrams, flowcharts, block diagrams, and patent figures, you MUST extract ALL visible text labels: box labels, reference numbers, axis labels, figure captions, page numbers, and any other readable text.
- CHEMICAL/TECHNICAL DIAGRAMS: Extract all text labels exactly.
- ITALIC & UNDERLINE DETECTION: You are prone to missing subtle italics or underlines in technical diagrams. You are explicitly instructed to meticulously check for slanted text (*italic*) and lines beneath text (__underline__). 
- COLOR: You MUST use [color:#RRGGBB]text[/color] for any clearly non-black text (blue links, colored dates, highlighted text). Do not omit color.""",
    'gemini': _OCR_PROMPT_COMMON + """
GEMINI-SPECIFIC RULES:
- LINE WRAPPING — MOST CRITICAL: Text displayed in a narrow column that wraps word-by-word or character-by-character across multiple visual lines MUST be joined into a single output line. Only start a new output line for a genuinely new and separate text element.
- NO DUPLICATE EXTRACTION: Extract each piece of text ONCE only.
- CHEMICAL/TECHNICAL DIAGRAMS: Extract all text labels exactly.
- ITALIC & UNDERLINE DETECTION: Pay extraordinary attention to text style. If a word or phrase is written in an oblique/slanted font (*text*), or has a line strictly under it (__text__), you must capture it using the markdown tags. DO NOT ignore these visual cues!
- BOLD DETECTION: Text that appears visually heavier/thicker/darker than surrounding body text is bold. Mark it **like this**. When text is white on a dark background AND appears with heavier strokes, mark it as both bold AND colored: **[color:#FFFFFF]text[/color]**
- SUPERSCRIPT/SUBSCRIPT — CRITICAL: Mark with ^super^ and ~sub~. Chemical subscript numbers MUST be marked.
- COLOR: Use [color:#RRGGBB]text[/color] for any clearly non-black text. White text on dark background = [color:#FFFFFF]text[/color].""",
    'mistral': _OCR_PROMPT_COMMON + """
MISTRAL-SPECIFIC RULES:
- COMPLETENESS — ABSOLUTELY CRITICAL: Extract EVERY piece of text in the image without exception.
- ITALIC & UNDERLINE: Meticulously inspect every word for slanting (*italic*) or lines beneath (__underline__). Do not hallucinate them, but NEVER miss them if they are physically there.
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
5. EXTREMELY IMPORTANT: Pay close attention to EVERY SINGLE character's style. If you see slanted text, mark it as *italic*. If there is a line drawn underneath the text, mark it as __underline__.
6. PRESERVE ALL RICH TEXT FORMATTING (**bold**, ~sub~, ^super^, [color:#RRGGBB]text[/color]). DO NOT drop any formatting markers.
7. The snippets often contain original document reference numbers and brackets (e.g., "[150]", "(180)", "601"). YOU MUST EXTRACT THEM. Do NOT confuse them with the RED tags."""

SETTINGS_KEYS = {
    'claude': {'key': 'claude_api_key', 'model': 'claude_model'},
    'gemini': {'key': 'gemini_api_key', 'model': 'gemini_model'},
    'mistral': {'key': 'mistral_api_key', 'model': 'mistral_model'},
    'proxy_url': 'proxy_url',
    'proxy_enabled': 'proxy_enabled',
}

CLAUDE_MODELS  = ['claude-sonnet-4-5', 'claude-opus-4-5', 'claude-haiku-4-5-20251001']
GEMINI_MODELS  = ['gemini-2.5-flash', 'gemini-2.5-pro', 'gemini-2.0-flash', 'gemini-1.5-flash', 'gemini-1.5-pro']
GEMINI_MODELS_ATF = ['gemini-3.1-pro-preview', 'gemini-3.7-flash', 'gemini-2.5-pro', 'gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-flash', 'gemini-1.5-pro']  # 仅ATF creator(Mask标签页)用，其它标签页仍用上面原来的GEMINI_MODELS
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

def parse_rich_text(text: str) -> list:
    """
    【行业金标准】基于“词法扫描 (Lexer) + 成对校验”的防御级富文本解析引擎。
    """
    import re
    
    # 0. 前置净化：修复大模型幻觉产生的全角符号
    text = text.replace('＊', '*').replace('～', '~').replace('＿', '_').replace('＾', '^')
    text = text.replace(r'\*', '*').replace(r'\_', '_').replace(r'\~', '~').replace(r'\^', '^')
    
    # 💡 修复问题 1：自动补全孤立的离子电荷上标 (如 ^+, ^2+, ^- -> ^+^, ^2+^, ^-^)
    # 只要 ^ 后面跟着可选的数字和正负号，且没有闭合，就强制给它套上闭合符
    text = re.sub(r'\^([0-9]*[+\-])(?!\^)', r'^\1^', text)

    # 防止 Markdown 与 Color 标签发生死锁干扰
    text = re.sub(r'\*\*\[color:(#[0-9A-Fa-f]{6})\](.*?)\[/color\]\*\*', r'[color:\1]**\2**[/color]', text, flags=re.IGNORECASE)
    text = re.sub(r'\*\[color:(#[0-9A-Fa-f]{6})\](.*?)\[/color\]\*', r'[color:\1]*\2*[/color]', text, flags=re.IGNORECASE)
    text = re.sub(r'__\[color:(#[0-9A-Fa-f]{6})\](.*?)\[/color\]__', r'[color:\1]__\2__[/color]', text, flags=re.IGNORECASE)
    text = re.sub(r'\^\[color:(#[0-9A-Fa-f]{6})\](.*?)\[/color\]\^', r'[color:\1]^\2^[/color]', text, flags=re.IGNORECASE)
    text = re.sub(r'~\[color:(#[0-9A-Fa-f]{6})\](.*?)\[/color\]~', r'[color:\1]~\2~[/color]', text, flags=re.IGNORECASE)

    tokens = []
    i = 0
    length = len(text)
    
    # ==========================================
    # 第 1 步：词法扫描 (Lexical Scanning) 
    # ==========================================
    while i < length:
        if text.startswith('**', i):
            tokens.append(('TAG', '**', 'bold'))
            i += 2
        elif text.startswith('__', i):
            # 防止表单填空线 "_____" 被当成下划线标签
            if (i > 0 and text[i-1] == '_') or (i+2 < length and text[i+2] == '_'):
                tokens.append(('TEXT', '_'))
                i += 1
            else:
                tokens.append(('TAG', '__', 'underline'))
                i += 2
        elif text.startswith('[color:#', i):
            end_idx = text.find(']', i)
            if end_idx != -1 and end_idx - i == 14:
                color_hex = text[i+7:end_idx]
                tokens.append(('TAG_OPEN', 'color', color_hex))
                i = end_idx + 1
            else:
                tokens.append(('TEXT', '['))
                i += 1
        elif text.startswith('[/color]', i):
            tokens.append(('TAG_CLOSE', 'color'))
            i += 8
        elif text[i] == '*':
            tokens.append(('TAG', '*', 'italic'))
            i += 1
        elif text[i] == '^':
            tokens.append(('TAG', '^', 'super'))
            i += 1
        elif text[i] == '~':
            tokens.append(('TAG', '~', 'sub'))
            i += 1
        else:
            if tokens and tokens[-1][0] == 'TEXT':
                tokens[-1] = ('TEXT', tokens[-1][1] + text[i])
            else:
                tokens.append(('TEXT', text[i]))
            i += 1

    # ==========================================
    # 第 2 步：严格成对校验 (Syntax Validation)
    # ==========================================
    paired_tokens = []
    active_tags = []
    
    for idx, token in enumerate(tokens):
        if token[0] == 'TEXT':
            paired_tokens.append(token)
        elif token[0] == 'TAG':
            tag_type = token[2]
            found_close = False
            for j in range(idx + 1, len(tokens)):
                # 💡 修复问题 2 的绝对核心：遇到物理换行符，立刻打断寻找闭合标签！
                # 这样孤立的尾注星号就绝对无法和下一段的星号配对了！
                if tokens[j][0] == 'TEXT' and '\n' in tokens[j][1]:
                    break
                if tokens[j][0] == 'TAG' and tokens[j][2] == tag_type:
                    found_close = True
                    break
            
            if tag_type in active_tags:
                active_tags.remove(tag_type)
                paired_tokens.append(('CMD_CLOSE', tag_type))
            elif found_close:
                active_tags.append(tag_type)
                paired_tokens.append(('CMD_OPEN', tag_type))
            else:
                # 孤儿标签，强制打回原形（普通文本星号/向上箭头等）
                paired_tokens.append(('TEXT', token[1]))
                
        elif token[0] == 'TAG_OPEN':
            active_tags.append(('color', token[2]))
            paired_tokens.append(('CMD_OPEN', 'color', token[2]))
        elif token[0] == 'TAG_CLOSE':
            for t in reversed(active_tags):
                if isinstance(t, tuple) and t[0] == 'color':
                    active_tags.remove(t)
                    paired_tokens.append(('CMD_CLOSE', 'color'))
                    break

    # ==========================================
    # 第 3 步：状态机渲染 (Render AST)
    # ==========================================
    runs = []
    current_state = {
        'bold': False, 'italic': False, 'underline': False,
        'super': False, 'sub': False, 'color': []
    }
    
    for cmd in paired_tokens:
        if cmd[0] == 'TEXT':
            color_val = current_state['color'][-1] if current_state['color'] else None
            runs.append({
                'text': cmd[1],
                'bold': current_state['bold'],
                'italic': current_state['italic'],
                'underline': current_state['underline'],
                'super': current_state['super'],
                'sub': current_state['sub'],
                'color': color_val
            })
        elif cmd[0] == 'CMD_OPEN':
            if cmd[1] == 'color': current_state['color'].append(cmd[2])
            else: current_state[cmd[1]] = True
        elif cmd[0] == 'CMD_CLOSE':
            if cmd[1] == 'color':
                if current_state['color']: current_state['color'].pop()
            else: current_state[cmd[1]] = False
            
    # ==========================================
    # 第 4 步：碎片合并优化
    # ==========================================
    merged = []
    for r in runs:
        if merged and all(merged[-1].get(k) == r.get(k) for k in ('color', 'bold', 'italic', 'underline', 'super', 'sub')):
            merged[-1]['text'] += r['text']
        else:
            merged.append(r)
            
    return merged

def rich_runs_to_html(text: str) -> str:
    import html as _html
    runs = parse_rich_text(text)
    parts = []
    for r in runs:
        seg = _html.escape(r.get('text', '')).replace('\n', '<br>')
        if r.get('color'):
            c = str(r['color'])
            if not c.startswith('#'): c = '#' + c
            seg = '<span style="color:' + c + '">' + seg + '</span>'
        if r.get('bold'): seg = '<b>' + seg + '</b>'
        if r.get('italic'): seg = '<i>' + seg + '</i>'
        if r.get('underline'): seg = '<u>' + seg + '</u>'
        if r.get('super'): seg = '<sup>' + seg + '</sup>'
        if r.get('sub'): seg = '<sub>' + seg + '</sub>'
        parts.append(seg)
    return '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;line-height:1.6;">' + ''.join(parts) + '</div>'


def _force_black_on_white(color_hex: str) -> str:
    try:
        h = color_hex.lstrip('#')
        r, g, b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
        luminance = 0.2126*r + 0.7152*g + 0.0722*b
        if luminance > 180: return '000000'
        return h
    except Exception: return '000000'

def build_atf_docx(items: list, output_path: str, dedup: bool = True):
    try:
        from docx import Document
        from docx.shared import Inches
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
    except ImportError: raise ImportError("python-docx not installed.")
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
            cleaned = re.sub(r'^#{1,6}\s+', '', stripped_raw, flags=re.MULTILINE)
            cleaned = re.sub(r'\[color:(#[0-9A-Fa-f]{6})\](.*?)\[/color\]', lambda m: '[color:' + m.group(1) + ']' + m.group(2).replace('\n', ' ') + '[/color]', cleaned, flags=re.DOTALL)
            cleaned = re.sub(r'^\s*\[\d+\]\s+', '', cleaned, flags=re.MULTILINE)
            lines = [l for l in cleaned.split('\n') if l.strip()]
            if dedup:
                seen_lines = set(); deduped = []
                for l in lines:
                    if l.strip() not in seen_lines: seen_lines.add(l.strip()); deduped.append(l)
                ocr_lines = deduped
            else:
                ocr_lines = lines
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
            tr_c = OxmlElement('w:tr')
            line = re.sub(r'(?i)(\[blank\]|\[illegible\])', r'[color:#FF0000]\1[/color]', line)
            rich_runs = parse_rich_text(line)
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
    def __init__(self, path, api_key, proxy, signals, model='gemini-3.1-flash-image-preview'):
        super().__init__(); self.path = path; self.api_key = api_key; self.proxy = proxy; self.signals = signals; self.model = model
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
                res = client.models.generate_content(model=self.model, config=types.GenerateContentConfig(temperature=0.0), contents=[detect_prompt, working_img])
                ans = res.text.strip().upper() if (res.candidates and res.candidates[0].content.parts) else ""
            else:
                buf = io.BytesIO(); working_img.save(buf, format='PNG'); b64 = base64.b64encode(buf.getvalue()).decode()
                resp = requests.post(f'https://generativelanguage.googleapis.com/v1beta/models/{self.model}:generateContent?key={self.api_key}', json={'contents': [{'parts': [{'inline_data': {'mime_type': 'image/png', 'data': b64}}, {'text': detect_prompt}]}], 'generationConfig': {'temperature': 0.0}}, proxies=self.proxy, timeout=60)
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
                res = client.models.generate_content(model="gemini-3.7-flash", config=types.GenerateContentConfig(temperature=0.0), contents=[prompt, working_img])
                text_res = res.text if (res.candidates and res.candidates[0].content.parts) else ""
            else:
                buf = io.BytesIO(); working_img.save(buf, format='PNG'); b64 = base64.b64encode(buf.getvalue()).decode()
                resp = requests.post(f'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.7-flash:generateContent?key={self.api_key}', json={'contents': [{'parts': [{'inline_data': {'mime_type': 'image/png', 'data': b64}}, {'text': prompt}]}], 'generationConfig': {'temperature': 0.0}}, proxies=self.proxy, timeout=60)
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


def _is_transient_error(exc: Exception) -> bool:
    """判断是不是"值得自动重试"的临时性错误：网络超时/连接中断，或者HTTP 429/500/502/503/504限流类错误。
    其他错误（比如 API Key 无效、图片格式错误、400 参数错误）第一次失败就直接抛出，不浪费时间重试。"""
    if isinstance(exc, (requests.exceptions.Timeout, requests.exceptions.ConnectionError)):
        return True
    if isinstance(exc, requests.exceptions.HTTPError):
        status = exc.response.status_code if exc.response is not None else None
        return status in (429, 500, 502, 503, 504)
    return False

def _call_with_retry(func, *args, max_retries: int = 3, base_delay: float = 2.0, tag: str = '', **kwargs):
    """对可能因网络抖动/接口限流而失败的调用做自动重试（指数退避 + 小幅抖动，避免多线程同时撞车）。
    只重试"看起来是临时性"的错误；其余错误立即抛出。"""
    last_exc = None
    for attempt in range(1, max_retries + 1):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            last_exc = e
            if attempt >= max_retries or not _is_transient_error(e):
                raise
            delay = base_delay * (2 ** (attempt - 1)) + random.uniform(0, 1.0)
            try: logger.warning(f'[retry] {tag} attempt {attempt}/{max_retries} failed ({type(e).__name__}: {e}), retrying in {delay:.1f}s')
            except Exception: pass
            time.sleep(delay)
    raise last_exc

class WorkerSignals(QObject):
    finished  = pyqtSignal(str, str); error = pyqtSignal(str, str); progress = pyqtSignal(int, int)
    excluded_info = pyqtSignal(str, int)  # (image_path, excluded_numeric_count) —— 让界面日志能显示Mask模式下被过滤掉的纯数字条目数

class OCRRunnable(QRunnable):
    def __init__(self, api, image_path, api_key, model, proxies, signals, done_counter, total, rotation_angle=0):
        super().__init__(); self.api = api; self.image_path = image_path; self.api_key = api_key; self.model = model
        self.proxies = proxies; self.signals = signals; self.done_counter = done_counter; self.total = total; self.rotation_angle = rotation_angle
    @pyqtSlot()
    def run(self):
        _tag = Path(self.image_path).name
        try:
            if self.api == 'claude': text = _call_with_retry(ocr_claude, self.image_path, self.api_key, self.model, self.proxies, self.rotation_angle, tag=_tag)
            elif self.api == 'gemini': text = _call_with_retry(ocr_gemini, self.image_path, self.api_key, self.model, self.proxies, self.rotation_angle, tag=_tag)
            else: text = _call_with_retry(ocr_mistral, self.image_path, self.api_key, self.model, self.proxies, self.rotation_angle, tag=_tag)
            self.signals.finished.emit(self.image_path, text)
        except Exception as e: self.signals.error.emit(self.image_path, str(e))
        finally:
            self.done_counter[0] += 1
            self.signals.progress.emit(self.done_counter[0], self.total)

def create_redacted_image_base64(image_path: str, polygons: list, rotation_angle: int = 0) -> tuple[str, str]:
    with _pil_draw_lock:
        with PILImage.open(image_path) as src: img = ImageOps.exif_transpose(src).convert('RGB')
        if rotation_angle == 90: img = img.transpose(PILImage.Transpose.ROTATE_270)
        elif rotation_angle == 180: img = img.transpose(PILImage.Transpose.ROTATE_180)
        elif rotation_angle == 270 or rotation_angle == -90: img = img.transpose(PILImage.Transpose.ROTATE_90)
        img_w, img_h = img.size; crops = []
    
        for poly_points in polygons:
            abs_points = [(int(x * img_w), int(y * img_h)) for x, y in poly_points]
            xs = [p[0] for p in abs_points]; ys = [p[1] for p in abs_points]
            x1, y1 = max(0, min(xs)-2), max(0, min(ys)-2)
            x2, y2 = min(img_w, max(xs)+2), min(img_h, max(ys)+2)
        
            mask = PILImage.new('L', (img_w, img_h), 0)
            ImageDraw.Draw(mask).polygon(abs_points, fill=255)
            white_bg = PILImage.new('RGB', (img_w, img_h), (255, 255, 255))
            composite = PILImage.composite(img, white_bg, mask)
            crops.append(composite.crop((x1, y1, x2, y2)))
        
        # 不再依赖当前这台电脑系统字体目录里的 arial.ttf —— 不同电脑上这个文件
        # 可能不是同一份数据（不同 Windows/语言包/更新历史），FreeType 解析到
        # 异常字体数据时可能触发原生崩溃，try/except 接不住。改用 Pillow 自带、
        # 随包附带、所有机器上保证一致的默认字体，彻底去掉这个变量。
        try: font = ImageFont.load_default(size=20)
        except TypeError: font = ImageFont.load_default()
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

def _is_pure_numeric_text(s: str) -> bool:
    """判断一段OCR结果是否是"纯数字"内容（专利图纸里常见的纯参考编号，如 150 / 23 / 150-152 / 45%）。
    允许数字常见的附属符号（正负号、百分号、连字符/范围号、括号、度数符号等），
    但只要出现任何字母或其他文字字符就不算纯数字。
    注意：带小数点或千分位逗号的数字（如 3.14 / 1,234 / 12,345.67）视为有实际含义的数值，
    不当作纯参考编号处理，会被保留，不会被过滤掉。"""
    s = s.strip()
    if not s or s.upper() in ('BLANK', '[BLANK]'): return False
    if re.search(r'\d[.,]\d', s): return False
    stripped = re.sub(r'[\s\-–—+%°′″()\[\]/]', '', s)
    return stripped.isdigit()

def _is_single_letter_text(s: str) -> bool:
    """判断一段OCR结果是否是"单个字母"内容（专利图纸里常见的截面/引出线标记，如 A / B / a）。
    允许外面套一层括号/方括号（如 (A) 、 [B]），但两个及以上字母、或字母和数字混在一起(如 A1、AB、A-A')，
    就不算单个字母，会被保留。"""
    s = s.strip()
    if not s or s.upper() in ('BLANK', '[BLANK]'): return False
    stripped = re.sub(r'[\s()\[\]]', '', s)
    return len(stripped) == 1 and stripped.isalpha()

def ocr_mask_engine(image_path: str, polygons: list, rotation_angle: int, api_type: str, api_key: str, model: str, proxies: dict, exclude_pure_numeric: bool = True) -> tuple[str, int]:
    _tag = Path(image_path).name
    if not polygons:
        try: logger.debug(f'[trace] {_tag} no polygons, skip')
        except Exception: pass
        return "", 0
    try: logger.debug(f'[trace] {_tag} composite:start')
    except Exception: pass
    b64, mime = create_redacted_image_base64(image_path, polygons, rotation_angle)
    try: logger.debug(f'[trace] {_tag} composite:done b64_len={len(b64)}')
    except Exception: pass
    dynamic_prompt = MASK_OCR_PROMPT + f"\n\nATTENTION: There are exactly {len(polygons)} snippets in this image. You MUST output exactly {len(polygons)} extracted paragraphs."
    try: logger.debug(f'[trace] {_tag} http:post api={api_type}')
    except Exception: pass
    if api_type == 'claude':
        resp = requests.post('https://api.anthropic.com/v1/messages', headers={'x-api-key': api_key, 'anthropic-version': '2023-06-01', 'content-type': 'application/json'}, json={'model': model, 'max_tokens': 4096, 'messages': [{'role': 'user', 'content': [{'type': 'image', 'source': {'type': 'base64', 'media_type': mime, 'data': b64}}, {'type': 'text', 'text': dynamic_prompt}]}]}, proxies=proxies, timeout=120)
        try: logger.debug(f'[trace] {_tag} http:resp status={resp.status_code}')
        except Exception: pass
        resp.raise_for_status()
        raw_text = resp.json()['content'][0]['text']
    else:
        resp = requests.post(f'https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}', json={'contents': [{'parts': [{'inline_data': {'mime_type': mime, 'data': b64}}, {'text': dynamic_prompt}]}], 'generationConfig': {'temperature': 0}}, proxies=proxies, timeout=120)
        try: logger.debug(f'[trace] {_tag} http:resp status={resp.status_code}')
        except Exception: pass
        resp.raise_for_status()
        raw_text = resp.json()['candidates'][0]['content']['parts'][0]['text']
    try: logger.debug(f'[trace] {_tag} parse:start len={len(raw_text)}')
    except Exception: pass
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
    excluded_numeric = 0
    for i in range(1, len(polygons) + 1):
        val = extracted.get(str(i), "").strip()
        if not val or val.upper() == "BLANK": val = "[BLANK]"
        if exclude_pure_numeric and (_is_pure_numeric_text(val) or _is_single_letter_text(val)):
            excluded_numeric += 1
            continue
        clean_lines.append(f"[{i}] {val}")
    try: logger.debug(f'[trace] {_tag} parse:done excluded_numeric={excluded_numeric}')
    except Exception: pass
    return '\n\n'.join(clean_lines), excluded_numeric

class MaskOCRRunnable(QRunnable):
    def __init__(self, api, image_path, polygons, rotation_angle, api_key, model, proxies, signals, done_counter, total, exclude_pure_numeric=True):
        super().__init__(); self.api = api; self.image_path = image_path; self.polygons = polygons; self.rotation_angle = rotation_angle
        self.api_key = api_key; self.model = model; self.proxies = proxies; self.signals = signals
        self.done_counter = done_counter; self.total = total; self.exclude_pure_numeric = exclude_pure_numeric
    @pyqtSlot()
    def run(self):
        _name = Path(self.image_path).name
        try:
            logger.debug(f'[trace] enter {_name} polygons={len(self.polygons)}')
        except Exception:
            pass
        try:
            text, excluded_numeric = _call_with_retry(ocr_mask_engine, self.image_path, self.polygons, self.rotation_angle, self.api, self.api_key, self.model, self.proxies, self.exclude_pure_numeric, tag=_name)
            try: logger.debug(f'[trace] exit-ok {_name}')
            except Exception: pass
            if excluded_numeric > 0:
                try: self.signals.excluded_info.emit(self.image_path, excluded_numeric)
                except Exception: pass
            self.signals.finished.emit(self.image_path, text)
        except Exception as e:
            try: logger.debug(f'[trace] exit-error {_name}: {e}')
            except Exception: pass
            self.signals.error.emit(self.image_path, str(e))
        finally:
            self.done_counter[0] += 1
            self.signals.progress.emit(self.done_counter[0], self.total)

class FileItem:
    def __init__(self, path: str):
        self.path = path; self.name = Path(path).name
        self.status = 'pending'; self.ocr_text = ''; self.error_msg = ''
        self.pixmap = None; self.boxes = []; self.polygons = []; self.box_detected = False; self.rotation_angle = 0 
    def set_rotation(self, angle: int):
        self.rotation_angle = angle % 360; self.pixmap = None; self.boxes.clear(); self.polygons.clear() 
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
        self.pan_x = 0; self.pan_y = 0 
        self.is_panning = False; self.pan_start_pos = None; self.pan_start_x = 0; self.pan_start_y = 0
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setStyleSheet('background: #1a1d27; border-radius: 6px; color: #64748b;')
        self.setText('← Select an image')
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
    @property
    def _source_pixmap(self): return self._internal_pixmap
    @_source_pixmap.setter
    def _source_pixmap(self, value): self._internal_pixmap = value; self._zoom_factor = 1.0; self.pan_x = 0; self.pan_y = 0
    def set_image(self, path: str):
        pix = QPixmap(path)
        if pix.isNull(): self._source_pixmap = None; self.setText('Cannot load image')
        else: self._source_pixmap = pix; self.setText(''); self.update()
    def clear_image(self): self._source_pixmap = None; self.clear(); self.setText('← Select an image'); self.update()
    def wheelEvent(self, event):
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            if event.angleDelta().y() > 0: self._zoom_factor = min(5.0, self._zoom_factor + 0.1)
            else: self._zoom_factor = max(0.1, self._zoom_factor - 0.1)
        elif event.modifiers() == Qt.KeyboardModifier.ShiftModifier:
            self.pan_x += event.angleDelta().y()
        else:
            self.pan_y += event.angleDelta().y()
            self.pan_x += event.angleDelta().x()
        self.update()
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.MiddleButton:
            self.is_panning = True; self.pan_start_pos = event.position(); self.pan_start_x = self.pan_x; self.pan_start_y = self.pan_y
    def mouseMoveEvent(self, event):
        if getattr(self, 'is_panning', False):
            delta = event.position() - self.pan_start_pos
            self.pan_x = self.pan_start_x + delta.x(); self.pan_y = self.pan_start_y + delta.y()
            self.update()
    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.MiddleButton: self.is_panning = False
    def paintEvent(self, event):
        super().paintEvent(event)
        if not self._internal_pixmap: return
        painter = QPainter(self); painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        v_w, v_h = self.width(), self.height()
        p_w, p_h = self._internal_pixmap.width(), self._internal_pixmap.height()
        base_scale = min(v_w / max(1, p_w), v_h / max(1, p_h)); scale_factor = base_scale * self._zoom_factor
        scaled_w, scaled_h = max(1, int(p_w * scale_factor)), max(1, int(p_h * scale_factor))
        offset_x = (v_w - scaled_w) // 2 + int(self.pan_x)
        offset_y = (v_h - scaled_h) // 2 + int(self.pan_y)
        scaled_pixmap = self._internal_pixmap.scaled(scaled_w, scaled_h, Qt.AspectRatioMode.IgnoreAspectRatio, Qt.TransformationMode.SmoothTransformation)
        painter.drawPixmap(offset_x, offset_y, scaled_pixmap)


class InteractiveMaskLabel(QLabel):
    boxes_changed = pyqtSignal()
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setStyleSheet('background: transparent; color: #64748b;')
        self.setText('← Select an image')
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setMouseTracking(True) 
        
        self.file_item = None; self.source_pixmap = None
        self.drawing_poly = []; self.drawing_rect_start = None 
        self.is_dragging = False; self.is_erasing = False        
        self.erase_start = None; self.current_mouse_pos = None  
        self.scale_factor = 1.0; self.offset_x = 0; self.offset_y = 0; self._zoom_factor = 1.0; self.pan_x = 0; self.pan_y = 0 
        self.is_panning = False; self.pan_start_pos = None; self.pan_start_x = 0; self.pan_start_y = 0

    def set_item(self, item):
        self._zoom_factor = 1.0; self.pan_x = 0; self.pan_y = 0; self.file_item = item; pix = QPixmap(item.path)
        if pix.isNull(): self.source_pixmap = None; self.setText('Cannot load image')
        else:
            if item.rotation_angle != 0: pix = pix.transformed(QTransform().rotate(item.rotation_angle), Qt.TransformationMode.SmoothTransformation)
            self.source_pixmap = pix; self.setText('')
        self.update()

    def clear_item(self):
        self._zoom_factor = 1.0; self.pan_x = 0; self.pan_y = 0; self.file_item = None; self.source_pixmap = None; self.setText('← Select an image'); self.update()

    def wheelEvent(self, event):
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            if event.angleDelta().y() > 0: self._zoom_factor = min(5.0, self._zoom_factor + 0.1)
            else: self._zoom_factor = max(0.1, self._zoom_factor - 0.1)
        elif event.modifiers() == Qt.KeyboardModifier.ShiftModifier:
            self.pan_x += event.angleDelta().y()
        else: 
            self.pan_y += event.angleDelta().y()
            self.pan_x += event.angleDelta().x()
        self.update()

    def paintEvent(self, event):
        super().paintEvent(event)
        if not self.source_pixmap or not self.file_item: return
        painter = QPainter(self); painter.setRenderHint(QPainter.RenderHint.Antialiasing); painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        v_w, v_h = self.width(), self.height(); p_w, p_h = self.source_pixmap.width(), self.source_pixmap.height()
        base_scale = min(v_w / max(1, p_w), v_h / max(1, p_h)); self.scale_factor = base_scale * self._zoom_factor
        scaled_w = max(1, int(p_w * self.scale_factor)); scaled_h = max(1, int(p_h * self.scale_factor))
        self.offset_x = (v_w - scaled_w) // 2 + self.pan_x; self.offset_y = (v_h - scaled_h) // 2 + self.pan_y
        scaled_pixmap = self.source_pixmap.scaled(scaled_w, scaled_h, Qt.AspectRatioMode.IgnoreAspectRatio, Qt.TransformationMode.SmoothTransformation)
        painter.drawPixmap(int(self.offset_x), int(self.offset_y), scaled_pixmap)
        
        pen_box = QPen(QColor(79, 142, 247, 220), 2, Qt.PenStyle.SolidLine); painter.setPen(pen_box)
        
        for i, poly_points in enumerate(self.file_item.polygons):
            qpoly = QPolygonF()
            for rx, ry in poly_points:
                qpoly.append(QPointF(rx * scaled_w + self.offset_x, ry * scaled_h + self.offset_y))
            painter.setBrush(QBrush(QColor(79, 142, 247, 40)))
            painter.drawPolygon(qpoly)
            if not qpoly.isEmpty():
                painter.setPen(QColor(255, 255, 255))
                top_point = min(qpoly, key=lambda p: p.y())
                painter.drawText(top_point + QPointF(5, -5), f"[{i+1}]")
                painter.setPen(pen_box) 

        if self.is_dragging and self.drawing_rect_start and self.current_mouse_pos:
            painter.setPen(QPen(QColor(34, 197, 94, 255), 2, Qt.PenStyle.DashLine))
            painter.setBrush(QBrush(QColor(34, 197, 94, 40)))
            painter.drawRect(QRectF(self.drawing_rect_start, self.current_mouse_pos).normalized())
            
        elif self.drawing_poly:
            painter.setPen(QPen(QColor(34, 197, 94, 255), 2, Qt.PenStyle.DashLine))
            painter.setBrush(QBrush(QColor(34, 197, 94, 40)))
            temp_poly = QPolygonF(self.drawing_poly)
            if self.current_mouse_pos: temp_poly.append(self.current_mouse_pos)
            painter.drawPolygon(temp_poly)
            
        if self.is_erasing and self.erase_start and self.current_mouse_pos:
            painter.setPen(QPen(QColor(239, 68, 68, 255), 2, Qt.PenStyle.DashLine))
            painter.setBrush(QBrush(QColor(239, 68, 68, 40)))
            painter.drawRect(QRectF(self.erase_start, self.current_mouse_pos).normalized())

    def mousePressEvent(self, event):
        if not self.file_item or not self.source_pixmap: return
        self.setFocus()
        if event.button() == Qt.MouseButton.LeftButton:
            self.drawing_rect_start = event.position(); self.is_dragging = False
        elif event.button() == Qt.MouseButton.RightButton:
            if self.drawing_poly:
                if len(self.drawing_poly) > 2:
                    p_w, p_h = self.source_pixmap.width(), self.source_pixmap.height()
                    base_scale = min(self.width() / max(1, p_w), self.height() / max(1, p_h))
                    scale_factor = base_scale * self._zoom_factor
                    scaled_w = max(1, int(p_w * scale_factor)); scaled_h = max(1, int(p_h * scale_factor))
                    poly_rel = []
                    for pt in self.drawing_poly:
                        poly_rel.append(((pt.x() - self.offset_x) / scaled_w, (pt.y() - self.offset_y) / scaled_h))
                    self.file_item.polygons.append(poly_rel); self.boxes_changed.emit()
                self.drawing_poly = []; self.drawing_rect_start = None; self.update()
            else:
                self.erase_start = event.position(); self.is_erasing = True; self.update()
        elif event.button() == Qt.MouseButton.MiddleButton:
            self.is_panning = True; self.pan_start_pos = event.position(); self.pan_start_x = self.pan_x; self.pan_start_y = self.pan_y

    def mouseMoveEvent(self, event):
        self.current_mouse_pos = event.position()
        if getattr(self, 'is_panning', False):
            delta = event.position() - self.pan_start_pos
            self.pan_x = self.pan_start_x + delta.x(); self.pan_y = self.pan_start_y + delta.y()
        elif event.buttons() & Qt.MouseButton.LeftButton and self.drawing_rect_start:
            diff = event.position() - self.drawing_rect_start
            if diff.manhattanLength() > 10: self.is_dragging = True
        self.update()

    def mouseReleaseEvent(self, event):
        if not self.file_item or not self.source_pixmap: return
        if event.button() == Qt.MouseButton.MiddleButton:
            self.is_panning = False; return
            
        p_w, p_h = self.source_pixmap.width(), self.source_pixmap.height()
        base_scale = min(self.width() / max(1, p_w), self.height() / max(1, p_h))
        scale_factor = base_scale * self._zoom_factor
        scaled_w = max(1, int(p_w * scale_factor)); scaled_h = max(1, int(p_h * scale_factor))
        
        if event.button() == Qt.MouseButton.LeftButton:
            if self.is_dragging and self.drawing_rect_start:
                rect = QRectF(self.drawing_rect_start, event.position()).normalized()
                pixmap_rect = QRectF(self.offset_x, self.offset_y, scaled_w, scaled_h)
                intersected = rect.intersected(pixmap_rect)
                if not intersected.isEmpty() and intersected.width() > 5 and intersected.height() > 5:
                    rx1 = (intersected.left() - self.offset_x) / scaled_w; ry1 = (intersected.top() - self.offset_y) / scaled_h
                    rx2 = (intersected.right() - self.offset_x) / scaled_w; ry2 = (intersected.bottom() - self.offset_y) / scaled_h
                    self.file_item.polygons.append([(rx1, ry1), (rx2, ry1), (rx2, ry2), (rx1, ry2)])
                    self.boxes_changed.emit()
                self.is_dragging = False; self.drawing_rect_start = None; self.drawing_poly = [] 
            else:
                if self.drawing_rect_start: self.drawing_poly.append(self.drawing_rect_start)
                self.is_dragging = False; self.drawing_rect_start = None
                
        elif event.button() == Qt.MouseButton.RightButton and self.is_erasing:
            self.is_erasing = False; self.current_pos = event.position()
            erase_rect = QRectF(self.erase_start, self.current_pos).normalized()
            erase_poly = QPolygonF(erase_rect)
            remaining = []
            for poly in self.file_item.polygons:
                qpoly = QPolygonF()
                for rx, ry in poly: qpoly.append(QPointF(rx * scaled_w + self.offset_x, ry * scaled_h + self.offset_y))
                if not erase_poly.intersects(qpoly): remaining.append(poly)
            if len(remaining) != len(self.file_item.polygons):
                self.file_item.polygons = remaining; self.boxes_changed.emit()
            self.erase_start = None
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
        self._rotation_model = 'gemini-3.1-flash-image-preview'  # 各标签页按需在自己的 __init__ 里覆盖这个值
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
        self._rich_preview_enabled = True
        self.ocr_edit = QPlainTextEdit(); self.ocr_edit.setPlaceholderText(T('ocr_placeholder')); self.ocr_edit.setStyleSheet('font-family: "Consolas","Monaco",monospace; font-size: 12px;')
        self.ocr_edit.textChanged.connect(self._save_ocr_edit); self.ocr_edit.setVisible(False); ocr_layout.addWidget(self.ocr_edit, 1)
        self.ocr_preview = QTextEdit(); self.ocr_preview.setReadOnly(True)
        self.ocr_preview.setToolTip('点击此处编辑标签源码 / Click to edit raw tags')
        self.ocr_preview.setStyleSheet('background:#3a3f4b; color:#f0f0f0; border:1px solid #2e3250; border-radius:4px;')
        ocr_layout.addWidget(self.ocr_preview, 1)
        self.ocr_preview.viewport().installEventFilter(self)
        self.ocr_edit.installEventFilter(self)
        preview_splitter.addWidget(ocr_frame)
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
        ab.addWidget(self.btn_start); ab.addWidget(self.btn_stop); ab.addWidget(self.btn_reset); pv = QVBoxLayout(); pv.setSpacing(2); pv.addWidget(self.progress_bar); pv.addWidget(self.progress_label); ab.addLayout(pv, 1); ab.addWidget(self.btn_gen_atf); log_layout.addLayout(ab); self.action_bar = ab
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
        try:
            (logger.error if is_error else logger.warning if is_warning else logger.info)(f'[{self.api_name}] {text}')
        except Exception:
            pass
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
                self.img_label._source_pixmap = pix; self.img_label.setText(''); self.img_label.update()
            else: self.img_label.set_image(item.path)
        else: self.img_label.set_image(item.path)
        self.img_name_label.setText(item.name); self.ocr_edit.blockSignals(True); self.ocr_edit.setPlainText(item.ocr_text); self.ocr_edit.blockSignals(False)
        self.btn_reocr.setEnabled(item.status != 'processing')
        self._refresh_ocr_preview()
    def _clear_preview(self): self.img_label.clear_image(); self.img_name_label.setText(''); self.ocr_edit.blockSignals(True); self.ocr_edit.clear(); self.ocr_edit.blockSignals(False); self.btn_reocr.setEnabled(False); self._refresh_ocr_preview()
    def _save_ocr_edit(self):
        if self.selected_idx is not None and self.selected_idx < len(self.file_items): self.file_items[self.selected_idx].ocr_text = self.ocr_edit.toPlainText()
        self._refresh_ocr_preview()
    def _refresh_ocr_preview(self):
        if not getattr(self, '_rich_preview_enabled', False): return
        if getattr(self, 'ocr_preview', None) is None or not self.ocr_preview.isVisible(): return
        self.ocr_preview.setHtml(rich_runs_to_html(self.ocr_edit.toPlainText()))
    def _enter_edit_mode(self):
        if not getattr(self, '_rich_preview_enabled', False): return
        self.ocr_preview.setVisible(False); self.ocr_edit.setVisible(True)
        self.ocr_edit.setFocus()
        cur = self.ocr_edit.textCursor(); cur.movePosition(cur.MoveOperation.End); self.ocr_edit.setTextCursor(cur)
    def _exit_edit_mode(self):
        if not getattr(self, '_rich_preview_enabled', False): return
        self._save_ocr_edit()
        self.ocr_edit.setVisible(False); self.ocr_preview.setVisible(True)
        self._refresh_ocr_preview()
    def eventFilter(self, obj, event):
        from PyQt6.QtCore import QEvent
        if getattr(self, '_rich_preview_enabled', False):
            if getattr(self, 'ocr_preview', None) is not None and obj is self.ocr_preview.viewport() and event.type() == QEvent.Type.MouseButtonPress:
                self._enter_edit_mode(); return True
            if getattr(self, 'ocr_edit', None) is not None and obj is self.ocr_edit and event.type() == QEvent.Type.FocusOut:
                self._exit_edit_mode()
        return super().eventFilter(obj, event)
    def _get_config(self): return self.config_widget.get_key(), self.config_widget.get_model(), self.proxy_widget.get_proxies()

    def _on_start(self):
        pending = [item for item in self.file_items if item.status in ('pending', 'error')]
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
            self.ocr_edit.blockSignals(True); self.ocr_edit.setPlainText(text); self.ocr_edit.blockSignals(False); self.btn_reocr.setEnabled(True); self._refresh_ocr_preview()
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
        self._tab_log(f"↺ Re-OCR started for {item.name}...")
        item.status = 'processing'; self._refresh_list(); self.btn_reocr.setEnabled(False)
        signals = WorkerSignals(); signals.finished.connect(self._on_ocr_done); signals.error.connect(self._on_ocr_error); signals.progress.connect(lambda d, t: None)
        worker = OCRRunnable(self.api_name, item.path, key, model, proxies, signals, [0], 1, item.rotation_angle)
        worker.setAutoDelete(True); self._thread_pool.start(worker); self._update_summary()
    def _on_gen_atf(self):
        done_items = [item for item in self.file_items if item.status == 'done' and item.ocr_text.strip()]
        if not done_items: QMessageBox.information(self, T('info_title'), T('no_ocr_results')); return
        skipped = [item for item in self.file_items if item.status == 'error']
        if skipped:
            reply = QMessageBox.question(self, T('info_title'), T('confirm_skip_failed', n=len(skipped)), QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.Yes)
            if reply != QMessageBox.StandardButton.Yes: return
        out_path, _ = QFileDialog.getSaveFileName(self, T('save_atf_title'), os.path.join(str(Path(done_items[0].path).parent), f'ATF_{self.api_name}.docx'), 'Word Documents (*.docx)')
        if not out_path: return
        from datetime import datetime
        ts = datetime.now().strftime('%Y%m%d_%H%M%S'); base, ext = os.path.splitext(out_path)
        if not any(c.isdigit() for c in Path(out_path).stem[-8:]): out_path = f'{base}_{ts}{ext}'
        
        try:
            build_atf_docx([{'name': i.name, 'image_path': i.path, 'ocr_text': i.ocr_text} for i in done_items], out_path)
            self.export_status = T('export_success')
            self._tab_log(T('log_atf_done', p=out_path))
            self._update_summary()
            
            # 强化修复：转换为系统绝对路径，确保 os.startfile 稳定调用
            import subprocess, sys as _sys
            try:
                abs_out = os.path.abspath(out_path)
                if _sys.platform == 'win32': os.startfile(abs_out)
                elif _sys.platform == 'darwin': subprocess.Popen(['open', abs_out])
                else: subprocess.Popen(['xdg-open', abs_out])
            except Exception as open_err:
                self._tab_log(f"File saved but auto-open failed: {open_err}", is_warning=True)
                
        except Exception as e:
            self.export_status = T('export_failed_s')
            self._tab_log(T('log_atf_fail', e=e), is_error=True)
            self._update_summary()
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
            self.img_label._source_pixmap = pix; self.img_label.setText(''); self.img_label.update()
        else: self.img_label.set_image(item.path)
        self._refresh_list(); self._tab_log(T('rot_left_log', n=item.name))
    def _on_rot_right(self):
        if self.selected_idx is None: return
        item = self.file_items[self.selected_idx]; item.set_rotation(item.rotation_angle + 90)
        pix = QPixmap(item.path)
        if not pix.isNull() and item.rotation_angle != 0:
            pix = pix.transformed(QTransform().rotate(item.rotation_angle), Qt.TransformationMode.SmoothTransformation)
            self.img_label._source_pixmap = pix; self.img_label.setText(''); self.img_label.update()
        else: self.img_label.set_image(item.path)
        self._refresh_list(); self._tab_log(T('rot_right_log', n=item.name))
    def _on_batch_auto_rot(self):
        if not self.file_items: return
        key, model, proxies = self._get_config()
        if not key: QMessageBox.warning(self, T('missing_key_title'), T('auto_rot_key_err', api=self.api_name)); return
        self.btn_auto_rot.setEnabled(False); self._tab_log(T('auto_rot_start', n=len(self.file_items)))
        for item in self.file_items:
            signals = DetectSignals(); signals.orientation_done.connect(self._on_ai_rot_res); signals.orientation_error.connect(self._on_ai_rot_err)
            worker = AIOrientationWorker(item.path, key, proxies, signals, model=self._rotation_model)
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

def _reading_order_index(polygons: list) -> list:
    """把一批矩形/多边形框按"阅读顺序"重新排列：先按行分组(从上到下)，同一行内再按从左到右排序。
    只用于ATF creator生成ATF文档前的最后一步重排，不影响画框编辑器里看到的框编号，也不影响发给OCR引擎的拼图编号。
    polygons: 每个元素是若干 (x, y) 归一化坐标点(0-1)组成的多边形。
    返回：按阅读顺序排列的原始下标列表(0-based)。"""
    if not polygons: return []
    boxes = []
    for idx, poly in enumerate(polygons):
        xs = [p[0] for p in poly]; ys = [p[1] for p in poly]
        boxes.append({'idx': idx, 'x0': min(xs), 'y0': min(ys), 'y1': max(ys)})
    boxes.sort(key=lambda b: b['y0'])
    rows = []
    for b in boxes:
        placed = False
        for row in rows:
            ry0 = min(r['y0'] for r in row); ry1 = max(r['y1'] for r in row)
            overlap = min(b['y1'], ry1) - max(b['y0'], ry0)
            row_h = ry1 - ry0; b_h = b['y1'] - b['y0']
            if overlap > 0.5 * min(b_h, row_h if row_h > 0 else b_h):
                row.append(b); placed = True; break
        if not placed: rows.append([b])
    rows.sort(key=lambda row: min(r['y0'] for r in row))
    ordered = []
    for row in rows:
        row.sort(key=lambda r: r['x0'])
        ordered.extend(r['idx'] for r in row)
    return ordered

def _reorder_ocr_text_by_reading_order(ocr_text: str, polygons: list) -> str:
    """按阅读顺序重排一段 "[1] ...\n\n[2] ..." 格式的OCR结果文本，只用于生成ATF文档这一步。
    不修改传入的原始文本/polygons，只返回一份重排后的新文本；解析失败或数量对不上时原样返回，不丢内容。"""
    if not polygons or not ocr_text.strip(): return ocr_text
    blocks = re.split(r'\n\n+', ocr_text.strip())
    parsed = {}
    for b in blocks:
        m = re.match(r'^\[(\d+)\]\s*(.*)', b, flags=re.DOTALL)
        if m: parsed[int(m.group(1))] = m.group(2)
    if not parsed: return ocr_text
    order = _reading_order_index(polygons)
    new_blocks = []; new_num = 1; covered_nums = set()
    for old_idx in order:
        old_num = old_idx + 1
        if old_num in parsed:
            new_blocks.append(f'[{new_num}] {parsed[old_num]}'); covered_nums.add(old_num); new_num += 1
    # 兜底：万一某个编号没能对上polygons(理论不该发生)，追加在末尾，避免静默丢内容
    for old_num, text in parsed.items():
        if old_num not in covered_nums:
            new_blocks.append(f'[{new_num}] {text}'); new_num += 1
    return '\n\n'.join(new_blocks)

class MaskApiTab(ApiTab):
    def __init__(self, api_name: str, models: list, proxy_widget: ProxyWidget, parent=None):
        super().__init__(api_name, models, proxy_widget, parent)
        self._rotation_model = 'gemini-3.7-flash'  # ATF creator专属：换成普通视觉模型做朝向判断，而不是image生成/编辑模型
        self.job_stats = {'rot_total': 0, 'rot_done': 0, 'rot_err': 0, 'box_total': 0, 'box_done': 0, 'box_empty': 0, 'box_err': 0}
    def _build_ui(self, models):
        super()._build_ui(models); self.api_left_toolbar.hide() 
        self.ui_cf_layout.takeAt(0).widget().deleteLater()
        mask_config = QWidget(); mc_layout = QHBoxLayout(mask_config); mc_layout.setContentsMargins(8, 0, 8, 0)
        lbl_engine = QLabel('Engine:'); lbl_engine.setStyleSheet("font-weight: bold; color: #4f8ef7;"); mc_layout.addWidget(lbl_engine)
        self.engine_combo = QComboBox(); self.engine_combo.addItems(['gemini', 'claude']); self.engine_combo.setMinimumWidth(100); mc_layout.addWidget(self.engine_combo)
        self.gemini_cfg = ApiConfigWidget('gemini', models); self.claude_cfg = ApiConfigWidget('claude', CLAUDE_MODELS)
        self.sep = QFrame(); self.sep.setFrameShape(QFrame.Shape.VLine); self.sep.setStyleSheet("color: #2e3250;")
        mc_layout.addWidget(self.sep); mc_layout.addWidget(self.gemini_cfg); mc_layout.addWidget(self.claude_cfg); mc_layout.addStretch()
        self.chk_dedup = QCheckBox(T('dedup_lines')); self.chk_dedup.setChecked(False)
        self.chk_dedup.setStyleSheet("QCheckBox { color: #e2e8f0; margin-right: 8px; } QCheckBox::indicator { width: 16px; height: 16px; border: 1px solid #94a3b8; border-radius: 3px; background: #2e3250; } QCheckBox::indicator:checked { background: #4f8ef7; border-color: #4f8ef7; }")
        self.action_bar.insertWidget(self.action_bar.indexOf(self.btn_gen_atf), self.chk_dedup)
        self.chk_exclude_numeric = QCheckBox(T('exclude_numeric')); self.chk_exclude_numeric.setChecked(False)
        self.chk_exclude_numeric.setStyleSheet("QCheckBox { color: #e2e8f0; margin-right: 8px; } QCheckBox::indicator { width: 16px; height: 16px; border: 1px solid #94a3b8; border-radius: 3px; background: #2e3250; } QCheckBox::indicator:checked { background: #4f8ef7; border-color: #4f8ef7; }")
        self.action_bar.insertWidget(self.action_bar.indexOf(self.btn_gen_atf), self.chk_exclude_numeric)
        self.ui_cf_layout.addWidget(mask_config)
        def update_cfg_vis(): 
            is_claude = (self.engine_combo.currentText() == 'claude')
            self.claude_cfg.setVisible(is_claude)
            self.gemini_cfg.setVisible(not is_claude)
        self.engine_combo.currentTextChanged.connect(update_cfg_vis); update_cfg_vis()
        
        # 👇 强力清扫：只要布局里有多余的画板，全部抽离并销毁
        while self.ui_img_layout.count() > 1:
            item = self.ui_img_layout.takeAt(1)
            if item.widget(): item.widget().deleteLater()
            
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
        self.ui_preview_splitter.setSizes([700, 300])
    def _retranslate_ui(self):
        super()._retranslate_ui()
        self.btn_batch_auto_rot.setText(T('mask_batch_rot')); self.btn_batch_mask.setText(T('mask_batch_box'))
        self.btn_autobox.setText(T('mask_autobox')); self.btn_clearbox.setText(T('mask_clearbox'))
        self.chk_dedup.setText(T('dedup_lines'))
        self.chk_exclude_numeric.setText(T('exclude_numeric'))
    def _update_summary(self, final=False):
        total = len(self.file_items)
        if total == 0: self.tab_summary.setPlainText(T("summary_no_files")); return
        rot_s = f"[1] ROTATION DETECTION\n    Processed: {self.job_stats['rot_done']}/{self.job_stats['rot_total']} | Errors: {self.job_stats['rot_err']}\n\n"
        box_s = f"[2] AUTO POLYGON MASK\n    Detected: {self.job_stats['box_done']}/{self.job_stats['box_total']} | Errors: {self.job_stats['box_err']}\n"
        if self.job_stats['box_empty'] > 0: box_s += f"    ⚠️ EMPTY (0 Masks): {self.job_stats['box_empty']} files!\n"
        ocr_done = len([i for i in self.file_items if i.status == 'done']); ocr_err = len([i for i in self.file_items if i.status == 'error'])
        ocr_pend = len([i for i in self.file_items if i.status == 'pending']); ocr_proc = len([i for i in self.file_items if i.status == 'processing'])
        ocr_s = f"\n[3] OCR EXTRACTION\n    Success: {ocr_done} | Failed: {ocr_err} | Processing: {ocr_proc} | Pending: {ocr_pend}\n\n"
        exp_s = f"[4] ATF EXPORT\n    Status: {self.export_status}\n"
        s = "=== WORKFLOW SUMMARY ===\n\n" + rot_s + box_s + ocr_s + exp_s
        err_items = [i for i in self.file_items if i.status == 'error' or (hasattr(i, 'box_detected') and i.box_detected and len(i.polygons)==0)]
        if err_items:
            s += "\n🚨 [ATTENTION REQUIRED]:\n"
            for i in err_items:
                if i.status == 'error': s += f"  - {i.name}: OCR Error\n"
                elif hasattr(i, 'box_detected') and i.box_detected and len(i.polygons)==0: s += f"  - {i.name}: 0 Masks Detected (Empty)\n"
        self.tab_summary.setPlainText(s)
    def _on_select(self, row):
        if row < 0 or row >= len(self.file_items): return
        self.selected_idx = row; item = self.file_items[row]
        self.img_label.set_item(item); self.img_name_label.setText(item.name)
        self.ocr_edit.blockSignals(True); self.ocr_edit.setPlainText(item.ocr_text); self.ocr_edit.blockSignals(False); self.btn_reocr.setEnabled(item.status != 'processing'); self._refresh_ocr_preview()
    def _clear_preview(self):
        self.img_label.clear_item(); self.img_name_label.setText(''); self.ocr_edit.blockSignals(True); self.ocr_edit.clear(); self.ocr_edit.blockSignals(False); self.btn_reocr.setEnabled(False); self._refresh_ocr_preview()
    def _get_gemini_key(self): return self.gemini_cfg.get_key()
    def _on_rot_left(self):
        if self.selected_idx is None: return
        item = self.file_items[self.selected_idx]; item.set_rotation(item.rotation_angle - 90)
        self.img_label.set_item(item); self._refresh_list(); self._tab_log(f'[Mask] {item.name} In-memory Virtual Rotate Left 90°')
    def _on_rot_right(self):
        if self.selected_idx is None: return
        item = self.file_items[self.selected_idx]; item.set_rotation(item.rotation_angle + 90)
        self.img_label.set_item(item); self._refresh_list(); self._tab_log(f'[Mask] {item.name} In-memory Virtual Rotate Right 90°')
    def _on_batch_auto_rot(self):
        if not self.file_items: return
        key = self._get_gemini_key()
        if not key: QMessageBox.warning(self, "API Key Missing", "AI Rotation requires the Gemini vision model."); return
        self.btn_batch_auto_rot.setEnabled(False); self._tab_log(f'[Mask] 🤖 AI is analyzing images orientation in ({len(self.file_items)} images)...')
        self.job_stats['rot_total'] = len(self.file_items); self.job_stats['rot_done'] = 0; self.job_stats['rot_err'] = 0; self._update_summary(); proxies = self.proxy_widget.get_proxies()
        for item in self.file_items:
            signals = DetectSignals(); signals.orientation_done.connect(self._on_ai_rot_res); signals.orientation_error.connect(self._on_ai_rot_err)
            worker = AIOrientationWorker(item.path, key, proxies, signals, model=self._rotation_model); worker.setAutoDelete(True); self._thread_pool.start(worker)
    def _on_ai_rot_res(self, path, angle):
        self.job_stats['rot_done'] += 1
        for item in self.file_items:
            if item.path == path:
                if angle != 0: item.set_rotation(item.rotation_angle + angle); self._tab_log(f'[Mask] Smart Detection: {item.name} Rotation suggested and applied')
                break
        self._update_summary(); self._finalize_ai_worker()
    def _on_ai_rot_err(self, path, msg):
        self.job_stats['rot_err'] += 1; self._tab_log(f'[Mask] Smart Rotation failed: {Path(path).name} - {msg}', is_error=True); self._update_summary(); self._finalize_ai_worker()
    def _on_auto_box(self):
        if self.selected_idx is None: return
        item = self.file_items[self.selected_idx]; key = self._get_gemini_key()
        if not key: QMessageBox.warning(self, "API Key Missing", "AI Rotation requires the Gemini API Key."); return
        self.btn_autobox.setEnabled(False); self._tab_log(f'[Mask] 🤖 AI Identifying semantic text blocks of {item.name} ...')
        self.job_stats['box_total'] += 1; self._update_summary()
        signals = DetectSignals(); signals.boxes_done.connect(self._on_ai_box_res); signals.boxes_error.connect(self._on_ai_box_err)
        worker = AIBoxDetectWorker(item.path, item.rotation_angle, key, self.proxy_widget.get_proxies(), signals)
        worker.setAutoDelete(True); self._thread_pool.start(worker)
    def _on_batch_auto_mask(self):
        pending = [i for i in self.file_items if i.status == 'pending']
        if not pending: return
        key = self._get_gemini_key()
        if not key: QMessageBox.warning(self, "API Key Missing", "AI requires the Gemini API Key."); return
        self.btn_batch_mask.setEnabled(False); self._tab_log(f'[Mask] 🤖 AI is performing batch detection of ({len(pending)} text boxes)...')
        self.job_stats['box_total'] = len(pending); self.job_stats['box_done'] = 0; self.job_stats['box_empty'] = 0; self.job_stats['box_err'] = 0; self._update_summary(); proxies = self.proxy_widget.get_proxies()
        for item in pending:
            signals = DetectSignals(); signals.boxes_done.connect(self._on_ai_box_res); signals.boxes_error.connect(self._on_ai_box_err)
            worker = AIBoxDetectWorker(item.path, item.rotation_angle, key, proxies, signals); worker.setAutoDelete(True); self._thread_pool.start(worker)
    def _on_ai_box_res(self, path, boxes):
        for item in self.file_items:
            if item.path == path:
                # 恢复 Tab 3 专属的多边形 (polygons) 转换逻辑
                polys = []
                for rx, ry, rw, rh in boxes:
                    polys.append([(rx, ry), (rx+rw, ry), (rx+rw, ry+rh), (rx, ry+rh)])
                item.polygons = polys
                item.box_detected = True
                
                if len(polys) == 0: 
                    self.job_stats['box_empty'] = self.job_stats.get('box_empty', 0) + 1
                    self._tab_log(f'⚠️ [Mask] Warning: {item.name} No text boxes detected!', is_warning=True)
                else: 
                    self.job_stats['box_done'] = self.job_stats.get('box_done', 0) + 1
                    self._tab_log(f'[Mask] Polygons created: {item.name} Found {len(polys)} paragraphs.')
                    
                if self.selected_idx is not None and self.file_items[self.selected_idx].path == path: 
                    self.img_label.update()
                break
        self.btn_autobox.setEnabled(True)
        self._update_summary()
        self._finalize_ai_worker()

    def _on_ai_box_err(self, path, msg):
        self.job_stats['box_err'] += 1; self._tab_log(f'[Mask] AI Boxing failed: {Path(path).name} - {msg}', is_error=True); self.btn_autobox.setEnabled(True); self._update_summary(); self._finalize_ai_worker()
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
        self.file_items[self.selected_idx].polygons.clear(); self.file_items[self.selected_idx].box_detected = False; self.img_label.update()
    def _get_config(self):
        api = self.engine_combo.currentText()
        if api == 'claude': key, model = self.claude_cfg.get_key(), self.claude_cfg.get_model()
        else: key, model = self.gemini_cfg.get_key(), self.gemini_cfg.get_model()
        return api, key, model, self.proxy_widget.get_proxies()
    def _on_start(self):
        pending = [i for i in self.file_items if i.status in ('pending', 'error') and len(i.polygons) > 0]
        if not pending: QMessageBox.information(self, 'Info', 'No pending or failed files with masks found.'); return
        api, key, model, proxies = self._get_config()
        if not key: QMessageBox.warning(self, 'Missing Key', f'Please ensure you have entered API Key of {api.upper()}.'); return
        self._stop_flag = False; self._batch_done = 0; self._batch_errors = 0; total = len(pending); done_counter = [0]
        signals = WorkerSignals(); signals.finished.connect(self._on_ocr_done); signals.error.connect(self._on_ocr_error); signals.progress.connect(lambda d, t: self._update_progress(d, t))
        signals.excluded_info.connect(self._on_excluded_numeric)
        self.progress_bar.setVisible(True); self.progress_bar.setMaximum(total); self.progress_bar.setValue(0); self.progress_label.setText(f'0 / {total}'); self.btn_gen_atf.setEnabled(False)
        self._tab_log(f'[Mask] Starting {total} files using {api}')
        for item in pending:
            if self._stop_flag: break
            item.status = 'processing'
            worker = MaskOCRRunnable(api, item.path, item.polygons, item.rotation_angle, key, model, proxies, signals, done_counter, total, self.chk_exclude_numeric.isChecked())
            worker.setAutoDelete(True); self._thread_pool.start(worker)
        self._refresh_list(); self._update_summary()
    def _on_reocr(self):
        if self.selected_idx is None: return
        item = self.file_items[self.selected_idx]
        if not item.polygons: QMessageBox.warning(self, 'Warning', 'Please draw masks on the image before extracting text.'); return
        api, key, model, proxies = self._get_config()
        if not key: QMessageBox.warning(self, 'Missing Key', 'Missing API Key.'); return
        
        self._tab_log(f"↺ Re-OCR started for {item.name} (Mask Mode)...")
        item.status = 'processing'; self._refresh_list(); self.btn_reocr.setEnabled(False)
        signals = WorkerSignals(); signals.finished.connect(self._on_ocr_done); signals.error.connect(self._on_ocr_error); signals.progress.connect(lambda d, t: None)
        signals.excluded_info.connect(self._on_excluded_numeric)
        worker = MaskOCRRunnable(api, item.path, item.polygons, item.rotation_angle, key, model, proxies, signals, [0], 1, self.chk_exclude_numeric.isChecked())
        worker.setAutoDelete(True); self._thread_pool.start(worker)
        
    def _on_excluded_numeric(self, path, count):
        self._tab_log(f'[Mask] {Path(path).name}: 已排除 {count} 个纯数字/单字母条目 (excluded {count} pure-number/single-letter snippet(s))')

    def _on_gen_atf(self):
        done_items = [item for item in self.file_items if item.status == 'done' and item.ocr_text.strip()]
        if not done_items: return
        skipped = [item for item in self.file_items if item.status == 'error']
        if skipped:
            reply = QMessageBox.question(self, 'Info', f'{len(skipped)} image(s) failed OCR and will be skipped (not included in the ATF).\n有 {len(skipped)} 张图片 OCR 失败，将被跳过，不会出现在 ATF 文档里。\n\nContinue? 是否继续？', QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.Yes)
            if reply != QMessageBox.StandardButton.Yes: return
        
        # 核心修复：动态获取当前下拉框中选中的引擎，取代写死的 self.api_name
        current_engine = self.engine_combo.currentText()
        
        out_path, _ = QFileDialog.getSaveFileName(self, 'Save ATF File', os.path.join(str(Path(done_items[0].path).parent), f'ATF_{current_engine}.docx'), 'Word Documents (*.docx)')
        if not out_path: return
        
        from datetime import datetime
        ts = datetime.now().strftime('%Y%m%d_%H%M%S'); base, ext = os.path.splitext(out_path)
        if not any(c.isdigit() for c in Path(out_path).stem[-8:]): out_path = f'{base}_{ts}{ext}'
        
        try:
            dedup = self.chk_dedup.isChecked()
            processed_items = [{'name': item.name, 'image_path': item.path, 'ocr_text': _reorder_ocr_text_by_reading_order(item.ocr_text, item.polygons)}
                for item in done_items]
            build_atf_docx(processed_items, out_path, dedup=dedup)
            
            self.export_status = 'Success'
            self._tab_log(f'ATF generated: {out_path}')
            self._update_summary()
            
            import subprocess, sys as _sys
            try:
                abs_out = os.path.abspath(out_path)
                if _sys.platform == 'win32': os.startfile(abs_out)
                elif _sys.platform == 'darwin': subprocess.Popen(['open', abs_out])
                else: subprocess.Popen(['xdg-open', abs_out])
            except Exception as open_err:
                self._tab_log(f"File saved but auto-open failed: {open_err}", is_warning=True)
                
        except Exception as e:
            self.export_status = 'Failed'
            self._tab_log(f'ATF generation failed: {e}', is_error=True)
            self._update_summary()
            QMessageBox.critical(self, 'Error', f'Generation failed:\n{e}\n\n{traceback.format_exc()}')

# ─────────────────────────────────────────────────────────────
# TAB 4: NEW IMAGE->PSD BUILDER (终极修复：精确透明检测 + 动态底图架构 + 纯正色彩)
# ─────────────────────────────────────────────────────────────

from PyQt6.QtCore import QRectF
from PIL import PngImagePlugin

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
    "visual_line_count": 2
  }
]

RULES:
1. "id" MUST match the numeric tag.
2. "text" is the exact text. NEVER swallow spaces.
3. "color_hex" must be a 6-character hex code representing the dominant color.
4. "is_bold" is true ONLY IF the text is a thick heading/title.
5. "font_style" MUST be either "sans-serif" (clean, straight edges) or "serif" (with decorative strokes at the ends).
6. "visual_line_count" MUST be the exact integer number of physical lines this text spans.
"""

GEMINI_ERASE_PROMPT = """【系统强制指令】请注意图片中被『品红色(#FF00FF)』涂满的区域。
你是一个图片文字处理专家，将蒙版范围全部抹除，并保留背景不变。
强制要求：绝对不要改变图片的亮度、对比度、颜色值、色彩空间、颜色模式等。
绝对不要重绘整个图片 
非蒙版范围内的文字一定不要删除！ 
不要生成新的文字"""

def _tab4_create_redacted_image_base64(image_path: str, boxes: list) -> tuple[str, str]:
    """专属 Tab4：适配带有对齐参数(长度为5)的 boxes 数据结构"""
    with _pil_draw_lock:
        with PILImage.open(image_path) as src: 
            img = ImageOps.exif_transpose(src).convert('RGB')
        
        img_w, img_h = img.size; crops = []
    
        for box in boxes:
            rx, ry, rw, rh = box[:4]
            x, y = int(rx * img_w), int(ry * img_h); w, h = int(rw * img_w), int(rh * img_h)
            x1, y1 = max(0, x-2), max(0, y-2); x2, y2 = min(img_w, x+w+2), min(img_h, y+h+2)
            crops.append(img.crop((x1, y1, x2, y2)))
        
        # 不再依赖当前这台电脑系统字体目录里的 arial.ttf —— 不同电脑上这个文件
        # 可能不是同一份数据（不同 Windows/语言包/更新历史），FreeType 解析到
        # 异常字体数据时可能触发原生崩溃，try/except 接不住。改用 Pillow 自带、
        # 随包附带、所有机器上保证一致的默认字体，彻底去掉这个变量。
        try: font = ImageFont.load_default(size=20)
        except TypeError: font = ImageFont.load_default()
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


class PsdAIBoxDetectWorker(QRunnable):
    def __init__(self, path, rot_state, api_key, proxy, signals):
        super().__init__()
        self.path = path; self.api_key = api_key; self.proxy = proxy; self.signals = signals

    @pyqtSlot()
    def run(self):
        try:
            if self.proxy: 
                os.environ['HTTP_PROXY'] = self.proxy.get('http', '').replace('http://', '')
                os.environ['HTTPS_PROXY'] = self.proxy.get('https', '').replace('http://', '')
            with PILImage.open(self.path) as src: 
                working_img = ImageOps.exif_transpose(src)
                if working_img.mode == 'CMYK': working_img = working_img.convert('RGB')
                working_img = working_img.copy()
                
            prompt = (
                "You are an expert layout analyzer. Detect all distinct SEMANTIC TEXT BLOCKS in this image. "
                "Return exactly ONE box per semantic block. "
                "CRITICAL: Also detect the text alignment for each block: 'L' (Left), 'C' (Center), or 'R' (Right). "
                "Format strictly as a JSON list of arrays: [[ymin, xmin, ymax, xmax, \"L\"], ...]. "
                "Scale coordinates to 0-1000."
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
                
            matches = re.findall(r'\[\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*["\']?([LRC])["\']?\s*\]', text_res, re.IGNORECASE)
            if not matches:
                matches = re.findall(r'\[\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*\]', text_res)
                matches = [(y1, x1, y2, x2, 'L') for y1, x1, y2, x2 in matches]

            raw_boxes = []
            for m in matches:
                ymin, xmin, ymax, xmax = map(float, m[:4])
                align = m[4].upper() if len(m) > 4 else 'L'
                if align not in ['L', 'C', 'R']: align = 'L'
                if ymax > ymin and xmax > xmin: raw_boxes.append([ymin, xmin, ymax, xmax, align])
                
            final_boxes = []
            for i, b1 in enumerate(raw_boxes):
                y1_min, x1_min, y1_max, x1_max, align1 = b1
                area1 = (y1_max - y1_min) * (x1_max - x1_min); is_contained = False
                for j, b2 in enumerate(raw_boxes):
                    if i == j: continue
                    y2_min, x2_min, y2_max, x2_max, align2 = b2
                    area2 = (y2_max - y2_min) * (x2_max - x2_min)
                    overlap_ymin = max(y1_min, y2_min); overlap_xmin = max(x1_min, x2_min)
                    overlap_ymax = min(y1_max, y2_max); overlap_xmax = min(x1_max, x2_max)
                    if overlap_ymax > overlap_ymin and overlap_xmax > overlap_xmin:
                        overlap_area = (overlap_ymax - overlap_ymin) * (overlap_xmax - overlap_xmin)
                        if overlap_area > 0.8 * area1 and area2 > area1: is_contained = True; break
                if not is_contained: final_boxes.append(b1)
                
            boxes = []
            for ymin, xmin, ymax, xmax, align in final_boxes:
                ymin = max(0, ymin)
                xmin = max(0, xmin)
                ymax = min(1000, ymax)
                xmax = min(1000, xmax)
                boxes.append([xmin/1000.0, ymin/1000.0, (xmax-xmin)/1000.0, (ymax-ymin)/1000.0, align])
            boxes.sort(key=lambda b: b[1])
            self.signals.boxes_done.emit(self.path, boxes)
        except Exception as e: self.signals.boxes_error.emit(self.path, str(e))


class PsdInteractiveMaskLabel(QLabel):
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
        
    def set_item(self, item):
        self._zoom_factor = 1.0; self.pan_y = 0; self.file_item = item; pix = QPixmap(item.path)
        if pix.isNull(): self.source_pixmap = None; self.setText('Cannot load image')
        else:
            self.source_pixmap = pix; self.setText('')
        self.update()
        
    def clear_item(self):
        self._zoom_factor = 1.0; self.pan_y = 0; self.file_item = None; self.source_pixmap = None; self.setText('← Select an image'); self.update()
        
    def wheelEvent(self, event):
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            if event.angleDelta().y() > 0: self._zoom_factor = min(5.0, self._zoom_factor + 0.1)
            else: self._zoom_factor = max(0.1, self._zoom_factor - 0.1)
        else: self.pan_y += event.angleDelta().y()
        self.update()
        
    def paintEvent(self, event):
        super().paintEvent(event)
        if not self.source_pixmap or not self.file_item: return
        
        painter = QPainter(self); painter.setRenderHint(QPainter.RenderHint.Antialiasing); painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        v_w, v_h = self.width(), self.height(); p_w, p_h = self.source_pixmap.width(), self.source_pixmap.height()
        base_scale = min(v_w / p_w, v_h / p_h); self.scale_factor = base_scale * self._zoom_factor
        
        scaled_w = p_w * self.scale_factor; scaled_h = p_h * self.scale_factor
        self.offset_x = (v_w - scaled_w) / 2.0; self.offset_y = (v_h - scaled_h) / 2.0 + self.pan_y
        
        scaled_pixmap = self.source_pixmap.scaled(
            max(1, int(scaled_w)), max(1, int(scaled_h)), 
            Qt.AspectRatioMode.IgnoreAspectRatio, Qt.TransformationMode.SmoothTransformation
        )
        painter.drawPixmap(QRectF(self.offset_x, self.offset_y, scaled_w, scaled_h), scaled_pixmap, QRectF(scaled_pixmap.rect()))
        
        pen_box = QPen(QColor(79, 142, 247, 220), 2, Qt.PenStyle.SolidLine); painter.setPen(pen_box)
        
        # --- 核心修改点 1：配置更小的加粗字体 ---
        small_font = painter.font()
        small_font.setPointSize(8) # 字号调小
        small_font.setBold(True)
        
        for i, box in enumerate(self.file_item.boxes):
            rx, ry, rw, rh = box[:4]
            align = box[4] if len(box) > 4 else 'L'
            x = rx * scaled_w + self.offset_x; y = ry * scaled_h + self.offset_y
            w = rw * scaled_w; h = rh * scaled_h
            
            painter.setBrush(QBrush(QColor(79, 142, 247, 40)))
            painter.drawRect(QRectF(x, y, w, h))
            
            # --- 核心修改点 2：应用小字体，并将文本绘制在框的左侧外边缘 ---
            painter.setFont(small_font)
            painter.setPen(QColor(245, 158, 11)) # 保持高对比度橙色
            
            # 在框外左侧留出 45 像素的绘制区域，高度与框一致，采用“右对齐 + 垂直居中”
            text_rect = QRectF(x - 45, y, 40, h)
            painter.drawText(text_rect, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, f"[{i+1}]{align}")
            
            painter.setPen(pen_box) 
            
        if self.drawing and self.start_pos and self.current_pos:
            painter.setPen(QPen(QColor(34, 197, 94, 255), 2, Qt.PenStyle.DashLine)); painter.setBrush(QBrush(QColor(34, 197, 94, 40)))
            painter.drawRect(QRectF(self.start_pos, self.current_pos).normalized())
        if self.is_erasing and self.start_pos and self.current_pos:
            painter.setPen(QPen(QColor(239, 68, 68, 255), 2, Qt.PenStyle.DashLine)); painter.setBrush(QBrush(QColor(239, 68, 68, 40)))
            painter.drawRect(QRectF(self.start_pos, self.current_pos).normalized())

    def _get_box_at(self, pos):
        if not self.file_item or not self.source_pixmap: return None
        scaled_w = self.source_pixmap.width() * self.scale_factor; scaled_h = self.source_pixmap.height() * self.scale_factor
        for box in reversed(self.file_item.boxes):
            rx, ry, rw, rh = box[:4]
            x = rx * scaled_w + self.offset_x; y = ry * scaled_h + self.offset_y
            w = rw * scaled_w; h = rh * scaled_h
            if QRectF(x, y, w, h).contains(pos): return box
        return None

    def mousePressEvent(self, event):
        if not self.file_item or not self.source_pixmap: return
        self.setFocus()
        if event.button() == Qt.MouseButton.LeftButton:
            clicked_box = self._get_box_at(event.position())
            if clicked_box:
                curr_align = clicked_box[4] if len(clicked_box) > 4 else 'L'
                next_align = {'L': 'C', 'C': 'R', 'R': 'L'}.get(curr_align, 'L')
                if len(clicked_box) > 4: clicked_box[4] = next_align
                else: clicked_box.append(next_align)
                self.update(); self.boxes_changed.emit()
            else: 
                self.start_pos = event.position(); self.current_pos = event.position(); self.drawing = True; self.update()
        elif event.button() == Qt.MouseButton.RightButton: 
            self.start_pos = event.position(); self.current_pos = event.position(); self.is_erasing = True; self.update()

    def mouseMoveEvent(self, event):
        if self.drawing or self.is_erasing: self.current_pos = event.position(); self.update()

    def mouseReleaseEvent(self, event):
        if not self.file_item or not self.source_pixmap: return
        scaled_w = self.source_pixmap.width() * self.scale_factor; scaled_h = self.source_pixmap.height() * self.scale_factor
        
        if event.button() == Qt.MouseButton.LeftButton and self.drawing:
            self.drawing = False; self.current_pos = event.position()
            rect = QRectF(self.start_pos, self.current_pos).normalized(); pixmap_rect = QRectF(self.offset_x, self.offset_y, scaled_w, scaled_h)
            intersected = rect.intersected(pixmap_rect)
            
            if not intersected.isEmpty() and intersected.width() > 5 and intersected.height() > 5:
                rx = (intersected.x() - self.offset_x) / scaled_w; ry = (intersected.y() - self.offset_y) / scaled_h
                rw = intersected.width() / scaled_w; rh = intersected.height() / scaled_h
                self.file_item.boxes.append([rx, ry, rw, rh, 'L'])
                self.file_item.boxes.sort(key=lambda b: b[1])
                self.boxes_changed.emit()
            self.update()
            
        elif event.button() == Qt.MouseButton.RightButton and self.is_erasing:
            self.is_erasing = False; self.current_pos = event.position()
            erase_rect = QRectF(self.start_pos, self.current_pos).normalized(); remaining = []
            for box in self.file_item.boxes:
                rx, ry, rw, rh = box[:4]
                x = rx * scaled_w + self.offset_x; y = ry * scaled_h + self.offset_y
                w = rw * scaled_w; h = rh * scaled_h
                if not erase_rect.intersects(QRectF(x, y, w, h)): remaining.append(box)
            if len(remaining) != len(self.file_item.boxes): self.file_item.boxes = remaining; self.boxes_changed.emit()
            self.update()


class ImageToPsdWorker(QRunnable):
    def __init__(self, api, image_path, boxes, rotation_angle, api_key, model, proxies, custom_fonts, signals, done_counter, total):
        super().__init__()
        self.api = api
        self.image_path = image_path
        self.boxes = boxes
        self.rotation_angle = rotation_angle 
        self.api_key = api_key
        self.model = model
        self.proxies = proxies
        self.custom_fonts = custom_fonts
        self.signals = signals
        self.done_counter = done_counter
        self.total = total

    def _check_has_alpha(self, img_path):
        """精准检测图片是否包含实质性的透明像素"""
        try:
            with PILImage.open(img_path) as img:
                if img.mode in ('RGBA', 'LA'):
                    extrema = img.getextrema()
                    if img.mode == 'RGBA':
                        return extrema[3][0] < 255
                    else:
                        return extrema[1][0] < 255
                elif img.mode == 'P':
                    if 'transparency' in img.info:
                        return True
            return False
        except:
            return False

    def _generate_background(self, temp_dir, unique_id, is_trans):
        try:
            with PILImage.open(self.image_path) as orig:
                orig_exif = ImageOps.exif_transpose(orig)
                orig_size = orig_exif.size
                orig_mode = orig_exif.mode
                orig_format = (orig_exif.format or "PNG").upper()
                orig_icc = orig_exif.info.get('icc_profile', None) 
                orig_exif_data = orig_exif.info.get('exif', None)
                orig_dpi = orig_exif.info.get('dpi', None)
                
                if is_trans:
                    working_img = orig_exif.convert("RGBA")
                    draw = ImageDraw.Draw(working_img)
                    w, h = working_img.size
                    for box in self.boxes:
                        rx, ry, rw, rh = box[:4]
                        x, y = int(rx * w), int(ry * h)
                        x2, y2 = int((rx+rw) * w), int((ry+rh) * h)
                        
                        box_height = y2 - y
                        # 核心修改：底部动态向下延伸 20% 的高度，用于包裹 g,p,y,q 等下沉字母
                        descender_padding = int(box_height * 0.20) 
                        
                        # 将框内全部置为透明 (0,0,0,0)，注意底部的 y2 + descender_padding
                        draw.rectangle([
                            max(0, x - 3),                 # 左侧多挖3像素防锯齿
                            max(0, y - 3),                 # 顶部多挖3像素防锯齿
                            min(w, x2 + 3),                # 右侧多挖3像素防锯齿
                            min(h, y2 + descender_padding) # 底部多挖20%高度包裹下沉字符
                        ], fill=(0,0,0,0))
                    
                    save_path = os.path.join(temp_dir, f"bg_{unique_id}.png")
                    save_kw = {"format": "PNG"}
                    if orig_dpi:
                        pnginfo = PngImagePlugin.PngInfo()
                        ppm_x, ppm_y = int(orig_dpi[0] * 39.3701), int(orig_dpi[1] * 39.3701)
                        pnginfo.add(b'pHYs', b'%s%s\x01' % (ppm_x.to_bytes(4, 'big'), ppm_y.to_bytes(4, 'big')))
                        save_kw["pnginfo"] = pnginfo
                    if orig_icc: save_kw["icc_profile"] = orig_icc
                    
                    working_img.save(save_path, **save_kw)
                    return save_path
                    
                working_img = orig_exif.convert('RGB')
                draw = ImageDraw.Draw(working_img)
                w, h = working_img.size
                for box in self.boxes:
                    rx, ry, rw, rh = box[:4]
                    x, y = int(rx * w), int(ry * h)
                    bw, bh = int(rw * w), int(rh * h)
                    
                    # 核心修改：底部动态向下延伸 20% 的高度
                    descender_padding = int(bh * 0.20)
                    
                    draw.rectangle([
                        max(0, x - 3), 
                        max(0, y - 3), 
                        min(w, x + bw + 3), 
                        min(h, y + bh + descender_padding) # 底部向下延伸
                    ], fill="#FF00FF")

            img_buf = io.BytesIO()
            working_img.save(img_buf, format="JPEG")
            img_bytes = img_buf.getvalue()

            from google import genai
            from google.genai import types
            client = genai.Client(api_key=self.api_key)
            if self.proxies:
                os.environ['HTTP_PROXY'] = self.proxies.get('http', '').replace('http://', '')
                os.environ['HTTPS_PROXY'] = self.proxies.get('https', '').replace('http://', '')

            res = client.models.generate_content(
                model="gemini-3.1-flash-image-preview",
                config=types.GenerateContentConfig(temperature=0.0),
                contents=[GEMINI_ERASE_PROMPT, types.Part.from_bytes(data=img_bytes, mime_type="image/jpeg")]
            )

            if not (res.candidates and res.candidates[0].content.parts):
                raise Exception("Empty response from Gemini.")
            
            raw_data = res.candidates[0].content.parts[0].inline_data.data
            if isinstance(raw_data, str): raw_data = base64.b64decode(raw_data)
            gemini_out = PILImage.open(io.BytesIO(raw_data))

            if gemini_out.size != orig_size:
                gemini_out = gemini_out.resize(orig_size, PILImage.Resampling.LANCZOS)
            
            if orig_mode == "CMYK":
                gemini_out = gemini_out.convert("RGB").convert("CMYK")
            elif gemini_out.mode != orig_mode:
                gemini_out = gemini_out.convert(orig_mode)

            ext = ".jpg" if orig_format in ("JPEG", "JPG") else ".png"
            save_path = os.path.join(temp_dir, f"bg_{unique_id}{ext}")

            save_kw = {}
            if orig_dpi: save_kw["dpi"] = orig_dpi
            if orig_exif_data: save_kw["exif"] = orig_exif_data
            
            if orig_icc:
                if ext == ".png" and orig_mode == "CMYK": pass
                else: save_kw["icc_profile"] = orig_icc

            if ext == ".jpg":
                save_kw["format"] = "JPEG"
                save_kw["quality"] = 100
                save_kw["subsampling"] = 0
            else:
                save_kw["format"] = "PNG"
                if orig_dpi:
                    pnginfo = PngImagePlugin.PngInfo()
                    ppm_x, ppm_y = int(orig_dpi[0] * 39.3701), int(orig_dpi[1] * 39.3701)
                    pnginfo.add(b'pHYs', b'%s%s\x01' % (ppm_x.to_bytes(4, 'big'), ppm_y.to_bytes(4, 'big')))
                    save_kw["pnginfo"] = pnginfo

            gemini_out.save(save_path, **save_kw)
            return save_path

        except Exception as e:
            self.signals.finished.emit(self.image_path, f"[Warning] Gemini failed to erase the text (Erase Error): {str(e)}")
            return None

    @pyqtSlot()
    def run(self):
        jsx_path = None
        json_path = None
        bg_path = None
        try:
            if not HAS_WIN32COM:
                raise Exception("Windows pywin32 is required to run Photoshop COM.")
            
            import tempfile, uuid
            temp_dir = os.path.abspath(tempfile.gettempdir())
            unique_id = str(uuid.uuid4())
            
            is_trans = self._check_has_alpha(self.image_path)
            
            self.signals.finished.emit(self.image_path, "Step 1/3: Preparing Background Layer...")
            bg_path = self._generate_background(temp_dir, unique_id, is_trans)
            
            self.signals.finished.emit(self.image_path, "Step 2/3: AI analyzing typography...")
            b64, mime = _tab4_create_redacted_image_base64(self.image_path, self.boxes)
            prompt = PSD_BUILDER_PROMPT + f"\n\nATTENTION: There are exactly {len(self.boxes)} snippets in this image. You MUST output exactly {len(self.boxes)} JSON objects."
            
            from google import genai
            from google.genai import types
            client = genai.Client(api_key=self.api_key)
            if self.proxies:
                os.environ['HTTP_PROXY'] = self.proxies.get('http', '').replace('http://', '')
                os.environ['HTTPS_PROXY'] = self.proxies.get('https', '').replace('http://', '')
                
            res = client.models.generate_content(
                model="gemini-3.1-flash-image-preview", 
                config=types.GenerateContentConfig(temperature=0.0), 
                contents=[prompt, types.Part.from_bytes(data=base64.b64decode(b64), mime_type=mime)]
            )
            raw_text = res.text if (res.candidates and res.candidates[0].content.parts) else ""

            clean_text = re.sub(r'(?:json)?\n?(.*?)\n? ', r'\1', raw_text, flags=re.DOTALL).strip()
            try: data = json.loads(clean_text)
            except:
                match = re.search(r'\[\s*{.*?}\s*\]', raw_text, flags=re.DOTALL)
                if match:
                    try: data = json.loads(match.group(0))
                    except: data = []
                else: data = []

            if not data:
                raise Exception("Failed to parse AI response into JSON format.")

            js_payload = []
            for d in data:
                idx = d.get('id', 0) - 1
                if 0 <= idx < len(self.boxes):
                    rx, ry, rw, rh = self.boxes[idx][:4]
                    align_char = self.boxes[idx][4] if len(self.boxes[idx]) > 4 else 'L'
                    
                    raw_text_str = d.get('text', '')
                    text_str = re.sub(r'[\r\n]+', ' ', raw_text_str).strip()
                    
                    visual_lines = int(d.get('visual_line_count', 1))
                    if visual_lines < 1: visual_lines = 1
                    
                    align_map = {'L': 'left', 'C': 'center', 'R': 'right'}
                    final_align = align_map.get(align_char, 'left')
                    
                    font_style = d.get('font_style', 'sans-serif').lower()
                    
                    js_payload.append({
                        'id': d.get('id'), 
                        'text': text_str,
                        'rx': rx, 'ry': ry, 'rw': rw, 'rh': rh,
                        'visual_lines': visual_lines, 'color': d.get('color_hex', '#FFFFFF'),
                        'align': final_align, 'is_bold': d.get('is_bold', False),  
                        'font_style': font_style,
                        'custom_fonts': self.custom_fonts
                    })

            json_path = os.path.join(temp_dir, f"psd_data_{unique_id}.json")
            jsx_path = os.path.join(temp_dir, f"psd_builder_{unique_id}.jsx")
            
            with open(json_path, 'w', encoding='utf-8-sig') as f:
                json.dump(js_payload, f)

            base_dir = os.path.abspath(os.path.dirname(self.image_path))
            psd_save_path = os.path.join(base_dir, f"{Path(self.image_path).stem}.psd")
            
            target_img_js = os.path.abspath(self.image_path).replace('\\', '/')
            psd_save_js = psd_save_path.replace('\\', '/')
            json_path_js = json_path.replace('\\', '/')
            bg_path_js = bg_path.replace('\\', '/') if bg_path else ""
            jsx_is_trans = "true" if is_trans else "false"

            # Step 3: Photoshop JSX (绝对单次执行！修复了设置内容的顺序，防止 PS 字体充置)
            self.signals.finished.emit(self.image_path, "Step 3/3: Building PSD via Adobe COM...")

            jsx_template = r"""
            var originalRuler = app.preferences.rulerUnits;
            var originalType = app.preferences.typeUnits;
            try {
                app.preferences.rulerUnits = Units.PIXELS;
                app.preferences.typeUnits = TypeUnits.PIXELS;

                var jsonFile = new File("__JSON_PATH__");
                jsonFile.open("r"); jsonFile.encoding = "UTF-8";
                var jsonString = jsonFile.read(); jsonFile.close();
                var data = eval("(" + jsonString + ")");

                var doc = app.open(new File("__TARGET_IMG__"));
                var docW = doc.width.value; 
                var docH = doc.height.value;
                var docRes = doc.resolution;
                var ptRatio = 72.0 / docRes;
                var isTransparent = __IS_TRANS__;

function resolveFont(familyInput, isBoldReq) {
                    if (!familyInput) return isBoldReq ? "Arial-BoldMT" : "ArialMT";
                    var lowerInput = familyInput.toLowerCase().replace(/\s+/g, '');
                    var variantWords = ["narrow", "condensed", "cond", "extended", "expanded", "wide"];
                    function hasUnwantedVariant(name) {
                        for (var v = 0; v < variantWords.length; v++) {
                            if (name.indexOf(variantWords[v]) !== -1 && lowerInput.indexOf(variantWords[v]) === -1) return true;
                        }
                        return false;
                    }
                    var familyMembers = [];
                    var prefixMembers = [];
                    var exactMatch = null;
                    
                    for (var f = 0; f < app.fonts.length; f++) {
                        var psFont = app.fonts[f];
                        var fam = (psFont.family || "").toLowerCase().replace(/\s+/g, '');
                        var pName = psFont.postScriptName || psFont.name || "";
                        var psName = pName.toLowerCase().replace(/\s+/g, '');
                        
                        if (psName === lowerInput) exactMatch = pName;
                        // 关键：家族名完全相等才算同一家族；并排除 narrow/condensed 等异形变体
                        if (fam === lowerInput) {
                            if (!hasUnwantedVariant(psName)) familyMembers.push(psFont);
                        } else if (psName.indexOf(lowerInput) === 0 && !hasUnwantedVariant(psName)) {
                            prefixMembers.push(psFont);
                        }
                    }
                    
                    // 只有在没有精确家族成员时，才退而用前缀匹配兜底
                    if (familyMembers.length === 0) familyMembers = prefixMembers;
                    
                    if (familyMembers.length > 0) {
                        for (var i = 0; i < familyMembers.length; i++) {
                            var style = (familyMembers[i].style || "").toLowerCase();
                            var fName = familyMembers[i].postScriptName || familyMembers[i].name || "";
                            if (isBoldReq === true || isBoldReq === "true") {
                                if (style.indexOf("bold") !== -1 || style.indexOf("black") !== -1 || 
                                    style.indexOf("heavy") !== -1 || style.indexOf("w6") !== -1 || 
                                    style.indexOf("700") !== -1 || style.indexOf("semibold") !== -1 || fName.toLowerCase().indexOf("bold") !== -1) {
                                    return fName;
                                }
                            } else {
                                if (style === "regular" || style === "normal" || style === "book" || 
                                    style === "medium" || style === "light" || fName.toLowerCase().indexOf("regular") !== -1) {
                                    return fName;
                                }
                            }
                        }
                        if (exactMatch) return exactMatch;
                        return familyMembers[0].postScriptName || familyMembers[0].name;
                    }
                    
                    if (exactMatch) return exactMatch;
                    return isBoldReq ? "Arial-BoldMT" : "ArialMT";
                }

                var bgLayer = doc.layers[doc.layers.length-1];
                if (bgLayer.isBackgroundLayer) bgLayer.isBackgroundLayer = false;
                var originalLayer = bgLayer;
                originalLayer.name = "1. 原图（隐藏）";
                originalLayer.visible = false;

                if (isTransparent) {
                    var bgFilePath = "__BG_PATH__";
                    if (bgFilePath !== "") {
                        var bgFile = new File(bgFilePath);
                        if (bgFile.exists) {
                            var bgDoc = app.open(bgFile);
                            bgDoc.selection.selectAll(); bgDoc.selection.copy();
                            bgDoc.close(SaveOptions.DONOTSAVECHANGES);
                            app.activeDocument = doc;
                            var newLayer = doc.paste();
                            newLayer.name = "2. 纯透明背景";
                            newLayer.move(originalLayer, ElementPlacement.PLACEBEFORE);
                            var b = newLayer.bounds;
                            newLayer.translate(0 - b[0].value, 0 - b[1].value);
                        }
                    }
                } else {
                    var psLayer = originalLayer.duplicate();
                    psLayer.name = "2. PS 内容识别背景";
                    psLayer.visible = true;
                    doc.activeLayer = psLayer;
                    for(var i=0; i<data.length; i++) {
                        var b = data[i];
                        var x1 = docW * b.rx; var y1 = docH * b.ry;
                        var x2 = docW * (b.rx + b.rw); var y2 = docH * (b.ry + b.rh);
                        doc.selection.select([[x1, y1], [x2, y1], [x2, y2], [x1, y2]]);
                        try {
                            var desc = new ActionDescriptor();
                            desc.putEnumerated(charIDToTypeID("Usng"), stringIDToTypeID("contentAware"), stringIDToTypeID("contentAware"));
                            executeAction(charIDToTypeID("Fl  "), desc, DialogModes.NO);
                        } catch(e) {}
                        doc.selection.deselect();
                    }
                    var geminiFilePath = "__BG_PATH__";
                    if (geminiFilePath !== "") {
                        var geminiFile = new File(geminiFilePath);
                        if (geminiFile.exists) {
                            var gDoc = app.open(geminiFile);
                            gDoc.selection.selectAll(); gDoc.selection.copy();
                            gDoc.close(SaveOptions.DONOTSAVECHANGES);
                            app.activeDocument = doc;
                            var gLayer = doc.paste();
                            gLayer.name = "3. Gemini 抹除背景";
                            gLayer.move(psLayer, ElementPlacement.PLACEBEFORE);
                            var gb = gLayer.bounds;
                            gLayer.translate(0 - gb[0].value, 0 - gb[1].value);
                        }
                    }
                }

                for(var i=0; i<data.length; i++) {
                    var b = data[i];
                    var x1 = docW * b.rx; var y1 = docH * b.ry;
                    var w = docW * b.rw; var h = docH * b.rh;
                    var layer = doc.artLayers.add(); layer.kind = LayerKind.TEXT;
                    
                    var safeName = b.text.substring(0, 20);
                    layer.name = safeName ? safeName : "Text_" + b.id;
                    
                    var textItem = layer.textItem; 
                    textItem.kind = TextType.PARAGRAPHTEXT;
                    
                    // 【绝对防御机制】：在 PS 中，必须首先填入文本内容，再更改其样式。
                    textItem.contents = b.text;
                    
                    textItem.position = Array(x1, y1);
                    textItem.width = w * ptRatio;  
                    textItem.height = h * ptRatio; 
                    
                    var pxLeading = h / b.visual_lines; 
                    var pxSize = (b.visual_lines > 1) ? (pxLeading * 0.78) : (pxLeading * 1.0);
                    
                    textItem.useAutoLeading = false;
                    try { textItem.leading = pxLeading; } catch(e) {}
                    try { textItem.size = pxSize * ptRatio; } catch(e) {}
                    
                    try {
                        var baseFamily = b.custom_fonts.base_font; 
                        var resolved = resolveFont(baseFamily, b.is_bold);
                        if (resolved) textItem.font = resolved;
                    } catch(e) { 
                        try { textItem.font = b.is_bold ? "Arial-BoldMT" : "ArialMT"; } catch(e2) {} 
                    }
                    
                    try {
                        var sc = new SolidColor();
                        sc.rgb.hexValue = b.color.replace("#", "");
                        textItem.color = sc;
                    } catch(e) {}
                    
                    try {
                        if (b.align == "center") textItem.justification = Justification.CENTER;
                        else if (b.align == "right") textItem.justification = Justification.RIGHT;
                        else textItem.justification = Justification.LEFT;
                    } catch(e) {}
            }

                var saveFile = new File("__SAVE_PATH__");
                var psdSaveOptions = new PhotoshopSaveOptions();
                psdSaveOptions.layers = true; psdSaveOptions.alphaChannels = true;
                doc.saveAs(saveFile, psdSaveOptions, true, Extension.LOWERCASE);
                doc.close(SaveOptions.DONOTSAVECHANGES);

            } finally {
                app.preferences.rulerUnits = originalRuler; 
                app.preferences.typeUnits = originalType;
            }
            """

            jsx_code = jsx_template.replace("__JSON_PATH__", json_path_js) \
                                   .replace("__TARGET_IMG__", target_img_js) \
                                   .replace("__SAVE_PATH__", psd_save_js) \
                                   .replace("__BG_PATH__", bg_path_js) \
                                   .replace("__IS_TRANS__", jsx_is_trans)

            with open(jsx_path, 'w', encoding='utf-8-sig') as f:
                f.write(jsx_code)

            import win32com.client, pythoncom
            pythoncom.CoInitialize()
            ps_app = win32com.client.Dispatch("Photoshop.Application")
            ps_app.DisplayDialogs = 3
            ps_app.DoJavaScriptFile(jsx_path)
            
            self.signals.finished.emit(self.image_path, "_SUCCESS_")

        except Exception as e:
            self.signals.error.emit(self.image_path, f"Error: {str(e)}")
        finally:
            for temp_file in [jsx_path, json_path, bg_path]:
                try:
                    if temp_file and os.path.exists(temp_file): 
                        os.remove(temp_file)
                except: pass
                
            self.done_counter[0] += 1
            self.signals.progress.emit(self.done_counter[0], self.total)
            try: pythoncom.CoUninitialize()
            except: pass


class ImageToPsdTab(MaskApiTab):
    def __init__(self, api_name: str, models: list, proxy_widget: ProxyWidget, parent=None):
        self._main_btn_style = 'background: #ec4899; color: white; font-weight: 600; border-radius: 6px; padding: 0 16px;'
        super().__init__(api_name, models, proxy_widget, parent)
        self._rotation_model = 'gemini-3.1-flash-image-preview'  # Image-to-PSD标签页保持原样，不跟着ATF creator换模型
        # 初始化主按钮文本
        self.btn_start.setText(T('img2psd_start'))
        self.btn_start.setStyleSheet(self._main_btn_style)
        self.btn_gen_atf.hide()
        self.btn_reocr.hide()
        self.ocr_hdr_lbl.setText("JSON ATTRIBUTES / PROGRESS")
        self._is_single_mode = False

    def _build_ui(self, models):
        super()._build_ui(models)
        # 保持原有的 img_label 替换逻辑不变
        if hasattr(self, 'ui_img_layout') and hasattr(self, 'img_label'):
            self.ui_img_layout.removeWidget(self.img_label)
            self.img_label.deleteLater()
        
        self.img_label = PsdInteractiveMaskLabel()
        self.img_label.setMinimumSize(300, 200)
        self.ui_img_layout.addWidget(self.img_label, 1)
        self.img_label.boxes_changed.connect(self._save_boxes)
        
        self.ocr_edit.setReadOnly(True)
        self._rich_preview_enabled = False
        if hasattr(self, 'ocr_preview'): self.ocr_preview.setVisible(False)
        self.ocr_edit.setVisible(True)
        
        # --- 字体设置区域 (升级为单基准字体搜索库) ---
        font_widget = QWidget()
        font_layout = QHBoxLayout(font_widget)
        font_layout.setContentsMargins(8, 0, 0, 0)
        font_layout.setSpacing(6)
        
        self.lbl_font_tip = QLabel("Base Font:") 
        self.lbl_font_tip.setStyleSheet("color: white; font-weight: bold; font-size: 11px;")
        font_layout.addWidget(self.lbl_font_tip)
        
        # 使用 QFontComboBox 自动加载系统所有字体，支持搜索过滤
        self.f_base_font = QFontComboBox()
        self.f_base_font.setCurrentText("Arial")
        self.f_base_font.setMaximumWidth(200)
        self.f_base_font.setMinimumWidth(120) 
        self.f_base_font.setEditable(True) # 允许手动微调
        # --- 字体下拉框美化版 (高对比度下拉按钮) ---
        self.f_base_font.setStyleSheet('''
            QComboBox {
                background-color: #0f1117;
                border: 1px solid #4f8ef7; /* 稍微加强边框色 */
                color: #ffffff;
                border-radius: 5px;
                padding: 2px 30px 2px 10px; /* 右侧留出足够空间给箭头 */
                font-size: 13px;
                min-height: 24px;
            }
            /* 自定义下拉按钮区域 */
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 25px;
                border-left: 1px solid #4f8ef7; /* 分割线 */
                border-top-right-radius: 5px;
                border-bottom-right-radius: 5px;
                background-color: #3b4150; /* 按钮区域背景稍亮 */
            }
            /* 绘制高对比度白色三角箭头 */
            QComboBox::down-arrow {
                image: url(data:image/svg+xml;charset=utf-8,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20viewBox%3D%220%200%2024%2024%22%3E%3Cpath%20fill%3D%22%23ffffff%22%20d%3D%22M7%2010l5%205%205-5z%22%2F%3E%3C%2Fsvg%3E);
                width: 18px;
                height: 18px;
            }
            /* 鼠标悬浮时的反馈 */
            QComboBox:hover {
                border-color: #34d399; /* 悬浮变绿 */
            }
            QComboBox::drop-down:hover {
                background-color: #2e3250;
            }
            /* 下拉列表弹窗样式 */
            QComboBox QAbstractItemView {
                background-color: #1a1d27;
                color: #e2e8f0;
                selection-background-color: #4f8ef7;
                selection-color: #ffffff;
                border: 1px solid #4f8ef7;
                outline: none;
            }
        ''')
        font_layout.addWidget(self.f_base_font)
            
        self.ui_right_toolbar.insertWidget(6, font_widget)

        # --- 单图生成 PSD 按钮 ---
        # 赋值给 self 变量，用于语言切换
        self.btn_build_single = QPushButton(T('tab4_build_single'))
        self.btn_build_single.setFixedHeight(28)
        self.btn_build_single.setStyleSheet(self._main_btn_style)
        self.ui_right_toolbar.insertWidget(5, self.btn_build_single)
        self.btn_build_single.clicked.connect(self._on_build_single)
        
        # 在 Toolbar 最后增加一个弹簧，防止控件在最大化时强撑窗口
        self.ui_right_toolbar.addStretch(1)

    # --- 新增：实现该 Tab 的中英文切换逻辑 ---
    def _retranslate_ui(self):
        # 1. 先调用父类翻译通用按钮
        super()._retranslate_ui()
        
        # 2. 刷新 Tab 4 独有的控件文字（必须加 hasattr 保护，防止初始化未完成时调用）
        if hasattr(self, 'lbl_font_tip'):
            self.lbl_font_tip.setText(T('tab4_ps_fonts'))
        
        if hasattr(self, 'btn_build_single'):
            self.btn_build_single.setText(T('tab4_build_single'))
        
        # 3. 补齐被 super() 覆盖的文字
        self.btn_start.setText(T('img2psd_start'))
        self.ocr_hdr_lbl.setText("JSON ATTRIBUTES / PROGRESS")

    def _update_summary(self, final=False):
        total = len(self.file_items)
        if total == 0:
            self.tab_summary.setPlainText("=== PSD BUILDER SUMMARY ===\n\nNo files loaded.")
            return

        box_total = getattr(self, 'job_stats', {}).get('box_total', 0)
        box_done = getattr(self, 'job_stats', {}).get('box_done', 0)
        box_err = getattr(self, 'job_stats', {}).get('box_err', 0)
        box_empty = getattr(self, 'job_stats', {}).get('box_empty', 0)

        box_s = f"[1] AI BOUNDING BOX\n    Detected: {box_done}/{box_total} | Errors: {box_err}\n"
        if box_empty > 0: box_s += f"    ⚠️ EMPTY (0 Boxes): {box_empty} files!\n"

        targets = [i for i in self.file_items if len(i.boxes) > 0]
        psd_total = len(targets)
        psd_success = len([i for i in targets if i.status == 'done'])
        psd_failed = len([i for i in targets if i.status == 'error'])
        psd_proc = len([i for i in targets if i.status == 'processing'])
        psd_pend = len([i for i in targets if i.status == 'pending'])

        psd_s = f"\n[2] PSD GENERATION\n    Target Files: {psd_total}\n    Success: {psd_success} | Failed: {psd_failed} | Processing: {psd_proc} | Pending: {psd_pend}\n\n"
        s = "=== PSD BUILDER SUMMARY ===\n\n" + box_s + psd_s

        err_items = [i for i in self.file_items if i.status == 'error' or (hasattr(i, 'box_detected') and i.box_detected and len(i.boxes)==0)]
        if err_items:
            s += "🚨 [ATTENTION REQUIRED]:\n"
            for i in err_items:
                if i.status == 'error': s += f"  - {i.name}: Generation Failed\n"
                elif hasattr(i, 'box_detected') and i.box_detected and len(i.boxes)==0: s += f"  - {i.name}: 0 Boxes Detected (Empty)\n"

        self.tab_summary.setPlainText(s)

    def _on_auto_box(self):
        if self.selected_idx is None: return
        item = self.file_items[self.selected_idx]; key = self._get_gemini_key()
        if not key: 
            QMessageBox.warning(self, "API Key Missing", "请在上方配置框中填入 API Key。")
            return
        self.btn_autobox.setEnabled(False); self._tab_log(f'[PSD] 🤖 AI 正在精确定位 {item.name} 的语义块与对齐方式...')
        if not hasattr(self, 'job_stats'): self.job_stats = {}
        self.job_stats['box_total'] = self.job_stats.get('box_total', 0) + 1
        self._update_summary()
        signals = DetectSignals(); signals.boxes_done.connect(self._on_ai_box_res); signals.boxes_error.connect(self._on_ai_box_err)
        worker = PsdAIBoxDetectWorker(item.path, item.rotation_angle, key, self.proxy_widget.get_proxies(), signals)
        worker.setAutoDelete(True); self._thread_pool.start(worker)

    def _on_clear_box(self):
        """重写：专属 Tab 4 的清空遮罩逻辑"""
        if self.selected_idx is None: return
        item = self.file_items[self.selected_idx]
        
        # 清空 Tab 4 专属的矩形框列表
        item.boxes.clear()
        item.box_detected = False
        
        # 刷新画布
        self.img_label.update()
        
        # 触发状态改变（可选，保持数据流完整）
        if hasattr(self.img_label, 'boxes_changed'):
            self.img_label.boxes_changed.emit()


    def _on_batch_auto_mask(self):
        pending = [i for i in self.file_items if i.status == 'pending']
        if not pending: 
            QMessageBox.information(self, "提示", "没有待处理的文件。")
            return
        key = self._get_gemini_key()
        if not key: 
            QMessageBox.warning(self, "API Key Missing", "请在上方配置框中填入 API Key。")
            return
        self.btn_batch_mask.setEnabled(False); self._tab_log(f'[PSD] 🤖 批量定位与对齐推断启动 ({len(pending)} 张)...')
        if not hasattr(self, 'job_stats'): self.job_stats = {}
        self.job_stats['box_total'] = len(pending); self.job_stats['box_done'] = 0; self.job_stats['box_empty'] = 0; self.job_stats['box_err'] = 0
        self._update_summary(); proxies = self.proxy_widget.get_proxies()
        for item in pending:
            signals = DetectSignals(); signals.boxes_done.connect(self._on_ai_box_res); signals.boxes_error.connect(self._on_ai_box_err)
            worker = PsdAIBoxDetectWorker(item.path, item.rotation_angle, key, proxies, signals)
            worker.setAutoDelete(True); self._thread_pool.start(worker)

    def _on_ai_box_res(self, path, boxes):
        for item in self.file_items:
            if item.path == path:
                item.boxes = boxes 
                item.box_detected = True
                if len(boxes) == 0: self.job_stats['box_empty'] = self.job_stats.get('box_empty', 0) + 1; self._tab_log(f'⚠️ [PSD] Warning: {item.name} No text boxes detected!', is_warning=True)
                else: self.job_stats['box_done'] = self.job_stats.get('box_done', 0) + 1; self._tab_log(f'[PSD] AI Boxes created: {item.name} Found {len(boxes)} blocks.')
                if self.selected_idx is not None and self.file_items[self.selected_idx].path == path: self.img_label.update()
                break
        self.btn_autobox.setEnabled(True); self._update_summary(); self._finalize_ai_worker()

    def _on_ocr_done(self, path: str, text: str):
        # --- 新增：智能拦截警告，绝不让文件提前判定为 failed ---
        if text.startswith("[Warning]"):
            self._tab_log(f"⚠️ {Path(path).name}: {text}", is_warning=True)
            for item in self.file_items:
                if item.path == path: 
                    item.bg_warning = True  # 打上缺失底图的标记，但维持 processing 状态
            return

        if text.startswith("Step "):
            return

        if text != "_SUCCESS_":
            return

        # 走到这里说明整个流程（无论带不带伤）都跑完了
        for item in self.file_items:
            if item.path == path: item.status = 'done'; break
            
        self._batch_done = getattr(self, '_batch_done', 0) + 1
        
        filename = Path(path).name
        if getattr(self, '_is_single_mode', False):
            if any(getattr(i, 'bg_warning', False) for i in self.file_items if i.path == path):
                self._tab_log(f"⚠️ [Single PSD] {filename} 生成完毕，但底图抹除失败！", is_warning=True)
            else:
                self._tab_log(f"✓ [Single PSD] {filename} 生成成功！")
        else:
            targets = [i for i in self.file_items if len(i.boxes) > 0]
            success = len([i for i in targets if i.status == 'done'])
            failed = len([i for i in targets if i.status == 'error'])
            pending = len([i for i in targets if i.status in ('pending', 'processing')])
            self._tab_log(f"✓ {filename} [Success: {success} | Failed: {failed} | Pending: {pending}]")
            
        self._refresh_list()
        self._check_batch_complete()
        self._update_summary()

    def _on_ocr_error(self, path: str, msg: str):
        for item in self.file_items:
            if item.path == path: item.status, item.error_msg = 'error', msg; break
            
        self._batch_errors = getattr(self, '_batch_errors', 0) + 1
        
        filename = Path(path).name
        if getattr(self, '_is_single_mode', False):
            self._tab_log(f"✗ [Single PSD] {filename} 生成失败: {msg}", is_error=True)
        else:
            targets = [i for i in self.file_items if len(i.boxes) > 0]
            success = len([i for i in targets if i.status == 'done'])
            failed = len([i for i in targets if i.status == 'error'])
            pending = len([i for i in targets if i.status in ('pending', 'processing')])
            self._tab_log(f"✗ {filename} [Success: {success} | Failed: {failed} | Pending: {pending}] (Err: {msg})", is_error=True)
            
        self._refresh_list(); self._check_batch_complete(); self._update_summary()

    def _check_batch_complete(self):
        if getattr(self, '_is_single_mode', False):
            return
            
        # 如果还有在处理中的文件，就绝对不触发“批量完成”逻辑
        if [i for i in self.file_items if i.status == 'processing']:
            return
            
        targets = [i for i in self.file_items if len(i.boxes) > 0]
        done_items = [i for i in targets if i.status == 'done']
        error_items = [i for i in targets if i.status == 'error']
        pending_items = [i for i in targets if i.status == 'pending']
        
        d = len(done_items)
        e = len(error_items)
        p = len(pending_items)
        
        base_msg = T('log_batch_done', d=d, e=e)
        if p > 0:
            base_msg += f', ⏸ {p} pending'
            
        # --- 核心：寻找有底图警告的成功文件，追加你要的专属红色提示 ---
        warning_names = [i.name for i in done_items if getattr(i, 'bg_warning', False)]
        if warning_names:
            warn_str = f"；⚠️ {', '.join(warning_names)} Although the PSD was created, Gemini failed to generate a background layout with the text removed"
            # 设为 is_warning 会自动标成显眼的颜色
            self._tab_log(base_msg + warn_str, is_warning=True)
        else:
            self._tab_log(base_msg)
            
        self._update_summary(final=True)


    def _get_custom_fonts(self):
        return {
            'base_font': self.f_base_font.currentText().strip() or "Arial"
        }

    def _on_build_single(self):
        if self.selected_idx is None:
            QMessageBox.information(self, '提示', '请先在左侧列表中选择一张图片。')
            return
            
        item = self.file_items[self.selected_idx]
        if len(item.boxes) == 0:
            QMessageBox.information(self, '提示', '当前图片还没有绘制任何文本遮罩框。')
            return
            
        api, key, model, proxies = self._get_config()
        if not key:
            QMessageBox.warning(self, 'API Key 缺失', f'请确保您在上方填入了 {api.upper()} 的 API Key。')
            return
            
        self._stop_flag = False
        self._is_single_mode = True
        
        total = 1
        done_counter = [0]
        
        signals = WorkerSignals()
        signals.finished.connect(self._on_ocr_done)
        signals.error.connect(self._on_ocr_error)
        signals.progress.connect(lambda d, t: self._update_progress(d, t))
        
        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(0)
        self.progress_label.setText(f'Processing: {item.name}')
        self._tab_log(f'[PSD Single] 开始生成单图: {item.name}')
        
        custom_fonts = self._get_custom_fonts()
        item.status = 'processing' 
        item.bg_warning = False  # <--- 新增：每次开始前重置底图警告标记
        
        worker = ImageToPsdWorker(api, item.path, item.boxes, item.rotation_angle, key, model, proxies, custom_fonts, signals, done_counter, total)
        worker.setAutoDelete(True)
        self._thread_pool.start(worker)
        
        self._refresh_list()
        self._update_summary()

    def _on_start(self):
        targets = [i for i in self.file_items if len(i.boxes) > 0]
        if not targets: QMessageBox.information(self, 'Info', '没有画好遮罩的图片可供生成。'); return
            
        api, key, model, proxies = self._get_config()
        if not key: QMessageBox.warning(self, 'Missing Key', f'请确保您在上方填入了 {api.upper()} 的 API Key。'); return
            
        self._stop_flag = False; self._batch_done = 0; self._batch_errors = 0; total = len(targets); done_counter = [0]
        self._is_single_mode = False
        
        signals = WorkerSignals()
        signals.finished.connect(self._on_ocr_done)
        signals.error.connect(self._on_ocr_error)
        signals.progress.connect(lambda d, t: self._update_progress(d, t))
        
        self.progress_bar.setVisible(True); self.progress_bar.setMaximum(total); self.progress_bar.setValue(0)
        self.progress_label.setText(f'0 / {total}')
        self._tab_log(f'[V5.1 Engine] Starting {total} files')
        
        custom_fonts = self._get_custom_fonts()
        
        for item in targets:
            if self._stop_flag: break
            item.status = 'processing' 
            item.bg_warning = False  # <--- 新增：每次开始前重置底图警告标记
            worker = ImageToPsdWorker(api, item.path, item.boxes, item.rotation_angle, key, model, proxies, custom_fonts, signals, done_counter, total)
            worker.setAutoDelete(True); self._thread_pool.start(worker)
            
        self._refresh_list(); self._update_summary()

# ==========================================
    # 👇 2. 追加 Tab 4 专属的物理旋转重写方法
    # ==========================================
    def _physical_rotate(self, item, angle):
        """Tab4 专属：物理旋转并覆盖硬盘原文件"""
        try:
            with PILImage.open(item.path) as img:
                img = ImageOps.exif_transpose(img)
                # PIL 的 ROTATE_90 是逆时针(向左)，ROTATE_270 是顺时针(向右)
                if angle == -90:
                    rotated = img.transpose(PILImage.Transpose.ROTATE_90)
                elif angle == 90:
                    rotated = img.transpose(PILImage.Transpose.ROTATE_270)
                elif angle in (180, -180):
                    rotated = img.transpose(PILImage.Transpose.ROTATE_180)
                else:
                    return True
                
                save_kw = {}
                fmt = img.format if img.format else 'PNG'
                if fmt in ('JPEG', 'JPG'): save_kw['quality'] = 100
                rotated.save(item.path, format=fmt, **save_kw)
            
            # 核心保护：物理覆盖后，强制清空该图片的所有坐标和缓存
            item.rotation_angle = 0
            item.pixmap = None
            item.boxes.clear()
            item.polygons.clear()
            item.box_detected = False
            return True
        except Exception as e:
            self._tab_log(f"✗ 物理旋转覆盖失败: {item.name} - {e}", is_error=True)
            return False

    def _on_rot_left(self):
        """重写：左转按钮点击事件"""
        if self.selected_idx is None: return
        item = self.file_items[self.selected_idx]
        if self._physical_rotate(item, -90):
            self.img_label.set_item(item)
            self._refresh_list()
            self._tab_log(f"↺ {item.name} 向左旋转 90° (已物理覆盖原图)")

    def _on_rot_right(self):
        """重写：右转按钮点击事件"""
        if self.selected_idx is None: return
        item = self.file_items[self.selected_idx]
        if self._physical_rotate(item, 90):
            self.img_label.set_item(item)
            self._refresh_list()
            self._tab_log(f"↻ {item.name} 向右旋转 90° (已物理覆盖原图)")

    def _on_ai_rot_res(self, path, angle):
        """重写：AI 自动批量识别旋转角度的接收事件"""
        self.job_stats['rot_done'] += 1
        for item in self.file_items:
            if item.path == path:
                if angle != 0: 
                    if self._physical_rotate(item, angle):
                        self._tab_log(f'[智能纠正] {item.name} 自动旋转 {angle}° (已物理覆盖原图)')
                break
        self._update_summary()
        self._finalize_ai_worker()

# ─────────────────────────────────────────────────────────────
# TAB 5: ADOBE DTP (PS/AI) 分治架构 - 独立导出 + 假翻 PDF
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
            elif p.lower().endswith(self.valid_exts): paths.append(p)
        if paths: self.files_dropped.emit(paths)
        event.acceptProposedAction()

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            for item in self.selectedItems():
                path = item.data(Qt.ItemDataRole.UserRole)
                self.file_deleted.emit(path)
                self.takeItem(self.row(item))
        else: super().keyPressEvent(event)


class DtpWorkerSignals(QObject):
    log = pyqtSignal(str, bool)       
    progress = pyqtSignal(int, int)   
    finished = pyqtSignal(str, str)   
    error = pyqtSignal(str)


class AdobeDtpWorker(QRunnable):
    def __init__(self, mode, files, out_dir="", is_pseudo=False, trans_paths=None, signals=None):
        super().__init__()
        self.mode = mode
        self.files = files
        self.out_dir = out_dir  
        self.is_pseudo = is_pseudo
        self.trans_paths = trans_paths
        self.signals = signals
        
        self.stats = {'total': len(files), 'ps_processed': 0, 'ai_processed': 0, 'fail': 0}
        self.hidden_layers = []
        self.locked_layers = []
        self.smart_objects = []
        self.embedded_objects = []
        self.linked_objects = []
        self.failed_layers = []
        self.skipped_layers = []

    def _apply_pseudo(self, text):
        """本地化假翻：保留结构，替换字符"""
        if not text: return ""
        res = ""
        for char in text:
            if char.isupper(): res += "X"
            elif char.islower(): res += "x"
            elif char.isdigit(): res += "0"
            else: res += char # 保留标点、空格、换行
        return res

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
                for idx, file_path in enumerate(self.files, 1):
                    p_obj = Path(file_path)
                    ext = p_obj.suffix.lower()
                    fname = p_obj.name
                    fstem = p_obj.stem  # 无后缀的文件名
                    
                    if self.signals:
                        self.signals.log.emit(f"[{idx}/{total_files}] Extracting {fname}...", False)
                        self.signals.progress.emit(idx, total_files)

                    try:
                        # ==========================================
                        # Photoshop (.psd, .psb) 导出逻辑
                        # ==========================================
                        if ext in ['.psd', '.psb']:
                            if not ps_app: 
                                ps_app = win32com.client.Dispatch("Photoshop.Application")
                                ps_app.DisplayDialogs = 3
                            
                            doc = ps_app.Open(file_path)
                            ps_extracted_items = []
                            text_counter = [0]
                            
                            def traverse_ps_export(layers):
                                for i in range(layers.Count):
                                    layer = layers.Item(i + 1)
                                    if not layer.Visible: self.hidden_layers.append(f"{fname} -> {layer.Name}")
                                    if layer.AllLocked:   self.locked_layers.append(f"{fname} -> {layer.Name}")
                                    
                                    if layer.LayerType == 1: 
                                        try:
                                            if layer.Kind == 17: 
                                                self.smart_objects.append(f"{fname} -> {layer.Name}")
                                        except: pass

                                    if layer.LayerType == 2:
                                        traverse_ps_export(layer.Layers)
                                    elif layer.LayerType == 1 and layer.Kind == 2:
                                        text_counter[0] += 1
                                        tid = f"TID_{text_counter[0]}"
                                        src = layer.TextItem.Contents.strip()
                                        if src: 
                                            ps_extracted_items.append({'file': fname, 'layer': tid, 'source': src})
                                        
                                        # 原地假翻替换
                                        if self.is_pseudo:
                                            try: layer.TextItem.Contents = self._apply_pseudo(src)
                                            except: pass
                            
                            traverse_ps_export(doc.Layers)
                            
                            if ps_extracted_items:
                                ps_out = os.path.join(self.out_dir, f"{fstem}.docx")
                                self._build_dtp_word(ps_extracted_items, ps_out)
                                if self.signals: self.signals.log.emit(f"  ✓ Saved Word: {Path(ps_out).name}", False)
                                
                            if self.is_pseudo:
                                try:
                                    pdf_path = os.path.abspath(os.path.join(self.out_dir, f"{fstem}.pdf")).replace('\\', '/')
                                    import tempfile
                                    temp_dir = tempfile.gettempdir()
                                    jsx_ps_path = os.path.abspath(os.path.join(temp_dir, "ps_pdf_export.jsx"))
                                    
                                    # 【核心修复1】通过 JSX 执行 PS 的 PDF 导出，彻底解决服务器运行失败的报错
                                    jsx_ps = f"""
                                    try {{
                                        var doc = app.activeDocument;
                                        var pdfFile = new File("{pdf_path}");
                                        var pdfOpts = new PDFSaveOptions();
                                        pdfOpts.layers = false;
                                        pdfOpts.embedFonts = true;
                                        pdfOpts.encoding = PDFEncoding.JPEG;
                                        pdfOpts.jpegQuality = 5;
                                        doc.saveAs(pdfFile, pdfOpts, true, Extension.LOWERCASE);
                                    }} catch(e) {{
                                        throw new Error(e.message);
                                    }}
                                    """
                                    with open(jsx_ps_path, 'w', encoding='utf-8-sig') as f:
                                        f.write(jsx_ps)
                                        
                                    ps_app.DoJavaScriptFile(jsx_ps_path)
                                    if self.signals: self.signals.log.emit(f"  ✓ Saved Pseudo PDF: {Path(pdf_path).name}", False)
                                    
                                    try: os.remove(jsx_ps_path)
                                    except: pass
                                except Exception as pdf_err:
                                    if self.signals: self.signals.log.emit(f"  -> PDF Save failed: {pdf_err}", True)

                            doc.Close(2) # 不保存修改，关闭文档
                            self.stats['ps_processed'] += 1

                        # ==========================================
                        # Illustrator (.ai, .eps, .ps) 导出逻辑
                        # ==========================================
                        elif ext in ['.ai', '.eps', '.ps']:
                            if not ai_app: 
                                ai_app = win32com.client.Dispatch("Illustrator.Application")
                                ai_app.UserInteractionLevel = -1 
                            
                            doc = ai_app.Open(file_path)
                            ai_extracted_items = []

                            def traverse_ai_layers(layers):
                                for i in range(layers.Count):
                                    layer = layers.Item(i + 1)
                                    try:
                                        if not layer.Visible: self.hidden_layers.append(f"{fname} -> [Layer] {layer.Name}")
                                        if layer.Locked: self.locked_layers.append(f"{fname} -> [Layer] {layer.Name}")
                                        if layer.Layers.Count > 0: traverse_ai_layers(layer.Layers)
                                    except: pass
                            try: traverse_ai_layers(doc.Layers)
                            except: pass

                            import tempfile
                            temp_dir = tempfile.gettempdir()
                            jsx_path = os.path.abspath(os.path.join(temp_dir, "ai_export_fast.jsx"))
                            out_txt_path = os.path.abspath(os.path.join(temp_dir, "ai_export_result.txt"))
                            out_txt_path_js = out_txt_path.replace('\\', '/')
                            
                            if self.signals: self.signals.log.emit(f"  -> Injecting 6D Rich Style extraction...", False)

                            jsx_template = r"""
                            var destFile = new File("__OUT_TXT_PATH__");
                            destFile.open("w");
                            destFile.encoding = "UTF-8";

                            var doc = app.activeDocument;
                            var frames = doc.textFrames;

                            for (var i = 0; i < frames.length; i++) {
                                var tf = frames[i];
                                var tid = "TID_" + (i + 1);
                                var tfName = tf.name || ("Unnamed TextFrame " + tid);

                                if (tf.hidden || (tf.layer && !tf.layer.visible)) {
                                    destFile.writeln("SKIP_HIDDEN|||" + tid + "|||" + tfName);
                                    continue;
                                }
                                if (tf.locked || (tf.layer && tf.layer.locked)) {
                                    destFile.writeln("SKIP_LOCKED|||" + tid + "|||" + tfName);
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

                                        txt = txt.replace(/\r/g, "[BR]").replace(/\n/g, "[BR]");
                                        txt = txt.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

                                        var font = "", size = "", colorStr = "", base = "", caps = "", just = "", under = "";
                                        
                                        try { font = tr.characterAttributes.textFont.name; } catch(e) {}
                                        try { size = Math.round(tr.characterAttributes.size * 100) / 100; } catch(e) {}
                                        
                                        try {
                                            var fColor = tr.characterAttributes.fillColor;
                                            var rVal = 0, gVal = 0, bVal = 0;
                                            var hasColor = false;
                                            
                                            var actualColor = fColor;
                                            if (fColor.typename === 'SpotColor') { actualColor = fColor.spot.color; }
                                            
                                            if (actualColor.typename === 'RGBColor') {
                                                rVal = Math.round(actualColor.red); 
                                                gVal = Math.round(actualColor.green); 
                                                bVal = Math.round(actualColor.blue);
                                                hasColor = true;
                                            } else if (actualColor.typename === 'CMYKColor') {
                                                var c = actualColor.cyan; var m = actualColor.magenta; 
                                                var y = actualColor.yellow; var k = actualColor.black;
                                                rVal = Math.round(255 * (1 - c/100) * (1 - k/100));
                                                gVal = Math.round(255 * (1 - m/100) * (1 - k/100));
                                                bVal = Math.round(255 * (1 - y/100) * (1 - k/100));
                                                hasColor = true;
                                            } else if (actualColor.typename === 'GrayColor') {
                                                var grayRGB = Math.round(255 - (actualColor.gray * 2.55));
                                                rVal = grayRGB; gVal = grayRGB; bVal = grayRGB;
                                                hasColor = true;
                                            }
                                            
                                            if (hasColor) {
                                                var hexR = rVal.toString(16); if(hexR.length==1) hexR="0"+hexR;
                                                var hexG = gVal.toString(16); if(hexG.length==1) hexG="0"+hexG;
                                                var hexB = bVal.toString(16); if(hexB.length==1) hexB="0"+hexB;
                                                colorStr = "#" + hexR.toUpperCase() + hexG.toUpperCase() + hexB.toUpperCase();
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
                                        
                                        if (font == current_font && size == current_size && colorStr == current_color && 
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
                                            current_font = font; current_size = size; current_color = colorStr;
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
                                    
                                    // 【核心修复2】AI 假翻完美保留样式：按碎片 (textRanges) 遍历，不动格式只换字
                                    if (__PSEUDO__) {
                                        try {
                                            for (var r = 0; r < tf.textRanges.length; r++) {
                                                var tr = tf.textRanges[r];
                                                tr.contents = tr.contents.replace(/[A-Z]/g, 'X').replace(/[a-z]/g, 'x').replace(/[0-9]/g, '0');
                                            }
                                        } catch(e) {}
                                    }

                                } catch(e) {
                                    destFile.writeln("ERR|||" + tid + "|||" + e.message);
                                }
                            }
                            destFile.close();
                            """
                            
                            is_p_js = "true" if self.is_pseudo else "false"
                            jsx_code = jsx_template.replace("__OUT_TXT_PATH__", out_txt_path_js).replace("__PSEUDO__", is_p_js)
                            
                            with open(jsx_path, 'w', encoding='utf-8-sig') as f:
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
                                            status, tid, content = parts[0], parts[1], parts[2]
                                            if status == "SKIP_HIDDEN": self.hidden_layers.append(f"{fname} -> {content}")
                                            elif status == "SKIP_LOCKED": self.locked_layers.append(f"{fname} -> {content}")
                                            elif status == "OK":
                                                final_xml = content.replace("[BR]", "\n")
                                                ai_extracted_items.append({'file': fname, 'layer': tid, 'source': final_xml})
                                                
                            try:
                                if os.path.exists(jsx_path): os.remove(jsx_path)
                                if os.path.exists(out_txt_path): os.remove(out_txt_path)
                            except: pass

                            if ai_extracted_items:
                                ai_out = os.path.join(self.out_dir, f"{fstem}.xlf")
                                self._build_dtp_xml(ai_extracted_items, ai_out)
                                if self.signals: self.signals.log.emit(f"  ✓ Saved XLF: {Path(ai_out).name}", False)

                            if self.is_pseudo:
                                try:
                                    pdf_path = os.path.abspath(os.path.join(self.out_dir, f"{fstem}.pdf"))
                                    pdf_opts = win32com.client.Dispatch("Illustrator.PDFSaveOptions")
                                    pdf_opts.PreserveEditability = False
                                    pdf_opts.AcrobatLayers = False
                                    doc.SaveAs(pdf_path, pdf_opts)
                                    if self.signals: self.signals.log.emit(f"  ✓ Saved Pseudo PDF: {Path(pdf_path).name}", False)
                                except Exception as e:
                                    if self.signals: self.signals.log.emit(f"  -> AI PDF Error: {e}", True)

                            doc.Close(2)
                            self.stats['ai_processed'] += 1

                    except Exception as e:
                        self.stats['fail'] += 1
                        if self.signals: self.signals.log.emit(f"Failed {fname}: {e}", True)

                summary = self._generate_summary("EXPORT")
                if self.signals: self.signals.finished.emit("DONE", summary)

            elif self.mode == 'import':
                master_translations = {}
                if self.trans_paths:
                    for trans_file in self.trans_paths:
                        import_ext = Path(trans_file).suffix.lower()
                        if import_ext == '.docx':
                            tmp_dict = self._read_dtp_word(trans_file)
                        elif import_ext in ['.xml', '.xlf', '.xliff']:
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
                            
                            import tempfile
                            temp_dir = tempfile.gettempdir()
                            data_path = os.path.abspath(os.path.join(temp_dir, "ai_import_data.txt"))
                            jsx_path = os.path.abspath(os.path.join(temp_dir, "ai_import_fast.jsx"))
                            
                            data_path_js = data_path.replace('\\', '/')
                            
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
                                            text = m.group(2).replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>').replace('\n', '[BR]').replace('\r', '[BR]')
                                            
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
                                            
                            jsx_template = r"""
                            var dataFile = new File("__DATA_PATH__");
                            dataFile.open("r");
                            dataFile.encoding = "UTF-8";
                            
                            var dict = {};
                            while (!dataFile.eof) {
                                var line = dataFile.readln();
                                if (!line) continue;
                                var parts = line.split("|||");
                                if (parts.length >= 9) {
                                    var tid = parts[0];
                                    var text = parts.slice(8).join("|||").replace(/\[BR\]/g, "\r");
                                    
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
                            
                            jsx_code = jsx_template.replace("__DATA_PATH__", data_path_js)
                            
                            with open(jsx_path, 'w', encoding='utf-8-sig') as f:
                                f.write(jsx_code)
                                
                            import time as pytime
                            start_time = pytime.time()
                            ai_app.DoJavaScriptFile(jsx_path)
                            elapsed = pytime.time() - start_time
                            
                            if self.signals: self.signals.log.emit(f"  -> Injection finished in {elapsed:.2f} seconds!", False)
                            
                            try:
                                if os.path.exists(jsx_path): os.remove(jsx_path)
                                if os.path.exists(data_path): os.remove(data_path)
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

        warnings_by_file = {}
        def add_to_dict(category, items_list):
            for item in items_list:
                if " -> " in item:
                    fname, lname = item.split(" -> ", 1)
                    if fname not in warnings_by_file: warnings_by_file[fname] = {}
                    if category not in warnings_by_file[fname]: warnings_by_file[fname][category] = []
                    warnings_by_file[fname][category].append(lname)

        add_to_dict("Hidden Layers", self.hidden_layers)
        add_to_dict("Locked Layers", self.locked_layers)
        add_to_dict("PSD Smart Objects", self.smart_objects)
        add_to_dict("AI Embedded Objects", self.embedded_objects)
        add_to_dict("AI Linked Objects", self.linked_objects)

        if warnings_by_file:
            s += "<br><b style='color:#ef4444;'>[WARNING] ADOBE OBJECTS ALERTS:</b><br>"
            for fname, cats in warnings_by_file.items():
                s += f"  <span style='color:#ef4444;'>{fname}</span><br>"
                for cat_name, layers in cats.items():
                    s += f"    <b style='color:#ffffff;'>{cat_name}:</b><br>"
                    unique_layers = sorted(list(set(layers)))
                    for lyr in unique_layers:
                        s += f"      <span style='color:#d4d4d4;'>- {lyr}</span><br>"
                        
        if self.skipped_layers:
            s += f"<br><b>[INFO] Skipped {len(self.skipped_layers)} empty translations.</b><br>"
            
        if self.failed_layers:
            s += f"<br><b style='color:#ef4444;'>[ERROR] Failed to inject specific layers:</b><br>"
            failed_dict = {}
            for item in self.failed_layers:
                if " -> " in item:
                    f, l = item.split(" -> ", 1)
                    if f not in failed_dict: failed_dict[f] = []
                    failed_dict[f].append(l)
            for fname, layers in failed_dict.items():
                s += f"  <span style='color:#ef4444;'>{fname}</span><br>"
                for lyr in set(layers):
                    s += f"    <span style='color:#d4d4d4;'>- {lyr}</span><br>"

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
        import re
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
                        tid_match = re.search(r'(TID_\d+)', parts[1])
                        if tid_match:
                            if fname not in translations: translations[fname] = {}
                            translations[fname][tid_match.group(1)] = target_text
        return translations

    def _build_dtp_xml(self, items, output_path):
        import re
        xml_out = []
        xml_out.append('<?xml version="1.0" encoding="UTF-8"?>')
        xml_out.append('<xliff version="1.2" xmlns="urn:oasis:names:tc:xliff:document:1.2" xmlns:okp="okapi-framework:xliff-extensions" xmlns:its="http://www.w3.org/2005/11/its" xmlns:itsxlf="http://www.w3.org/ns/its-xliff/" its:version="2.0">')

        files_dict = {}
        for item in items:
            fname = item['file']
            if fname not in files_dict: files_dict[fname] = []
            files_dict[fname].append(item)

        for fname, contents in files_dict.items():
            xml_out.append(f'<file original="{fname}" source-language="en-US" datatype="xml">')
            xml_out.append('<body>')

            for idx, item in enumerate(contents, 1):
                tid = item['layer']
                source_str = item['source']

                xml_out.append(f'<trans-unit id="{idx}" resname="{tid}" xml:space="preserve">')

                # 解析 <cf> 标签并替换为带有良好缩进的 bpt / ept 格式
                pattern = re.compile(r'(<cf[^>]*>)(.*?)(</cf>)', re.IGNORECASE | re.DOTALL)
                bpt_id = 1
                last_end = 0
                formatted_source = "\n"

                for match in pattern.finditer(source_str):
                    prefix = source_str[last_end:match.start()]
                    if prefix:
                        formatted_source += prefix.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

                    open_tag = match.group(1).replace('<', '&lt;').replace('>', '&gt;')
                    inner_text = match.group(2).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                    close_tag = match.group(3).replace('<', '&lt;').replace('>', '&gt;')

                    formatted_source += f'            <bpt id="{bpt_id}">{open_tag}</bpt>{inner_text}<ept id="{bpt_id}">{close_tag}</ept>\n'
                    bpt_id += 1
                    last_end = match.end()

                suffix = source_str[last_end:]
                if suffix:
                    formatted_source += suffix.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;') + "\n"

                formatted_source += "        "

                xml_out.append(f'<source xml:lang="en-US">{formatted_source}</source>')
                xml_out.append('</trans-unit>')

            xml_out.append('</body>')
            xml_out.append('</file>')

        xml_out.append('</xliff>')

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(xml_out))

    
    def _read_dtp_xml(self, input_path):
        import xml.etree.ElementTree as ET
        import html

        translations = {}
        try:
            tree = ET.parse(input_path)
            root = tree.getroot()
            ns = {'xliff': 'urn:oasis:names:tc:xliff:document:1.2'}

            def local_name(tag):
                return tag.split('}', 1)[-1] if '}' in tag else tag

            for file_node in root.findall('.//xliff:file', ns) or root.findall('.//file'):
                fname = file_node.get('original')
                if not fname:
                    continue

                translations.setdefault(fname, {})

                for tu_node in file_node.findall('.//xliff:trans-unit', ns) or file_node.findall('.//trans-unit'):
                    tid = tu_node.get('resname')
                    if not tid:
                        continue

                    target_node = tu_node.find('xliff:target', ns)
                    if target_node is None:
                        target_node = tu_node.find('target')
                    if target_node is None:
                        continue

                    parts = []

                    # target 起始文本
                    if target_node.text:
                        parts.append(target_node.text)

                    # 逐个 child 处理，忽略 bpt/ept 外壳，只保留里面的转义 cf 文本
                    for child in list(target_node):
                        lname = local_name(child.tag)

                        if lname in ('bpt', 'ept'):
                            if child.text:
                                parts.append(html.unescape(child.text))
                        else:
                            # 其他标签如果以后有，原样保留
                            parts.append(ET.tostring(child, encoding='unicode', method='xml'))

                        if child.tail:
                            parts.append(child.tail)

                    clean_xml = ''.join(parts).strip()
                    if clean_xml:
                        translations[fname][tid] = clean_xml

        except Exception as e:
            if self.signals:
                self.signals.log.emit(f"Failed to parse XML: {e}", True)

        return translations


class AdobeDtpTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_items = []
        self._thread_pool = QThreadPool()
        self.valid_exts = ('.psd', '.psb', '.ai', '.eps', '.ps')
        self._build_ui()

    def _build_ui(self):
        main_layout = QVBoxLayout(self); main_layout.setContentsMargins(0, 0, 0, 0)
        toolbar = QFrame(); toolbar.setFixedHeight(55); toolbar.setStyleSheet('background: #1a1d27; border-bottom: 1px solid #2e3250;')
        tb_layout = QHBoxLayout(toolbar); tb_layout.setContentsMargins(12, 0, 12, 0)
        
        self.btn_add_files = QPushButton("📄 Add Files"); self.btn_add_files.setFixedHeight(30)
        self.btn_add_folder = QPushButton("📁 Add Folder"); self.btn_add_folder.setFixedHeight(30)
        
        # 假翻复选框（已修改为白色）
        self.cb_pseudo = QCheckBox("假翻 (Pseudo)")
        self.cb_pseudo.setStyleSheet("color: white; font-weight: bold; font-size: 13px;")
        self.cb_pseudo.setToolTip("开启后：导出独立文件，并将PDF中的文字替换为XXXX，用于检查排版布局")
        
        self.btn_export = QPushButton(T('dtp_export')); self.btn_export.setStyleSheet('background: #4f8ef7; color: white; padding: 6px 16px; font-weight: bold;')
        self.btn_import = QPushButton(T('dtp_import')); self.btn_import.setStyleSheet('background: #34d399; color: #0a1a12; padding: 6px 16px; font-weight: bold;')
        self.btn_reset = QPushButton("↻ Reset"); self.btn_reset.setStyleSheet('background: #f59e0b; color: #1a1000; padding: 6px 16px; font-weight: bold;')
        
        # 布局调整：把弹簧(addStretch)放在复选框前面，让复选框紧贴 Export 按钮
        tb_layout.addWidget(self.btn_add_files)
        tb_layout.addWidget(self.btn_add_folder)
        tb_layout.addStretch() 
        tb_layout.addWidget(self.cb_pseudo)
        tb_layout.addSpacing(15) 
        tb_layout.addWidget(self.btn_export)
        tb_layout.addSpacing(10)
        tb_layout.addWidget(self.btn_import)
        tb_layout.addSpacing(10)
        tb_layout.addWidget(self.btn_reset)
        
        main_layout.addWidget(toolbar)

        progress_bar_container = QFrame(); progress_bar_container.setStyleSheet('background: #1a1d27; border-bottom: 1px solid #2e3250; padding: 4px 12px;')
        pb_layout = QHBoxLayout(progress_bar_container); self.progress_bar = QProgressBar(); self.progress_bar.setFixedHeight(8)
        self.progress_bar.setTextVisible(False); self.progress_label = QLabel("Ready")
        pb_layout.addWidget(self.progress_bar); pb_layout.addWidget(self.progress_label); main_layout.addWidget(progress_bar_container)

        splitter = QSplitter(Qt.Orientation.Horizontal); self.list_widget = DtpFileListWidget()
        self.list_widget.files_dropped.connect(self._add_files); self.list_widget.file_deleted.connect(self._remove_file)
        splitter.addWidget(self.list_widget); self.log_console = QPlainTextEdit(); self.log_console.setReadOnly(True)
        self.log_console.setStyleSheet("background-color: #1e1e1e; color: #d4d4d4; font-family: 'Consolas'; font-size: 12px; border: none;")
        splitter.addWidget(self.log_console); splitter.setSizes([350, 850]); main_layout.addWidget(splitter, 1)

        self.btn_add_files.clicked.connect(self._on_add_files_btn); self.btn_add_folder.clicked.connect(self._on_add_folder_btn)
        self.btn_export.clicked.connect(self._on_export); self.btn_import.clicked.connect(self._on_import); self.btn_reset.clicked.connect(self._on_reset)

    def _add_files(self, paths):
        for p in paths:
            if p not in self.file_items:
                self.file_items.append(p); item = QListWidgetItem(Path(p).name)
                item.setData(Qt.ItemDataRole.UserRole, p); self.list_widget.addItem(item)
        self._log(f"Added {len(paths)} Adobe file(s).")

    def _on_add_files_btn(self):
        paths, _ = QFileDialog.getOpenFileNames(self, 'Select Adobe DTP Files', '', 'Adobe Files (*.psd *.psb *.ai *.eps *.ps)')
        if paths: self._add_files(paths)

    def _on_add_folder_btn(self):
        folder = QFileDialog.getExistingDirectory(self, 'Select Folder')
        if folder: self._add_files([os.path.join(r, f) for r, _, fs in os.walk(folder) for f in fs if f.lower().endswith(self.valid_exts)])

    def _remove_file(self, path): 
        if path in self.file_items: self.file_items.remove(path)
    def _on_reset(self): 
        self.file_items.clear(); self.list_widget.clear(); self.log_console.clear(); self.progress_bar.setValue(0)
    def _log(self, text, is_warn=False):
        color = "#ef4444" if is_warn else "#d4d4d4"
        self.log_console.appendHtml(f"<span style='color: {color};'>[{time.strftime('%H:%M:%S')}] {text}</span>")
    def _update_progress(self, current, total):
        self.progress_bar.setMaximum(total); self.progress_bar.setValue(current); self.progress_label.setText(f"File: {current} / {total}")
    def _on_worker_finished(self, status, summary_text):
        self.btn_export.setEnabled(True); self.btn_import.setEnabled(True); self._log(summary_text)

    def _on_export(self):
        if not self.file_items: return
        self.btn_export.setEnabled(False); self.btn_import.setEnabled(False); self.progress_bar.setValue(0)
        self._log(f"Starting Independent Export (Pseudo: {self.cb_pseudo.isChecked()})...")
        signals = DtpWorkerSignals()
        signals.log.connect(self._log); signals.progress.connect(self._update_progress); signals.finished.connect(self._on_worker_finished)
        worker = AdobeDtpWorker(mode='export', files=self.file_items, out_dir=str(Path(self.file_items[0]).parent), is_pseudo=self.cb_pseudo.isChecked(), signals=signals)
        worker.setAutoDelete(True); self._thread_pool.start(worker)

    def _on_import(self):
        if not self.file_items: return
        file_paths, _ = QFileDialog.getOpenFileNames(self, 'Select Translations', str(Path(self.file_items[0]).parent), 'Translation Files (*.docx *.xml *.xlf)')
        if not file_paths: return
        self.btn_export.setEnabled(False); self.btn_import.setEnabled(False)
        signals = DtpWorkerSignals(); signals.log.connect(self._log); signals.progress.connect(self._update_progress); signals.finished.connect(self._on_worker_finished)
        worker = AdobeDtpWorker(mode='import', files=self.file_items, trans_paths=file_paths, signals=signals)
        worker.setAutoDelete(True); self._thread_pool.start(worker)

    def _retranslate_ui(self):
        self.btn_add_files.setText(T('add_files')); self.btn_add_folder.setText(T('add_folder'))
        self.btn_export.setText(T('dtp_export')); self.btn_import.setText(T('dtp_import')); self.btn_reset.setText(T('reset'))


#屏蔽了tab1-2 非蒙版模式
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
            self.tabs.setTabVisible(self.tabs.indexOf(tab), False)

        self.mask_tab = MaskApiTab('gemini', GEMINI_MODELS_ATF, self.proxy_widget)
        self.tabs.addTab(self.mask_tab, T('tab_mask'))

        self.image_to_psd_tab = ImageToPsdTab('gemini', GEMINI_MODELS, self.proxy_widget)
        self.tabs.addTab(self.image_to_psd_tab, T('tab_img2psd'))

        self.adobe_tab = AdobeDtpTab()
        self.tabs.addTab(self.adobe_tab, T('tab_adobe'))

        # 👇 核心修复：强制激活第一个可见的标签页 (索引为 2 的 Mask 模式)
        self.tabs.setCurrentIndex(2)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage(T('status_ready', t=MAX_THREADS))

# 将这个函数加在 MainWindow 类的内部，比如 _build_ui 函数的下面
    def closeEvent(self, event):
        """
        拦截窗口关闭事件，执行“三步走”彻底清退策略
        """
        # 第一步：优雅地通知所有 Tab 停止正在进行的工作
        all_tabs = list(self.api_tabs.values()) + [self.mask_tab, self.image_to_psd_tab, self.adobe_tab]
        for tab in all_tabs:
            # 触发你原本写好的停止标志
            if hasattr(tab, '_stop_flag'):
                tab._stop_flag = True
            # 清空线程池中排队但还未开始的任务
            if hasattr(tab, '_thread_pool'):
                tab._thread_pool.clear()

        # 第二步：尝试释放 Windows COM 接口占用（防止锁死 Photoshop）
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except:
            pass

        # 接受关闭事件，销毁 UI
        event.accept()

        # 第三步：终极必杀技 —— 强制结束进程
        # 很多时候 requests 请求或 win32com 的死锁会无视正常的退出指令。
        # os._exit(0) 会绕过 Python 的常规垃圾回收，直接向操作系统发送“立刻杀死我自己”的系统级指令，瞬间释放所有内存和后台死线程。
        import os
        os._exit(0)

    def _switch_lang(self):
        _LANG[0] = 'ZH' if _LANG[0] == 'EN' else 'EN'
        self.btn_lang.setText(T('lang_btn'))
        self.proxy_lbl.setText(T('proxy_lbl'))
        self.proxy_widget.checkbox.setText(T('use_proxy'))
        self.status_bar.showMessage(T('status_ready', t=MAX_THREADS))
        
        # 索引依然完美对应
        tab_keys = ['tab_gemini', 'tab_claude']
        for i, key in enumerate(tab_keys): self.tabs.setTabText(i, T(key))
        
        self.tabs.setTabText(2, T('tab_mask'))
        self.tabs.setTabText(3, T('tab_img2psd'))
        self.tabs.setTabText(4, T('tab_adobe'))
        
        for tab in list(self.api_tabs.values()) + [self.mask_tab, self.image_to_psd_tab, self.adobe_tab]: 
            tab._retranslate_ui()

    def _show_readme(self):
        import webbrowser
        lang_suffix = _LANG[0].lower() 
        filename = f"readme_{lang_suffix}.html"
        readme_path = os.path.abspath(filename)
        
        if os.path.exists(readme_path):
            webbrowser.open(f"file:///{readme_path.replace(chr(92), '/')}")
        else:
            err_msg = f"Cannot find manual file:\n{filename}\n\nPlease ensure it is in the same directory." if lang_suffix == "en" else f"Cannot find manual file:\n{filename}\n\nPlease ensure it is in the same directory as the application."
            err_title = "File Not Found" if lang_suffix == "en" else "File Not Found"
            QMessageBox.warning(self, err_title, err_msg)


# ─────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────
def main():
    setup_crash_logging()
    missing = []
    try: import requests
    except ImportError: missing.append('requests')
    try: from PIL import Image
    except ImportError: missing.append('Pillow')
    try: from docx import Document
    except ImportError: missing.append('python-docx')

    if missing:
        logger.error(f'Missing dependencies: {missing}')
        app = QApplication(sys.argv)
        QMessageBox.critical(None, 'Missing Dependencies', f'Please install the following:\n\npip install {" ".join(missing)}\n\npython-docx is required for Word output')
        sys.exit(1)

    app = QApplication(sys.argv)
    QImageReader.setAllocationLimit(0)
    app.setApplicationName(APP_TITLE)
    app.setOrganizationName(ORG_NAME)
    font = QFont('Segoe UI', 10)
    app.setFont(font)
    win = MainWindow()
    win.show()
    try:
        exit_code = app.exec()
    except Exception:
        logger.critical('Exception escaped app.exec()', exc_info=True)
        raise
    logger.info(f'App exit normally, code={exit_code}')
    sys.exit(exit_code)

if __name__ == '__main__':
    main()