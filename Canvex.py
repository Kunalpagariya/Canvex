import ctypes
import sys

# 🔴 REQUIRED FOR WINDOWS TASKBAR ICON
def set_app_id():
    if sys.platform == "win32":
        app_id = "com.canvex.app"  # any unique string
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)

set_app_id()


"""
Canva Image Excel Creator — Ultra Mode

This module implements a PyQt5 GUI application that reads an input Excel file,
searches the web for images matching values in specified input columns, and
inserts those images into an output Excel workbook alongside the original
data.

Key features and components:
- Multiple search backends (Bing, Google, DuckDuckGo) via Selenium scraping
- Parallel image downloading and high-quality resizing using Pillow
- Image content filters to reduce black/white/cartoon results
- Robust worker thread with checkpointing and safe shutdown
- Persistent settings and mappings stored in `canva_last_settings.json`

Notes for maintainers:
- Keep UI wiring isolated in `CanvaImageExcelCreator` class.
- Heavy I/O (Selenium, HTTP, file writes) happens in `WorkerUltra` thread.
- Settings are stored minimally to allow re-use across runs.
"""

import sys, os, time, json, random, traceback, subprocess, re
import concurrent.futures, requests
import pandas as pd

from io import BytesIO
from datetime import datetime
from PIL import Image
import xlsxwriter

from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QVBoxLayout,
    QFileDialog, QTableWidget, QTableWidgetItem, QComboBox,
    QLineEdit, QHeaderView, QAbstractItemView, QMessageBox,
    QHBoxLayout, QProgressBar, QListWidget, QListWidgetItem,
    QStackedWidget, QSplitter, QTextBrowser, QTabWidget, QDialog
)
from PyQt5.QtGui import QIcon, QPixmap, QFont, QColor
try:
    from PyQt5.QtSvg import QSvgRenderer
except Exception:
    QSvgRenderer = None

# Modern Font Awesome icons for Qt
try:
    import qtawesome as qta
    HAS_ICONS = True
except ImportError:
    HAS_ICONS = False

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


# ============================================================
# SYSTEM THEME DETECTION
# ============================================================

def is_windows_dark_mode():
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize"
        )
        val, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        return val == 0
    except:
        return False


def is_macos_dark_mode():
    try:
        out = subprocess.run(
            ["defaults", "read", "-g", "AppleInterfaceStyle"],
            capture_output=True, text=True
        )
        return "Dark" in out.stdout
    except:
        return False


def system_dark_mode():
    if sys.platform.startswith("win"):
        return is_windows_dark_mode()
    if sys.platform == "darwin":
        return is_macos_dark_mode()
    return True  # default dark for Linux


# ============================================================
# THEMES
# ============================================================

THEMES = [
    "headshot portrait closeup face",
    "cinematic lighting portrait",
    "studio headshot clean background",
    "dramatic portrait closeup",
    "smiling closeup face",
    "full body portrait",
    "natural daylight portrait",
    "magazine cover portrait",
    "Custom Theme..."
]

BAD_SITES = [
    "shutterstock","alamy","getty","adobe","dreamstime","depositphotos",
    "123rf","bigstock","vectorstock","istock"
]


# ============================================================
# STYLE SHEETS
# ============================================================

DARK_STYLE = """
/* macOS Dark Theme - Simplified for native rendering */
QWidget {
    background: #1e1e1e;
    color: #ffffff;
    font-size: 13px;
}
QLabel {
    color: #ffffff;
    padding: 2px;
    background: transparent;
}
QLineEdit {
    background: #2d2d2d;
    border: 1px solid #404040;
    border-radius: 6px;
    color: white;
    padding: 6px 10px;
    min-height: 22px;
}
QLineEdit:focus {
    border: 1px solid #0a84ff;
}
QComboBox {
    border: 1px solid #404040;
    border-radius: 6px;
    padding: 6px 10px;
    min-height: 22px;
    background: #2d2d2d;
    color: white;
}
QComboBox:focus {
    border: 2px solid #0a84ff;
}
QComboBox:disabled {
    color: #999;
}
QComboBox::drop-down {
    border: none;
    background: transparent;
    width: 30px;
}
QComboBox::down-arrow {
    image: none;
    width: 12px;
    height: 12px;
}
QTableWidget {
    background: #252525;
    alternate-background-color: #2a2a2a;
    gridline-color: #3a3a3a;
    border: 1px solid #404040;
    border-radius: 6px;
}
QTableWidget::item {
    padding: 6px;
}
QTableWidget::item:selected {
    background: #0a84ff;
}
QListWidget {
    background: #252525;
    border: 1px solid #404040;
    border-radius: 6px;
    outline: none;
}
QListWidget::item {
    padding: 8px 12px;
    margin: 2px 0px;
    border-radius: 4px;
}
QListWidget::item:hover {
    background: #3a3a3c;
}
QListWidget::item:selected {
    background: #0a84ff;
    color: white;
}
QHeaderView::section {
    background: #2d2d2d;
    color: #999;
    font-weight: 500;
    padding: 8px;
    border: none;
    border-bottom: 1px solid #404040;
}
QPushButton {
    background: #0a84ff;
    padding: 8px 16px;
    border-radius: 6px;
    color: white;
    font-weight: 500;
    border: none;
    min-height: 20px;
}
QPushButton:hover {
    background: #409cff;
}
QPushButton:pressed {
    background: #0060c0;
}
QPushButton:disabled {
    background: #404040;
    color: #666;
}
QPushButton#toolbarBtn {
    background: #2d2d2d;
    border: 1px solid #404040;
    color: #ffffff;
    padding: 6px 14px;
    min-width: 80px;
}
QPushButton#toolbarBtn:hover {
    background: #3d3d3d;
    border-color: #505050;
}
QPushButton#toolbarBtn:pressed {
    background: #252525;
}
QProgressBar {
    border: none;
    background: #333;
    border-radius: 6px;
    height: 16px;
    text-align: center;
}
QProgressBar::chunk {
    background: #0a84ff;
    border-radius: 6px;
}
QScrollBar:vertical {
    background: #2d2d2d;
    width: 12px;
}
QScrollBar::handle:vertical {
    background: #555;
    border-radius: 6px;
    min-height: 20px;
}
QScrollBar::handle:vertical:hover {
    background: #666;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}
QFrame#toolbar {
    background: #252525;
    border: 1px solid #404040;
    border-radius: 8px;
}
QCheckBox {
    spacing: 8px;
    color: #ffffff;
}
QCheckBox::indicator {
    width: 16px;
    height: 16px;
}
QSpinBox {
    background: #2d2d2d;
    border: 1px solid #404040;
    border-radius: 6px;
    padding: 6px 10px;
    min-height: 22px;
    color: white;
}
QSpinBox:hover {
    border: 1px solid #0a84ff;
}
QSpinBox::up-button, QSpinBox::down-button {
    width: 20px;
    border: none;
    background: #404040;
}
QSpinBox::up-button:hover, QSpinBox::down-button:hover {
    background: #505050;
}
QSpinBox::up-arrow {
    width: 0;
    height: 0;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-bottom: 5px solid #aaa;
}
QSpinBox::down-arrow {
    width: 0;
    height: 0;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-top: 5px solid #aaa;
}
QTextBrowser {
    background: #252525;
    border: none;
    padding: 10px;
}
QDialog {
    background: #1e1e1e;
}
"""

LIGHT_STYLE = """
/* macOS Light Theme - Simplified for native rendering */
QWidget {
    background: #f5f5f7;
    color: #1d1d1f;
    font-size: 13px;
}
QLabel {
    color: #1d1d1f;
    padding: 2px;
    background: transparent;
}
QLineEdit {
    background: #ffffff;
    border: 1px solid #d2d2d7;
    border-radius: 6px;
    color: #1d1d1f;
    padding: 6px 10px;
    min-height: 22px;
}
QLineEdit:focus {
    border: 1px solid #0a84ff;
}
QComboBox {
    border: 1px solid #d2d2d7;
    border-radius: 6px;
    padding: 6px 10px;
    min-height: 22px;
    background: #ffffff;
    color: #1d1d1f;
}
QComboBox:focus {
    border: 2px solid #0a84ff;
}
QComboBox:disabled {
    color: #999;
}
QComboBox::drop-down {
    border: none;
    background: transparent;
    width: 30px;
}
QComboBox::down-arrow {
    image: none;
    width: 12px;
    height: 12px;
}
QTableWidget {
    background: #ffffff;
    alternate-background-color: #fafafa;
    gridline-color: #e5e5e5;
    border: 1px solid #d2d2d7;
    border-radius: 6px;
}
QTableWidget::item {
    padding: 6px;
}
QTableWidget::item:selected {
    background: #0a84ff;
    color: white;
}
QListWidget {
    background: #ffffff;
    border: 1px solid #d2d2d7;
    border-radius: 6px;
    outline: none;
}
QListWidget::item {
    padding: 8px 12px;
    margin: 2px 0px;
    border-radius: 4px;
}
QListWidget::item:hover {
    background: #f0f0f0;
}
QListWidget::item:selected {
    background: #0a84ff;
    color: white;
}
QHeaderView::section {
    background: #f5f5f7;
    color: #86868b;
    font-weight: 500;
    padding: 8px;
    border: none;
    border-bottom: 1px solid #d2d2d7;
}
QPushButton {
    background: #0a84ff;
    padding: 8px 16px;
    border-radius: 6px;
    color: white;
    font-weight: 500;
    border: none;
    min-height: 20px;
}
QPushButton:hover {
    background: #409cff;
}
QPushButton:pressed {
    background: #0060c0;
}
QPushButton:disabled {
    background: #e5e5e5;
    color: #999;
}
QPushButton#toolbarBtn {
    background: #ffffff;
    border: 1px solid #d2d2d7;
    color: #1d1d1f;
    padding: 6px 14px;
    min-width: 80px;
}
QPushButton#toolbarBtn:hover {
    background: #e8e8ed;
    border-color: #c7c7cc;
}
QPushButton#toolbarBtn:pressed {
    background: #d2d2d7;
}
QProgressBar {
    border: none;
    background: #e5e5e5;
    border-radius: 6px;
    height: 16px;
    text-align: center;
}
QProgressBar::chunk {
    background: #0a84ff;
    border-radius: 6px;
}
QScrollBar:vertical {
    background: #f5f5f7;
    width: 12px;
}
QScrollBar::handle:vertical {
    background: #c7c7cc;
    border-radius: 6px;
    min-height: 20px;
}
QScrollBar::handle:vertical:hover {
    background: #a8a8ad;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}
QFrame#toolbar {
    background: #ffffff;
    border: 1px solid #d2d2d7;
    border-radius: 8px;
}
QCheckBox {
    spacing: 8px;
    color: #1d1d1f;
}
QCheckBox::indicator {
    width: 16px;
    height: 16px;
}
QSpinBox {
    background: #ffffff;
    border: 1px solid #d2d2d7;
    border-radius: 6px;
    padding: 6px 10px;
    min-height: 22px;
    color: #1d1d1f;
}
QSpinBox:hover {
    border: 1px solid #0a84ff;
}
QSpinBox::up-button, QSpinBox::down-button {
    width: 20px;
    border: none;
    background: #e8e8ed;
}
QSpinBox::up-button:hover, QSpinBox::down-button:hover {
    background: #d8d8dd;
}
QSpinBox::up-arrow {
    width: 0;
    height: 0;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-bottom: 5px solid #666;
}
QSpinBox::down-arrow {
    width: 0;
    height: 0;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-top: 5px solid #666;
}
QTextBrowser {
    background: #ffffff;
    border: none;
    padding: 10px;
}
QDialog {
    background: #f5f5f7;
}
"""


# ============================================================
# URL VALIDATION
# ============================================================

def is_valid_image_url(url):
    if not url or not url.startswith("http"):
        return False

    u = url.lower()

    # Block watermark/stock sites
    for bad in BAD_SITES:
        if bad in u:
            return False

    # Accept common image extensions
    patterns = [
        r"\.jpg(\?|$)", r"\.jpeg(\?|$)",
        r"\.png(\?|$)", r"\.webp(\?|$)"
    ]
    for p in patterns:
        if re.search(p, u):
            return True

    # Accept Bing redirect /imgres?url=...
    if "imgres?url=" in u:
        return True

    return False


# ============================================================
# WAIT HELPER (HUMAN-LIKE)
# ============================================================

def wait(a=0.25, b=0.45):
    time.sleep(random.uniform(a, b))


# ============================================================
# SELENIUM — OPTIMIZED DRIVER CREATION
# ============================================================

def create_driver():
    """Create and return a Selenium Chrome WebDriver configured for
    performance and basic anti-bot stability.

    The returned driver has a short page load timeout and several
    options to reduce resource usage when running many parallel
    image searches.
    """
    opts = webdriver.ChromeOptions()

    # Performance + Anti-bot stability
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--disable-popup-blocking")
    opts.add_argument("--disable-infobars")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--incognito")
    opts.add_argument("--window-size=1280,1000")
    opts.add_argument("--disable-features=VizDisplayCompositor")
    opts.add_argument("--renderer-process-limit=3")
    opts.add_argument("--blink-settings=imagesEnabled=true")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=opts
    )

    driver.set_page_load_timeout(12)
    return driver


# ============================================================
# BING IMAGE SCRAPER — PORTRAIT-FIRST + RETRY
# ============================================================

def bing_urls(driver, term, theme, limit=36):
    """Scrape Bing Images for candidate image URLs.

    Parameters:
    - driver: Selenium WebDriver instance (already created)
    - term: search term (string)
    - theme: theme string appended to the query to bias results
    - limit: maximum number of URLs to return

    Returns a list of image URLs (may be fewer than `limit`).
    """
    q = term.replace(" ", "+")
    t = theme.replace(" ", "+")
    url = f"https://www.bing.com/images/search?q={q}+{t}&form=HDRSC2"

    # Try load twice if needed
    try:
        driver.get(url)
    except:
        time.sleep(1)
        driver.get(url)

    time.sleep(1)

    items = driver.find_elements(By.CSS_SELECTOR, "a.iusc")
    urls = []

    # Parse JSON metadata
    for e in items:
        try:
            meta = json.loads(e.get_attribute("m"))
            link = meta.get("murl", "")
            if is_valid_image_url(link):
                urls.append(link)
                if len(urls) >= limit:
                    break
        except:
            continue

    # Retry once if Bing did not load fully
    if not urls:
        try:
            driver.get(url)
            time.sleep(1)
            items = driver.find_elements(By.CSS_SELECTOR, "a.iusc")
            for e in items:
                try:
                    meta = json.loads(e.get_attribute("m"))
                    link = meta.get("murl", "")
                    if is_valid_image_url(link):
                        urls.append(link)
                        if len(urls) >= limit:
                            break
                except:
                    continue
        except:
            pass

    random.shuffle(urls)
    return urls[:limit]


def google_urls(driver, term, theme, limit=36):
    """Scrape Google Images for candidate URLs.

    Google DOM and anti-bot measures change frequently; this implementation
    uses a simple selector search and is intended as a best-effort.
    If Google returns no results, the worker will attempt a Bing fallback.
    """
    q = term.replace(" ", "+")
    t = theme.replace(" ", "+")
    url = f"https://www.google.com/search?tbm=isch&q={q}+{t}"
    try:
        driver.get(url)
    except:
        time.sleep(1)
        driver.get(url)

    time.sleep(1)
    urls = []
    # Try common img selectors
    imgs = driver.find_elements(By.CSS_SELECTOR, "img")
    for im in imgs:
        try:
            src = im.get_attribute("src") or im.get_attribute("data-src") or ""
            if src and is_valid_image_url(src):
                urls.append(src)
                if len(urls) >= limit:
                    break
        except:
            continue

    random.shuffle(urls)
    return urls[:limit]


def ddg_urls(driver, term, theme, limit=36):
    """Scrape DuckDuckGo image results for candidate URLs.

    DuckDuckGo can provide alternative results where Google/Bing are
    blocked or rate-limited.
    """
    q = term.replace(" ", "+")
    t = theme.replace(" ", "+")
    url = f"https://duckduckgo.com/?q={q}+{t}&iax=images&ia=images"
    try:
        driver.get(url)
    except:
        time.sleep(1)
        driver.get(url)

    time.sleep(1)
    urls = []
    imgs = driver.find_elements(By.CSS_SELECTOR, "img.tile--img__img, img")
    for im in imgs:
        try:
            src = im.get_attribute("src") or im.get_attribute("data-src") or im.get_attribute("data-iurl") or ""
            if src and is_valid_image_url(src):
                urls.append(src)
                if len(urls) >= limit:
                    break
        except:
            continue

    random.shuffle(urls)
    return urls[:limit]


def fetch_image_urls(driver, term, theme, browser="Bing Images", limit=36):
    """Dispatch to a specific search backend based on `browser` string.

    This function normalizes the browser name and selects the best
    scraping function. Defaults to Bing when unrecognized.
    """
    b = (browser or "Bing Images").lower()
    if "google" in b:
        return google_urls(driver, term, theme, limit)
    if "duck" in b or "duckduckgo" in b:
        return ddg_urls(driver, term, theme, limit)
    # default to bing
    return bing_urls(driver, term, theme, limit)


# ============================================================
# PARALLEL DOWNLOAD + RESIZE — HIGH QUALITY, SMART FILTER
# ============================================================

DOWNLOAD_CACHE = {}


def dl_resize(url, target):
    """Download an image from `url` and resize it so its longest side is
    `target` pixels while applying heuristic filters to avoid unwanted
    images (near-black/white, low color variance, cartoon-like graphics).

    Returns a Pillow Image object on success, otherwise returns None.

    Implementation notes:
    - Uses a small in-memory cache to avoid duplicate downloads for the same
      URL during a single run.
    - Performs a lightweight content-type check when available but still
      attempts to open image bytes even if headers are missing.
    - Uses a small thumbnail (32x32) to estimate unique colors quickly.
    - Retries twice before giving up.
    """
    for _ in range(2):     # retry twice
        try:
            # simple in-memory cache to avoid re-downloading same URL during run
            if url in DOWNLOAD_CACHE:
                content = DOWNLOAD_CACHE[url]
                response = None
            else:
                response = requests.get(url, timeout=7)
                content = response.content
                if len(content) <= 1024 * 1024:  # cache up to ~1MB
                    DOWNLOAD_CACHE[url] = content

            # best-effort content-type check
            try:
                if response is not None:
                    ctype = response.headers.get("content-type", "").lower()
                    if "image" not in ctype:
                        # still attempt to open from bytes
                        pass
            except:
                pass

            img = Image.open(BytesIO(content))
            img.load()

            # Auto-convert unsupported modes
            if img.mode not in ("RGB", "RGBA"):
                img = img.convert("RGB")

            # --- Image content filters to avoid black/white, cartoons, graphics ---
            try:
                # Reject near-black or near-white images
                gray = img.convert("L")
                stat = Image.Stat.Stat(gray)
                mean_brightness = stat.mean[0]
                if mean_brightness < 10 or mean_brightness > 245:
                    return None

                # Reject near-grayscale / very low color variance
                if img.mode == "RGB":
                    statc = Image.Stat.Stat(img)
                    # channel standard deviations
                    stds = statc.stddev
                    if max(stds) < 12:    # very little color variance
                        return None

                # Reject cartoon/graphic-like images with very few unique colors
                # Use smaller thumbnail (32x32) for faster analysis
                small = img.convert("RGB").resize((32, 32))
                colors = small.getcolors(maxcolors=1024)
                unique_colors = len(colors) if colors else 0
                if unique_colors and unique_colors < 24:
                    return None
            except Exception:
                # If analysis fails, continue — don't block valid images
                pass

            w, h = img.size
            scale = target / max(w, h)
            new_w, new_h = int(w * scale), int(h * scale)

            # Skip tiny/garbage images
            if new_w < 80 or new_h < 80:
                continue

            return img.resize((new_w, new_h), Image.LANCZOS)

        except Exception:
            # small backoff then retry
            time.sleep(0.2)

    return None


# ============================================================
# ULTRA MODE WORKER — PARALLEL ENGINE + SOFT CANCEL
# ============================================================

class WorkerUltra(QThread):
    sig_overall = pyqtSignal(int)      # overall progress
    sig_step = pyqtSignal(int)         # per-download progress
    sig_log = pyqtSignal(str)          # text log
    sig_done = pyqtSignal(str)         # finished path
    sig_error = pyqtSignal(str)        # error message

    def __init__(
        self,
        excel_path,
        mappings,
        save_path,
        theme,
        custom_theme,
        resolution,
        custom_res,
        fmt,
        jpg_quality,
        browser,
        selected_sheet=None,
        split_enabled=False,
        records_per_file=None,
    ):
        super().__init__()
        self.excel_path = excel_path
        self.mappings = mappings
        self.save_path = save_path
        self.theme = theme
        self.custom_theme = custom_theme.strip()
        self.resolution = resolution
        self.custom_res = custom_res.strip()
        self.format = fmt.lower()
        self.jpg_quality = jpg_quality
        self.browser = browser
        self.selected_sheet = selected_sheet
        self.split_enabled = split_enabled
        self.records_per_file = records_per_file or 20

        self.cancel_requested = False
        self.logs = []

        self.start_time = None
        self.end_time = None
        self.created_files = []  # Track files created in split mode

    # ---------------------------------------------------------
    def log(self, msg):
        self.logs.append(msg)
        self.sig_log.emit(msg)

        def __doc__(self):
                """WorkerThread: performs the heavy work off the GUI thread.

                Responsibilities:
                - Read input Excel using pandas
                - Create an output Excel workbook with xlsxwriter
                - For each mapping and row: search images, download, filter, resize
                - Insert final images into the output workbook
                - Emit progress and log messages to the UI via signals
                - Ensure safe checkpointing by closing the workbook and quitting
                    Selenium even when exceptions occur
                """

    # ---------------------------------------------------------
    def get_resolution(self):
        if self.resolution == "Custom…":
            try:
                px = int(self.custom_res)
                return max(240, min(px, 4000))
            except:
                return 720

        table = {
            "240p": 240,
            "360p": 360,
            "480p": 480,
            "720p": 720,
            "1080p": 1080,
            "1440p": 1440,
            "2160p": 2160,
            "3840p": 3840,
        }
        return table.get(self.resolution, 720)

    # ---------------------------------------------------------
    def run(self):
        # Use structured try/except/finally so we always close workbook and
        # quit Selenium — this ensures a checkpointed file on failures.
        workbook = None
        workbooks = []  # For split mode
        driver = None
        pool = None
        temp_files = []
        error_exc = None
        completed = False

        try:
            # Timing
            self.start_time = datetime.now()
            self.log(f"[START] {self.start_time}")

            # Load input Excel into a pandas DataFrame. Columns are normalized
            # to string names so downstream UI combo boxes can display them.
            if self.selected_sheet:
                df = pd.read_excel(self.excel_path, sheet_name=self.selected_sheet)
            else:
                df = pd.read_excel(self.excel_path)

            # Clean column names
            clean = []
            for i, c in enumerate(df.columns):
                c2 = str(c).strip()
                if c2 == "" or c2.lower() in ("none", "nan"):
                    c2 = f"col_{i+1}"
                clean.append(c2)
            df.columns = clean

            # Theme selection: if the UI asked for a custom theme, use that
            # string. The theme is appended to search queries to bias results
            # toward portrait/headshot images.
            theme = self.custom_theme if self.theme == "Custom Theme..." else self.theme

            self.log(f"[LOG] Theme: {theme}")
            self.log(f"[LOG] Search Browser: {self.browser}")
            self.log(f"[LOG] Format: {self.format}")
            self.log(f"[LOG] JPG Quality: {self.jpg_quality}")

            res = self.get_resolution()
            self.log(f"[LOG] Resolution: {res}px")
            
            # Handle split mode
            if self.split_enabled:
                self.log(f"[LOG] Split mode enabled: {self.records_per_file} records per file")
                self.created_files = self._run_split_mode(df, theme, res, temp_files)
                completed = True
            else:
                # Original single-file mode
                workbook = xlsxwriter.Workbook(self.save_path)
                # Name the sheet based on the output file name (cleaned)
                base_name = os.path.splitext(os.path.basename(self.save_path))[0]
                # Clean sheet name (Excel limits: 31 chars, no special chars)
                sheet_name = base_name[:31].replace('/', '-').replace('\\', '-').replace('*', '').replace('?', '').replace('[', '').replace(']', '')
                if not sheet_name:
                    sheet_name = "Data"
                sheet = workbook.add_worksheet(sheet_name)

                # Write headers
                for i, col in enumerate(df.columns):
                    sheet.write(0, i, col)

                # Determine output columns
                out_map = {}
                next_i = len(df.columns)

                for inp, out in self.mappings:
                    if out not in df.columns:
                        out_map[out] = next_i
                        sheet.write(0, next_i, out)
                        next_i += 1
                    else:
                        out_map[out] = df.columns.tolist().index(out)

                for idx in out_map.values():
                    sheet.set_column(idx, idx, 22)

                sheet.set_default_row(120)

                # Selenium driver: created once per run. The driver is used for
                # scraping image search result pages. Keep it alive for the
                # duration of the processing loop to amortize startup cost.
                self.log("[LOG] Starting Selenium...")
                driver = create_driver()

                # Thread pool for images (auto-scale by CPU for parallel downloads)
                try:
                    cpus = os.cpu_count() or 4
                    maxw = min(20, max(6, cpus * 2))
                except:
                    maxw = 8
                pool = concurrent.futures.ThreadPoolExecutor(max_workers=maxw)
                total_rows = len(df)

                # ======================================================
                # PROCESS EACH ROW
                # ======================================================
                for ri, row in df.iterrows():

                    if self.cancel_requested:
                        self.log("[CANCEL] User cancelled.")
                        break

                    # Write row text data
                    for ci, col in enumerate(df.columns):
                        val = "" if pd.isna(row[col]) else str(row[col])
                        sheet.write(ri + 1, ci, val)

                    # Overall progress
                    self.sig_overall.emit(int((ri + 1) / total_rows * 100))

                    # Each mapping per row
                    for inp, out in self.mappings:

                        if self.cancel_requested:
                            break

                        val = row[inp]
                        if pd.isna(val):
                            continue

                        term = str(val).strip()

                        self.log(f"[SEARCH] {term}")
                        self.sig_step.emit(10)

                        # 1. Image search scrape (browser selectable)
                        # If using Google, give the page a chance to load more images
                        try:
                            if "google" in (self.browser or "").lower():
                                try:
                                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                                    wait(0.4, 0.9)
                                except:
                                    pass

                        except:
                            pass

                        # Dispatcher chooses the scraping backend based on the
                        # selected browser (Bing/Google/DuckDuckGo). We limit to
                        # a small candidate set for speed — the fastest portrait
                        # image is chosen among these candidates.
                        urls = fetch_image_urls(driver, term, theme, browser=self.browser, limit=24)
                        # Log first few fetched URLs for debugging
                        try:
                            self.log(f"[URLS] ({self.browser}) {len(urls)} found: {urls[:6]}")
                        except:
                            pass

                        # If Google returned nothing, try Bing as a fallback
                        if (not urls) and "google" in (self.browser or "").lower():
                            self.log("[WARN] Google returned no URLs; trying Bing fallback...")
                            try:
                                urls = fetch_image_urls(driver, term, theme, browser="Bing Images", limit=36)
                                self.log(f"[URLS] (Bing fallback) {len(urls)} found: {urls[:6]}")
                            except Exception as e:
                                self.log(f"[ERROR] Bing fallback failed: {e}")
                        self.sig_step.emit(30)

                        if not urls:
                            self.log(f"[WARN] No images found: {term}")
                            continue

                        urls = urls[:8]

                        # 2. Parallel download: submit candidate URL downloads to
                        # the thread pool. Each task returns a Pillow Image or
                        # None if download/filters fail.
                        futures = [pool.submit(dl_resize, u, res) for u in urls]

                        portrait_img = None
                        fallback_img = None

                        for fut in concurrent.futures.as_completed(futures):

                            if self.cancel_requested:
                                break

                            img = fut.result()
                            if img:
                                w, h = img.size
                                if h > w and portrait_img is None:
                                    portrait_img = img
                                if fallback_img is None:
                                    fallback_img = img
                                if portrait_img:
                                    break

                        final_img = portrait_img or fallback_img
                        if not final_img:
                            self.log(f"[WARN] Failed downloading image for {term}")
                            continue

                        self.sig_step.emit(70)

                        # 3. Save temp image to disk prior to insertion. On
                        # successful completion these temps are removed, but if an
                        # error occurs we keep them for debugging (saved with
                        # _ERROR_log.txt).
                        ext = self.format
                        # Use get_temp_dir() to ensure we write to a writable location
                        # (macOS bundled apps have read-only app bundles)
                        temp_dir = get_temp_dir()
                        fname = os.path.join(temp_dir, f"temp_{ri}_{out}.{ext}")
                        temp_files.append(fname)

                        if ext == "png":
                            final_img.save(fname, "PNG")
                        elif ext == "jpg":
                            # JPEG does not support alpha. If image has an alpha
                            # channel (RGBA/LA) composite it over white first.
                            try:
                                mode = final_img.mode
                            except:
                                mode = None

                            if mode in ("RGBA", "LA") or (mode == "P" and "transparency" in getattr(final_img, "info", {})):
                                bg = Image.new("RGB", final_img.size, (255, 255, 255))
                                alpha = final_img.split()[-1]
                                bg.paste(final_img, mask=alpha)
                                bg.save(fname, "JPEG", quality=self.jpg_quality)
                            else:
                                final_img.convert("RGB").save(fname, "JPEG", quality=self.jpg_quality)
                        else:
                            final_img.save(fname, "WEBP", quality=95)

                        self.sig_step.emit(90)

                        # 4. Insert image
                        out_col = out_map[out]
                        sheet.insert_image(
                            ri + 1,
                            out_col,
                            fname,
                            {"object_position": 1, "x_scale": 0.20, "y_scale": 0.20}
                        )

                        self.sig_step.emit(100)

                # Mark successful completion
                self.end_time = datetime.now()
                completed = True

        except Exception as e:
            error_exc = e
            # Keep going to finally block to save checkpoint and logs

        finally:
            # Always try to close workbook to finalize file (checkpoint)
            try:
                if workbook is not None:
                    workbook.close()
            except Exception:
                pass

            # Ensure Selenium is quit
            try:
                if driver is not None:
                    driver.quit()
            except Exception:
                pass

            # Cleanup temp files only on successful completion; on failure keep
            # them so user can inspect images for debugging.
            if completed:
                for f in temp_files:
                    try:
                        if os.path.exists(f):
                            os.remove(f)
                    except:
                        pass

            # Timing and logs
            try:
                if self.end_time is None:
                    self.end_time = datetime.now()
                dur = self.end_time - self.start_time if self.start_time else None
                if dur:
                    secs = int(dur.total_seconds())
                    mins, secs = divmod(secs, 60)
                    hrs, mins = divmod(mins, 60)
                    duration_text = f"{hrs}h {mins}m {secs}s"
                else:
                    duration_text = "N/A"
            except:
                duration_text = "N/A"

            # Save log file (always)
            try:
                log_path = os.path.splitext(self.save_path)[0] + "_log.txt"
                with open(log_path, "w", encoding="utf-8") as f:
                    f.write("\n".join(self.logs))
                    f.write(f"\n\nTime taken: {duration_text}")
            except:
                pass

            # If there was an exception, write detailed error log and emit
            if error_exc:
                try:
                    err_path = os.path.splitext(self.save_path)[0] + "_ERROR_log.txt"
                    with open(err_path, "w", encoding="utf-8") as f:
                        f.write("\n".join(self.logs))
                        f.write("\n\n=== ERROR ===\n")
                        f.write(traceback.format_exc())
                except:
                    pass

                # Inform UI
                if self.cancel_requested:
                    self.sig_error.emit("Process cancelled safely.")
                else:
                    self.sig_error.emit(str(error_exc))
            else:
                # Success - emit with created files if in split mode
                if self.split_enabled and self.created_files:
                    # Format: "SPLIT|file1|file2|file3" so we can distinguish from single file
                    files_string = "SPLIT|" + "|".join(self.created_files)
                    self.log(f"[SPLIT] Emitting signal with {len(self.created_files)} files")
                    self.sig_done.emit(files_string)
                else:
                    self.log(f"[DEBUG] split_enabled={self.split_enabled}, created_files={len(self.created_files) if self.created_files else 0}")
                    self.sig_done.emit(self.save_path)
    
    # ---------------------------------------------------------
    def _run_split_mode(self, df, theme, res, temp_files):
        """Process in split mode - create multiple Excel files."""
        total_rows = len(df)
        num_files = (total_rows + self.records_per_file - 1) // self.records_per_file
        
        # Get base filename
        base_path = self.save_path
        base_name = os.path.splitext(os.path.basename(base_path))[0]
        base_dir = os.path.dirname(base_path)
        
        driver = create_driver()
        try:
            cpus = os.cpu_count() or 4
            maxw = min(20, max(6, cpus * 2))
        except:
            maxw = 8
        pool = concurrent.futures.ThreadPoolExecutor(max_workers=maxw)
        
        # Create files for each chunk
        for file_idx in range(num_files):
            if self.cancel_requested:
                self.log("[CANCEL] User cancelled.")
                break
            
            start_row = file_idx * self.records_per_file
            end_row = min(start_row + self.records_per_file, total_rows)
            
            # Create filename with row range
            start_num = start_row + 1
            end_num = end_row
            output_filename = f"{base_name}_{start_num:05d}-{end_num:05d}.xlsx"
            output_path = os.path.join(base_dir, output_filename)
            
            self.log(f"[SPLIT] Creating file {file_idx + 1}/{num_files}: {output_filename}")
            
            # Create workbook for this chunk
            workbook = xlsxwriter.Workbook(output_path)
            sheet = workbook.add_worksheet("Data")
            
            # Write headers
            for i, col in enumerate(df.columns):
                sheet.write(0, i, col)
            
            # Determine output columns
            out_map = {}
            next_i = len(df.columns)
            
            for inp, out in self.mappings:
                if out not in df.columns:
                    out_map[out] = next_i
                    sheet.write(0, next_i, out)
                    next_i += 1
                else:
                    out_map[out] = df.columns.tolist().index(out)
            
            for idx in out_map.values():
                sheet.set_column(idx, idx, 22)
            
            sheet.set_default_row(120)
            
            # Process rows in this chunk
            for local_ri, (global_ri, row) in enumerate(df.iloc[start_row:end_row].iterrows()):
                if self.cancel_requested:
                    break
                
                # Write row text data
                for ci, col in enumerate(df.columns):
                    val = "" if pd.isna(row[col]) else str(row[col])
                    sheet.write(local_ri + 1, ci, val)
                
                # Overall progress
                progress = int((start_row + local_ri + 1) / total_rows * 100)
                self.sig_overall.emit(progress)
                
                # Each mapping per row
                for inp, out in self.mappings:
                    if self.cancel_requested:
                        break
                    
                    val = row[inp]
                    if pd.isna(val):
                        continue
                    
                    term = str(val).strip()
                    self.log(f"[SEARCH] {term}")
                    self.sig_step.emit(10)
                    
                    # Image search
                    try:
                        if "google" in (self.browser or "").lower():
                            try:
                                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                                wait(0.4, 0.9)
                            except:
                                pass
                    except:
                        pass
                    
                    urls = fetch_image_urls(driver, term, theme, browser=self.browser, limit=24)
                    try:
                        self.log(f"[URLS] ({self.browser}) {len(urls)} found: {urls[:6]}")
                    except:
                        pass
                    
                    # Fallback to Bing if Google returns nothing
                    if (not urls) and "google" in (self.browser or "").lower():
                        self.log("[WARN] Google returned no URLs; trying Bing fallback...")
                        try:
                            urls = fetch_image_urls(driver, term, theme, browser="Bing Images", limit=36)
                            self.log(f"[URLS] (Bing fallback) {len(urls)} found: {urls[:6]}")
                        except Exception as e:
                            self.log(f"[ERROR] Bing fallback failed: {e}")
                    self.sig_step.emit(30)
                    
                    if not urls:
                        self.log(f"[WARN] No images found: {term}")
                        continue
                    
                    urls = urls[:8]
                    
                    # Download images
                    futures = [pool.submit(dl_resize, u, res) for u in urls]
                    
                    portrait_img = None
                    fallback_img = None
                    
                    for fut in concurrent.futures.as_completed(futures):
                        if self.cancel_requested:
                            break
                        
                        img = fut.result()
                        if img:
                            w, h = img.size
                            if h > w and portrait_img is None:
                                portrait_img = img
                            if fallback_img is None:
                                fallback_img = img
                            if portrait_img:
                                break
                    
                    final_img = portrait_img or fallback_img
                    if not final_img:
                        self.log(f"[WARN] Failed downloading image for {term}")
                        continue
                    
                    self.sig_step.emit(70)
                    
                    # Save temp image
                    ext = self.format
                    temp_dir = get_temp_dir()
                    fname = os.path.join(temp_dir, f"temp_{global_ri}_{out}.{ext}")
                    temp_files.append(fname)
                    
                    if ext == "png":
                        final_img.save(fname, "PNG")
                    elif ext == "jpg":
                        try:
                            mode = final_img.mode
                        except:
                            mode = None
                        
                        if mode in ("RGBA", "LA") or (mode == "P" and "transparency" in getattr(final_img, "info", {})):
                            bg = Image.new("RGB", final_img.size, (255, 255, 255))
                            alpha = final_img.split()[-1]
                            bg.paste(final_img, mask=alpha)
                            bg.save(fname, "JPEG", quality=self.jpg_quality)
                        else:
                            final_img.convert("RGB").save(fname, "JPEG", quality=self.jpg_quality)
                    else:
                        final_img.save(fname, "WEBP", quality=95)
                    
                    self.sig_step.emit(90)
                    
                    # Insert image
                    out_col = out_map[out]
                    sheet.insert_image(
                        local_ri + 1,
                        out_col,
                        fname,
                        {"object_position": 1, "x_scale": 0.20, "y_scale": 0.20}
                    )
                    
                    self.sig_step.emit(100)
            
            # Close workbook for this chunk
            try:
                workbook.close()
                self.created_files.append(output_path)
            except:
                pass
        
        # Cleanup
        try:
            driver.quit()
        except:
            pass
        
        self.end_time = datetime.now()
        return self.created_files  # Return list of created files


# ============================================================
# CUSTOM HOVER WIDGETS
# ============================================================

class HoverListWidget(QListWidget):
    """Custom QListWidget with mouse hover effects on items."""
    
    def __init__(self, is_dark_mode=True):
        super().__init__()
        self.is_dark_mode = is_dark_mode
        # Much more visible hover colors
        self.hover_color = QColor("#454545") if is_dark_mode else QColor("#e8e8eb")
        self.normal_color = QColor("#2a2a2a") if is_dark_mode else QColor("#ffffff")
        self.hover_text_color = QColor("#ffffff") if is_dark_mode else QColor("#000000")
        self.normal_text_color = QColor("#ffffff") if is_dark_mode else QColor("#000000")
        self.setMouseTracking(True)
    
    def mouseMoveEvent(self, event):
        super().mouseMoveEvent(event)
        item = self.itemAt(event.pos())
        
        # Update background colors for all items
        for i in range(self.count()):
            current_item = self.item(i)
            # Only apply hover to non-selected items
            if current_item == item and current_item != self.currentItem():
                current_item.setBackground(self.hover_color)
                current_item.setForeground(self.hover_text_color)
            elif current_item != self.currentItem():
                current_item.setBackground(self.normal_color)
                current_item.setForeground(self.normal_text_color)
    
    def leaveEvent(self, event):
        super().leaveEvent(event)
        # Reset all non-selected items to normal color
        for i in range(self.count()):
            item = self.item(i)
            if item != self.currentItem():
                item.setBackground(self.normal_color)
                item.setForeground(self.normal_text_color)
    
    def setCurrentRow(self, row):
        super().setCurrentRow(row)
        # Update all item colors after selection change
        for i in range(self.count()):
            item = self.item(i)
            if item == self.currentItem():
                # Selected items keep their blue color (handled by stylesheet)
                pass
            else:
                item.setBackground(self.normal_color)
                item.setForeground(self.normal_text_color)


# ============================================================
# GUI APPLICATION
# ============================================================

class HoverListWidget(QListWidget):
    """Custom QListWidget with mouse hover effects on items."""
    
    def __init__(self, is_dark_mode=True):
        super().__init__()
        self.is_dark_mode = is_dark_mode
        # Much more visible hover colors
        self.hover_color = QColor("#454545") if is_dark_mode else QColor("#e8e8eb")
        self.normal_color = QColor("#2a2a2a") if is_dark_mode else QColor("#ffffff")
        self.hover_text_color = QColor("#ffffff") if is_dark_mode else QColor("#000000")
        self.normal_text_color = QColor("#ffffff") if is_dark_mode else QColor("#000000")
        self.setMouseTracking(True)
    
    def mouseMoveEvent(self, event):
        super().mouseMoveEvent(event)
        item = self.itemAt(event.pos())
        
        # Update background colors for all items
        for i in range(self.count()):
            current_item = self.item(i)
            # Only apply hover to non-selected items
            if current_item == item and current_item != self.currentItem():
                current_item.setBackground(self.hover_color)
                current_item.setForeground(self.hover_text_color)
            elif current_item != self.currentItem():
                current_item.setBackground(self.normal_color)
                current_item.setForeground(self.normal_text_color)
    
    def leaveEvent(self, event):
        super().leaveEvent(event)
        # Reset all non-selected items to normal color
        for i in range(self.count()):
            item = self.item(i)
            if item != self.currentItem():
                item.setBackground(self.normal_color)
                item.setForeground(self.normal_text_color)
    
    def setCurrentRow(self, row):
        super().setCurrentRow(row)
        # Update all item colors after selection change
        for i in range(self.count()):
            item = self.item(i)
            if item == self.currentItem():
                # Selected items keep their blue color (handled by stylesheet)
                pass
            else:
                item.setBackground(self.normal_color)
                item.setForeground(self.normal_text_color)


# ============================================================
# GUI APPLICATION
# ============================================================

class CanvaImageExcelCreator(QWidget):
    """Main GUI window for the Canvex application.

    Responsibilities:
    - Provide a small, focused UI to select an input Excel file and configure
      image search/mapping options.
    - Store and restore UI settings and mappings to disk.
    - Start/stop the background `WorkerUltra` thread and display progress.
    """

    def __init__(self):
        super().__init__()

        self.setWindowTitle("Canvex")

        # Allow window resizing
        self.resize(900, 700)
        self.setMinimumSize(700, 600)

        self.setAcceptDrops(True)

        # DATA
        self.excel_path = None
        self.columns = []
        self.worker = None
        self.session_running = False
        self.manual_theme_override = None  # None = auto, "light", "dark"
        
        # Filter settings (for Settings dialog)
        self.filter_portrait = True
        self.filter_bw = True
        self.filter_cartoon = True
        
        # Split settings (for Settings dialog)
        self.split_enabled_setting = False  # Enable/disable split mode
        self.records_per_file_setting = 20  # Records per output file
        
        # Settings file to persist last mappings and options
        # Use get_writable_dir() to ensure we can write to this location
        # even when running as a bundled macOS app
        try:
            base = get_writable_dir()
        except:
            base = os.getcwd()
        self.settings_path = os.path.join(base, "canva_last_settings.json")
        print(f"[DEBUG] Settings path initialized: {self.settings_path}")
        # Last directory used for opening/saving Excel files
        self.last_dir = None
        # Try set window icon from app_icon.ico or logo.svg in the same folder.
        # Store icon for later use in taskbar and splash
        self.app_icon = None
        try:
            # First, try app_icon.ico (best for Windows taskbar)
            # Use resource_path() so the icon is found when bundled by PyInstaller
            icon_path = resource_path("app_icon.ico")
            if os.path.exists(icon_path):
                self.app_icon = QIcon(icon_path)
                if not self.app_icon.isNull():
                    self.setWindowIcon(self.app_icon)
            else:
                # Fallback to logo.svg
                icon_path = os.path.join(base, "logo.svg")
                if not os.path.exists(icon_path):
                    icon_path = os.path.join(base, "Applogo.svg")
                if os.path.exists(icon_path):
                    icon = QIcon()
                    # Preferred sizes for window/titlebar/taskbar
                    sizes = [64, 48, 32, 24, 16]
                    if QSvgRenderer is not None:
                        renderer = QSvgRenderer(icon_path)
                        for s in sizes:
                            pix = QPixmap(s, s)
                            pix.fill(Qt.transparent)
                            painter = None
                            try:
                                from PyQt5.QtGui import QPainter
                                painter = QPainter(pix)
                                renderer.render(painter)
                            except Exception:
                                pass
                            finally:
                                if painter:
                                    painter.end()
                            if not pix.isNull():
                                icon.addPixmap(pix)
                    else:
                        # Fallback: try loading as pixmap directly (PyQt may support SVG natively)
                        for s in sizes:
                            pix = QPixmap(icon_path)
                            if not pix.isNull():
                                icon.addPixmap(pix.scaled(s, s))

                    # If we built an icon, set it
                    if not icon.isNull():
                        self.app_icon = icon
                        self.setWindowIcon(icon)
        except Exception:
            pass

        # Initial theme
        self.apply_theme_auto()

        # MAIN LAYOUT
        layout = QVBoxLayout()
        self.setLayout(layout)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        # =============================================================
        # TOOLBAR (File, Settings, Help, About, Theme)
        # =============================================================
        from PyQt5.QtWidgets import QFrame
        
        toolbar_frame = QFrame()
        toolbar_frame.setObjectName("toolbar")
        toolbar_frame.setFixedHeight(44)
        toolbar_layout = QHBoxLayout(toolbar_frame)
        toolbar_layout.setSpacing(6)
        toolbar_layout.setContentsMargins(8, 0, 8, 0)  # Zero vertical margins for centering
        toolbar_layout.setAlignment(Qt.AlignVCenter)  # Center buttons vertically
        layout.addWidget(toolbar_frame)

        # File Menu Button
        self.btn_file = QPushButton(" File")
        if HAS_ICONS:
            self.btn_file.setIcon(qta.icon('fa5s.folder-open', color='#0a84ff'))
        self.btn_file.setObjectName("toolbarBtn")
        self.btn_file.setFixedHeight(32)
        self.btn_file.clicked.connect(self.show_file_menu)
        toolbar_layout.addWidget(self.btn_file)

        # Settings Menu Button
        self.btn_settings = QPushButton(" Settings")
        if HAS_ICONS:
            self.btn_settings.setIcon(qta.icon('fa5s.cog', color='#0a84ff'))
        self.btn_settings.setObjectName("toolbarBtn")
        self.btn_settings.setFixedHeight(32)
        self.btn_settings.clicked.connect(self.show_settings_dialog)
        toolbar_layout.addWidget(self.btn_settings)

        # Help Menu Button
        self.btn_help = QPushButton(" Help")
        if HAS_ICONS:
            self.btn_help.setIcon(qta.icon('fa5s.question-circle', color='#0a84ff'))
        self.btn_help.setObjectName("toolbarBtn")
        self.btn_help.setFixedHeight(32)
        self.btn_help.clicked.connect(self.show_help)
        toolbar_layout.addWidget(self.btn_help)

        # About Button
        self.btn_about = QPushButton(" About")
        if HAS_ICONS:
            self.btn_about.setIcon(qta.icon('fa5s.info-circle', color='#0a84ff'))
        self.btn_about.setObjectName("toolbarBtn")
        self.btn_about.setFixedHeight(32)
        self.btn_about.clicked.connect(self.show_about)
        toolbar_layout.addWidget(self.btn_about)

        toolbar_layout.addStretch()

        # Theme Toggle on toolbar
        self.btn_theme = QPushButton(" Theme")
        if HAS_ICONS:
            self.btn_theme.setIcon(qta.icon('fa5s.adjust', color='#0a84ff'))
        self.btn_theme.setObjectName("toolbarBtn")
        self.btn_theme.setFixedHeight(32)
        self.btn_theme.clicked.connect(self.show_theme_dialog)
        toolbar_layout.addWidget(self.btn_theme)

        # =============================================================
        # FILE SECTION (always visible)
        # - Placeholder shown initially (large drop area)
        # - Compact file info shown after an Excel is loaded
        # =============================================================
        self.btn_select = QPushButton("  Select Excel File")
        if HAS_ICONS:
            self.btn_select.setIcon(qta.icon('fa5s.file-excel', color='white'))
            self.btn_select.setIconSize(QSize(20, 20))
        self.btn_select.setFixedHeight(44)
        self.btn_select.setStyleSheet("QPushButton { font-size: 14px; font-weight: 600; }")
        self.btn_select.clicked.connect(self.load_excel)
        layout.addWidget(self.btn_select)

        # Big placeholder area (shown when no file loaded)
        self.placeholder_widget = QLabel("Drop Excel file here\nor click the button above")
        self.placeholder_widget.setAlignment(Qt.AlignCenter)
        self.placeholder_widget.setStyleSheet("QLabel { padding: 30px; font-size: 15px; border: 2px dashed #555; border-radius: 12px; color: #888; }")
        layout.addWidget(self.placeholder_widget)

        # Compact file info (hidden until a file is loaded)
        self.file_info_widget = QWidget()
        fi_layout = QHBoxLayout(self.file_info_widget)
        self.lbl_file_compact = QLabel("")
        self.lbl_file_compact.setStyleSheet("QLabel { padding: 6px; }")
        fi_layout.addWidget(self.lbl_file_compact)
        fi_layout.addStretch()
        self.btn_change_file = QPushButton("  Change File")
        if HAS_ICONS:
            self.btn_change_file.setIcon(qta.icon('fa5s.sync-alt', color='white'))
            self.btn_change_file.setIconSize(QSize(14, 14))
        self.btn_change_file.clicked.connect(self.load_excel)
        fi_layout.addWidget(self.btn_change_file)
        self.file_info_widget.setVisible(False)
        layout.addWidget(self.file_info_widget)

        # Load saved basic settings (theme, browser, last dir) but do NOT restore
        # column mappings until an Excel file is loaded and `self.columns` exists.
        # NOTE: Will be called at end of __init__ after all UI elements are created

        # compact spacing between file area and content
        # (no extra stretch here to avoid large gaps)

        # =============================================================
        # CONTENT AREA (hidden until Excel is loaded)
        # =============================================================
        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout(self.content_widget)
        self.content_widget.setVisible(False)
        layout.addWidget(self.content_widget)

        # IMAGE THEME
        prm_row = QHBoxLayout()
        self.content_layout.addLayout(prm_row)
        prm_row.addWidget(QLabel("Image Theme:"))
        self.dd_theme = QComboBox()
        self.dd_theme.addItems(THEMES)
        self.dd_theme.currentIndexChanged.connect(self.toggle_custom_theme)
        prm_row.addWidget(self.dd_theme)
        self.txt_custom_theme = QLineEdit()
        self.txt_custom_theme.setPlaceholderText("Custom theme…")
        self.txt_custom_theme.setVisible(False)
        prm_row.addWidget(self.txt_custom_theme)

        # SEARCH BROWSER
        browser_row = QHBoxLayout()
        self.content_layout.addLayout(browser_row)
        browser_row.addWidget(QLabel("Search Browser:"))
        self.dd_browser = QComboBox()
        self.dd_browser.addItems(["Bing Images", "Google Images", "DuckDuckGo"])
        browser_row.addWidget(self.dd_browser)

        # RESOLUTION
        res_row = QHBoxLayout()
        self.content_layout.addLayout(res_row)
        res_row.addWidget(QLabel("Resolution:"))
        self.dd_res = QComboBox()
        self.dd_res.addItems([
            "240p", "360p", "480p", "720p", "1080p",
            "1440p", "2160p", "3840p", "Custom…"
        ])
        self.dd_res.currentIndexChanged.connect(self.toggle_custom_res)
        res_row.addWidget(self.dd_res)
        self.txt_custom_res = QLineEdit()
        self.txt_custom_res.setPlaceholderText("240–4000")
        self.txt_custom_res.setVisible(False)
        res_row.addWidget(self.txt_custom_res)

        # FORMAT
        fmt_row = QHBoxLayout()
        self.content_layout.addLayout(fmt_row)
        fmt_row.addWidget(QLabel("Format:"))
        self.dd_fmt = QComboBox()
        self.dd_fmt.addItems(["PNG", "JPG", "WEBP"])
        self.dd_fmt.currentIndexChanged.connect(self.toggle_jpg_quality)
        fmt_row.addWidget(self.dd_fmt)

        q_row = QHBoxLayout()
        self.content_layout.addLayout(q_row)
        self.jpg_quality_label = QLabel("JPG Quality:")
        self.jpg_quality_label.setVisible(False)
        q_row.addWidget(self.jpg_quality_label)
        self.dd_jpg_quality = QComboBox()
        self.dd_jpg_quality.addItems([
            "60 (Low)",
            "75 (Medium)",
            "90 (High)",
            "100 (Ultra)"
        ])
        self.dd_jpg_quality.setVisible(False)
        q_row.addWidget(self.dd_jpg_quality)

        # MAPPING TABLE with Add Mapping button
        map_row = QHBoxLayout()
        lbl_map = QLabel("Column Mappings:")
        lbl_map.setStyleSheet("font-weight: 600; font-size: 14px;")
        map_row.addWidget(lbl_map)
        map_row.addStretch()
        
        self.btn_load_previous = QPushButton("  Load Previous")
        if HAS_ICONS:
            self.btn_load_previous.setIcon(qta.icon('fa5s.history', color='white'))
            self.btn_load_previous.setIconSize(QSize(16, 16))
        self.btn_load_previous.setFixedHeight(34)
        self.btn_load_previous.setMinimumWidth(130)
        self.btn_load_previous.setToolTip("Load a previously saved mapping")
        self.btn_load_previous.clicked.connect(self.show_previous_mappings)
        map_row.addWidget(self.btn_load_previous)
        
        self.btn_add = QPushButton("  Add Mapping")
        if HAS_ICONS:
            self.btn_add.setIcon(qta.icon('fa5s.plus-circle', color='white'))
            self.btn_add.setIconSize(QSize(16, 16))
        self.btn_add.setFixedHeight(34)
        self.btn_add.setMinimumWidth(130)
        self.btn_add.setToolTip("Add a new column mapping row")
        self.btn_add.clicked.connect(self.add_mapping)
        map_row.addWidget(self.btn_add)
        self.content_layout.addLayout(map_row)

        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels([
            "#", "Input Column", "Output Column",
            "New Column Name", ""
        ])
        # Set proper column widths
        self.table.setColumnWidth(0, 40)   # Row number
        self.table.setColumnWidth(1, 180)  # Input Column
        self.table.setColumnWidth(2, 180)  # Output Column
        self.table.setColumnWidth(3, 140)  # New Column Name
        self.table.setColumnWidth(4, 50)   # Delete
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setDefaultSectionSize(44)  # Row height for dropdown visibility
        self.table.verticalHeader().setVisible(False)  # Hide row numbers on left
        self.table.setMinimumHeight(180)
        self.table.setAlternatingRowColors(True)
        self.content_layout.addWidget(self.table)

        # PROGRESS BAR (single)
        self.progress_container = QWidget()
        self.progress_layout = QVBoxLayout(self.progress_container)
        self.progress_layout.setContentsMargins(0, 4, 0, 4)
        self.progress_layout.setSpacing(2)
        self.progress_container.setVisible(False)
        self.content_layout.addWidget(self.progress_container)

        self.lbl_overall = QLabel("Progress:")
        self.lbl_overall.setStyleSheet("font-weight: 600;")
        self.progress_layout.addWidget(self.lbl_overall)
        self.pb_overall = QProgressBar()
        self.pb_overall.setFixedHeight(14)
        self.progress_layout.addWidget(self.pb_overall)

        # START / CANCEL BUTTONS
        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)
        self.content_layout.addLayout(btn_row)
        self.btn_start = QPushButton("  Start Processing")
        if HAS_ICONS:
            self.btn_start.setIcon(qta.icon('fa5s.play-circle', color='white'))
            self.btn_start.setIconSize(QSize(20, 20))
        self.btn_start.setFixedHeight(44)
        self.btn_start.setStyleSheet("QPushButton { font-size: 14px; font-weight: 600; background: #34c759; } QPushButton:hover { background: #30d158; }")
        self.btn_start.clicked.connect(self.start_session)
        btn_row.addWidget(self.btn_start)
        self.btn_cancel = QPushButton("  Cancel")
        if HAS_ICONS:
            self.btn_cancel.setIcon(qta.icon('fa5s.stop-circle', color='white'))
            self.btn_cancel.setIconSize(QSize(20, 20))
        self.btn_cancel.setFixedHeight(44)
        self.btn_cancel.clicked.connect(self.cancel_session)
        self.btn_cancel.setVisible(False)
        self.btn_cancel.setStyleSheet("QPushButton { font-size: 14px; font-weight: 600; background: #ff3b30; } QPushButton:hover { background: #ff6961; }")
        btn_row.addWidget(self.btn_cancel)

        # NOW load saved basic settings after all UI elements are created
        try:
            self.load_basic_settings()
        except Exception as e:
            print(f"[DEBUG] Error loading basic settings on startup: {e}")

    # ============================================================
    # THEME CONTROLS
    # ============================================================

    def apply_theme_auto(self):
        dark = system_dark_mode()
        if dark:
            QApplication.instance().setStyleSheet(DARK_STYLE)
        else:
            QApplication.instance().setStyleSheet(LIGHT_STYLE)

    # ============================================================
    # CUSTOM FIELDS HANDLING
    # ============================================================

    def toggle_custom_theme(self):
        self.txt_custom_theme.setVisible(
            self.dd_theme.currentText() == "Custom Theme..."
        )

    def toggle_custom_res(self):
        self.txt_custom_res.setVisible(
            self.dd_res.currentText() == "Custom…"
        )

    def toggle_jpg_quality(self):
        is_jpg = self.dd_fmt.currentText() == "JPG"
        self.jpg_quality_label.setVisible(is_jpg)
        self.dd_jpg_quality.setVisible(is_jpg)

    # ============================================================
    # LOAD EXCEL + RESET SESSION
    # ============================================================

    def load_excel(self):
        """Open a file dialog to choose an input Excel workbook and load
        its column headers into the mapping UI.

        This method intentionally only loads headers (not full data) to
        populate the column combo boxes quickly.
        """
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel", self.last_dir or "", "Excel (*.xlsx)"
        )
        if not path:
            return

        # Get available sheets
        try:
            xl = pd.ExcelFile(path)
            sheet_names = xl.sheet_names
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to read Excel file:\n{str(e)}")
            return
        
        # If multiple sheets, let user select
        selected_sheet = None
        if len(sheet_names) > 1:
            selected_sheet = self._select_sheet_dialog(sheet_names)
            if selected_sheet is None:
                return  # User cancelled
        else:
            selected_sheet = sheet_names[0]
        
        # Store selected sheet for later use
        self.selected_sheet = selected_sheet
        
        df = pd.read_excel(path, sheet_name=selected_sheet)
        self.columns = [str(c).strip() for c in df.columns]

        # Reset ONLY table + progress
        self.table.setRowCount(0)
        self.pb_overall.setValue(0)

        self.excel_path = path
        # Persist last directory for future dialogs
        try:
            self.last_dir = os.path.dirname(path)
            # Save minimal settings (mappings may be empty) so next run opens same folder
            try:
                self.save_settings([])
            except:
                pass
        except:
            pass
        # Switch UI from placeholder to compact file info
        self.placeholder_widget.setVisible(False)
        self.btn_select.setVisible(False)
        # Show filename and sheet name
        self.lbl_file_compact.setText(f"{os.path.basename(path)} [{selected_sheet}]")
        self.file_info_widget.setVisible(True)

        # Add to recent files
        self._add_to_recent_files(path)

        # Restore last mappings/settings (if available)
        self.load_settings()

        # Show content area now that Excel is loaded
        self.content_widget.setVisible(True)
        self.session_running = False

    # ============================================================
    # MAPPING ROW CONTROL
    # ============================================================

    def add_mapping(self):
        """Append a fresh mapping row to the mapping `QTableWidget`.

        Each mapping row contains:
        - row number (read-only)
        - input column selector (combo)
        - output column selector (combo with 'Create New Column...')
        - optional text field for new column name
        - delete button
        """
        r = self.table.rowCount()
        self.table.insertRow(r)

        self.table.setItem(r, 0, QTableWidgetItem(str(r + 1)))

        dd_in = QComboBox()
        dd_in.addItems(self.columns)
        self.table.setCellWidget(r, 1, dd_in)

        dd_out = QComboBox()
        dd_out.addItems(self.columns)
        dd_out.addItem("Create New Column...")
        dd_out.currentIndexChanged.connect(lambda _, row=r: self.toggle_new_col(row))
        self.table.setCellWidget(r, 2, dd_out)

        txt_new = QLineEdit()
        txt_new.setPlaceholderText("Enter new column name...")
        txt_new.setVisible(False)  # Start hidden
        txt_new.setText("")  # Clear any text
        self.table.setCellWidget(r, 3, txt_new)

        btn_del = QPushButton()
        if HAS_ICONS:
            btn_del.setIcon(qta.icon('fa5s.trash-alt', color='#ff3b30'))
            btn_del.setIconSize(QSize(16, 16))
        else:
            btn_del.setText("✕")
        btn_del.setFixedSize(34, 34)
        btn_del.setToolTip("Remove this mapping")
        btn_del.setStyleSheet("QPushButton { background: transparent; border: 1px solid #ff3b30; border-radius: 6px; } QPushButton:hover { background: #ff3b30; }")
        # Use cellWidget to find the row dynamically instead of capturing it
        btn_del.clicked.connect(lambda _, button=btn_del: self.delete_row_by_button(button))
        self.table.setCellWidget(r, 4, btn_del)

    def load_basic_settings(self):
        """Load only basic UI settings (theme, browser, last_dir, filters) on app startup.
        
        This does NOT load mappings or try to access self.columns.
        This is called in __init__ to restore the last used directory immediately.
        """
        try:
            if not os.path.exists(self.settings_path):
                print(f"[DEBUG] Settings file not found: {self.settings_path}")
                return
            with open(self.settings_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            print(f"[DEBUG] Basic settings loaded from: {self.settings_path}")
        except Exception as e:
            print(f"[DEBUG] Error reading settings: {e}")
            return

        try:
            # Restore theme
            theme = data.get("theme")
            if theme:
                idx = self.dd_theme.findText(theme)
                if idx != -1:
                    self.dd_theme.setCurrentIndex(idx)
                else:
                    self.dd_theme.setCurrentIndex(self.dd_theme.findText("Custom Theme..."))
                    self.txt_custom_theme.setText(data.get("custom_theme", ""))

            # Restore resolution
            res = data.get("resolution")
            if res:
                idx = self.dd_res.findText(res)
                if idx != -1:
                    self.dd_res.setCurrentIndex(idx)
                else:
                    self.dd_res.setCurrentIndex(self.dd_res.findText("Custom…"))
                    self.txt_custom_res.setText(data.get("custom_res", ""))

            # Restore format
            fmt = data.get("format")
            if fmt:
                idx = self.dd_fmt.findText(fmt)
                if idx != -1:
                    self.dd_fmt.setCurrentIndex(idx)
            
            # Restore JPG quality
            jpgq = data.get("jpg_quality")
            if jpgq:
                try:
                    qlist = [self.dd_jpg_quality.itemText(i).split()[0] for i in range(self.dd_jpg_quality.count())]
                    if str(jpgq) in qlist:
                        idx = qlist.index(str(jpgq))
                        self.dd_jpg_quality.setCurrentIndex(idx)
                except:
                    pass
            
            # Restore browser
            browser = data.get("browser")
            if browser and hasattr(self, 'dd_browser'):
                bi = self.dd_browser.findText(browser)
                if bi != -1:
                    self.dd_browser.setCurrentIndex(bi)
            
            # Restore last directory - THIS IS KEY!
            lastd = data.get("last_excel_dir")
            if lastd and os.path.exists(lastd):
                self.last_dir = lastd
                print(f"[DEBUG] Restored last directory: {self.last_dir}")
            
            # Restore filter settings
            if "filter_portrait" in data:
                self.filter_portrait = data.get("filter_portrait", True)
            if "filter_bw" in data:
                self.filter_bw = data.get("filter_bw", True)
            if "filter_cartoon" in data:
                self.filter_cartoon = data.get("filter_cartoon", True)
        except Exception as e:
            print(f"[DEBUG] Error restoring basic settings: {e}")

    def load_settings(self):
        """Load column mappings after an Excel file has been loaded.
        
        Only restores mappings if self.columns exists (Excel is loaded).
        Basic settings should have been loaded already via load_basic_settings().
        """
        try:
            if not os.path.exists(self.settings_path):
                print(f"[DEBUG] Settings file not found: {self.settings_path}")
                return
            with open(self.settings_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            print(f"[DEBUG] Settings loaded from: {self.settings_path}")
        except Exception as e:
            print(f"[DEBUG] Error reading settings from {self.settings_path}: {e}")
            return

        # Only restore mappings if we have loaded an Excel (columns exist)
        maps = data.get("mappings") or []
        if not maps or not self.columns:
            print(f"[DEBUG] Skipping mappings restore: maps={bool(maps)}, columns={bool(self.columns)}")
            return

        self.table.setRowCount(0)
        for m in maps:
            src = m[0]
            dst = m[1]
            self.add_mapping()
            r = self.table.rowCount() - 1

            dd_in = self.table.cellWidget(r, 1)
            dd_out = self.table.cellWidget(r, 2)
            txt_new = self.table.cellWidget(r, 3)

            if src and dd_in.findText(src) != -1:
                dd_in.setCurrentIndex(dd_in.findText(src))

            if dst:
                if dst in self.columns and dd_out.findText(dst) != -1:
                    dd_out.setCurrentIndex(dd_out.findText(dst))
                else:
                    # Use create-new option
                    idx_create = dd_out.findText("Create New Column...")
                    if idx_create != -1:
                        dd_out.setCurrentIndex(idx_create)
                        txt_new.setText(dst)
                        txt_new.setVisible(True)

    def auto_save_mappings(self):
        """Automatically save current mappings whenever they change.
        
        This is called whenever a mapping row is added, modified, or deleted.
        """
        if not self.excel_path:
            return  # Only save if an Excel file is loaded
        
        # Collect current mappings from table
        mappings = []
        for r in range(self.table.rowCount()):
            dd_in = self.table.cellWidget(r, 1)
            dd_out = self.table.cellWidget(r, 2)
            txt_new = self.table.cellWidget(r, 3)
            
            if not dd_in or not dd_out:
                continue
            
            src = dd_in.currentText()
            out_col = dd_out.currentText()
            
            if out_col == "Create New Column...":
                dst = txt_new.text()
                if not dst:
                    continue  # Skip if no name provided
            else:
                dst = out_col
            
            mappings.append((src, dst))
        
        # Save mappings silently in the background
        try:
            self.save_settings(mappings)
            print(f"[DEBUG] Auto-saved mappings: {len(mappings)} rows")
        except Exception as e:
            print(f"[DEBUG] Error auto-saving mappings: {e}")

    def save_settings(self, mappings):
        """Persist current UI settings and mappings to `self.settings_path`.

        The `mappings` argument should be a list of (input, output)
        tuples provided by the caller (the UI collects and validates these).
        
        This method merges with existing data to preserve recent_files and other
        settings that may have been added separately. It also maintains a history
        of the last 5 mapping configurations.
        """
        try:
            # Load existing data to preserve recent_files and other settings
            existing_data = {}
            if os.path.exists(self.settings_path):
                try:
                    with open(self.settings_path, "r", encoding="utf-8") as f:
                        existing_data = json.load(f)
                except Exception:
                    existing_data = {}
            
            # Update mapping history - keep last 5 configurations
            mapping_history = existing_data.get("mapping_history", [])
            
            # Only add to history if mappings are not empty and different from the last one
            if mappings and (not mapping_history or mapping_history[0]["mappings"] != mappings):
                new_entry = {
                    "timestamp": datetime.now().isoformat(),
                    "mappings": mappings
                }
                mapping_history.insert(0, new_entry)
                mapping_history = mapping_history[:5]  # Keep only last 5
            
            # Merge new settings with existing data
            data = existing_data.copy()
            data.update({
                "theme": self.dd_theme.currentText(),
                "custom_theme": self.txt_custom_theme.text(),
                "resolution": self.dd_res.currentText(),
                "custom_res": self.txt_custom_res.text(),
                "format": self.dd_fmt.currentText(),
                "jpg_quality": int(self.dd_jpg_quality.currentText().split(" ")[0]) if self.dd_jpg_quality.currentIndex() != -1 else None,
                "browser": self.dd_browser.currentText() if hasattr(self, 'dd_browser') else "Bing Images",
                "last_excel_dir": self.last_dir,
                "mappings": mappings,
                "mapping_history": mapping_history,
                "filter_portrait": self.filter_portrait,
                "filter_bw": self.filter_bw,
                "filter_cartoon": self.filter_cartoon
            })
            
            # Ensure directory exists
            settings_dir = os.path.dirname(self.settings_path)
            if settings_dir and not os.path.exists(settings_dir):
                os.makedirs(settings_dir, exist_ok=True)
            
            with open(self.settings_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            print(f"[DEBUG] Settings saved to: {self.settings_path}")
        except Exception as e:
            print(f"[DEBUG] Error saving settings to {self.settings_path}: {e}")

    def load_mapping_from_history(self, mapping_config):
        """Load a saved mapping configuration into the table.
        
        Args:
            mapping_config: A list of (input_col, output_col) tuples
        """
        if not self.columns:
            QMessageBox.warning(self, "Error", "Please load an Excel file first")
            return
        
        # Clear existing mappings
        self.table.setRowCount(0)
        
        # Add each mapping from history
        for src, dst in mapping_config:
            self.add_mapping_without_save()
            r = self.table.rowCount() - 1
            
            dd_in = self.table.cellWidget(r, 1)
            dd_out = self.table.cellWidget(r, 2)
            txt_new = self.table.cellWidget(r, 3)
            
            # Set input column
            if dd_in and src in self.columns:
                dd_in.setCurrentIndex(dd_in.findText(src))
            
            # Set output column
            if dd_out:
                if dst in self.columns:
                    dd_out.setCurrentIndex(dd_out.findText(dst))
                else:
                    # This is a new column name
                    idx_create = dd_out.findText("Create New Column...")
                    if idx_create != -1:
                        dd_out.setCurrentIndex(idx_create)
                        txt_new.setText(dst)
                        txt_new.setVisible(True)
        
        QMessageBox.information(self, "Success", "Mapping loaded successfully!")

    def add_mapping_without_save(self):
        """Same as add_mapping() but without triggering auto-save.
        
        Used when loading from history to avoid multiple saves.
        """
        r = self.table.rowCount()
        self.table.insertRow(r)

        self.table.setItem(r, 0, QTableWidgetItem(str(r + 1)))

        dd_in = QComboBox()
        dd_in.addItems(self.columns)
        self.table.setCellWidget(r, 1, dd_in)

        dd_out = QComboBox()
        dd_out.addItems(self.columns)
        dd_out.addItem("Create New Column...")
        dd_out.currentIndexChanged.connect(lambda _, row=r: self.toggle_new_col(row))
        self.table.setCellWidget(r, 2, dd_out)

        txt_new = QLineEdit()
        txt_new.setPlaceholderText("Enter new column name...")
        txt_new.setVisible(False)  # Start hidden
        txt_new.setText("")  # Clear any text
        self.table.setCellWidget(r, 3, txt_new)

        btn_del = QPushButton()
        if HAS_ICONS:
            btn_del.setIcon(qta.icon('fa5s.trash-alt', color='#ff3b30'))
            btn_del.setIconSize(QSize(16, 16))
        else:
            btn_del.setText("✕")
        btn_del.setFixedSize(34, 34)
        btn_del.setToolTip("Remove this mapping")
        btn_del.setStyleSheet("QPushButton { background: transparent; border: 1px solid #ff3b30; border-radius: 6px; } QPushButton:hover { background: #ff3b30; }")
        # Use cellWidget to find the row dynamically instead of capturing it
        btn_del.clicked.connect(lambda _, button=btn_del: self.delete_row_by_button(button))
        self.table.setCellWidget(r, 4, btn_del)

    def show_previous_mappings(self):
        """Show a dialog with the last 5 saved mapping configurations."""
        try:
            if not os.path.exists(self.settings_path):
                QMessageBox.information(self, "No Recent Mappings", "No recent mappings found")
                return
            
            with open(self.settings_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            
            mapping_history = data.get("mapping_history", [])
            
            if not mapping_history:
                QMessageBox.information(self, "No Recent Mappings", "No recent mappings found")
                return
            
            # Create dialog
            is_dark = self.manual_theme_override == "dark" or (self.manual_theme_override is None and system_dark_mode())
            
            dlg = QDialog(self)
            dlg.setWindowTitle("Previous Mappings")
            dlg.resize(700, 600)  # Larger initial size
            dlg.setMinimumSize(600, 500)  # Minimum to prevent cramping
            
            layout = QVBoxLayout(dlg)
            layout.setSpacing(12)
            layout.setContentsMargins(20, 20, 20, 20)
            
            title = QLabel("Select a Previous Mapping")
            title.setStyleSheet("font-weight: 600; font-size: 14px;")
            layout.addWidget(title)
            
            # Create the custom list widget with hover effects
            mapping_list = HoverListWidget(is_dark)
            mapping_list.setMinimumHeight(200)
            
            # Apply styling
            if is_dark:
                mapping_list.setStyleSheet("""
                    QListWidget {
                        background: #2a2a2a;
                        border: 1px solid #404040;
                        border-radius: 6px;
                        outline: none;
                    }
                    QListWidget::item {
                        padding: 8px 12px;
                        margin: 2px 0px;
                        border-radius: 4px;
                        color: #ffffff;
                    }
                    QListWidget::item:selected {
                        background: #0a84ff;
                        color: white;
                    }
                """)
            else:
                mapping_list.setStyleSheet("""
                    QListWidget {
                        background: #f9f9f9;
                        border: 1px solid #d2d2d7;
                        border-radius: 6px;
                        outline: none;
                    }
                    QListWidget::item {
                        padding: 8px 12px;
                        margin: 2px 0px;
                        border-radius: 4px;
                        color: #000000;
                    }
                    QListWidget::item:selected {
                        background: #0a84ff;
                        color: white;
                    }
                """)
            
            # Populate with items
            for idx, entry in enumerate(mapping_history):
                timestamp = entry.get("timestamp", "Unknown")
                mappings = entry.get("mappings", [])
                
                # Parse timestamp to show human-readable format
                try:
                    dt = datetime.fromisoformat(timestamp)
                    time_str = dt.strftime("%Y-%m-%d %H:%M:%S")
                except:
                    time_str = timestamp
                
                # Create display text showing number of mappings
                display_text = f"Mapping #{idx + 1} - {time_str} ({len(mappings)} mappings)"
                item = QListWidgetItem(display_text)
                item.setData(Qt.UserRole, json.dumps(mappings))
                
                # Set initial background color
                if is_dark:
                    item.setBackground(QColor("#2a2a2a"))
                    item.setForeground(QColor("#ffffff"))
                else:
                    item.setBackground(QColor("#ffffff"))
                    item.setForeground(QColor("#000000"))
                
                mapping_list.addItem(item)
            
            layout.addWidget(mapping_list, 1)  # Give it stretch space
            
            # Preview section
            preview_label = QLabel("Preview of Selected Mapping:")
            preview_label.setStyleSheet("font-weight: 600; margin-top: 10px;")
            layout.addWidget(preview_label)
            
            preview_table = QTableWidget()
            preview_table.setColumnCount(2)
            preview_table.setHorizontalHeaderLabels(["Input Column", "Output Column"])
            preview_table.setMinimumHeight(200)  # More space for preview
            preview_table.setMaximumHeight(250)
            preview_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
            preview_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
            preview_table.verticalHeader().setVisible(False)
            preview_table.setAlternatingRowColors(True)  # Better visibility
            layout.addWidget(preview_table, 1)  # Give it stretch space
            
            # Function to update preview
            def update_preview():
                item = mapping_list.currentItem()
                preview_table.setRowCount(0)
                
                if not item:
                    return
                
                mapping_json = item.data(Qt.UserRole)
                try:
                    mappings = json.loads(mapping_json)
                    
                    for src, dst in mappings:
                        row = preview_table.rowCount()
                        preview_table.insertRow(row)
                        preview_table.setItem(row, 0, QTableWidgetItem(src))
                        preview_table.setItem(row, 1, QTableWidgetItem(dst))
                except Exception as e:
                    print(f"[DEBUG] Error updating preview: {e}")
            
            mapping_list.itemSelectionChanged.connect(update_preview)
            
            # Auto-select first item and update preview
            if mapping_list.count() > 0:
                mapping_list.setCurrentRow(0)
                update_preview()  # Explicitly call to show preview
            
            # Buttons
            btn_row = QHBoxLayout()
            btn_row.setSpacing(10)
            
            btn_cancel = QPushButton("Cancel")
            btn_cancel.clicked.connect(dlg.reject)
            btn_row.addWidget(btn_cancel)
            
            btn_reset = QPushButton("Reset All Mappings")
            btn_reset.setStyleSheet("QPushButton { background: #ff9500; color: white; font-weight: 500; padding: 6px 16px; border-radius: 6px; }")
            btn_reset.clicked.connect(lambda: self.reset_all_mappings() or dlg.accept())
            btn_row.addWidget(btn_reset)
            
            btn_row.addStretch()
            
            btn_load = QPushButton("Load Selected")
            btn_load.setStyleSheet("QPushButton { background: #0a84ff; color: white; font-weight: 500; padding: 6px 16px; border-radius: 6px; }")
            btn_row.addWidget(btn_load)
            
            layout.addLayout(btn_row)
            
            dlg.setStyleSheet(DARK_STYLE if is_dark else LIGHT_STYLE)
            
            def load_selected():
                item = mapping_list.currentItem()
                if not item:
                    QMessageBox.warning(dlg, "Error", "Please select a mapping")
                    return
                
                mapping_json = item.data(Qt.UserRole)
                mappings = json.loads(mapping_json)
                dlg.accept()
                self.load_mapping_from_history(mappings)
            
            btn_load.clicked.connect(load_selected)
            mapping_list.itemDoubleClicked.connect(load_selected)
            
            dlg.exec_()
            
            
        except Exception as e:
            print(f"[DEBUG] Error showing previous mappings: {e}")
            QMessageBox.critical(self, "Error", f"Error loading previous mappings:\n{str(e)}")

    def toggle_new_col(self, row):
        """Show or hide the new column name field based on dropdown selection.
        
        Only show txt_new when "Create New Column..." is selected.
        Otherwise, hide and clear the field.
        """
        dd_out = self.table.cellWidget(row, 2)
        txt = self.table.cellWidget(row, 3)
        
        if not dd_out or not txt:
            return
        
        is_create_new = dd_out.currentText() == "Create New Column..."
        
        if is_create_new:
            # Show and enable field for new column name
            txt.setVisible(True)
            txt.setEnabled(True)
        else:
            # Hide and clear field - not creating new column
            txt.setVisible(False)
            txt.setEnabled(False)
            txt.setText("")  # Clear any existing text

    def delete_row(self, row):
        self.table.removeRow(row)
        # Renumber
        for i in range(self.table.rowCount()):
            self.table.setItem(i, 0, QTableWidgetItem(str(i + 1)))

    def delete_row_by_button(self, button):
        """Delete a row by finding which row contains the button."""
        for row in range(self.table.rowCount()):
            if self.table.cellWidget(row, 4) == button:
                self.table.removeRow(row)
                # Renumber remaining rows
                for i in range(self.table.rowCount()):
                    self.table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
                print(f"[DEBUG] Deleted mapping row {row + 1}")
                return
        print(f"[DEBUG] Could not find button in table")

    def reset_all_mappings(self):
        """Clear all mappings from the table."""
        reply = QMessageBox.question(
            self,
            "Reset Mappings",
            "Are you sure you want to clear all mappings?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            # Remove all rows from the table
            while self.table.rowCount() > 0:
                self.table.removeRow(0)
            print("[DEBUG] All mappings cleared")

    # ============================================================
    # UI LOCK / UNLOCK
    # ============================================================

    def lock_ui(self):
        self.session_running = True
        # Disable ALL interactive elements during processing
        widgets = [
            self.btn_select, self.btn_add, self.table,
            self.dd_theme, self.dd_res, self.dd_fmt,
            self.txt_custom_theme, self.txt_custom_res,
            self.dd_jpg_quality, self.btn_theme,
            self.btn_file, self.btn_settings, self.btn_help, self.btn_about,
            self.btn_start
        ]
        if hasattr(self, 'dd_browser'):
            widgets.append(self.dd_browser)
        if hasattr(self, 'btn_change_file'):
            widgets.append(self.btn_change_file)
            
        for w in widgets:
            w.setEnabled(False)

        self.btn_start.setVisible(False)
        # Ensure Cancel button is visible and reset to default state
        self.btn_cancel.setText("Cancel")
        self.btn_cancel.setEnabled(True)
        self.btn_cancel.setVisible(True)

        # Show progress container
        self.progress_container.setVisible(True)

    def unlock_ui(self):
        self.session_running = False

        widgets = [
            self.btn_select, self.btn_add, self.table,
            self.dd_theme, self.dd_res, self.dd_fmt,
            self.txt_custom_theme, self.txt_custom_res,
            self.dd_jpg_quality, self.btn_theme,
            self.btn_file, self.btn_settings, self.btn_help, self.btn_about,
            self.btn_start
        ]
        if hasattr(self, 'dd_browser'):
            widgets.append(self.dd_browser)
        if hasattr(self, 'btn_change_file'):
            widgets.append(self.btn_change_file)
            
        for w in widgets:
            w.setEnabled(True)

        # Reset cancel button state so next run starts clean
        self.btn_cancel.setText("Cancel")
        self.btn_cancel.setEnabled(True)

        self.btn_start.setVisible(True)
        self.btn_cancel.setVisible(False)
        
        # Hide progress container
        self.progress_container.setVisible(False)

    # ============================================================
    # CANCEL PROCESS
    # ============================================================

    def cancel_session(self):
        if self.worker:
            self.worker.cancel_requested = True
            self.btn_cancel.setText("Cancelling…")
            self.btn_cancel.setEnabled(False)

    # ============================================================
    # START SESSION
    # ============================================================

    def start_session(self):
        """Validate input mappings and start the `WorkerUltra` thread.

        This method performs lightweight validation (e.g. new column names
        not empty) and then persists the settings before launching the
        background worker.
        """
        if self.session_running:
            QMessageBox.warning(self, "Running", "Process already running.")
            return

        if not self.excel_path:
            QMessageBox.warning(self, "Missing File", "Select an Excel file first.")
            return

        # Collect mappings
        mappings = []
        for r in range(self.table.rowCount()):
            dd_in = self.table.cellWidget(r, 1)
            dd_out = self.table.cellWidget(r, 2)
            txt_new = self.table.cellWidget(r, 3)

            in_col = dd_in.currentText()
            out_col = dd_out.currentText()

            if out_col == "Create New Column...":
                new_col_name = txt_new.text().strip()
                if not new_col_name:
                    QMessageBox.warning(self, "Error", f"Row {r+1}: Enter new column name.")
                    return
                out_col = new_col_name

            mappings.append((in_col, out_col))

        # Save location - use custom dialog with split options
        save_path, split_enabled, records_per_file = self._get_save_path_with_split_options()
        if not save_path:
            return
        
        # Prepare settings
        theme = self.dd_theme.currentText()
        custom_theme = self.txt_custom_theme.text()
        res = self.dd_res.currentText()
        custom_res = self.txt_custom_res.text()
        fmt = self.dd_fmt.currentText()
        quality = int(self.dd_jpg_quality.currentText().split(" ")[0])

        # Persist mappings + settings so user doesn't need to re-enter next time
        try:
            self.save_settings(mappings)
        except Exception:
            pass

        # Lock UI
        self.lock_ui()

        # Start worker (pass browser selection)
        browser = self.dd_browser.currentText() if hasattr(self, 'dd_browser') else "Bing Images"
        selected_sheet = getattr(self, 'selected_sheet', None)
        # Start worker with split parameters
        self.worker = WorkerUltra(
            self.excel_path, mappings, save_path,
            theme, custom_theme, res, custom_res,
            fmt, quality, browser, selected_sheet,
            split_enabled, records_per_file
        )

        self.worker.sig_overall.connect(self.pb_overall.setValue)
        self.worker.sig_log.connect(lambda m: print(m))
        self.worker.sig_done.connect(self.finish_success)
        self.worker.sig_error.connect(self.finish_error)

        self.worker.start()

    # ============================================================
    # FINISH CALLBACKS
    # ============================================================

    def finish_success(self, path):
        try:
            st = self.worker.start_time
            et = self.worker.end_time
            dur = et - st
            secs = int(dur.total_seconds())
            mins, secs = divmod(secs, 60)
            hrs, mins = divmod(mins, 60)
            duration = f"{hrs}h {mins}m {secs}s"
        except:
            duration = "N/A"

        # Check if this is split mode (path starts with "SPLIT|")
        if path.startswith("SPLIT|"):
            # Extract all created files
            files = path.split("|")[1:]
            self._show_split_files_dialog(files, duration)
        else:
            # Single file mode - original behavior
            reply = QMessageBox.question(
                self,
                "Success",
                f"Excel created:\n{path}\n\nTime taken: {duration}\n\nWould you like to open the file?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )

            if reply == QMessageBox.Yes:
                try:
                    if sys.platform == "darwin":
                        subprocess.run(["open", path])
                    elif sys.platform == "win32":
                        os.startfile(path)
                    else:
                        subprocess.run(["xdg-open", path])
                except Exception as e:
                    QMessageBox.warning(self, "Warning", f"Could not open file:\n{e}")

        self.unlock_ui()
        # Reset progress bar
        self.pb_overall.setValue(0)
    
    def _get_save_path_with_split_options(self):
        """Show file save dialog with split options included."""
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QSpinBox, QLabel, QPushButton, QFileDialog, QMessageBox
        from PyQt5.QtCore import Qt
        
        # First get the file path
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Save Output Excel", self.last_dir or "", "Excel (*.xlsx)"
        )
        if not save_path:
            return None, False, 20
        
        is_dark = self.manual_theme_override == "dark" or (self.manual_theme_override is None and system_dark_mode())
        
        # Step 1: Ask if user wants to split
        reply = QMessageBox.question(
            self,
            "Split Output Files",
            "Do you want to split the output into multiple files?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.No:
            # No split
            return save_path, False, 20
        
        # Step 2: If yes, ask for records per file
        dlg = QDialog(self)
        dlg.setWindowTitle("Records Per File")
        dlg.setFixedSize(350, 150)
        
        layout = QVBoxLayout()
        
        # Instructions
        lbl_info = QLabel("How many records per file?")
        layout.addWidget(lbl_info)
        
        # Spinbox for records per file
        h_layout = QHBoxLayout()
        h_layout.addWidget(QLabel("Records per file:"))
        spin_records = QSpinBox()
        spin_records.setMinimum(1)
        spin_records.setMaximum(10000)
        spin_records.setValue(getattr(self, 'records_per_file_setting', 20))
        h_layout.addWidget(spin_records)
        h_layout.addStretch()
        layout.addLayout(h_layout)
        
        layout.addSpacing(20)
        
        # Buttons
        btn_layout = QHBoxLayout()
        btn_ok = QPushButton("OK")
        btn_cancel = QPushButton("Cancel")
        btn_layout.addStretch()
        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)
        
        dlg.setLayout(layout)
        
        # Style dialog based on theme
        if is_dark:
            dlg.setStyleSheet("""
                QDialog { background-color: #2b2b2b; color: white; }
                QLabel { color: white; }
                QSpinBox { background-color: #1e1e1e; color: white; border: 1px solid #555; }
                QPushButton { background-color: #0e639c; color: white; border: none; padding: 5px; border-radius: 3px; }
                QPushButton:hover { background-color: #1177bb; }
            """)
        
        btn_ok.clicked.connect(dlg.accept)
        btn_cancel.clicked.connect(dlg.reject)
        
        result = dlg.exec_()
        
        if result == QDialog.Accepted:
            records_per_file = spin_records.value()
            # Update settings for next time
            self.split_enabled_setting = True
            self.records_per_file_setting = records_per_file
            return save_path, True, records_per_file
        else:
            # User cancelled - proceed without split
            return save_path, False, 20
    
    def _show_split_files_dialog(self, files, duration):
        """Show dialog with all generated split files for user to open."""
        from PyQt5.QtWidgets import QDialog, QListWidget, QListWidgetItem, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QMessageBox
        
        is_dark = self.manual_theme_override == "dark" or (self.manual_theme_override is None and system_dark_mode())
        
        dlg = QDialog(self)
        dlg.setWindowTitle("Split Files Generated")
        dlg.setFixedSize(500, 400)
        
        if is_dark:
            dlg_style = """
                QDialog { background-color: #1e1e1e; }
                QLabel { color: #ffffff; font-size: 13px; }
                QLabel#title { font-size: 16px; font-weight: 600; color: #0a84ff; }
                QLabel#info { font-size: 12px; color: #8e8e93; }
                QListWidget {
                    background-color: #2a2a2c; border: 1px solid #3a3a3c;
                    border-radius: 8px; color: #ffffff; font-size: 13px;
                }
                QListWidget::item { padding: 8px 12px; }
                QListWidget::item:hover { background-color: #3a3a3c; }
                QListWidget::item:selected { background-color: #0a84ff; color: #ffffff; }
                QPushButton {
                    background-color: #0a84ff; color: white; border: none;
                    border-radius: 6px; padding: 8px 16px; font-weight: 500;
                }
                QPushButton:hover { background-color: #409cff; }
                QPushButton:disabled { background-color: #404040; color: #666; }
            """
        else:
            dlg_style = """
                QDialog { background-color: #f5f5f7; }
                QLabel { color: #1d1d1f; font-size: 13px; }
                QLabel#title { font-size: 16px; font-weight: 600; color: #007aff; }
                QLabel#info { font-size: 12px; color: #6e6e73; }
                QListWidget {
                    background-color: #ffffff; border: 1px solid #d2d2d7;
                    border-radius: 8px; color: #1d1d1f; font-size: 13px;
                }
                QListWidget::item { padding: 8px 12px; }
                QListWidget::item:hover { background-color: #f0f0f0; }
                QListWidget::item:selected { background-color: #007aff; color: #ffffff; }
                QPushButton {
                    background-color: #007aff; color: white; border: none;
                    border-radius: 6px; padding: 8px 16px; font-weight: 500;
                }
                QPushButton:hover { background-color: #0066d6; }
            """
        
        dlg.setStyleSheet(dlg_style)
        layout = QVBoxLayout(dlg)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Title
        title = QLabel("✓ Files Successfully Created!")
        title.setObjectName("title")
        layout.addWidget(title)
        
        # Info
        info = QLabel(f"Time taken: {duration}\nGenerated {len(files)} files")
        info.setObjectName("info")
        layout.addWidget(info)
        
        # Instructions
        instruction = QLabel("Click on a file to open it:")
        instruction.setObjectName("info")
        layout.addWidget(instruction)
        
        # List of files
        file_list = QListWidget()
        for file_path in files:
            filename = os.path.basename(file_path)
            item = QListWidgetItem(filename)
            item.setData(Qt.UserRole, file_path)
            file_list.addItem(item)
        
        file_list.setMaximumHeight(250)
        layout.addWidget(file_list)
        
        # Double-click to open
        def open_file(item):
            file_path = item.data(Qt.UserRole)
            if file_path and os.path.exists(file_path):
                try:
                    if sys.platform == "darwin":
                        subprocess.run(["open", file_path])
                    elif sys.platform == "win32":
                        os.startfile(file_path)
                    else:
                        subprocess.run(["xdg-open", file_path])
                except Exception as e:
                    QMessageBox.warning(self, "Warning", f"Could not open file:\n{e}")
        
        file_list.itemDoubleClicked.connect(open_file)
        
        # Buttons
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        
        btn_close = QPushButton("Close")
        btn_close.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_close)
        
        btn_row.addStretch()
        
        btn_open_folder = QPushButton("Open Folder")
        def open_folder():
            if files:
                folder = os.path.dirname(files[0])
                try:
                    if sys.platform == "darwin":
                        subprocess.run(["open", folder])
                    elif sys.platform == "win32":
                        os.startfile(folder)
                    else:
                        subprocess.run(["xdg-open", folder])
                except:
                    pass
            dlg.accept()
        btn_open_folder.clicked.connect(open_folder)
        btn_row.addWidget(btn_open_folder)
        
        layout.addLayout(btn_row)
        
        dlg.exec_()


    def finish_error(self, msg):
        QMessageBox.critical(self, "Error", msg)
        self.unlock_ui()
        # Reset progress bar
        self.pb_overall.setValue(0)

    # ============================================================
    # DRAG & DROP
    # ============================================================

    def show_file_menu(self):
        from PyQt5.QtWidgets import QDialog, QListWidget, QListWidgetItem
        
        # Detect theme
        is_dark = self.manual_theme_override == "dark" or (self.manual_theme_override is None and system_dark_mode())
        
        # macOS-native theme styles
        if is_dark:
            dialog_style = """
                QDialog {
                    background-color: #1e1e1e;
                    color: #ffffff;
                    font-family: -apple-system, 'SF Pro Text', 'Helvetica Neue', sans-serif;
                }
                QLabel {
                    color: #ffffff;
                    font-size: 13px;
                }
                QLabel#header {
                    font-size: 16px;
                    font-weight: 600;
                    color: #ffffff;
                    padding: 4px 0 8px 0;
                }
                QLabel#sectionLabel {
                    font-size: 12px;
                    font-weight: 500;
                    color: #8e8e93;
                    padding-top: 8px;
                }
                QLabel#currentFile {
                    color: #0a84ff;
                    font-size: 12px;
                    padding: 8px;
                    background-color: #2a2a2c;
                    border-radius: 6px;
                }
                QPushButton {
                    background-color: #3a3a3c;
                    color: #ffffff;
                    border: none;
                    border-radius: 6px;
                    padding: 8px 16px;
                    font-size: 13px;
                    font-weight: 500;
                }
                QPushButton:hover {
                    background-color: #48484a;
                }
                QPushButton:pressed {
                    background-color: #2c2c2e;
                }
                QPushButton#primaryBtn {
                    background-color: #0a84ff;
                    color: #ffffff;
                }
                QPushButton#primaryBtn:hover {
                    background-color: #0077e6;
                }
                QListWidget {
                    background-color: #2a2a2c;
                    border: 1px solid #3a3a3c;
                    border-radius: 8px;
                    padding: 4px;
                    color: #ffffff;
                    font-size: 13px;
                }
                QListWidget::item {
                    padding: 8px 12px;
                    border-radius: 6px;
                    margin: 2px 4px;
                }
                QListWidget::item:hover {
                    background-color: #3a3a3c;
                }
                QListWidget::item:selected {
                    background-color: #0a84ff;
                    color: #ffffff;
                }
            """
        else:
            dialog_style = """
                QDialog {
                    background-color: #f5f5f7;
                    color: #1d1d1f;
                    font-family: -apple-system, 'SF Pro Text', 'Helvetica Neue', sans-serif;
                }
                QLabel {
                    color: #1d1d1f;
                    font-size: 13px;
                }
                QLabel#header {
                    font-size: 16px;
                    font-weight: 600;
                    color: #1d1d1f;
                    padding: 4px 0 8px 0;
                }
                QLabel#sectionLabel {
                    font-size: 12px;
                    font-weight: 500;
                    color: #6e6e73;
                    padding-top: 8px;
                }
                QLabel#currentFile {
                    color: #007aff;
                    font-size: 12px;
                    padding: 8px;
                    background-color: #ffffff;
                    border: 1px solid #d2d2d7;
                    border-radius: 6px;
                }
                QPushButton {
                    background-color: #ffffff;
                    color: #1d1d1f;
                    border: 1px solid #d2d2d7;
                    border-radius: 6px;
                    padding: 8px 16px;
                    font-size: 13px;
                    font-weight: 500;
                }
                QPushButton:hover {
                    background-color: #f0f0f0;
                }
                QPushButton:pressed {
                    background-color: #e0e0e0;
                }
                QPushButton#primaryBtn {
                    background-color: #007aff;
                    color: #ffffff;
                    border: none;
                }
                QPushButton#primaryBtn:hover {
                    background-color: #0066d6;
                }
                QListWidget {
                    background-color: #ffffff;
                    border: 1px solid #d2d2d7;
                    border-radius: 8px;
                    padding: 4px;
                    color: #1d1d1f;
                    font-size: 13px;
                }
                QListWidget::item {
                    padding: 8px 12px;
                    border-radius: 6px;
                    margin: 2px 4px;
                }
                QListWidget::item:hover {
                    background-color: #f0f0f0;
                }
                QListWidget::item:selected {
                    background-color: #007aff;
                    color: #ffffff;
                }
            """
        
        dlg = QDialog(self)
        dlg.setWindowTitle("File")
        dlg.resize(420, 380)
        dlg.setStyleSheet(dialog_style)
        layout = QVBoxLayout(dlg)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Header label
        hdr = QLabel("File Operations")
        hdr.setObjectName("header")
        layout.addWidget(hdr)
        
        # Action buttons
        btn_open = QPushButton("Open Excel File...")
        btn_open.setObjectName("primaryBtn")
        btn_open.clicked.connect(lambda: (dlg.accept(), self.load_excel()))
        layout.addWidget(btn_open)
        
        # Recent files section
        recent_label = QLabel("Recent Files (Last 10)")
        recent_label.setObjectName("sectionLabel")
        layout.addWidget(recent_label)
        
        # Use HoverListWidget for better UX
        recent_list = HoverListWidget(is_dark)
        recent_list.setMaximumHeight(200)
        
        # Load recent files from settings (all 10)
        recent_files = []
        try:
            if os.path.exists(self.settings_path):
                with open(self.settings_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    recent_files = data.get("recent_files", [])  # Get all, not just [:5]
        except Exception as e:
            print(f"[DEBUG] Error loading recent files: {e}")
        
        if recent_files:
            for f in recent_files:
                file_exists = os.path.exists(f)
                file_name = os.path.basename(f)
                
                if file_exists:
                    # File exists - show normally
                    item = QListWidgetItem(f"✓ {file_name}")
                    item.setToolTip(f"Click to open\n{f}")
                    if is_dark:
                        item.setBackground(QColor("#2a2a2a"))
                        item.setForeground(QColor("#ffffff"))
                    else:
                        item.setBackground(QColor("#ffffff"))
                        item.setForeground(QColor("#000000"))
                else:
                    # File missing - show with warning
                    item = QListWidgetItem(f"✗ {file_name}")
                    item.setToolTip(f"File not found at:\n{f}\n(Double-click to try locating it)")
                    item.setForeground(QColor("#ff3b30") if is_dark else QColor("#ff6961"))
                    if is_dark:
                        item.setBackground(QColor("#2a2a2a"))
                    else:
                        item.setBackground(QColor("#ffffff"))
                
                item.setData(Qt.UserRole, f)
                recent_list.addItem(item)
        
        if recent_list.count() == 0:
            item = QListWidgetItem("No recent files")
            item.setFlags(item.flags() & ~Qt.ItemIsSelectable)
            if is_dark:
                item.setBackground(QColor("#2a2a2a"))
                item.setForeground(QColor("#999999"))
            else:
                item.setBackground(QColor("#ffffff"))
                item.setForeground(QColor("#999999"))
            recent_list.addItem(item)
        
        layout.addWidget(recent_list)
        
        # Open recent file on double-click
        def open_recent(item):
            path = item.data(Qt.UserRole)
            if path:
                if os.path.exists(path):
                    dlg.accept()
                    self._load_excel_from_path(path)
                else:
                    # File not found - show popup
                    QMessageBox.warning(
                        dlg,
                        "File Not Found",
                        f"The file is no longer at:\n{path}\n\nWould you like to browse for it?",
                        QMessageBox.Ok,
                        QMessageBox.Ok
                    )
        
        recent_list.itemDoubleClicked.connect(open_recent)
        
        # Buttons row
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        
        btn_clear_recent = QPushButton("Clear Recent")
        btn_clear_recent.clicked.connect(lambda: self._clear_recent_files(recent_list))
        btn_row.addWidget(btn_clear_recent)
        
        btn_row.addStretch()
        
        btn_reveal = QPushButton("Reveal Settings")
        btn_reveal.clicked.connect(self._reveal_settings_file)
        btn_row.addWidget(btn_reveal)
        
        layout.addLayout(btn_row)
        
        layout.addStretch()
        
        # Current file info
        if self.excel_path:
            current_lbl = QLabel(f"Current: {os.path.basename(self.excel_path)}")
            current_lbl.setObjectName("currentFile")
            current_lbl.setToolTip(self.excel_path)
            layout.addWidget(current_lbl)
        
        # Close button
        btn_close = QPushButton("Close")
        btn_close.clicked.connect(dlg.accept)
        layout.addWidget(btn_close)
        
        dlg.exec_()
    
    def _load_excel_from_path(self, path):
        """Load an Excel file from a given path."""
        try:
            # Get available sheets
            xl = pd.ExcelFile(path)
            sheet_names = xl.sheet_names
            
            # If multiple sheets, let user select
            selected_sheet = None
            if len(sheet_names) > 1:
                selected_sheet = self._select_sheet_dialog(sheet_names)
                if selected_sheet is None:
                    return  # User cancelled
            else:
                selected_sheet = sheet_names[0]
            
            self.selected_sheet = selected_sheet
            df = pd.read_excel(path, sheet_name=selected_sheet)
            self.columns = [str(c).strip() for c in df.columns]
            self.table.setRowCount(0)
            self.pb_overall.setValue(0)
            self.excel_path = path
            self.last_dir = os.path.dirname(path)
            
            # Add to recent files
            self._add_to_recent_files(path)
            
            self.placeholder_widget.setVisible(False)
            self.btn_select.setVisible(False)
            self.lbl_file_compact.setText(f"{os.path.basename(path)} [{selected_sheet}]")
            self.file_info_widget.setVisible(True)
            self.load_settings()
            self.content_widget.setVisible(True)
            self.session_running = False
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load file:\n{str(e)}")
    
    def _show_split_config_dialog(self):
        """Show dialog to configure how many records per file for splitting output."""
        from PyQt5.QtWidgets import QDialog, QSpinBox
        
        is_dark = self.manual_theme_override == "dark" or (self.manual_theme_override is None and system_dark_mode())
        
        dlg = QDialog(self)
        dlg.setWindowTitle("Split Output Files")
        dlg.setFixedWidth(380)
        
        if is_dark:
            dlg_style = """
                QDialog { background-color: #1e1e1e; }
                QLabel { color: #ffffff; font-size: 13px; }
                QLabel#title { font-size: 16px; font-weight: 600; }
                QLabel#info { font-size: 12px; color: #8e8e93; }
                QSpinBox {
                    background-color: #2d2d2d; color: white; border: 1px solid #404040;
                    border-radius: 6px; padding: 6px 10px; min-height: 22px;
                }
                QPushButton {
                    background-color: #0a84ff; color: white; border: none;
                    border-radius: 6px; padding: 8px 16px; font-weight: 500;
                }
                QPushButton:hover { background-color: #409cff; }
            """
        else:
            dlg_style = """
                QDialog { background-color: #f5f5f7; }
                QLabel { color: #1d1d1f; font-size: 13px; }
                QLabel#title { font-size: 16px; font-weight: 600; }
                QLabel#info { font-size: 12px; color: #6e6e73; }
                QSpinBox {
                    background-color: #ffffff; color: #1d1d1f; border: 1px solid #d2d2d7;
                    border-radius: 6px; padding: 6px 10px; min-height: 22px;
                }
                QPushButton {
                    background-color: #007aff; color: white; border: none;
                    border-radius: 6px; padding: 8px 16px; font-weight: 500;
                }
                QPushButton:hover { background-color: #0066d6; }
            """
        
        dlg.setStyleSheet(dlg_style)
        layout = QVBoxLayout(dlg)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)
        
        title = QLabel("Configure Output Splitting")
        title.setObjectName("title")
        layout.addWidget(title)
        
        info = QLabel("How many records should each file contain?")
        info.setObjectName("info")
        layout.addWidget(info)
        
        spin = QSpinBox()
        spin.setMinimum(1)
        spin.setMaximum(10000)
        spin.setValue(20)  # Default to 20 records per file
        layout.addWidget(spin)
        
        hint = QLabel("For example: 20 = each file will have 20 rows of data")
        hint.setObjectName("info")
        layout.addWidget(hint)
        
        layout.addStretch()
        
        btn_row = QHBoxLayout()
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(dlg.reject)
        btn_row.addWidget(btn_cancel)
        
        btn_row.addStretch()
        
        btn_ok = QPushButton("OK")
        btn_ok.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_ok)
        
        layout.addLayout(btn_row)
        
        if dlg.exec_() == QDialog.Accepted:
            return spin.value()
        return None
    
    def _select_sheet_dialog(self, sheet_names):
        """Show a dialog to select which sheet to load."""
        from PyQt5.QtWidgets import QDialog, QListWidget, QListWidgetItem
        
        is_dark = self.manual_theme_override == "dark" or (self.manual_theme_override is None and system_dark_mode())
        
        if is_dark:
            bg, card_bg, border, text, accent = "#1e1e1e", "#2a2a2c", "#3a3a3c", "#ffffff", "#0a84ff"
        else:
            bg, card_bg, border, text, accent = "#f5f5f7", "#ffffff", "#d2d2d7", "#1d1d1f", "#007aff"
        
        dlg = QDialog(self)
        dlg.setWindowTitle("Select Sheet")
        dlg.resize(350, 300)
        dlg.setStyleSheet(f"""
            QDialog {{ background-color: {bg}; font-family: -apple-system, sans-serif; }}
            QLabel {{ color: {text}; font-size: 14px; font-weight: 600; }}
            QListWidget {{
                background-color: {card_bg};
                border: 1px solid {border};
                border-radius: 8px;
                color: {text};
                font-size: 13px;
                padding: 4px;
            }}
            QListWidget::item {{
                padding: 10px 12px;
                border-radius: 6px;
                margin: 2px 4px;
            }}
            QListWidget::item:hover {{ background-color: {border}; }}
            QListWidget::item:selected {{ background-color: {accent}; color: white; }}
            QPushButton {{
                background-color: {card_bg};
                color: {text};
                border: 1px solid {border};
                border-radius: 6px;
                padding: 8px 20px;
                font-size: 13px;
            }}
            QPushButton:hover {{ background-color: {border}; }}
            QPushButton#primary {{
                background-color: {accent};
                color: white;
                border: none;
            }}
        """)
        
        layout = QVBoxLayout(dlg)
        layout.setSpacing(16)
        layout.setContentsMargins(20, 20, 20, 20)
        
        layout.addWidget(QLabel("Select a sheet to load:"))
        
        sheet_list = QListWidget()
        for name in sheet_names:
            item = QListWidgetItem(name)
            sheet_list.addItem(item)
        sheet_list.setCurrentRow(0)
        layout.addWidget(sheet_list)
        
        btn_row = QHBoxLayout()
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(dlg.reject)
        btn_row.addWidget(btn_cancel)
        btn_row.addStretch()
        btn_select = QPushButton("Select")
        btn_select.setObjectName("primary")
        btn_select.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_select)
        layout.addLayout(btn_row)
        
        # Double-click to select
        sheet_list.itemDoubleClicked.connect(dlg.accept)
        
        if dlg.exec_() == QDialog.Accepted:
            current = sheet_list.currentItem()
            if current:
                return current.text()
        return None
    
    def _add_to_recent_files(self, path):
        """Add a file to the recent files list."""
        try:
            data = {}
            if os.path.exists(self.settings_path):
                with open(self.settings_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            
            recent = data.get("recent_files", [])
            if path in recent:
                recent.remove(path)
            recent.insert(0, path)
            data["recent_files"] = recent[:10]  # Keep only 10 recent files
            
            # Ensure directory exists
            settings_dir = os.path.dirname(self.settings_path)
            if settings_dir and not os.path.exists(settings_dir):
                os.makedirs(settings_dir, exist_ok=True)
            
            with open(self.settings_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            print(f"[DEBUG] Added to recent files: {path}")
        except Exception as e:
            print(f"[DEBUG] Error adding to recent files: {e}")
    
    def _clear_recent_files(self, list_widget):
        """Clear the recent files list."""
        try:
            if os.path.exists(self.settings_path):
                with open(self.settings_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                data["recent_files"] = []
                with open(self.settings_path, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2)
            
            list_widget.clear()
            item = QListWidgetItem("No recent files")
            item.setFlags(item.flags() & ~Qt.ItemIsSelectable)
            list_widget.addItem(item)
        except:
            pass
    
    def _reveal_settings_file(self):
        """Open the folder containing the settings file."""
        try:
            folder = os.path.dirname(self.settings_path)
            if sys.platform == "darwin":
                subprocess.run(["open", folder])
            elif sys.platform == "win32":
                subprocess.run(["explorer", folder])
            else:
                subprocess.run(["xdg-open", folder])
        except:
            QMessageBox.information(self, "Settings Location", f"Settings file:\n{self.settings_path}")



    def show_settings_dialog(self):
        from PyQt5.QtWidgets import QDialog, QFrame, QScrollArea, QSpinBox, QCheckBox
        
        # Detect theme
        is_dark = self.manual_theme_override == "dark" or (self.manual_theme_override is None and system_dark_mode())
        
        # Colors based on theme
        if is_dark:
            bg, card_bg, border, text, text2, accent = "#1e1e1e", "#2a2a2c", "#3a3a3c", "#ffffff", "#8e8e93", "#0a84ff"
        else:
            bg, card_bg, border, text, text2, accent = "#f5f5f7", "#ffffff", "#d2d2d7", "#1d1d1f", "#6e6e73", "#007aff"
        
        dlg = QDialog(self)
        dlg.setWindowTitle("Settings")
        dlg.resize(480, 580)
        dlg.setStyleSheet(f"""
            QDialog {{ background-color: {bg}; }}
            QLabel {{ color: {text}; font-size: 13px; background: transparent; }}
            QLabel#title {{ font-size: 20px; font-weight: 600; }}
            QLabel#section {{ font-size: 11px; font-weight: 600; color: {text2}; text-transform: uppercase; letter-spacing: 0.5px; }}
            QFrame#card {{ background-color: {card_bg}; border: 1px solid {border}; border-radius: 10px; }}
            QComboBox {{
                background-color: {card_bg}; color: {text}; border: 1px solid {border};
                border-radius: 6px; padding: 5px 10px; min-width: 130px; font-size: 13px;
            }}
            QSpinBox {{
                background-color: {card_bg}; color: {text}; border: 1px solid {border};
                border-radius: 6px; padding: 5px 10px; min-width: 70px; font-size: 13px;
            }}
            QCheckBox {{ color: {text}; font-size: 13px; spacing: 8px; }}
            QPushButton {{
                background-color: {card_bg}; color: {text}; border: 1px solid {border};
                border-radius: 6px; padding: 8px 16px; font-size: 13px; font-weight: 500;
            }}
            QPushButton:hover {{ background-color: {border}; }}
            QPushButton#primary {{ background-color: {accent}; color: white; border: none; }}
            QPushButton#primary:hover {{ background-color: {'#0077e6' if is_dark else '#0066d6'}; }}
            QScrollArea {{ border: none; background: transparent; }}
        """)
        
        main_layout = QVBoxLayout(dlg)
        main_layout.setSpacing(12)
        main_layout.setContentsMargins(24, 24, 24, 24)
        
        title = QLabel("Settings")
        title.setObjectName("title")
        main_layout.addWidget(title)
        
        # Scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_content = QWidget()
        layout = QVBoxLayout(scroll_content)
        layout.setSpacing(8)
        layout.setContentsMargins(0, 0, 8, 0)
        
        # === OUTPUT SECTION ===
        layout.addWidget(self._settings_section_label("OUTPUT"))
        card1 = self._settings_card()
        card1_layout = QVBoxLayout(card1)
        card1_layout.setContentsMargins(16, 12, 16, 12)
        card1_layout.setSpacing(10)
        
        # Resolution (synced)
        dd_res = QComboBox()
        dd_res.addItems(["240p", "360p", "480p", "720p", "1080p", "1440p", "2160p", "3840p"])
        current_res = self.dd_res.currentText()
        if current_res != "Custom…":
            idx = dd_res.findText(current_res)
            if idx >= 0: dd_res.setCurrentIndex(idx)
        card1_layout.addLayout(self._settings_row("Resolution", dd_res))
        
        # Format (synced)
        dd_fmt = QComboBox()
        dd_fmt.addItems(["PNG", "JPG", "WEBP"])
        dd_fmt.setCurrentText(self.dd_fmt.currentText())
        card1_layout.addLayout(self._settings_row("Format", dd_fmt))
        
        # JPG Quality (synced)
        dd_quality = QComboBox()
        dd_quality.addItems(["60 (Low)", "75 (Medium)", "90 (High)", "100 (Ultra)"])
        dd_quality.setCurrentText(self.dd_jpg_quality.currentText())
        card1_layout.addLayout(self._settings_row("JPG Quality", dd_quality))
        
        layout.addWidget(card1)
        
        # === SEARCH SECTION ===
        layout.addWidget(self._settings_section_label("SEARCH"))
        card2 = self._settings_card()
        card2_layout = QVBoxLayout(card2)
        card2_layout.setContentsMargins(16, 12, 16, 12)
        card2_layout.setSpacing(10)
        
        # Search Engine (synced)
        dd_browser = QComboBox()
        dd_browser.addItems(["Bing Images", "Google Images", "DuckDuckGo"])
        dd_browser.setCurrentText(self.dd_browser.currentText())
        card2_layout.addLayout(self._settings_row("Search Engine", dd_browser))
        
        # Theme Suffix (synced)
        dd_theme = QComboBox()
        dd_theme.addItems(["Portrait", "Wallpaper", "High Quality", "Photo", "Custom Theme..."])
        current_theme = self.dd_theme.currentText()
        idx = dd_theme.findText(current_theme)
        if idx >= 0: dd_theme.setCurrentIndex(idx)
        card2_layout.addLayout(self._settings_row("Theme Suffix", dd_theme))
        
        layout.addWidget(card2)
        
        # === PERFORMANCE SECTION ===
        layout.addWidget(self._settings_section_label("PERFORMANCE"))
        card3 = self._settings_card()
        card3_layout = QVBoxLayout(card3)
        card3_layout.setContentsMargins(16, 12, 16, 12)
        card3_layout.setSpacing(10)
        
        spin_threads = QSpinBox()
        spin_threads.setMinimum(2)
        spin_threads.setMaximum(20)
        spin_threads.setValue(6)
        card3_layout.addLayout(self._settings_row("Download Threads", spin_threads))
        
        spin_timeout = QSpinBox()
        spin_timeout.setMinimum(3)
        spin_timeout.setMaximum(30)
        spin_timeout.setValue(7)
        spin_timeout.setSuffix(" sec")
        card3_layout.addLayout(self._settings_row("Request Timeout", spin_timeout))
        
        layout.addWidget(card3)
        
        # === OUTPUT SPLITTING SECTION ===
        layout.addWidget(self._settings_section_label("OUTPUT SPLITTING"))
        card_split = self._settings_card()
        card_split_layout = QVBoxLayout(card_split)
        card_split_layout.setContentsMargins(16, 12, 16, 12)
        card_split_layout.setSpacing(10)
        
        # Enable split checkbox
        chk_split = QCheckBox("Split output into multiple files")
        chk_split.setChecked(self.split_enabled_setting)
        card_split_layout.addWidget(chk_split)
        
        # Records per file spinbox
        spin_records = QSpinBox()
        spin_records.setMinimum(1)
        spin_records.setMaximum(10000)
        spin_records.setValue(self.records_per_file_setting)
        spin_records.setEnabled(chk_split.isChecked())
        
        # Connect checkbox to enable/disable spinbox
        def toggle_records_spin(checked):
            spin_records.setEnabled(checked)
        chk_split.toggled.connect(toggle_records_spin)
        
        card_split_layout.addLayout(self._settings_row("Records per file", spin_records))
        
        hint_split = QLabel("Each output file will contain this many rows of data")
        hint_split.setObjectName("info")
        hint_split.setStyleSheet(f"color: {text2}; font-size: 11px;")
        card_split_layout.addWidget(hint_split)
        
        layout.addWidget(card_split)
        
        # === FILTERS SECTION ===
        layout.addWidget(self._settings_section_label("IMAGE FILTERS"))
        card4 = self._settings_card()
        card4_layout = QVBoxLayout(card4)
        card4_layout.setContentsMargins(16, 12, 16, 12)
        card4_layout.setSpacing(8)
        
        chk_portrait = QCheckBox("Prioritize portrait-oriented images")
        chk_portrait.setChecked(self.filter_portrait)
        card4_layout.addWidget(chk_portrait)
        
        chk_bw = QCheckBox("Filter out black & white images")
        chk_bw.setChecked(self.filter_bw)
        card4_layout.addWidget(chk_bw)
        
        chk_cartoon = QCheckBox("Filter out cartoon/graphic images")
        chk_cartoon.setChecked(self.filter_cartoon)
        card4_layout.addWidget(chk_cartoon)
        
        layout.addWidget(card4)
        layout.addStretch()
        
        scroll.setWidget(scroll_content)
        main_layout.addWidget(scroll)
        
        # Buttons
        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)
        
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(dlg.reject)
        btn_row.addWidget(btn_cancel)
        
        btn_row.addStretch()
        
        btn_apply = QPushButton("Apply")
        btn_apply.setObjectName("primary")
        def apply_settings():
            # Sync back to main UI
            res_idx = self.dd_res.findText(dd_res.currentText())
            if res_idx >= 0: self.dd_res.setCurrentIndex(res_idx)
            
            fmt_idx = self.dd_fmt.findText(dd_fmt.currentText())
            if fmt_idx >= 0: self.dd_fmt.setCurrentIndex(fmt_idx)
            
            quality_idx = self.dd_jpg_quality.findText(dd_quality.currentText())
            if quality_idx >= 0: self.dd_jpg_quality.setCurrentIndex(quality_idx)
            
            browser_idx = self.dd_browser.findText(dd_browser.currentText())
            if browser_idx >= 0: self.dd_browser.setCurrentIndex(browser_idx)
            
            theme_idx = self.dd_theme.findText(dd_theme.currentText())
            if theme_idx >= 0: self.dd_theme.setCurrentIndex(theme_idx)
            
            # Save filter settings
            self.filter_portrait = chk_portrait.isChecked()
            self.filter_bw = chk_bw.isChecked()
            self.filter_cartoon = chk_cartoon.isChecked()
            
            # Save split settings
            self.split_enabled_setting = chk_split.isChecked()
            self.records_per_file_setting = spin_records.value()
            
            dlg.accept()
        
        btn_apply.clicked.connect(apply_settings)
        btn_row.addWidget(btn_apply)
        
        main_layout.addLayout(btn_row)
        dlg.exec_()
    
    def _settings_section_label(self, text):
        lbl = QLabel(text)
        lbl.setObjectName("section")
        return lbl
    
    def _settings_card(self):
        from PyQt5.QtWidgets import QFrame
        card = QFrame()
        card.setObjectName("card")
        return card
    
    def _settings_row(self, label_text, widget):
        row = QHBoxLayout()
        row.addWidget(QLabel(label_text))
        row.addStretch()
        row.addWidget(widget)
        return row

    def _get_help_base_style(self, is_dark):
        """Generate CSS styles for help content based on theme"""
        if is_dark:
            bg_color = "#252525"
            text_color = "#ffffff"
            accent_color = "#0a84ff"
            header_bg = "#333333"
            table_border = "#404040"
            table_alt_bg = "#2a2a2a"
            code_bg = "#1e1e1e"
        else:
            bg_color = "#ffffff"
            text_color = "#1d1d1f"
            accent_color = "#0a84ff"
            header_bg = "#f5f5f7"
            table_border = "#d2d2d7"
            table_alt_bg = "#fafafa"
            code_bg = "#f0f0f0"
        
        return f"""
            body {{
                background-color: {bg_color};
                color: {text_color};
                font-family: -apple-system, BlinkMacSystemFont, 'SF Pro Text', 'Helvetica Neue', sans-serif;
                font-size: 13px;
                line-height: 1.6;
                padding: 20px;
                margin: 0;
            }}
            h1 {{
                color: {accent_color};
                font-size: 24px;
                font-weight: 700;
                margin-bottom: 20px;
                margin-top: 0;
                border-bottom: 2px solid {accent_color};
                padding-bottom: 10px;
            }}
            h2 {{
                color: {accent_color};
                font-size: 18px;
                font-weight: 600;
                margin-bottom: 15px;
                margin-top: 25px;
            }}
            h3 {{
                color: {text_color};
                font-size: 15px;
                font-weight: 600;
                margin-top: 20px;
                margin-bottom: 10px;
            }}
            h4 {{
                color: {text_color};
                font-size: 14px;
                font-weight: 600;
                margin-top: 15px;
                margin-bottom: 8px;
            }}
            p {{ margin-bottom: 12px; }}
            b, strong {{ color: {text_color}; }}
            a {{ color: {accent_color}; text-decoration: none; }}
            a:hover {{ text-decoration: underline; }}
            table {{
                border-collapse: collapse;
                width: 100%;
                margin: 15px 0;
                border-radius: 8px;
                overflow: hidden;
                border: 1px solid {table_border};
            }}
            th {{
                background-color: {header_bg};
                color: {text_color};
                padding: 10px 12px;
                text-align: left;
                font-weight: 600;
                border-bottom: 1px solid {table_border};
            }}
            td {{
                padding: 10px 12px;
                border-bottom: 1px solid {table_border};
                background-color: {bg_color};
            }}
            tr:nth-child(even) td {{
                background-color: {table_alt_bg};
            }}
            tr:last-child td {{ border-bottom: none; }}
            ul, ol {{ margin-left: 20px; padding-left: 0; margin-bottom: 15px; }}
            li {{ margin-bottom: 8px; }}
            code {{
                background: {code_bg};
                padding: 2px 6px;
                border-radius: 4px;
                font-family: 'SF Mono', Monaco, Consolas, monospace;
                font-size: 12px;
            }}
            pre {{
                background: {code_bg};
                padding: 12px;
                border-radius: 8px;
                overflow-x: auto;
                font-family: 'SF Mono', Monaco, Consolas, monospace;
                font-size: 12px;
                line-height: 1.4;
            }}
            .tip {{
                background: {table_alt_bg};
                border-left: 4px solid {accent_color};
                padding: 12px 15px;
                margin: 15px 0;
                border-radius: 0 8px 8px 0;
            }}
            .warning {{
                background: #fff3cd;
                border-left: 4px solid #ffc107;
                padding: 12px 15px;
                margin: 15px 0;
                border-radius: 0 8px 8px 0;
            }}
            .section-divider {{
                border-top: 1px solid {table_border};
                margin: 25px 0;
            }}
        """

    def show_help(self):
        from PyQt5.QtWidgets import QDialog, QTextBrowser, QStackedWidget, QListWidget, QListWidgetItem, QSplitter, QScrollArea
        from PyQt5.QtCore import QSize
        dlg = QDialog(self)
        dlg.setWindowTitle("Canvex — User Guide & Documentation")
        dlg.resize(950, 700)
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # Detect current theme for styling
        is_dark = self.manual_theme_override == "dark" or (self.manual_theme_override is None and system_dark_mode())
        
        # Get base style from helper method
        base_style = self._get_help_base_style(is_dark)
        
        # Theme-aware colors for sidebar
        if is_dark:
            bg_color = "#252525"
            text_color = "#ffffff"
            accent_color = "#0a84ff"
            table_border = "#404040"
            sidebar_bg = "#1e1e1e"
            sidebar_selected = "#0a84ff"
            sidebar_hover = "#333333"
        else:
            bg_color = "#ffffff"
            text_color = "#1d1d1f"
            accent_color = "#0a84ff"
            table_border = "#d2d2d7"
            sidebar_bg = "#f5f5f7"
            sidebar_selected = "#0a84ff"
            sidebar_hover = "#e8e8ed"
        
        # Create sidebar navigation (macOS style)
        splitter = QSplitter(Qt.Horizontal)
        layout.addWidget(splitter)
        
        # Sidebar
        sidebar = QListWidget()
        sidebar.setFixedWidth(200)
        sidebar.setIconSize(QSize(20, 20))
        sidebar.setSpacing(2)
        
        # Sidebar styling
        sidebar_style = f"""
            QListWidget {{
                background: {sidebar_bg};
                border: none;
                border-right: 1px solid {table_border};
                outline: none;
                padding: 10px 5px;
            }}
            QListWidget::item {{
                padding: 10px 15px;
                border-radius: 6px;
                margin: 2px 5px;
                color: {text_color};
            }}
            QListWidget::item:selected {{
                background: {sidebar_selected};
                color: white;
            }}
            QListWidget::item:hover:!selected {{
                background: {sidebar_hover};
            }}
        """
        sidebar.setStyleSheet(sidebar_style)
        
        # Add sidebar items - comprehensive navigation
        items = [
            ("📖", "Introduction"),
            ("🚀", "Getting Started"),
            ("🖥️", "Interface Overview"),
            ("📝", "Step-by-Step Guide"),
            ("⚙", "Configuration"),
            ("🔗", "Column Mappings"),
            ("🎛️", "Settings Panel"),
            ("✨", "New Features"),
            ("📁", "Output Files"),
            ("💡", "Tips & Tricks"),
            ("🔧", "Troubleshooting"),
            ("❓", "FAQ"),
        ]
        for icon, text in items:
            item = QListWidgetItem(f"{icon}  {text}")
            item.setSizeHint(QSize(180, 38))
            sidebar.addItem(item)
        
        splitter.addWidget(sidebar)
        
        # Content stack
        content_stack = QStackedWidget()
        content_stack.setStyleSheet(f"background: {bg_color};")
        splitter.addWidget(content_stack)
        
        # Helper to create scrollable content pages
        def create_page(html_content):
            page = QWidget()
            page_layout = QVBoxLayout(page)
            page_layout.setContentsMargins(0, 0, 0, 0)
            browser = QTextBrowser()
            browser.setOpenExternalLinks(True)
            browser.setHtml(f"<html><head><style>{base_style}</style></head><body>{html_content}</body></html>")
            page_layout.addWidget(browser)
            return page
        
        # ===== Page 1: Introduction =====
        intro_html = f"""
        <h1>🖼️ Canvex User Guide</h1>
        <p style="font-size: 15px; color: {text_color}; margin-bottom: 20px;">
            <em>Automatically search and insert images into Excel files</em>
        </p>
        
        <h2>What is Canvex?</h2>
        <p>Canvex is a powerful desktop application that <b>automatically searches the web for images</b> 
        based on text in your Excel spreadsheet and <b>inserts them directly into a new Excel file</b>.</p>
        
        <h3>Perfect for:</h3>
        <ul>
            <li>📸 Creating employee directories with headshots</li>
            <li>🎬 Building cast lists with actor photos</li>
            <li>🏢 Generating product catalogs with images</li>
            <li>📊 Any data visualization requiring images</li>
        </ul>
        
        <h2>Key Features</h2>
        <table>
            <tr><th>Feature</th><th>Description</th></tr>
            <tr><td><b>🔍 Multi-Engine Search</b></td><td>Search using Bing, Google, or DuckDuckGo</td></tr>
            <tr><td><b>🎨 Smart Filtering</b></td><td>Automatically removes low-quality, B&W, and cartoon images</td></tr>
            <tr><td><b>⚡ Parallel Processing</b></td><td>Downloads multiple images simultaneously</td></tr>
            <tr><td><b>💾 Auto-Save Settings</b></td><td>Your preferences are remembered between sessions</td></tr>
            <tr><td><b>🌓 Theme Support</b></td><td>Light, Dark, or System-following themes</td></tr>
            <tr><td><b>📐 Flexible Resolution</b></td><td>From 240p to 4K, or custom values</td></tr>
            <tr><td><b>🎯 Portrait Priority</b></td><td>Prefers portrait-oriented images for headshots</td></tr>
        </table>
        
        <h2>System Requirements</h2>
        <table>
            <tr><th>Requirement</th><th>Specification</th></tr>
            <tr><td><b>Operating System</b></td><td>macOS 10.14+ or Windows 10+</td></tr>
            <tr><td><b>Internet Connection</b></td><td>Required for image searches</td></tr>
            <tr><td><b>Chrome Browser</b></td><td>Must be installed (used for web scraping)</td></tr>
            <tr><td><b>RAM</b></td><td>4GB minimum, 8GB recommended</td></tr>
            <tr><td><b>Storage</b></td><td>100MB for app + space for output files</td></tr>
        </table>
        """
        content_stack.addWidget(create_page(intro_html))
        
        # ===== Page 2: Getting Started =====
        getting_started_html = f"""
        <h1>🚀 Getting Started</h1>
        
        <h2>Quick Start (5 Minutes)</h2>
        <p style="font-size: 15px; text-align: center; padding: 15px; background: {'#2a2a2a' if is_dark else '#f0f0f5'}; border-radius: 8px;">
            <b>1. Load Excel</b> → <b>2. Set Theme</b> → <b>3. Add Mappings</b> → <b>4. Start</b> → <b>5. Save</b>
        </p>
        
        <h2>Step 1: Load Your Excel File</h2>
        <h3>Option A: Click to Browse</h3>
        <ol>
            <li>Click the <span style="color:{accent_color}"><b>"Select Excel File"</b></span> button</li>
            <li>Navigate to your <code>.xlsx</code> file</li>
            <li>Click <b>Open</b></li>
        </ol>
        
        <h3>Option B: Drag and Drop</h3>
        <ol>
            <li>Open your file explorer/finder</li>
            <li>Drag the <code>.xlsx</code> file onto the Canvex window</li>
            <li>Release to load</li>
        </ol>
        
        <div class="tip">
            <b>💡 Sheet Selection:</b> If your Excel has multiple sheets, a dialog will appear to let you choose which sheet to process.
        </div>
        
        <div class="tip">
            <b>💡 Tip:</b> Only <code>.xlsx</code> files are supported. Convert older <code>.xls</code> files first.
        </div>
        
        <h2>Step 2: Configure Image Settings</h2>
        <p>After loading your file, configure the following:</p>
        <ul>
            <li><b>Image Theme:</b> Choose a search style (e.g., "headshot portrait" for faces)</li>
            <li><b>Search Browser:</b> Select Bing (recommended), Google, or DuckDuckGo</li>
            <li><b>Resolution:</b> Select image quality (720p recommended)</li>
            <li><b>Format:</b> PNG (best quality), JPG (smaller size), or WEBP</li>
        </ul>
        
        <h2>Step 3: Set Up Column Mappings</h2>
        <ol>
            <li>Click <b>"+ Add Mapping"</b> button</li>
            <li>Select <b>Input Column</b> (column with search terms)</li>
            <li>Select <b>Output Column</b> (where images go) or create new</li>
            <li>Repeat for additional mappings if needed</li>
        </ol>
        
        <h2>Step 4: Start Processing</h2>
        <ol>
            <li>Click <span style="color:#34c759"><b>"▶ Start Processing"</b></span></li>
            <li>Choose where to save the output file</li>
            <li>Enter a filename and click <b>Save</b></li>
            <li>Wait for processing to complete</li>
        </ol>
        
        <h2>Step 5: Review Output</h2>
        <p>When complete, you'll be asked if you want to open the output file. Your output includes:</p>
        <ul>
            <li><b>your_output.xlsx</b> — Excel with images inserted</li>
            <li><b>your_output_log.txt</b> — Processing log (always created)</li>
        </ul>
        """
        content_stack.addWidget(create_page(getting_started_html))
        
        # ===== Page 3: Interface Overview =====
        interface_html = f"""
        <h1>🖥️ Main Interface Overview</h1>
        
        <h2>Toolbar Buttons</h2>
        <table>
            <tr><th>Button</th><th>Function</th></tr>
            <tr><td><b>📁 File</b></td><td>Open files, view recent files, reveal settings location</td></tr>
            <tr><td><b>⚙️ Settings</b></td><td>Configure search engine, filters, resolution, and format</td></tr>
            <tr><td><b>❓ Help</b></td><td>View this user guide</td></tr>
            <tr><td><b>ℹ️ About</b></td><td>Application version and contact information</td></tr>
            <tr><td><b>🎨 Theme</b></td><td>Switch between Light, Dark, or System theme</td></tr>
        </table>
        
        <h2>Configuration Area</h2>
        <table>
            <tr><th>Component</th><th>Description</th></tr>
            <tr><td><b>Image Theme</b></td><td>Dropdown with 8 presets + custom option</td></tr>
            <tr><td><b>Search Browser</b></td><td>Choose between Bing, Google, or DuckDuckGo</td></tr>
            <tr><td><b>Resolution</b></td><td>240p to 4K, or enter custom value</td></tr>
            <tr><td><b>Format</b></td><td>PNG, JPG (with quality), or WEBP</td></tr>
        </table>
        
        <h2>Mapping Table</h2>
        <p>The mapping table shows your column mappings with these columns:</p>
        <table>
            <tr><th>Column</th><th>Description</th></tr>
            <tr><td><b>#</b></td><td>Row number (auto-generated)</td></tr>
            <tr><td><b>Input Column</b></td><td>Column containing search terms</td></tr>
            <tr><td><b>Output Column</b></td><td>Where images will be inserted</td></tr>
            <tr><td><b>New Column Name</b></td><td>Name for new columns (if creating)</td></tr>
            <tr><td><b>🗑️</b></td><td>Delete this mapping row</td></tr>
        </table>
        
        <h2>Progress Area</h2>
        <p>During processing, you'll see:</p>
        <ul>
            <li><b>Progress bar</b> showing overall completion percentage</li>
            <li><b>Cancel button</b> to safely stop processing at any time</li>
        </ul>
        """
        content_stack.addWidget(create_page(interface_html))
        
        # ===== Page 4: Step-by-Step Guide =====
        stepbystep_html = f"""
        <h1>📝 Step-by-Step Workflow</h1>
        
        <h2>Image Theme Options</h2>
        <table>
            <tr><th>Theme</th><th>Best For</th></tr>
            <tr><td>headshot portrait closeup face</td><td>Professional headshots, ID photos</td></tr>
            <tr><td>cinematic lighting portrait</td><td>Dramatic, artistic portraits</td></tr>
            <tr><td>studio headshot clean background</td><td>Corporate/LinkedIn style photos</td></tr>
            <tr><td>dramatic portrait closeup</td><td>High-contrast artistic shots</td></tr>
            <tr><td>smiling closeup face</td><td>Friendly, approachable photos</td></tr>
            <tr><td>full body portrait</td><td>Full-length photos</td></tr>
            <tr><td>natural daylight portrait</td><td>Outdoor, natural lighting</td></tr>
            <tr><td>magazine cover portrait</td><td>High-fashion style</td></tr>
            <tr><td>Custom Theme...</td><td>Enter your own search keywords</td></tr>
        </table>
        
        <h2>Search Browser Comparison</h2>
        <table>
            <tr><th>Engine</th><th>Speed</th><th>Reliability</th><th>Notes</th></tr>
            <tr><td><b>Bing Images</b></td><td>●●●●</td><td>●●●●</td><td>⭐ Recommended - fastest and most reliable</td></tr>
            <tr><td><b>Google Images</b></td><td>●●○○</td><td>●●●○</td><td>Alternative results, may be slower</td></tr>
            <tr><td><b>DuckDuckGo</b></td><td>●●●○</td><td>●●●○</td><td>Privacy-focused, good backup option</td></tr>
        </table>
        
        <div class="tip">
            <b>📝 Note:</b> If Google returns no results, Canvex automatically tries Bing as a fallback.
        </div>
        
        <h2>Resolution Guide</h2>
        <table>
            <tr><th>Setting</th><th>Pixels</th><th>Use Case</th><th>Speed</th></tr>
            <tr><td>240p</td><td>240</td><td>Thumbnails</td><td>●●●● Fastest</td></tr>
            <tr><td>360p</td><td>360</td><td>Small previews</td><td>●●●○</td></tr>
            <tr><td><b>480p</b></td><td>480</td><td>Standard docs</td><td>●●○○</td></tr>
            <tr><td><b>720p</b></td><td>720</td><td>⭐ Recommended</td><td>●●○○</td></tr>
            <tr><td>1080p</td><td>1080</td><td>High-quality</td><td>●○○○</td></tr>
            <tr><td>1440p</td><td>1440</td><td>Large displays</td><td>○○○○ Slowest</td></tr>
            <tr><td>2160p</td><td>2160</td><td>4K quality</td><td>○○○○</td></tr>
            <tr><td>Custom...</td><td>240-4000</td><td>Your choice</td><td>Varies</td></tr>
        </table>
        
        <h2>Image Format Comparison</h2>
        <table>
            <tr><th>Format</th><th>Quality</th><th>File Size</th><th>Transparency</th><th>Best For</th></tr>
            <tr><td><b>PNG</b></td><td>★★★ Best</td><td>Large</td><td>✓ Yes</td><td>Quality priority</td></tr>
            <tr><td><b>JPG</b></td><td>★★ Good</td><td>Medium</td><td>✗ No</td><td>Balanced</td></tr>
            <tr><td><b>WEBP</b></td><td>★★ Good</td><td>Smallest</td><td>✓ Yes</td><td>Modern apps</td></tr>
        </table>
        
        <h3>JPG Quality Options</h3>
        <ul>
            <li><b>60 (Low)</b> — Smallest files, visible compression</li>
            <li><b>75 (Medium)</b> — Good balance</li>
            <li><b>90 (High)</b> — High quality, moderate size</li>
            <li><b>100 (Ultra)</b> — Maximum quality, larger files</li>
        </ul>
        """
        content_stack.addWidget(create_page(stepbystep_html))
        
        # ===== Page 5: Configuration =====
        config_html = f"""
        <h1>⚙ Configuration Options</h1>
        
        <h2>Image Theme Details</h2>
        <p>The theme affects how search queries are constructed:</p>
        <pre>Search Query = [Cell Value] + [Theme]
Example: "Tom Hanks" + "headshot portrait closeup face"</pre>
        
        <h3>Custom Theme Tips</h3>
        <ul>
            <li>Use descriptive words: <code>professional</code>, <code>corporate</code>, <code>natural</code></li>
            <li>Add style modifiers: <code>high quality</code>, <code>HD</code>, <code>portrait</code></li>
            <li>Specify background: <code>white background</code>, <code>studio</code></li>
        </ul>
        
        <h2>Browser Selection Details</h2>
        <table>
            <tr><th>Feature</th><th>Bing</th><th>Google</th><th>DuckDuckGo</th></tr>
            <tr><td>Speed</td><td>●●●●</td><td>●●○○</td><td>●●●○</td></tr>
            <tr><td>Reliability</td><td>●●●●</td><td>●●●○</td><td>●●●○</td></tr>
            <tr><td>Image Quality</td><td>●●●●</td><td>●●●●</td><td>●●●○</td></tr>
            <tr><td>Rate Limiting</td><td>Low</td><td>Medium</td><td>Low</td></tr>
            <tr><td>Fallback</td><td>—</td><td>Bing</td><td>—</td></tr>
        </table>
        
        <h2>Settings Persistence</h2>
        <p>All settings are <b>automatically saved</b> and restored on next launch:</p>
        <ul>
            <li>Selected theme and custom theme text</li>
            <li>Search browser preference</li>
            <li>Resolution and format settings</li>
            <li>Image filters (portrait, B&W, cartoon)</li>
            <li>UI theme preference (light/dark/system)</li>
            <li>Last used directory</li>
            <li>Column mappings (for same file structure)</li>
            <li>Recent files list</li>
        </ul>
        
        <h3>Settings File Location</h3>
        <table>
            <tr><th>Platform</th><th>Location</th></tr>
            <tr><td><b>macOS</b></td><td>~/Library/Application Support/Canvex/</td></tr>
            <tr><td><b>Windows</b></td><td>%APPDATA%/Canvex/</td></tr>
        </table>
        """
        content_stack.addWidget(create_page(config_html))
        
        # ===== Page 6: Column Mappings =====
        mappings_html = f"""
        <h1>🔗 Column Mappings</h1>
        
        <h2>Understanding Mappings</h2>
        <p>Column mappings tell Canvex which columns contain search terms and where to put the images.</p>
        
        <table>
            <tr><th>Field</th><th>Description</th><th>Example</th></tr>
            <tr><td><b>Input Column</b></td><td>Column with search text</td><td>actor_name</td></tr>
            <tr><td><b>Output Column</b></td><td>Where to insert images</td><td>actor_image</td></tr>
            <tr><td><b>New Column Name</b></td><td>For new columns only</td><td>photo</td></tr>
        </table>
        
        <h2>Creating New Columns</h2>
        <ol>
            <li>In <b>Output Column</b>, select <b>"Create New Column..."</b></li>
            <li>A text field appears</li>
            <li>Enter the <b>new column name</b></li>
            <li>The new column is added to the right of existing data</li>
        </ol>
        
        <h2>Multiple Mappings Example</h2>
        <p><b>Input Excel:</b></p>
        <table>
            <tr><th>lead_actor</th><th>supporting_actor</th><th>director</th></tr>
            <tr><td>Tom Cruise</td><td>Val Kilmer</td><td>Tony Scott</td></tr>
            <tr><td>Keanu Reeves</td><td>Laurence Fishburne</td><td>The Wachowskis</td></tr>
        </table>
        
        <p><b>Mappings:</b></p>
        <ol>
            <li><code>lead_actor</code> → <code>lead_photo</code> (new)</li>
            <li><code>supporting_actor</code> → <code>support_photo</code> (new)</li>
            <li><code>director</code> → <code>director_photo</code> (new)</li>
        </ol>
        
        <p><b>Result:</b> Output Excel has 6 columns with images in the last 3.</p>
        
        <h2>Column Naming Best Practices</h2>
        <table>
            <tr><th>Good ✓</th><th>Avoid ✗</th></tr>
            <tr><td>actor_photo</td><td>photo 1</td></tr>
            <tr><td>employee_headshot</td><td>image</td></tr>
            <tr><td>product_img</td><td>column_A</td></tr>
        </table>
        """
        content_stack.addWidget(create_page(mappings_html))
        
        # ===== Page 7: Settings Panel =====
        settings_html = f"""
        <h1>🎛️ Settings Panel</h1>
        <p>Access via <b>Settings</b> button in the toolbar.</p>
        
        <h2>Image Filters</h2>
        <table>
            <tr><th>Filter</th><th>Effect</th><th>Recommended For</th></tr>
            <tr><td><b>Prioritize portrait images</b></td><td>Prefers taller-than-wide images</td><td>Headshots, portraits</td></tr>
            <tr><td><b>Filter out B&W images</b></td><td>Excludes grayscale images</td><td>Modern, colorful photos</td></tr>
            <tr><td><b>Filter out cartoon images</b></td><td>Excludes illustrations/graphics</td><td>Real photographs only</td></tr>
        </table>
        
        <h2>Filter Recommendation Matrix</h2>
        <table>
            <tr><th>Use Case</th><th>Portrait</th><th>B&W Filter</th><th>Cartoon Filter</th></tr>
            <tr><td>Professional headshots</td><td>✓ On</td><td>✓ On</td><td>✓ On</td></tr>
            <tr><td>Product photos</td><td>✗ Off</td><td>✓ On</td><td>✓ On</td></tr>
            <tr><td>Artistic portraits</td><td>✓ On</td><td>✗ Off</td><td>✓ On</td></tr>
            <tr><td>Character illustrations</td><td>✗ Off</td><td>✗ Off</td><td>✗ Off</td></tr>
        </table>
        
        <h2>Output Section</h2>
        <ul>
            <li><b>Resolution:</b> Quick access to resolution selector</li>
            <li><b>Format:</b> PNG/JPG/WEBP format selector</li>
            <li><b>JPG Quality:</b> Quality slider when JPG is selected</li>
        </ul>
        
        <h2>Search Section</h2>
        <ul>
            <li><b>Search Engine:</b> Bing/Google/DuckDuckGo</li>
            <li><b>Theme Suffix:</b> Default theme for new sessions</li>
        </ul>
        
        <h2>Performance Section</h2>
        <ul>
            <li><b>Download Threads:</b> Number of parallel downloads (2-20)</li>
            <li><b>Request Timeout:</b> Seconds before download fails (3-30)</li>
        </ul>
        """
        content_stack.addWidget(create_page(settings_html))
        
        # ===== Page 8: New Features (December 2025) =====
        newfeatures_html = f"""
        <h1>✨ New Features & Enhancements</h1>
        <p style="font-size: 14px; color: {text_color}; font-style: italic;">Latest updates from December 2025</p>
        
        <h2>📂 Session Persistence</h2>
        <p>Canvex now remembers everything about your work sessions!</p>
        
        <h3>Auto-Save Last Directory</h3>
        <p>When you open an Excel file, Canvex remembers that folder. Next time you open a file, it starts in the same location. No more browsing through folders!</p>
        
        <h3>Recent Files List</h3>
        <p>Click <b>File</b> → <b>Recent Files</b> to see your last 10 opened Excel files. File indicators show:</p>
        <ul>
            <li><b>✓</b> File still exists and is ready to open</li>
            <li><b>✗</b> File was deleted (remove from list)</li>
        </ul>
        <p>Just click any file to open it instantly!</p>
        
        <h3>Automatic Settings Backup</h3>
        <p>Your settings are automatically saved:</p>
        <ul>
            <li>🎨 Theme preference (Light/Dark)</li>
            <li>📐 Resolution setting</li>
            <li>🔍 Search browser choice</li>
            <li>🎯 Image filters & theme</li>
            <li>📂 Last used directory</li>
        </ul>
        <p>All these settings load automatically when you start the app!</p>
        
        <h2>🔄 Smart Mapping Management</h2>
        <p>Never recreate the same mappings twice!</p>
        
        <h3>Automatic Mapping History</h3>
        <p>Every time you start processing, your mappings are automatically saved. The last <b>5 configurations</b> are kept with timestamps, so you can quickly restore previous setups.</p>
        
        <h3>Previous Mappings Dialog</h3>
        <p>Click <b>File</b> → <b>Load Previous Mappings</b> to see all saved configurations.</p>
        <table>
            <tr><th>Feature</th><th>Description</th></tr>
            <tr><td><b>📋 Browse History</b></td><td>See all 5 previous mapping configurations with dates and times</td></tr>
            <tr><td><b>👁️ Live Preview</b></td><td>Click any mapping to see which columns are mapped</td></tr>
            <tr><td><b>⚡ Quick Load</b></td><td>Double-click or click "Load Selected" to restore that configuration</td></tr>
            <tr><td><b>🗑️ Reset Option</b></td><td>Clear all mappings with a single click (with confirmation)</td></tr>
        </table>
        
        <h3>Smart Column Detection</h3>
        <p>If you load a previous mapping and some columns have been deleted from your Excel file, Canvex automatically detects this and switches those columns to "Create New Column..." mode. No errors, no confusion!</p>
        
        <h2>🎯 Enhanced Column Mapping</h2>
        
        <h3>Better Delete Functionality</h3>
        <p>Delete buttons now work perfectly, even when you load mappings from history. Click the trash icon (✕) next to any mapping to remove it instantly. Remaining rows renumber automatically.</p>
        
        <h3>Smart Text Field Visibility</h3>
        <p>The text field for new column names only appears when you select "Create New Column..." — keeping the interface clean and uncluttered. When you switch to another option, the field automatically hides and clears.</p>
        
        <h2>🎨 Visual Improvements</h2>
        
        <h3>Native Theme Integration</h3>
        <p>Dropdowns now follow your system theme automatically:</p>
        <ul>
            <li>In <b>Light mode</b>, dropdowns use light backgrounds with dark text</li>
            <li>In <b>Dark mode</b>, dropdowns use dark backgrounds with light text</li>
            <li>Selected values are now clearly visible and easy to read</li>
            <li>Native macOS rendering for authentic appearance</li>
        </ul>
        
        <h3>List Hover Effects</h3>
        <p>Mapping lists now have smooth, interactive hover effects. Move your mouse over any item to see it highlight, giving you clear visual feedback.</p>
        <ul>
            <li><b>Dark Theme:</b> Items highlight with a subtle gray background</li>
            <li><b>Light Theme:</b> Items highlight with a light gray background</li>
            <li>Smooth transitions make the interface feel responsive</li>
        </ul>
        
        <h2>💾 Where Settings Are Stored</h2>
        <p>All your settings and history are stored in:</p>
        <table>
            <tr><th>Platform</th><th>Location</th></tr>
            <tr><td><b>macOS</b></td><td><code>~/Library/Application Support/Canvex/canva_last_settings.json</code></td></tr>
            <tr><td><b>Windows</b></td><td><code>%APPDATA%/Canvex/canva_last_settings.json</code></td></tr>
        </table>
        <p>The file is created automatically if it doesn't exist. You can view it via <b>File</b> → <b>Show Settings Location</b>.</p>
        
        <h2>🎯 Use Cases for New Features</h2>
        
        <h3>Quick Recurring Tasks</h3>
        <p><b>Scenario:</b> You process employee headshots every week with the same setup.<br/>
        <b>Solution:</b> Your mappings are automatically saved. Just click "Load Previous Mappings" and load last week's configuration instantly!</p>
        
        <h3>Multiple Projects</h3>
        <p><b>Scenario:</b> You switch between different projects (actors, products, employees).<br/>
        <b>Solution:</b> Each time you finish a project, the mappings are saved. When you return to an old project file, load the previous configuration that matches.</p>
        
        <h3>Team Usage</h3>
        <p><b>Scenario:</b> Your team shares computers and uses Canvex.<br/>
        <b>Solution:</b> Each user's recent files and settings are saved separately, making it easy for anyone to pick up where they left off.</p>
        """
        content_stack.addWidget(create_page(newfeatures_html))
        
        # ===== Page 9: Output Files =====
        output_html = f"""
        <h1>📁 Output Files</h1>
        
        <h2>Excel Output Structure</h2>
        <p>The output Excel file contains:</p>
        <ol>
            <li><b>All original data</b> from the input file</li>
            <li><b>New image columns</b> based on your mappings</li>
            <li><b>Images embedded</b> directly in cells</li>
        </ol>
        
        <h3>Image Properties</h3>
        <table>
            <tr><th>Property</th><th>Value</th></tr>
            <tr><td>Scale</td><td>20% of original size</td></tr>
            <tr><td>Position</td><td>Anchored to cell</td></tr>
            <tr><td>Row Height</td><td>120 pixels (auto-set)</td></tr>
            <tr><td>Column Width</td><td>22 characters (auto-set)</td></tr>
        </table>
        
        <h2>Log Files</h2>
        <table>
            <tr><th>File</th><th>Contents</th><th>Created</th></tr>
            <tr><td><b>output.xlsx</b></td><td>Excel with images</td><td>Always</td></tr>
            <tr><td><b>output_log.txt</b></td><td>Processing log</td><td>Always</td></tr>
            <tr><td><b>output_ERROR_log.txt</b></td><td>Error details + stack trace</td><td>Only on errors</td></tr>
        </table>
        
        <h3>Sample Log Output</h3>
        <pre>[START] 2025-01-15 10:30:00
[LOG] Theme: headshot portrait closeup face
[LOG] Search Browser: Bing Images
[LOG] Resolution: 720px
[SEARCH] Tom Hanks
[URLS] (Bing Images) 24 found: [...]
[SEARCH] Brad Pitt
[URLS] (Bing Images) 24 found: [...]

Time taken: 0h 5m 23s</pre>
        """
        content_stack.addWidget(create_page(output_html))
        
        # ===== Page 9: Tips & Tricks =====
        tips_html = f"""
        <h1>💡 Tips & Best Practices</h1>
        
        <h2>For Best Image Results</h2>
        <table>
            <tr><td width="30">✓</td><td><b>Use specific search terms</b> — "John Smith CEO Microsoft" works better than just "John Smith"</td></tr>
            <tr><td>✓</td><td><b>Choose appropriate themes</b> — Match theme to content type</td></tr>
            <tr><td>✓</td><td><b>Enable all filters for headshots</b> — Removes unwanted image types</td></tr>
            <tr><td>✓</td><td><b>Start with 720p</b> — Good balance of quality and speed</td></tr>
            <tr><td>✓</td><td><b>Use PNG format</b> — Best quality, no compression artifacts</td></tr>
        </table>
        
        <h2>For Faster Processing</h2>
        <table>
            <tr><td width="30">▶</td><td>Use <b>Bing Images</b> — Fastest and most reliable</td></tr>
            <tr><td>▶</td><td>Lower resolution = faster processing</td></tr>
            <tr><td>▶</td><td>Stable internet connection avoids timeout retries</td></tr>
            <tr><td>▶</td><td>Close other browsers for more resources</td></tr>
        </table>
        
        <h2>For Large Files (1000+ rows)</h2>
        <ol>
            <li><b>Process in batches</b> — Split into smaller files</li>
            <li><b>Use lower resolution</b> — 480p is sufficient for previews</li>
            <li><b>Choose JPG format</b> — Smaller file sizes</li>
            <li><b>Monitor progress</b> — Cancel if stuck</li>
        </ol>
        
        <h2>Pro Tips</h2>
        <div class="tip">
            <b>💡 Specific Names:</b> Add context like job title, company, or location to get better matches for common names.
        </div>
        
        <div class="tip">
            <b>💡 Custom Themes:</b> Create custom themes for specific industries — e.g., "professional linkedin corporate headshot" for business profiles.
        </div>
        
        <div class="tip">
            <b>💡 Batch Processing:</b> For very large files, process 500 rows at a time to avoid memory issues.
        </div>
        """
        content_stack.addWidget(create_page(tips_html))
        
        # ===== Page 10: Troubleshooting =====
        troubleshooting_html = f"""
        <h1>🔧 Troubleshooting</h1>
        
        <h2>Common Issues</h2>
        
        <h3>❌ "No images found"</h3>
        <p><b>Causes:</b> Search term too vague, name misspelled, person/item not well-known</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Make search terms more specific</li>
            <li>Try a different search engine</li>
            <li>Simplify the theme</li>
            <li>Check spelling in Excel</li>
        </ul>
        
        <h3>❌ "Wrong images appearing"</h3>
        <p><b>Causes:</b> Common name (e.g., "John Smith"), theme not matching content</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Add context: "John Smith actor" or "John Smith CEO"</li>
            <li>Try a different theme</li>
            <li>Use custom theme with specific keywords</li>
        </ul>
        
        <h3>❌ "Processing is very slow"</h3>
        <p><b>Causes:</b> High resolution selected, slow internet, many rows</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Lower resolution to 480p or 720p</li>
            <li>Check internet speed</li>
            <li>Process in smaller batches</li>
        </ul>
        
        <h3>❌ "App appears frozen"</h3>
        <p><b>Causes:</b> Chrome/Selenium starting up, large batch processing</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Wait 30 seconds — Selenium needs time to start</li>
            <li>Check task manager — If CPU active, processing continues</li>
            <li>If truly frozen, force quit and check _ERROR_log.txt</li>
        </ul>
        
        <h2>Error Messages</h2>
        <table>
            <tr><th>Error</th><th>Meaning</th><th>Solution</th></tr>
            <tr><td>Chrome not found</td><td>ChromeDriver issue</td><td>Install/update Chrome browser</td></tr>
            <tr><td>Connection timeout</td><td>Network issue</td><td>Check internet, try again</td></tr>
            <tr><td>Permission denied</td><td>File locked</td><td>Close the Excel file</td></tr>
            <tr><td>Out of memory</td><td>Too many images</td><td>Process smaller batches</td></tr>
        </table>
        """
        content_stack.addWidget(create_page(troubleshooting_html))
        
        # ===== Page 11: FAQ =====
        faq_html = f"""
        <h1>❓ Frequently Asked Questions</h1>
        
        <h2>General Questions</h2>
        
        <h3>Q: What file formats are supported?</h3>
        <p>A: Input must be <code>.xlsx</code> (Excel 2007+). Output is always <code>.xlsx</code>.</p>
        
        <h3>Q: Can I process multiple Excel files at once?</h3>
        <p>A: No, process one file at a time. For batch processing, run Canvex multiple times.</p>
        
        <h3>Q: Are my images saved locally?</h3>
        <p>A: Yes, images are embedded directly in the output Excel file. Temporary files are deleted after processing.</p>
        
        <h3>Q: Does Canvex work offline?</h3>
        <p>A: No, internet connection is required for image searches.</p>
        
        <h2>Technical Questions</h2>
        
        <h3>Q: Why does Canvex need Chrome?</h3>
        <p>A: Canvex uses Selenium with Chrome to scrape image search results. This provides more reliable results than API-only approaches.</p>
        
        <h3>Q: Where are settings saved?</h3>
        <p>A:</p>
        <ul>
            <li><b>macOS:</b> ~/Library/Application Support/Canvex/canva_last_settings.json</li>
            <li><b>Windows:</b> %APPDATA%/Canvex/canva_last_settings.json</li>
        </ul>
        
        <h3>Q: Can I customize the image size in Excel?</h3>
        <p>A: Currently, images are inserted at 20% scale with fixed row height (120px). For different sizes, edit the output in Excel.</p>
        
        <h2>Performance Questions</h2>
        
        <h3>Q: How long does processing take?</h3>
        <p>A: Depends on rows and resolution. Typical: ~2-5 seconds per row at 720p.</p>
        
        <h3>Q: Can I process 10,000+ rows?</h3>
        <p>A: Technically yes, but recommended to split into batches of 500-1000 for reliability.</p>
        
        <h3>Q: Does resolution affect processing time?</h3>
        <p>A: Yes. Higher resolution = larger downloads = slower processing.</p>
        
        <div class="section-divider"></div>
        
        <h2>Contact & Support</h2>
        <p><b>Publisher:</b> Kunal Pagariya</p>
        <p><b>Email:</b> <a href="mailto:kunal.pagariya@outlook.com">kunal.pagariya@outlook.com</a></p>
        <p><b>Version:</b> 1.0</p>
        <p style="color: {text_color}; opacity: 0.7;">© 2025 Kunal Pagariya</p>
        """
        content_stack.addWidget(create_page(faq_html))
        
        # Connect sidebar to content stack
        sidebar.currentRowChanged.connect(content_stack.setCurrentIndex)
        sidebar.setCurrentRow(0)
        
        # Set splitter proportions
        splitter.setSizes([200, 750])
        
        dlg.exec_()

    def show_about(self):
        from PyQt5.QtWidgets import QDialog
        dlg = QDialog(self)
        dlg.setWindowTitle("About Canvex")
        dlg.resize(400, 350)
        layout = QVBoxLayout(dlg)
        layout.setSpacing(15)
        
        # Logo/Splash image at top
        logo_label = QLabel()
        logo_label.setAlignment(Qt.AlignCenter)
        splash_path = resource_path("splash.png")
        if os.path.exists(splash_path):
            logo_pixmap = QPixmap(splash_path)
            if not logo_pixmap.isNull():
                # Scale to fit nicely in dialog
                logo_pixmap = logo_pixmap.scaledToWidth(200, Qt.SmoothTransformation)
                logo_label.setPixmap(logo_pixmap)
        layout.addWidget(logo_label)
        
        # About text
        about_text = (
            "<b style='font-size:18px;'></b><br>"
            "<span style='color:#666;'>Image Excel Inserter</span><br><br>"
            "Version: <b>1.0</b><br>"
            "Publisher: <b>Kunal Pagariya</b><br>"
            "Contact: <a href='mailto:kunal.pagariya@outlook.com'>kunal.pagariya@outlook.com</a><br><br>"
            "<span style='color:#888;'>© 2025 Kunal Pagariya</span>"
        )
        lbl = QLabel(about_text)
        lbl.setTextFormat(Qt.RichText)
        lbl.setAlignment(Qt.AlignCenter)
        lbl.setOpenExternalLinks(True)
        layout.addWidget(lbl)
        
        layout.addStretch()
        
        # OK button
        btn_ok = QPushButton("OK")
        btn_ok.setFixedWidth(100)
        btn_ok.clicked.connect(dlg.accept)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(btn_ok)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        dlg.exec_()

    def show_theme_dialog(self):
        from PyQt5.QtWidgets import QDialog
        dlg = QDialog(self)
        dlg.setWindowTitle("Select Theme")
        dlg.resize(300, 150)
        layout = QVBoxLayout(dlg)
        
        # Header label
        hdr = QLabel("Choose a theme:")
        hdr.setStyleSheet("font-weight:600; font-size:14px; padding-bottom:6px;")
        layout.addWidget(hdr)
        
        # Button row
        btn_row = QHBoxLayout()
        layout.addLayout(btn_row)
        
        btn_light = QPushButton(" Light")
        btn_dark = QPushButton(" Dark")
        btn_system = QPushButton(" System")
        
        if HAS_ICONS:
            btn_light.setIcon(qta.icon('fa5s.sun', color='#ffcc00'))
            btn_dark.setIcon(qta.icon('fa5s.moon', color='#5e5ce6'))
            btn_system.setIcon(qta.icon('fa5s.desktop', color='#0a84ff'))
        
        btn_row.addWidget(btn_light)
        btn_row.addWidget(btn_dark)
        btn_row.addWidget(btn_system)
        
        btn_light.clicked.connect(lambda: (self.set_theme("light"), dlg.accept()))
        btn_dark.clicked.connect(lambda: (self.set_theme("dark"), dlg.accept()))
        btn_system.clicked.connect(lambda: (self.set_theme(None), dlg.accept()))
        
        dlg.exec_()

    def set_theme(self, mode):
        """Set theme to light, dark, or system (None)"""
        self.manual_theme_override = mode
        if mode is None:
            # System
            dark = system_dark_mode()
            if dark:
                QApplication.instance().setStyleSheet(DARK_STYLE)
            else:
                QApplication.instance().setStyleSheet(LIGHT_STYLE)
        elif mode == "light":
            QApplication.instance().setStyleSheet(LIGHT_STYLE)
        elif mode == "dark":
            QApplication.instance().setStyleSheet(DARK_STYLE)

    # ============================================================
    # DRAG & DROP
    # ============================================================

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            u = e.mimeData().urls()[0].toLocalFile().lower()
            if u.endswith(".xlsx"):
                e.acceptProposedAction()

    def dropEvent(self, e):
        path = e.mimeData().urls()[0].toLocalFile()
        df = pd.read_excel(path)
        self.columns = [str(c).strip() for c in df.columns]

        self.excel_path = path

        # Persist last directory when file is dropped
        try:
            self.last_dir = os.path.dirname(path)
            try:
                self.save_settings([])
            except:
                pass
        except:
            pass

        # Switch UI from placeholder to compact file info
        try:
            self.placeholder_widget.setVisible(False)
        except Exception:
            pass
        try:
            self.btn_select.setVisible(False)
        except Exception:
            pass
        self.lbl_file_compact.setText(os.path.basename(path))
        self.file_info_widget.setVisible(True)

        # Restore last mappings/settings (if available)
        self.load_settings()

        self.table.setRowCount(0)
        # Show content area now that Excel is loaded
        self.content_widget.setVisible(True)


# ============================================================
# PERFORMANCE & STABILITY TWEAKS (FINAL + PATCHED)
# ============================================================

# Allow PIL to load very large images without errors
try:
    Image.MAX_IMAGE_PIXELS = None
except:
    pass

# Global requests session for faster HTTP calls
requests_session = requests.Session()
requests_session.headers.update({
    "User-Agent":
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/125.0 Safari/537.36"
})

# Improve connection pooling and retries for faster parallel downloads
try:
    from requests.adapters import HTTPAdapter
    from urllib3.util.retry import Retry
    retry = Retry(total=2, backoff_factor=0.2, status_forcelist=(500,502,503,504))
    adapter = HTTPAdapter(pool_connections=100, pool_maxsize=100, max_retries=retry)
    requests_session.mount("http://", adapter)
    requests_session.mount("https://", adapter)
except Exception:
    pass

# ------------------------------------------------------------
# FIXED FAST_GET PATCH — supports ALL kwargs including verify=
# ------------------------------------------------------------
def fast_get(url, **kwargs):
    """
    Replaces requests.get()
    Ensures compatibility with verify=, timeout= and all other arguments.
    """
    return requests_session.get(url, **kwargs)

# Apply monkey patch
requests.get = fast_get


# ------------------------------------------------------------
# SAFE EXIT (prevents crash on forced quit or Ctrl+C)
# ------------------------------------------------------------
def safe_exit():
    try:
        app = QApplication.instance()
        if app:
            app.quit()
    except:
        pass


import os
import tempfile

def app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def get_writable_dir():
    """Get a writable directory for temp files and settings.
    Saves in the same directory where the app is located."""
    if getattr(sys, 'frozen', False):
        # For bundled apps, save alongside the app bundle
        # sys.executable is inside the app bundle, so go up to find the app directory
        exe_path = sys.executable
        # Navigate: .../Canvex.app/Contents/MacOS/Canvex -> .../
        app_bundle = exe_path
        # Go up until we find .app folder
        while app_bundle and not app_bundle.endswith('.app'):
            app_bundle = os.path.dirname(app_bundle)
        if app_bundle.endswith('.app'):
            # Return the directory containing the .app bundle
            return os.path.dirname(app_bundle)
        else:
            # Fallback
            return os.path.dirname(exe_path)
    else:
        # For development, use the app directory
        return app_dir()

def get_temp_dir():
    """Get a temp directory for temporary image files during processing."""
    if getattr(sys, 'frozen', False):
        # Use system temp directory for bundled apps
        temp_dir = os.path.join(tempfile.gettempdir(), 'Canvex_temp')
        os.makedirs(temp_dir, exist_ok=True)
        return temp_dir
    else:
        return app_dir()

def resource_path(filename):
    local = os.path.join(app_dir(), filename)
    if os.path.exists(local):
        return local
    try:
        return os.path.join(sys._MEIPASS, filename)
    except Exception:
        return local


# ============================================================
# MAIN APPLICATION RUNNER
# ============================================================

if __name__ == "__main__":
    try:
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    except:
        pass

    app = QApplication(sys.argv)

    # Set application icon for taskbar and window decorations
    try:
        # Use resource_path so the icon is loaded correctly from frozen builds
        icon_path = resource_path("app_icon.ico")
        if os.path.exists(icon_path):
            app.setApplicationIcon(QIcon(icon_path))
    except Exception:
        pass

    # Show splash screen if available
    splash = None
    splash_start_time = None
    try:
        # Resolve splash via resource_path so PyInstaller bundles are supported
        splash_path = resource_path("splash.png")
        if os.path.exists(splash_path):
            from PyQt5.QtWidgets import QSplashScreen
            from PyQt5.QtCore import Qt
            from PyQt5.QtGui import QImage
            import time

            # Load as QImage first for proper high-DPI handling
            splash_image = QImage(splash_path)
            
            # Ensure image loaded correctly
            if not splash_image.isNull():
                # Get device pixel ratio for high-DPI displays (Retina)
                device_pixel_ratio = app.devicePixelRatio()
                
                # Use larger size for crisp display (700x560 logical pixels)
                # This keeps the high-res PNG sharp on Retina displays
                target_width = int(700 * device_pixel_ratio)
                target_height = int(560 * device_pixel_ratio)
                
                # Scale only if the image is larger than target
                if splash_image.width() > target_width or splash_image.height() > target_height:
                    # Scale to fit while maintaining aspect ratio
                    splash_image = splash_image.scaled(
                        target_width, target_height,
                        Qt.KeepAspectRatio, Qt.SmoothTransformation
                    )
                
                # Convert to pixmap and set device pixel ratio for crisp rendering
                splash_pixmap = QPixmap.fromImage(splash_image)
                splash_pixmap.setDevicePixelRatio(device_pixel_ratio)
                
                splash = QSplashScreen(splash_pixmap)
                # Set window flags: frameless, always on top, no taskbar
                splash.setWindowFlags(splash.windowFlags() | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
                # Center on desktop (use logical size for positioning)
                logical_width = splash_pixmap.width() / device_pixel_ratio
                logical_height = splash_pixmap.height() / device_pixel_ratio
                splash.move(
                    int((app.desktop().screenGeometry().width() - logical_width) // 2),
                    int((app.desktop().screenGeometry().height() - logical_height) // 2)
                )
                splash.show()
                # Process events multiple times to ensure splash is rendered
                for _ in range(5):
                    app.processEvents()
                    time.sleep(0.05)
                splash_start_time = time.time()
    except Exception as e:
        pass

    # High-DPI Windows fix
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass

    # Ensure splash is visible for at least 2 seconds
    if splash is not None and splash_start_time is not None:
        import time
        elapsed = time.time() - splash_start_time
        if elapsed < 2.0:
            time.sleep(2.0 - elapsed)

    win = CanvaImageExcelCreator()
    win.show()
    
    # Close splash screen when main window appears
    if splash is not None:
        try:
            splash.finish(win)
        except:
            pass

    try:
        sys.exit(app.exec_())
    except KeyboardInterrupt:
        safe_exit()
