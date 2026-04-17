# -*- coding: utf-8 -*-
"""
企微自动建群引擎 v1.0
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
技术: SendMessage (后台点击) + PrintWindow (后台截图)
特点: 不抢鼠标、不抢键盘、企微被盖住也能操作
要求: 企微不能最小化 (被盖住可以)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
import os, sys, time, datetime, traceback, subprocess, ctypes, json

if sys.platform != "win32":
    print("This script only runs on Windows"); sys.exit(1)

# ━━━ 自动安装依赖 ━━━
def _ensure_pip():
    """确保 pip 可用 (嵌入式 Python 没有自带 pip)"""
    try:
        import pip
        return True
    except ImportError:
        pass
    print("  pip not found, bootstrapping...")
    try:
        import urllib.request
        pip_url = "https://bootstrap.pypa.io/get-pip.py"
        pip_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_get_pip.py")
        urllib.request.urlretrieve(pip_url, pip_path)
        subprocess.call([sys.executable, pip_path, "--no-warn-script-location"])
        os.remove(pip_path)
        return True
    except Exception as e:
        print(f"  pip bootstrap failed: {e}")
        return False

def _install_deps():
    for pkg, mod in [("pywin32", "win32gui"), ("Pillow", "PIL")]:
        try:
            __import__(mod)
        except ImportError:
            print(f"  Installing {pkg}...")
            if not _ensure_pip():
                print(f"  ERROR: Cannot install {pkg} without pip")
                sys.exit(1)
            subprocess.call([sys.executable, "-m", "pip", "install", pkg, "-q", "--no-warn-script-location"])

# 打包成 EXE 后不需要检查依赖 (已内置)
if not getattr(sys, 'frozen', False):
    _install_deps()
else:
    print("  [wecom_auto] ✅ frozen mode, 跳过依赖检查")

import win32gui, win32process, win32api, win32con, win32ui
from PIL import Image
from ctypes import windll
import zipfile, shutil

# ━━━ 智谱 AI OCR (高精度云端) ━━━
ZHIPU_OCR_URL = "https://open.bigmodel.cn/api/paas/v4/files/ocr"
_zhipu_api_key = ""  # 由 app.py 通过 set_zhipu_api_key() 注入

def set_zhipu_api_key(key):
    """外部设置智谱 AI OCR API Key (由 app.py 调用)"""
    global _zhipu_api_key
    _zhipu_api_key = key
    log(f"  [OCR] 智谱 API Key 已设置 ({key[:8]}...)")

def _zhipu_ocr(png_bytes, api_key=None):
    """调用智谱 AI OCR, 返回统一格式的识别结果列表。
    每项: {'text': str, 'cx': int, 'cy': int, 'conf': float,
           'x1': int, 'y1': int, 'x2': int, 'y2': int}
    """
    import urllib.request, json
    boundary = '----ZhipuOCR' + str(int(time.time()))
    parts = []
    for name, val in [('tool_type', 'hand_write'), ('language_type', 'CHN_ENG'), ('probability', 'true')]:
        parts.append(f'--{boundary}\r\nContent-Disposition: form-data; name="{name}"\r\n\r\n{val}'.encode())
    parts.append(
        f'--{boundary}\r\nContent-Disposition: form-data; name="file"; filename="s.png"\r\nContent-Type: image/png\r\n\r\n'.encode()
        + png_bytes
    )
    body = b'\r\n'.join(parts) + f'\r\n--{boundary}--\r\n'.encode()
    
    key = api_key or _zhipu_api_key
    if not key:
        log("  [OCR] ❌ 智谱 API Key 未配置!", "ERROR")
        return []
    req = urllib.request.Request(ZHIPU_OCR_URL, data=body, headers={
        'Authorization': f'Bearer {key}',
        'Content-Type': f'multipart/form-data; boundary={boundary}'
    })
    resp = urllib.request.urlopen(req, timeout=15)
    data = json.loads(resp.read())
    
    items = []
    for w in data.get('words_result', []):
        loc = w.get('location', {})
        x1 = loc.get('left', 0)
        y1 = loc.get('top', 0)
        w2 = loc.get('width', 0)
        h2 = loc.get('height', 0)
        prob = w.get('probability', {})
        items.append({
            'text': w.get('words', ''),
            'cx': x1 + w2 // 2,
            'cy': y1 + h2 // 2,
            'conf': prob.get('average', 0),
            'x1': x1, 'y1': y1,
            'x2': x1 + w2, 'y2': y1 + h2,
        })
    return items

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#   配置
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
if getattr(sys, 'frozen', False):
    SCRIPT_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
OUTPUT_DIR = os.path.join(DESKTOP, "企微建群日志")
os.makedirs(OUTPUT_DIR, exist_ok=True)

RUN_ID = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
RUN_DIR = os.path.join(OUTPUT_DIR, f"run_{RUN_ID}")
os.makedirs(RUN_DIR, exist_ok=True)

# 操作间隔 (秒) — 太快企微反应不过来
CLICK_DELAY = 0.3       # 点击后等待
TYPE_DELAY = 0.05        # 每个字符间隔
SEARCH_WAIT = 1.0        # 搜索结果加载等待
STEP_DELAY = 0.8         # 步骤之间等待
SCREENSHOT_DELAY = 0.5   # 截图前等待渲染

log_lines = []

def log(msg, level="INFO"):
    ts = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
    line = f"[{ts}] [{level:5s}] {msg}"
    print(line); log_lines.append(line)

def log_sep(title):
    log(""); log("=" * 60); log(f"  {title}"); log("=" * 60)

def get_run_path(filename):
    return os.path.join(RUN_DIR, filename)

def save_log(filename="auto_group_log.txt"):
    path = os.path.join(RUN_DIR, filename)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(log_lines))
    return path

def pack_results():
    zip_name = f"group_result_{RUN_ID}.zip"
    zip_path = os.path.join(OUTPUT_DIR, zip_name)
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(RUN_DIR):
            for f in files:
                fp = os.path.join(root, f)
                zf.write(fp, f)
    shutil.rmtree(RUN_DIR, ignore_errors=True)
    size_kb = os.path.getsize(zip_path) / 1024
    print(f"\n{'='*55}")
    print(f"  Results: {zip_path}")
    print(f"  Size: {size_kb:.1f} KB")
    print(f"{'='*55}")
    return zip_path


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#   窗口管理
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class WeComWindow:
    """企微窗口操作封装"""

    def __init__(self):
        self.hwnd = None
        self.pid = None
        self.width = 0
        self.height = 0

    def find(self):
        """查找企微主窗口 (包含最小化的窗口)"""
        windows = []
        def _enum(hwnd, _):
            cls = win32gui.GetClassName(hwnd)
            if cls == "WeWorkWindow":
                title = win32gui.GetWindowText(hwnd)
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                r = win32gui.GetWindowRect(hwnd)
                is_visible = win32gui.IsWindowVisible(hwnd)
                is_iconic = win32gui.IsIconic(hwnd)
                windows.append({"hwnd": hwnd, "pid": pid, "title": title,
                                "w": r[2]-r[0], "h": r[3]-r[1],
                                "visible": is_visible, "iconic": is_iconic,
                                "rect": r})
        win32gui.EnumWindows(_enum, None)

        if not windows:
            log("ERROR: WeCom not found!", "ERROR")
            log("  Please open WeCom and try again", "ERROR")
            return False

        win = windows[0]
        self.hwnd = win["hwnd"]
        self.pid = win["pid"]

        log(f"  [DEBUG] hwnd={win['hwnd']} visible={win['visible']} iconic={win['iconic']}")
        log(f"  [DEBUG] rect=({win['rect'][0]},{win['rect'][1]},{win['rect'][2]},{win['rect'][3]})")

        if win32gui.IsIconic(self.hwnd):
            log("  Window is minimized, restoring...")
            win32gui.ShowWindow(self.hwnd, win32con.SW_RESTORE)
            time.sleep(0.8)

        cr = win32gui.GetClientRect(self.hwnd)
        self.width = cr[2]
        self.height = cr[3]

        log(f"OK: {win['title']} hwnd={self.hwnd} client={self.width}x{self.height}")
        return True

    def sink_to_bottom(self):
        """将企微窗口压到最底层"""
        try:
            win32gui.SetWindowPos(self.hwnd, win32con.HWND_BOTTOM, 0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE)
            log(f"  sink_to_bottom OK, iconic={win32gui.IsIconic(self.hwnd)}")
        except Exception as e:
            log(f"  sink_to_bottom FAILED: {e}", "WARN")

    def ensure_not_minimized(self):
        """确保窗口没有最小化"""
        if win32gui.IsIconic(self.hwnd):
            log("  [WARN] Window is minimized! Restoring...", "WARN")
            win32gui.ShowWindow(self.hwnd, win32con.SW_RESTORE)
            time.sleep(0.5)
            self.sink_to_bottom()
            time.sleep(0.3)

    def log_window_state(self, label=""):
        """记录窗口状态 (调试)"""
        try:
            ic = win32gui.IsIconic(self.hwnd)
            vis = win32gui.IsWindowVisible(self.hwnd)
            r = win32gui.GetWindowRect(self.hwnd)
            cr = win32gui.GetClientRect(self.hwnd)
            log(f"  [STATE:{label}] iconic={ic} vis={vis} rect=({r[0]},{r[1]},{r[2]},{r[3]}) client={cr[2]}x{cr[3]}")
        except Exception as e:
            log(f"  [STATE:{label}] error: {e}")

    def screenshot(self, label="step"):
        """后台截图 (PrintWindow)"""
        self.ensure_not_minimized()
        time.sleep(SCREENSHOT_DELAY)
        try:
            w, h = self.width, self.height
            if w <= 0 or h <= 0:
                cr = win32gui.GetClientRect(self.hwnd)
                w, h = cr[2], cr[3]

            hwnd_dc = win32gui.GetWindowDC(self.hwnd)
            mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
            save_dc = mfc_dc.CreateCompatibleDC()
            bmp = win32ui.CreateBitmap()
            bmp.CreateCompatibleBitmap(mfc_dc, w, h)
            save_dc.SelectObject(bmp)

            # PW_CLIENTONLY=1 只截客户区
            windll.user32.PrintWindow(self.hwnd, save_dc.GetSafeHdc(), 1)

            bmpinfo = bmp.GetInfo()
            bmpstr = bmp.GetBitmapBits(True)
            img = Image.frombuffer('RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                                   bmpstr, 'raw', 'BGRX', 0, 1)

            fname = f"step_{label}.png"
            img.save(get_run_path(fname))

            # 清理
            win32gui.DeleteObject(bmp.GetHandle())
            save_dc.DeleteDC()
            mfc_dc.DeleteDC()
            win32gui.ReleaseDC(self.hwnd, hwnd_dc)

            log(f"  📸 截图: {fname} ({w}x{h})")
            return img
        except Exception as e:
            log(f"  截图失败: {e}", "ERROR")
            return None

    # ━━━ 后台操作 ━━━
    def _make_lparam(self, x, y):
        """构造 LPARAM (坐标参数)"""
        return (y << 16) | (x & 0xFFFF)

    def _full_click_seq(self, hwnd, x, y):
        """完整5步后台点击序列 (适配 Chromium/Qt hit-testing)
        Chromium 的 hit-testing 依赖 WM_MOUSEMOVE 先到达目标坐标，
        否则 LBUTTONDOWN 不知道要点击哪个元素，会静默失败。
        """
        lp = self._make_lparam(x, y)
        # Step 1: WM_MOUSEACTIVATE — 通知窗口即将收到鼠标消息
        win32gui.SendMessage(hwnd, win32con.WM_MOUSEACTIVATE,
                             hwnd, (win32con.WM_LBUTTONDOWN << 16) | 1)
        time.sleep(0.03)
        # Step 2: WM_SETCURSOR — 触发光标设置，准备 hit-test
        win32gui.SendMessage(hwnd, 0x0020,  # WM_SETCURSOR
                             hwnd, (win32con.WM_LBUTTONDOWN << 16) | 1)
        time.sleep(0.03)
        # Step 3: WM_MOUSEMOVE — hover 预热 (关键!)
        win32gui.SendMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lp)
        time.sleep(0.15)  # 留够时间让 Chromium 完成 hit-test
        # Step 4: WM_LBUTTONDOWN
        win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN,
                             win32con.MK_LBUTTON, lp)
        time.sleep(0.08)
        # Step 5: WM_LBUTTONUP
        win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lp)
        time.sleep(CLICK_DELAY)

    def click(self, x, y, description=""):
        """后台点击 (带 hover 预热, 适配 Chromium hit-testing)"""
        self.ensure_not_minimized()
        desc = f" ({description})" if description else ""
        log(f"  🖱️ 点击 ({x},{y}){desc}")
        self._full_click_seq(self.hwnd, x, y)

    def double_click(self, x, y, description=""):
        """后台双击"""
        self.ensure_not_minimized()
        desc = f" ({description})" if description else ""
        log(f"  🖱️ 双击 ({x},{y}){desc}")
        lp = self._make_lparam(x, y)
        self._full_click_seq(self.hwnd, x, y)
        time.sleep(0.05)
        win32gui.SendMessage(self.hwnd, win32con.WM_LBUTTONDBLCLK,
                             win32con.MK_LBUTTON, lp)
        time.sleep(0.08)
        win32gui.SendMessage(self.hwnd, win32con.WM_LBUTTONUP, 0, lp)
        time.sleep(CLICK_DELAY)

    def type_text(self, text, description=""):
        """后台输入文字 (逐字符)"""
        desc = f" ({description})" if description else ""
        log(f"  ⌨️ 输入: '{text}'{desc}")
        for ch in text:
            # 使用 WM_IME_CHAR 发送中文字符
            win32api.PostMessage(self.hwnd, 0x0286, ord(ch), 0)  # WM_IME_CHAR
            time.sleep(TYPE_DELAY)
        time.sleep(0.2)

    def send_key(self, vk_code, description=""):
        """后台发送按键"""
        desc = f" ({description})" if description else ""
        log(f"  ⌨️ 按键: VK={hex(vk_code)}{desc}")
        win32api.PostMessage(self.hwnd, win32con.WM_KEYDOWN, vk_code, 0)
        time.sleep(0.05)
        win32api.PostMessage(self.hwnd, win32con.WM_KEYUP, vk_code, 0)
        time.sleep(CLICK_DELAY)

    def send_ctrl_key(self, key_char, description=""):
        """后台发送 Ctrl+Key 组合键"""
        desc = f" ({description})" if description else ""
        vk = ord(key_char.upper())
        log(f"  ⌨️ Ctrl+{key_char}{desc}")
        win32api.PostMessage(self.hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
        time.sleep(0.02)
        win32api.PostMessage(self.hwnd, win32con.WM_KEYDOWN, vk, 0)
        time.sleep(0.02)
        win32api.PostMessage(self.hwnd, win32con.WM_KEYUP, vk, 0)
        time.sleep(0.02)
        win32api.PostMessage(self.hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
        time.sleep(CLICK_DELAY)

    def _find_popup_hwnd(self, popup_class):
        """查找弹窗句柄"""
        popups = []
        def _enum(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                cls = win32gui.GetClassName(hwnd)
                if cls == popup_class:
                    popups.append(hwnd)
        win32gui.EnumWindows(_enum, None)
        return popups[0] if popups else None

    def click_popup(self, popup_class, x, y, description=""):
        """点击弹窗上的按钮 (完整5步序列, 适配 Chromium hit-testing)"""
        phwnd = self._find_popup_hwnd(popup_class)
        if phwnd:
            log(f"  🖱️ 弹窗点击 [{popup_class}] ({x},{y}) ({description})")
            self._full_click_seq(phwnd, x, y)
        else:
            log(f"  ⚠️ 弹窗 [{popup_class}] 未找到", "WARN")
            time.sleep(CLICK_DELAY)

    def screenshot_popup(self, popup_class, label="popup"):
        """截取弹窗截图"""
        popups = []
        def _enum(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                cls = win32gui.GetClassName(hwnd)
                if cls == popup_class:
                    popups.append(hwnd)
        win32gui.EnumWindows(_enum, None)

        if not popups:
            log(f"  ⚠️ 弹窗 [{popup_class}] 未找到", "WARN")
            return None

        phwnd = popups[0]
        try:
            cr = win32gui.GetClientRect(phwnd)
            w, h = cr[2], cr[3]
            if w <= 0 or h <= 0:
                r = win32gui.GetWindowRect(phwnd)
                w, h = r[2]-r[0], r[3]-r[1]

            hwnd_dc = win32gui.GetWindowDC(phwnd)
            mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
            save_dc = mfc_dc.CreateCompatibleDC()
            bmp = win32ui.CreateBitmap()
            bmp.CreateCompatibleBitmap(mfc_dc, w, h)
            save_dc.SelectObject(bmp)
            windll.user32.PrintWindow(phwnd, save_dc.GetSafeHdc(), 1)

            bmpinfo = bmp.GetInfo()
            bmpstr = bmp.GetBitmapBits(True)
            img = Image.frombuffer('RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                                   bmpstr, 'raw', 'BGRX', 0, 1)

            fname = f"step_{label}.png"
            img.save(get_run_path(fname))

            win32gui.DeleteObject(bmp.GetHandle())
            save_dc.DeleteDC()
            mfc_dc.DeleteDC()
            win32gui.ReleaseDC(phwnd, hwnd_dc)

            log(f"  📸 弹窗截图: {fname} ({w}x{h})")
            return img
        except Exception as e:
            log(f"  弹窗截图失败: {e}", "ERROR")
            return None

    def type_text_to_popup(self, popup_class, text, description=""):
        """向弹窗输入文字"""
        popups = []
        def _enum(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                cls = win32gui.GetClassName(hwnd)
                if cls == popup_class:
                    popups.append(hwnd)
        win32gui.EnumWindows(_enum, None)

        if not popups:
            log(f"  ⚠️ 弹窗未找到，无法输入", "WARN")
            return

        phwnd = popups[0]
        desc = f" ({description})" if description else ""
        log(f"  ⌨️ 弹窗输入: '{text}'{desc}")
        for ch in text:
            win32api.PostMessage(phwnd, 0x0286, ord(ch), 0)  # WM_IME_CHAR
            time.sleep(TYPE_DELAY)
        time.sleep(0.2)

    def send_key_to_popup(self, popup_class, vk_code, description=""):
        """向弹窗发送按键"""
        popups = []
        def _enum(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                cls = win32gui.GetClassName(hwnd)
                if cls == popup_class:
                    popups.append(hwnd)
        win32gui.EnumWindows(_enum, None)

        if not popups:
            log(f"  ⚠️ 弹窗未找到", "WARN")
            return

        phwnd = popups[0]
        win32api.PostMessage(phwnd, win32con.WM_KEYDOWN, vk_code, 0)
        time.sleep(0.05)
        win32api.PostMessage(phwnd, win32con.WM_KEYUP, vk_code, 0)
        desc = f" ({description})" if description else ""
        log(f"  ⌨️ 弹窗按键: VK={hex(vk_code)}{desc}")
        time.sleep(CLICK_DELAY)

    # ━━━ OCR 动态元素定位 ━━━
    def _screenshot_to_image(self, hwnd):
        """对任意窗口做 PrintWindow 截图，返回 PIL Image。
        修复: 使用 GetDC (客户区 DC) 而非 GetWindowDC (含标题栏),
        确保 OCR 坐标与窗口客户区坐标一致。
        """
        # DPI 感知: 确保截图分辨率与实际像素一致
        try:
            windll.user32.SetProcessDPIAware()
        except Exception:
            pass
        
        cr = win32gui.GetClientRect(hwnd)
        w, h = cr[2], cr[3]
        if w <= 0 or h <= 0:
            return None
        
        # 使用 GetDC (客户区) 而非 GetWindowDC (整个窗口含标题栏)
        hwnd_dc = win32gui.GetDC(hwnd)
        mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
        save_dc = mfc_dc.CreateCompatibleDC()
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(mfc_dc, w, h)
        save_dc.SelectObject(bmp)
        # PW_CLIENTONLY = 1, 只截取客户区
        windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 1)
        bmpinfo = bmp.GetInfo()
        bmpstr = bmp.GetBitmapBits(True)
        img = Image.frombuffer('RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                               bmpstr, 'raw', 'BGRX', 0, 1)
        win32gui.DeleteObject(bmp.GetHandle())
        save_dc.DeleteDC()
        mfc_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwnd_dc)
        return img

    def _screenshot_foreground(self, hwnd):
        """前台截图: 捕捉 Chromium overlay (如群管理面板)。
        PrintWindow 无法捕获 CSS overlay, 需要从屏幕 DC 截取。
        """
        try:
            windll.user32.SetProcessDPIAware()
        except Exception:
            pass
        
        # 强制窗口到最前台 (绕过 Windows 前台锁)
        try:
            import ctypes
            user32 = ctypes.windll.user32
            kernel32 = ctypes.windll.kernel32
            fg_hwnd = user32.GetForegroundWindow()
            fg_tid = user32.GetWindowThreadProcessId(fg_hwnd, None)
            my_tid = kernel32.GetCurrentThreadId()
            if fg_tid != my_tid:
                user32.AttachThreadInput(my_tid, fg_tid, True)
            user32.keybd_event(0x12, 0, 0x0002, 0)  # ALT up
            user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 0x0040 | 0x0001 | 0x0002)
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            user32.SetForegroundWindow(hwnd)
            user32.BringWindowToTop(hwnd)
            if fg_tid != my_tid:
                user32.AttachThreadInput(my_tid, fg_tid, False)
            time.sleep(0.8)
        except Exception:
            pass
        
        # 获取窗口客户区在屏幕上的位置
        cr = win32gui.GetClientRect(hwnd)
        w, h = cr[2], cr[3]
        if w <= 0 or h <= 0:
            return None
        
        import ctypes.wintypes
        pt = ctypes.wintypes.POINT(0, 0)
        ctypes.windll.user32.ClientToScreen(hwnd, ctypes.byref(pt))
        sx, sy = pt.x, pt.y
        
        # 从屏幕 DC 截取 (能捕获 overlay)
        screen_dc = win32gui.GetDC(0)  # 屏幕 DC
        mfc_dc = win32ui.CreateDCFromHandle(screen_dc)
        save_dc = mfc_dc.CreateCompatibleDC()
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(mfc_dc, w, h)
        save_dc.SelectObject(bmp)
        save_dc.BitBlt((0, 0), (w, h), mfc_dc, (sx, sy), 0x00CC0020)  # SRCCOPY
        
        bmpinfo = bmp.GetInfo()
        bmpstr = bmp.GetBitmapBits(True)
        img = Image.frombuffer('RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                               bmpstr, 'raw', 'BGRX', 0, 1)
        win32gui.DeleteObject(bmp.GetHandle())
        save_dc.DeleteDC()
        mfc_dc.DeleteDC()
        win32gui.ReleaseDC(0, screen_dc)
        
        # 恢复正常 z-order (取消 TOPMOST)
        try:
            import ctypes
            ctypes.windll.user32.SetWindowPos(
                hwnd, -2, 0, 0, 0, 0, 0x0001 | 0x0002)  # HWND_NOTOPMOST
        except Exception:
            pass
        return img

    def ocr_scan_foreground(self, hwnd, label="fg"):
        """前台 OCR: 用屏幕截取捕捉 Chromium overlay 面板内容。"""
        img = self._screenshot_foreground(hwnd)
        if img is None:
            log("  OCR[前台]: screenshot failed", "ERROR")
            return [], img
        try:
            img.save(get_run_path(f"ocr_{label}.png"))
        except Exception:
            pass
        import io
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        items = _zhipu_ocr(buf.getvalue())
        log(f"  OCR[前台/智谱]: {len(items)} items ({label})")
        return items, img

    def ocr_scan(self, hwnd, label="ocr"):
        """对窗口截图并 OCR (智谱AI优先, 本地回退)。
        返回识别结果列表, 每项:
        {'text': str, 'cx': int, 'cy': int, 'conf': float,
         'x1': int, 'y1': int, 'x2': int, 'y2': int}
        """
        img = self._screenshot_to_image(hwnd)
        if img is None:
            log("  OCR: screenshot failed", "ERROR")
            return [], img
        try:
            img.save(get_run_path(f"ocr_{label}.png"))
        except Exception:
            pass
        
        import io
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        items = _zhipu_ocr(buf.getvalue())
        log(f"  OCR[智谱]: {len(items)} items ({label})")
        return items, img

    def ocr_scan_region(self, hwnd, x1, y1, x2, y2, label="region"):
        """区域 OCR: 先截全图裁剪到指定区域, 再识别。
        坐标已转换回窗口坐标系。智谱AI优先, 本地回退。
        """
        img = self._screenshot_to_image(hwnd)
        if img is None:
            return [], None
        
        iw, ih = img.size
        rx1 = max(0, min(x1, iw))
        ry1 = max(0, min(y1, ih))
        rx2 = max(rx1, min(x2, iw))
        ry2 = max(ry1, min(y2, ih))
        cropped = img.crop((rx1, ry1, rx2, ry2))
        
        try:
            cropped.save(get_run_path(f"ocr_{label}.png"))
        except Exception:
            pass
        
        import io
        buf = io.BytesIO()
        cropped.save(buf, format='PNG')
        items = _zhipu_ocr(buf.getvalue())
        for it in items:
            it['cx'] += rx1; it['cy'] += ry1
            it['x1'] += rx1; it['y1'] += ry1
            it['x2'] += rx1; it['y2'] += ry1
        log(f"  OCR[智谱] region [{rx1},{ry1}-{rx2},{ry2}]: {len(items)} items ({label})")
        return items, cropped

    def ocr_find(self, items, keyword):
        """在 OCR 结果中查找匹配项 (支持模糊匹配)。
        
        匹配策略 (按优先级):
        1. 精确子串匹配
        2. 首字匹配 + 长度接近
        3. 字符重叠率 >= 50%
        """
        if not items or not keyword:
            return None
        
        # 策略1: 精确子串
        for item in items:
            if keyword in item['text']:
                return item
        
        # 策略2: 首字匹配 + text 长度合理 (±3字以内)
        for item in items:
            t = item['text']
            if len(t) > 0 and keyword[0] == t[0] and abs(len(t) - len(keyword)) <= 3:
                return item
        
        # 策略3: 字符重叠率 (keyword 的字有多少出现在 text 中)
        if len(keyword) >= 2:
            best_item, best_ratio = None, 0
            for item in items:
                t = item['text']
                if len(t) < 1:
                    continue
                match_count = sum(1 for c in keyword if c in t)
                ratio = match_count / len(keyword)
                if ratio > best_ratio and ratio >= 0.5:
                    best_ratio = ratio
                    best_item = item
            if best_item:
                return best_item
        
        return None

    def ocr_click_text(self, hwnd, keyword, label="", items=None):
        """OCR 定位文字并点击: 截图 → OCR → 找文字 → 点击中心坐标
        返回 (success, items, img)
        """
        if items is None:
            items, img = self.ocr_scan(hwnd, label or keyword)
        match = self.ocr_find(items, keyword)
        if match:
            log(f"  OCR click: '{keyword}' -> ({match['cx']},{match['cy']}) conf={match['conf']:.2f}")
            self._full_click_seq(hwnd, match['cx'], match['cy'])
            return True, items
        else:
            log(f"  OCR click: '{keyword}' NOT FOUND", "WARN")
            return False, items

    def find_process_window(self, window_class):
        """查找与企微同进程的指定类名窗口"""
        if not self.pid:
            return None
        result = []
        def _enum(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                cls = win32gui.GetClassName(hwnd)
                if cls == window_class:
                    try:
                        _, wpid = win32process.GetWindowThreadProcessId(hwnd)
                        if wpid == self.pid:
                            result.append(hwnd)
                    except Exception:
                        pass
        win32gui.EnumWindows(_enum, None)
        return result[0] if result else None

    def setup_group_privacy(self):
        """开启「禁止互相添加为联系人」— 使用 OCR 动态定位

        群管理面板是独立弹窗 ExternalConversationManagerWindow,
        不在主窗口 sidebar 内。
        """
        POPUP_CLASS = "ExternalConversationManagerWindow"
        log("  [Privacy] 开启「禁止互相添加为联系人」...")

        # Step 1: 查找群管理弹窗 — 可能在建群后自动打开
        popup = self.find_process_window(POPUP_CLASS)
        if not popup:
            # 弹窗未打开，需要通过主窗口操作打开
            log("  [Privacy] 群管理弹窗未找到，从主窗口打开...")
            # 点击 "..." 打开聊天信息面板
            ok, items = self.ocr_click_text(self.hwnd, "...", "main_dots")
            if not ok:
                # fallback: 用侧边栏右上角大致位置
                self.click(self.width - 66, 38, "... fallback")
            time.sleep(2.0)
            # 点击 "群管理"
            ok, items = self.ocr_click_text(self.hwnd, "群管理", "main_sidebar")
            time.sleep(3.0)
            popup = self.find_process_window(POPUP_CLASS)

        if not popup:
            log("  [Privacy] ERROR: 无法打开群管理弹窗", "ERROR")
            return False

        log(f"  [Privacy] 群管理弹窗 HWND={popup}")

        # Step 2: OCR 扫描弹窗，查找「禁止互相添加」
        items, img = self.ocr_scan(popup, "privacy_popup")
        target = self.ocr_find(items, "禁止互相添加")

        if not target:
            log("  [Privacy] 未找到「禁止互相添加」文字", "ERROR")
            return False

        # Step 3: 检查复选框状态
        # 复选框在文字左侧 (x1-21 ~ x1-5)，与文字同高
        # 勾选态: 蓝色填充 RGB≈(76,149,243)
        # 未勾选: 白色内部 + 灰色边框
        def _is_checked(img, target_item):
            """检测复选框是否已勾选 — 采样多个像素"""
            if img is None:
                return False
            pixels = img.load()
            # 复选框中心约在 (x1-10, cy)
            chk_cx = target_item['x1'] - 10
            chk_cy = target_item['cy']
            blue_count = 0
            total = 0
            # 在复选框区域采样 3x3 格点
            for dx in [-5, 0, 5]:
                for dy in [-3, 0, 3]:
                    px = chk_cx + dx
                    py = chk_cy + dy
                    if 0 < px < img.width and 0 < py < img.height:
                        r, g, b = pixels[px, py][:3]
                        total += 1
                        # WeCom 勾选蓝色: RGB≈(76,149,243), b > 200 且 b 远大于 r
                        if b > 200 and b > r + 50:
                            blue_count += 1
            return blue_count >= 3  # 至少 3 个蓝色像素

        if _is_checked(img, target):
            log("  [Privacy] 已经勾选，无需操作")
            return True

        # Step 4: 点击复选框文字（点击文字区域也能触发复选框）
        log(f"  [Privacy] 点击「禁止互相添加为联系人」at ({target['cx']},{target['cy']})")
        self._full_click_seq(popup, target['cx'], target['cy'])
        time.sleep(1.5)

        # Step 5: 验证 — 重新截图确认变化
        items2, img2 = self.ocr_scan(popup, "privacy_after")
        if _is_checked(img2, target):
            log("  [Privacy] OK 已勾选「禁止互相添加为联系人」")
            return True

        log("  [Privacy] 已尝试勾选 (请查看 ocr_privacy_*.png 验证)")
        return True


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#   坐标计算 (基于比例, 适配不同窗口大小)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# 参考分辨率 (从截图获取): 客户区 1046x705 (去掉28px标题栏)
REF_W = 1046
REF_H = 705

# 关键坐标 (相对于客户区, 基于截图测量)
COORDS = {
    # 左侧导航栏 (固定55px宽)
    "nav_message":     (28, 80),     # 消息 tab
    "nav_contacts":    (28, 560),    # 通讯录 tab

    # 搜索区域
    "search_box":      (170, 27),    # 搜索框
    "plus_button":     (283, 27),    # [+] 按钮

    # 聊天列表 (第一个结果)
    "chat_list_1st":   (170, 90),    # 聊天列表第一项
    "search_result_1": (170, 90),    # 搜索结果第一项

    # 聊天内容区右上角
    "chat_more":       (940, 12),    # ⋯ 更多按钮
    "chat_add_member": (970, 12),    # + 邀请成员

    # 发起群聊弹窗 (weWorkSelectUser, 相对于弹窗本身)
    # 弹窗大小 640x560 (从截图确认)
    # 左侧: 搜索框+成员列表, 右侧: "发起群聊"标题
    # 底部: [完成](蓝色,偏左) [取消](偏右)
    "popup_search":     (160, 40),   # 搜索框 (左上)
    "popup_result_1":   (160, 170),  # 搜索结果第一项 (带checkbox)
    "popup_confirm":    (415, 495),  # "完成" 按钮 (蓝色, 偏左!) 
    "popup_cancel":     (555, 495),  # "取消" 按钮 (偏右)
}

def calc_pos(wc, key):
    """根据窗口实际大小计算坐标"""
    ref_x, ref_y = COORDS[key]
    actual_x = int(ref_x * wc.width / REF_W)
    actual_y = int(ref_y * wc.height / REF_H)
    return actual_x, actual_y


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#   建群流程
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def create_group(wc, customer_name, members_to_add):
    """
    v1.1: No ESC key (ESC minimizes WeCom!)
    Uses [+] -> popup -> search all members in popup
    """
    all_members = [customer_name] + members_to_add
    log_sep(f"Create group: {' + '.join(all_members)}")

    # Step 0: Prepare
    log("\n>> Step 0: Prepare")
    wc.log_window_state("start")
    wc.sink_to_bottom()
    wc.screenshot("00_initial")
    wc.log_window_state("after_sink")

    # Step 1: Click message tab
    log("\n>> Step 1: Message tab")
    x, y = calc_pos(wc, "nav_message")
    wc.click(x, y, "msg_tab")
    time.sleep(STEP_DELAY)
    wc.log_window_state("after_msg_tab")
    wc.screenshot("01_msg_tab")

    # Step 2: Click [+] button (top left, next to search)
    log("\n>> Step 2: Click [+] button")
    x, y = calc_pos(wc, "plus_button")
    wc.click(x, y, "plus_btn")
    time.sleep(STEP_DELAY)
    wc.log_window_state("after_plus")
    wc.screenshot("02_plus_clicked")

    # Step 3: Check for popup
    log("\n>> Step 3: Check popup")
    time.sleep(0.5)
    popup_img = wc.screenshot_popup("weWorkSelectUser", "03_popup")

    if popup_img is None:
        log("  No popup yet, try clicking menu item...")
        # The [+] shows a dropdown menu, group chat is usually the 1st item
        menu_x, menu_y = x, y + 50
        wc.click(menu_x, menu_y, "menu_group_chat")
        time.sleep(STEP_DELAY)
        wc.screenshot("03a_menu")
        time.sleep(0.5)
        popup_img = wc.screenshot_popup("weWorkSelectUser", "03b_popup")

    if popup_img is None:
        log("  WARN: Still no popup!", "WARN")
        wc.log_window_state("no_popup")

    # Step 4: Search and select ALL members in popup
    for i, member in enumerate(all_members):
        step_n = i + 1
        log(f"\n>> Step 4.{step_n}: Search '{member}'")

        # Click search box in popup
        px, py = COORDS["popup_search"]
        wc.click_popup("weWorkSelectUser", px, py, "popup_search")
        time.sleep(0.2)

        # Type member name
        wc.type_text_to_popup("weWorkSelectUser", member, f"search_{member}")
        time.sleep(SEARCH_WAIT)
        wc.screenshot_popup("weWorkSelectUser", f"04_{step_n}_search_{member}")

        # Click first result (check it)
        rx, ry = COORDS["popup_result_1"]
        wc.click_popup("weWorkSelectUser", rx, ry, f"select_{member}")
        time.sleep(0.3)
        wc.screenshot_popup("weWorkSelectUser", f"04_{step_n}_selected_{member}")

        # Clear search with BACKSPACE (NOT ESC! ESC closes/minimizes!)
        clear_count = len(member) * 2 + 5  # extra for safety
        log(f"  Clearing search (backspace x{clear_count})")
        for _ in range(clear_count):
            wc.send_key_to_popup("weWorkSelectUser", win32con.VK_BACK, "")
            time.sleep(0.02)
        time.sleep(0.3)

    # Step 5: Confirm
    log("\n>> Step 5: Confirm group creation")
    wc.screenshot_popup("weWorkSelectUser", "05_before_confirm")
    cx, cy = COORDS["popup_confirm"]
    wc.click_popup("weWorkSelectUser", cx, cy, "confirm_btn")
    time.sleep(STEP_DELAY * 2)

    # Step 6: Verify
    log("\n>> Step 6: Verify result")
    wc.log_window_state("after_confirm")
    wc.screenshot("06_result")
    time.sleep(0.5)
    wc.screenshot("06_result_final")

    log_sep("DONE!")
    log(f"  Customer: {customer_name}")
    log(f"  Members: {', '.join(members_to_add)}")
    log(f"  Output: {RUN_DIR}")
    wc.log_window_state("final")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#   主入口
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def main():
    print()
    print("=" * 55)
    print("  WeCom Auto Group Creator v1.0")
    print("  Silent mode: SendMessage + PrintWindow")
    print("=" * 55)
    print()

    log_sep(f"v1.0 | {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    ver = sys.getwindowsversion()
    log(f"System: Win{'11' if ver.build >= 22000 else '10'} Build {ver.build}")
    log(f"Screen: {ctypes.windll.user32.GetSystemMetrics(0)}x{ctypes.windll.user32.GetSystemMetrics(1)}")

    # 查找企微
    wc = WeComWindow()
    if not wc.find():
        save_log(); pack_results()
        return

    print()
    print("  Select mode:")
    print("  [1] Test mode - take screenshots only (safe)")
    print("  [2] Create test group (JunHuai + WuZeHua + WuTianYu)")
    print("  [3] Custom group")
    print()

    mode = input("  Enter mode (1/2/3): ").strip()

    if mode == "1":
        # Safe test mode
        log_sep("Test Mode: Screenshots Only")
        wc.log_window_state("initial")
        wc.sink_to_bottom()
        wc.screenshot("test_main")
        wc.log_window_state("after_sink")

        # Test click message tab
        x, y = calc_pos(wc, "nav_message")
        wc.click(x, y, "msg_tab")
        time.sleep(STEP_DELAY)
        wc.log_window_state("after_click_msg")
        wc.screenshot("test_after_click")

        # Test click [+] button
        x, y = calc_pos(wc, "plus_button")
        wc.click(x, y, "plus_btn")
        time.sleep(STEP_DELAY)
        wc.log_window_state("after_click_plus")
        wc.screenshot("test_plus_clicked")

        # Check if popup appeared
        time.sleep(0.3)
        wc.screenshot_popup("weWorkSelectUser", "test_popup")
        wc.log_window_state("before_close")

        # Click away to close any menu (click on chat area)
        wc.click(500, 400, "click_away_to_close")
        time.sleep(0.3)
        wc.log_window_state("after_close")
        wc.screenshot("test_closed")

        log_sep("Test completed!")

    elif mode == "2":
        # 测试建群 (使用预设账号)
        create_group(wc,
                     customer_name="君怀",
                     members_to_add=["吴泽华", "吴天宇"])

    elif mode == "3":
        # 自定义建群
        customer = input("  Customer name: ").strip()
        members_str = input("  Members (comma separated): ").strip()
        members = [m.strip() for m in members_str.split(",") if m.strip()]
        if customer and members:
            create_group(wc, customer, members)
        else:
            log("Invalid input", "ERROR")

    save_log()
    pack_results()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"\nERROR: {e}", "ERROR")
        log(traceback.format_exc(), "ERROR")
        save_log()
        pack_results()
    input("\nPress Enter to exit...")
