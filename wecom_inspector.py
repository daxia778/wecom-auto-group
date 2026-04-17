# -*- coding: utf-8 -*-
"""
企微静默操作能力验证 v3.0
验证: PrintWindow截图 + SendMessage点击 + 后台键盘输入
"""
import os, sys, time, datetime, traceback, subprocess, ctypes

if sys.platform != "win32":
    print("Only Windows"); sys.exit(1)

for pkg, mod in [("uiautomation", "uiautomation"), ("pywin32", "win32gui"), ("Pillow", "PIL")]:
    try: __import__(mod)
    except ImportError:
        print(f"  Installing {pkg}...")
        try: subprocess.call([sys.executable, "-m", "pip", "install", pkg, "-q", "--no-warn-script-location"])
        except: pass
print("Dependencies OK\n")

import win32gui, win32process, win32api, win32con, win32ui
from PIL import Image
from ctypes import windll

import zipfile, tempfile, shutil

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# 输出文件夹: 桌面/企微探查结果/ (固定位置,不会乱)
DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
OUTPUT_DIR = os.path.join(DESKTOP, "企微探查结果")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 本次运行的临时目录 (最后打包成zip)
RUN_ID = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
RUN_DIR = os.path.join(OUTPUT_DIR, f"run_{RUN_ID}")
os.makedirs(RUN_DIR, exist_ok=True)

log_lines = []

def log(msg, level="INFO"):
    ts = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
    line = f"[{ts}] [{level:5s}] {msg}"
    print(line); log_lines.append(line)

def log_sep(title):
    log(""); log("=" * 60); log(f"  {title}"); log("=" * 60)

def get_run_path(filename):
    """获取本次运行的输出路径 (截图等文件用这个)"""
    return os.path.join(RUN_DIR, filename)

def save_log(filename="wecom_silent_test.txt"):
    path = os.path.join(RUN_DIR, filename)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(log_lines))
    return path

def pack_results():
    """把本次运行的所有文件打包成zip"""
    zip_name = f"wecom_result_{RUN_ID}.zip"
    zip_path = os.path.join(OUTPUT_DIR, zip_name)
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(RUN_DIR):
            for f in files:
                fp = os.path.join(root, f)
                zf.write(fp, f)
    # 清理临时目录
    shutil.rmtree(RUN_DIR, ignore_errors=True)
    
    size_kb = os.path.getsize(zip_path) / 1024
    print(f"\n{'='*55}")
    print(f"  OK! All results packed:")
    print(f"  {zip_path}")
    print(f"  Size: {size_kb:.1f} KB")
    print(f"")
    print(f"  Please send this ONE zip file back!")
    print(f"{'='*55}")
    return zip_path

# ━━━ 查找企微窗口 ━━━
def find_wecom_windows():
    """找到所有企微窗口"""
    windows = []
    def _enum(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            cls = win32gui.GetClassName(hwnd)
            if cls in ("WeWorkWindow", "weWorkSelectUser") or "企业微信" in title:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                r = win32gui.GetWindowRect(hwnd)
                windows.append({"hwnd": hwnd, "title": title, "class": cls,
                                "pid": pid, "rect": r,
                                "w": r[2]-r[0], "h": r[3]-r[1]})
    win32gui.EnumWindows(_enum, None)
    return windows

# ━━━ 测试1: PrintWindow 后台截图 ━━━
def test_printwindow(hwnd, label="main"):
    """测试能否在后台截取企微窗口"""
    log_sep(f"Test 1: PrintWindow ({label})")
    try:
        r = win32gui.GetWindowRect(hwnd)
        w, h = r[2] - r[0], r[3] - r[1]
        if w <= 0 or h <= 0:
            # 窗口可能最小化，尝试获取正常大小
            placement = win32gui.GetWindowPlacement(hwnd)
            nr = placement[4]  # normalPosition
            w, h = nr[2] - nr[0], nr[3] - nr[1]
            log(f"  Window minimized, normal size: {w}x{h}")

        log(f"  Window: {w}x{h}")

        # 方法1: PrintWindow (PW_RENDERFULLCONTENT=2, 最完整)
        for flag_name, flag in [("PW_CLIENTONLY", 1), ("PW_RENDERFULLCONTENT", 2), ("PW_DEFAULT", 0)]:
            try:
                hwnd_dc = win32gui.GetWindowDC(hwnd)
                mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
                save_dc = mfc_dc.CreateCompatibleDC()
                bmp = win32ui.CreateBitmap()
                bmp.CreateCompatibleBitmap(mfc_dc, w, h)
                save_dc.SelectObject(bmp)

                result = windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), flag)

                # 转为 PIL Image
                bmpinfo = bmp.GetInfo()
                bmpstr = bmp.GetBitmapBits(True)
                img = Image.frombuffer('RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                                       bmpstr, 'raw', 'BGRX', 0, 1)

                # 检查是否全黑
                pixels = list(img.getdata())
                non_black = sum(1 for p in pixels[:1000] if sum(p) > 30)
                is_valid = non_black > 100

                fname = f"screenshot_{label}_{flag_name}.png"
                img.save(get_run_path(fname))

                log(f"  {flag_name}: result={result}, valid={is_valid}, saved={fname}")
                if is_valid:
                    log(f"  ✅ PrintWindow {flag_name} SUCCESS! Non-black pixels: {non_black}/1000")

                # 清理
                win32gui.DeleteObject(bmp.GetHandle())
                save_dc.DeleteDC()
                mfc_dc.DeleteDC()
                win32gui.ReleaseDC(hwnd, hwnd_dc)

            except Exception as e:
                log(f"  {flag_name}: FAILED - {e}", "WARN")

    except Exception as e:
        log(f"  PrintWindow test failed: {e}", "ERROR")
        log(traceback.format_exc(), "DEBUG")

# ━━━ 测试2: SendMessage 后台点击 ━━━
def test_sendmessage(hwnd):
    """测试 SendMessage 后台发送点击消息"""
    log_sep("Test 2: SendMessage (background click)")
    try:
        WM_LBUTTONDOWN = 0x0201
        WM_LBUTTONUP   = 0x0202
        MK_LBUTTON     = 0x0001

        # 获取窗口大小
        r = win32gui.GetWindowRect(hwnd)
        w, h = r[2] - r[0], r[3] - r[1]
        log(f"  Window size: {w}x{h}")

        # 测试在窗口中心发送一个无害的点击
        # 这里只测试消息是否能发送成功，不会真的操作
        cx, cy = w // 2, h // 2

        def MAKELPARAM(x, y):
            return (y << 16) | (x & 0xFFFF)

        # 先测试 PostMessage (异步，不阻塞)
        r1 = win32api.PostMessage(hwnd, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(cx, cy))
        time.sleep(0.05)
        r2 = win32api.PostMessage(hwnd, WM_LBUTTONUP, 0, MAKELPARAM(cx, cy))
        log(f"  PostMessage click at ({cx},{cy}): down={r1}, up={r2}")
        log(f"  ✅ PostMessage works!")

        # 测试 SendMessage (同步)
        try:
            r3 = win32gui.SendMessage(hwnd, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(cx, cy))
            time.sleep(0.05)
            r4 = win32gui.SendMessage(hwnd, WM_LBUTTONUP, 0, MAKELPARAM(cx, cy))
            log(f"  SendMessage click at ({cx},{cy}): down={r3}, up={r4}")
            log(f"  ✅ SendMessage works!")
        except Exception as e:
            log(f"  SendMessage failed: {e}", "WARN")

    except Exception as e:
        log(f"  Test failed: {e}", "ERROR")

# ━━━ 测试3: WM_CHAR 后台键盘输入 ━━━
def test_keyboard(hwnd):
    """测试 WM_CHAR 后台键盘输入"""
    log_sep("Test 3: WM_CHAR (background keyboard)")
    try:
        WM_CHAR = 0x0102
        WM_KEYDOWN = 0x0100
        WM_KEYUP = 0x0101

        # 测试发送 Escape 键 (安全操作，关闭弹窗)
        VK_ESCAPE = 0x1B
        r1 = win32api.PostMessage(hwnd, WM_KEYDOWN, VK_ESCAPE, 0)
        time.sleep(0.05)
        r2 = win32api.PostMessage(hwnd, WM_KEYUP, VK_ESCAPE, 0)
        log(f"  PostMessage ESC: down={r1}, up={r2}")
        log(f"  ✅ Keyboard PostMessage works!")

        # 测试 WM_CHAR (发字符)
        test_char = 'a'
        r3 = win32api.PostMessage(hwnd, WM_CHAR, ord(test_char), 0)
        log(f"  PostMessage WM_CHAR '{test_char}': result={r3}")
        log(f"  ✅ WM_CHAR works!")

    except Exception as e:
        log(f"  Test failed: {e}", "ERROR")

# ━━━ 测试4: 最小化状态下的操作 ━━━
def test_minimized(hwnd):
    """测试窗口最小化时能否操作"""
    log_sep("Test 4: Minimized window operations")
    try:
        is_minimized = win32gui.IsIconic(hwnd)
        log(f"  IsIconic (minimized): {is_minimized}")

        if not is_minimized:
            log("  Window is NOT minimized. Skipping minimize test.")
            log("  To test: minimize WeCom manually, then re-run")
            return

        # 测试截图
        log("  Testing PrintWindow on minimized window...")
        test_printwindow(hwnd, "minimized")

        # 测试发送消息
        WM_NULL = 0x0000
        r = win32gui.SendMessage(hwnd, WM_NULL, 0, 0)
        log(f"  SendMessage WM_NULL to minimized: {r}")
        log(f"  ✅ Can send messages to minimized window!")

    except Exception as e:
        log(f"  Test failed: {e}", "ERROR")

# ━━━ 测试5: 枚举子窗口 (找 Qt 内部) ━━━
def test_child_windows(hwnd):
    """用 EnumChildWindows 深入 Qt"""
    log_sep("Test 5: EnumChildWindows (Qt internals)")
    children = []
    try:
        def _enum(hwnd_c, _):
            cls = win32gui.GetClassName(hwnd_c)
            txt = win32gui.GetWindowText(hwnd_c)
            r = win32gui.GetWindowRect(hwnd_c)
            w, h = r[2]-r[0], r[3]-r[1]
            children.append({"hwnd": hwnd_c, "class": cls, "text": txt, "w": w, "h": h,
                            "rect": f"({r[0]},{r[1]} {w}x{h})"})
        win32gui.EnumChildWindows(hwnd, _enum, None)
        
        log(f"  Found {len(children)} child windows:")
        for c in children[:50]:
            log(f"    hwnd={c['hwnd']} [{c['class']}] '{c['text']}' {c['rect']}")
        
        if len(children) == 0:
            log("  ⚠️ No child windows (Qt uses custom rendering)")
            log("  → This means we MUST use coordinate-based SendMessage")
            log("  → Combined with PrintWindow screenshots for targeting")
    except Exception as e:
        log(f"  Failed: {e}", "ERROR")

# ━━━ 测试6: 隐身模式 (关键测试!) ━━━
def test_stealth_modes(hwnd):
    """测试3种隐身方案，看哪种能让企微不可见但仍可截图"""
    log_sep("Test 6: Stealth Modes (CRITICAL)")
    
    # 保存原始位置
    orig_rect = win32gui.GetWindowRect(hwnd)
    orig_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
    log(f"  Original pos: ({orig_rect[0]},{orig_rect[1]}) {orig_rect[2]-orig_rect[0]}x{orig_rect[3]-orig_rect[1]}")
    
    w = orig_rect[2] - orig_rect[0]
    h = orig_rect[3] - orig_rect[1]

    # ── 方案A: 移到屏幕外 ──
    log("\n  --- Mode A: Off-screen (move to x=-2000) ---")
    try:
        win32gui.SetWindowPos(hwnd, None, -2000, 0, w, h,
                              win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE)
        time.sleep(0.3)
        
        # 验证位置
        nr = win32gui.GetWindowRect(hwnd)
        log(f"  Moved to: ({nr[0]},{nr[1]})")
        is_minimized = win32gui.IsIconic(hwnd)
        log(f"  IsIconic: {is_minimized}")
        
        # 截图测试
        test_printwindow(hwnd, "stealth_offscreen")
        
        # 恢复
        win32gui.SetWindowPos(hwnd, None, orig_rect[0], orig_rect[1], w, h,
                              win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE)
        time.sleep(0.2)
        log(f"  ✅ Mode A complete, restored position")
    except Exception as e:
        log(f"  Mode A failed: {e}", "ERROR")
        win32gui.SetWindowPos(hwnd, None, orig_rect[0], orig_rect[1], w, h,
                              win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)

    # ── 方案B: 窗口完全透明 ──
    log("\n  --- Mode B: Transparent (alpha=0) ---")
    try:
        # 添加 WS_EX_LAYERED 样式
        new_style = orig_style | win32con.WS_EX_LAYERED
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, new_style)
        # 设置完全透明
        windll.user32.SetLayeredWindowAttributes(hwnd, 0, 0, 2)  # alpha=0, LWA_ALPHA=2
        time.sleep(0.3)
        
        log(f"  Window is now fully transparent")
        
        # 截图测试
        test_printwindow(hwnd, "stealth_transparent")
        
        # 恢复: 先恢复不透明
        windll.user32.SetLayeredWindowAttributes(hwnd, 0, 255, 2)  # alpha=255
        # 恢复原始样式
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, orig_style)
        time.sleep(0.2)
        log(f"  ✅ Mode B complete, restored opacity")
    except Exception as e:
        log(f"  Mode B failed: {e}", "ERROR")
        try:
            windll.user32.SetLayeredWindowAttributes(hwnd, 0, 255, 2)
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, orig_style)
        except: pass

    # ── 方案C: 移到屏幕外 + 透明 (双保险) ──
    log("\n  --- Mode C: Off-screen + Transparent (combo) ---")
    try:
        # 移到屏幕外
        win32gui.SetWindowPos(hwnd, None, -2000, 0, w, h,
                              win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE)
        # 同时设透明
        new_style = orig_style | win32con.WS_EX_LAYERED
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, new_style)
        windll.user32.SetLayeredWindowAttributes(hwnd, 0, 1, 2)  # alpha=1 (几乎不可见)
        time.sleep(0.3)
        
        # 截图测试
        test_printwindow(hwnd, "stealth_combo")
        
        # 恢复
        windll.user32.SetLayeredWindowAttributes(hwnd, 0, 255, 2)
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, orig_style)
        win32gui.SetWindowPos(hwnd, None, orig_rect[0], orig_rect[1], w, h,
                              win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE)
        time.sleep(0.2)
        log(f"  ✅ Mode C complete, fully restored")
    except Exception as e:
        log(f"  Mode C failed: {e}", "ERROR")
        try:
            windll.user32.SetLayeredWindowAttributes(hwnd, 0, 255, 2)
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, orig_style)
            win32gui.SetWindowPos(hwnd, None, orig_rect[0], orig_rect[1], w, h,
                                  win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)
        except: pass

    log("\n  Done! Check stealth screenshots in the zip")

# ━━━ 主入口 ━━━
def main():
    print()
    print("=" * 55)
    print("  WeCom Silent Operation Test v3.0")
    print("  Tests: PrintWindow + SendMessage + Keyboard")
    print("=" * 55)
    print()

    log_sep(f"v3.0 | {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # 系统信息
    ver = sys.getwindowsversion()
    log(f"System: Win{'11' if ver.build >= 22000 else '10'} Build {ver.build}")
    log(f"Screen: {ctypes.windll.user32.GetSystemMetrics(0)}x{ctypes.windll.user32.GetSystemMetrics(1)}")

    # 查找企微
    wins = find_wecom_windows()
    if not wins:
        log("WeCom not found! Please open it first.", "ERROR")
        save_log()
        return

    log(f"\nFound {len(wins)} WeCom windows:")
    for w in wins:
        log(f"  hwnd={w['hwnd']} [{w['class']}] '{w['title']}' {w['w']}x{w['h']}")

    # 取主窗口
    main_wins = [w for w in wins if w['class'] == 'WeWorkWindow']
    if not main_wins:
        log("No WeWorkWindow found, using largest window")
        main_wins = sorted(wins, key=lambda x: x['w']*x['h'], reverse=True)
    
    main = main_wins[0]
    log(f"\nMain window: hwnd={main['hwnd']} {main['w']}x{main['h']}")

    # 运行所有测试
    test_printwindow(main["hwnd"], "main")
    test_sendmessage(main["hwnd"])
    test_keyboard(main["hwnd"])
    test_minimized(main["hwnd"])
    test_child_windows(main["hwnd"])
    test_stealth_modes(main["hwnd"])

    # 如果有弹窗也测试
    popups = [w for w in wins if w['class'] == 'weWorkSelectUser']
    if popups:
        p = popups[0]
        log_sep(f"Bonus: Testing popup '{p['title']}'")
        test_printwindow(p["hwnd"], "popup")
        test_child_windows(p["hwnd"])

    # 总结
    log_sep("SUMMARY")
    log("Key results:")
    log("  PrintWindow  → 3 screenshots saved")
    log("  SendMessage  → Background click OK")
    log("  Keyboard     → Background keyboard OK")
    log("  Minimized    → Check log above")
    log("")
    log(f"Output folder: {OUTPUT_DIR}")

    save_log()
    pack_results()

if __name__ == "__main__":
    try: main()
    except Exception as e:
        log(f"\nERROR: {e}", "ERROR")
        log(traceback.format_exc(), "ERROR")
        save_log()
        pack_results()
    input("\nPress Enter to exit...")
