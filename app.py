# -*- coding: utf-8 -*-
import os, sys, time, datetime, json, random, threading
import urllib.request, urllib.error, urllib.parse
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

if sys.platform != "win32": sys.exit(1)
try: import win32gui, win32api, win32con
except: pass

# 预加载 OCR 模块 (避免建群时等待)
print("  [启动] 加载 wecom_auto 模块...")
try:
    _t0 = time.time()
    from wecom_auto import WeComWindow as _WeComOCR, set_zhipu_api_key as _set_ocr_key
    print(f"  [启动] ✅ wecom_auto 加载完成 ({time.time()-_t0:.1f}s)")
except Exception as _e:
    print(f"  [启动] ⚠️ wecom_auto 加载失败: {_e}")
    _WeComOCR = None
    _set_ocr_key = None

# EXE 打包后 __file__ 指向临时目录, config 应在 EXE 同级目录
if getattr(sys, 'frozen', False):
    SCRIPT_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.json")
STATE_FILE = os.path.join(SCRIPT_DIR, "local_state.json")

class ConfigManager:
    def __init__(self):
        self.config = {
            "corp_id": "",
            "secret": "",
            "zhipu_api_key": "",
            "fixed_members": [],
            "check_interval": 60,
            "group_owner": "",
            "target_userid": "",
            "auto_start": False
        }
        self.state = {"processed_customers": []}
        if os.path.exists(CONFIG_FILE):
            try: self.config.update(json.load(open(CONFIG_FILE, "r", encoding="utf-8")))
            except: pass
        if os.path.exists(STATE_FILE):
            try: self.state.update(json.load(open(STATE_FILE, "r", encoding="utf-8")))
            except: pass
        # 注入智谱 OCR API Key 到 wecom_auto 模块
        if self.config.get("zhipu_api_key") and _set_ocr_key:
            _set_ocr_key(self.config["zhipu_api_key"])
    def save_config(self):
        json.dump(self.config, open(CONFIG_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        # 同步更新 OCR Key
        if self.config.get("zhipu_api_key") and _set_ocr_key:
            _set_ocr_key(self.config["zhipu_api_key"])
    def save_state(self): json.dump(self.state, open(STATE_FILE, "w", encoding="utf-8"), ensure_ascii=False)
    def is_processed(self, uid): return uid in self.state["processed_customers"]
    def mark_processed(self, uid):
        if uid not in self.state["processed_customers"]:
            self.state["processed_customers"].append(uid)
            self.save_state()

class APIClient:
    """企微官方 API 客户端 (直接调用 qyapi.weixin.qq.com)"""
    WECOM_BASE = "https://qyapi.weixin.qq.com/cgi-bin"

    def __init__(self, config):
        self.corp_id = config.get("corp_id", "")
        self.secret = config.get("secret", "")
        self._access_token = ""
        self._token_expires = 0
        ready = "✅ 已配置" if (self.corp_id and self.secret) else "❌ 未配置 (需要 corp_id + secret)"
        print(f"  [API] 初始化 APIClient: {ready}")

    def _get_token(self):
        """获取/刷新 access_token (缓存 7200 秒)"""
        import time as _t
        if self._access_token and _t.time() < self._token_expires:
            return self._access_token
        if not self.corp_id or not self.secret:
            print("  [API] ❌ corp_id 或 secret 未配置")
            return None
        url = f"{self.WECOM_BASE}/gettoken?corpid={self.corp_id}&corpsecret={self.secret}"
        try:
            t0 = _t.time()
            raw = urllib.request.urlopen(url, timeout=15).read().decode("utf-8")
            res = json.loads(raw)
            if res.get("errcode", 0) != 0:
                print(f"  [API] ❌ 获取 token 失败: {res.get('errmsg', 'unknown')}")
                return None
            self._access_token = res["access_token"]
            self._token_expires = _t.time() + res.get("expires_in", 7200) - 300  # 提前5分钟刷新
            print(f"  [API] ✅ access_token 获取成功 ({_t.time()-t0:.1f}s, 有效期={res.get('expires_in',0)}s)")
            return self._access_token
        except Exception as e:
            print(f"  [API] ❌ 获取 token 异常: {str(e)[:80]}")
            return None

    def _api_get(self, path, params=None, timeout=30):
        """GET 请求企微 API"""
        token = self._get_token()
        if not token:
            return None
        url = f"{self.WECOM_BASE}{path}?access_token={token}"
        if params:
            url += "&" + urllib.parse.urlencode(params)
        try:
            raw = urllib.request.urlopen(url, timeout=timeout).read().decode("utf-8")
            res = json.loads(raw)
            if res.get("errcode", 0) != 0:
                print(f"  [API] ❌ {path}: errcode={res.get('errcode')}, {res.get('errmsg', '')}")
                return None
            return res
        except Exception as e:
            print(f"  [API] ❌ {path}: {str(e)[:80]}")
            return None

    def _api_post(self, path, body, timeout=30):
        """POST 请求企微 API"""
        token = self._get_token()
        if not token:
            return None
        url = f"{self.WECOM_BASE}{path}?access_token={token}"
        data = json.dumps(body).encode("utf-8")
        req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
        try:
            raw = urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8")
            res = json.loads(raw)
            if res.get("errcode", 0) != 0:
                print(f"  [API] ❌ {path}: errcode={res.get('errcode')}, {res.get('errmsg', '')}")
                return None
            return res
        except Exception as e:
            print(f"  [API] ❌ {path}: {str(e)[:80]}")
            return None

    def get_members(self):
        """获取所有在职员工列表 (通讯录)"""
        print(f"  [API] 🔍 获取员工列表...")
        res = self._api_get("/user/simplelist", {"department_id": 1, "fetch_child": 1})
        if not res:
            return []
        members = res.get("userlist", [])
        # 统一字段格式: 确保每个成员有 'name', 'userid', 'status'
        result = []
        for m in members:
            result.append({
                "userid": m.get("userid", ""),
                "name": m.get("name", ""),
                "status": m.get("status", 1)  # 1=已激活
            })
        print(f"  [API] 📋 员工列表: {len(result)} 人")
        return result

    def get_groups(self):
        """获取客户群列表"""
        print(f"  [API] 🔍 获取群聊列表...")
        all_groups = []
        cursor = ""
        while True:
            body = {"status_filter": 0, "limit": 100}
            if cursor:
                body["cursor"] = cursor
            res = self._api_post("/externalcontact/groupchat/list", body)
            if not res:
                break
            group_list = res.get("group_chat_list", [])
            all_groups.extend(group_list)
            cursor = res.get("next_cursor", "")
            if not cursor:
                break
        # 获取每个群的详情 (群名等)
        result = []
        for g in all_groups:
            chat_id = g.get("chat_id", "")
            detail = self._api_post("/externalcontact/groupchat/get", {"chat_id": chat_id})
            if detail and "group_chat" in detail:
                gc = detail["group_chat"]
                result.append({
                    "chat_id": chat_id,
                    "name": gc.get("name", ""),
                    "owner": gc.get("owner", ""),
                    "member_count": len(gc.get("member_list", []))
                })
            else:
                result.append({"chat_id": chat_id, "name": "", "owner": "", "member_count": 0})
        print(f"  [API] 📋 群聊列表: {len(result)} 个")
        return result

    def get_contacts(self, uid, log_fn=None):
        """获取员工的外部联系人 (客户列表)"""
        import time as _t
        _log = log_fn or (lambda m: print(f"  [API] {m}"))
        _log(f"📡 正在获取 {uid} 的外部联系人...")
        t0 = _t.time()

        # Step 1: 获取 external_userid 列表
        res = self._api_get("/externalcontact/list", {"userid": uid})
        if not res:
            _log(f"⚠️ {uid}: 获取联系人列表失败 (可能不在「客户联系」权限范围内)")
            return []
        ext_ids = res.get("external_userid", [])
        _log(f"📋 {uid}: 共 {len(ext_ids)} 个外部联系人, 开始获取详情...")

        # Step 2: 逐个获取详情
        contacts = []
        for i, ext_id in enumerate(ext_ids):
            detail = self._api_get("/externalcontact/get", {"external_userid": ext_id}, timeout=10)
            if detail and "external_contact" in detail:
                ec = detail["external_contact"]
                contacts.append({
                    "external_userid": ext_id,
                    "name": ec.get("name", ""),
                    "type": ec.get("type", 0),
                    "corp_name": ec.get("corp_name", ""),
                })
            if (i + 1) % 20 == 0:
                _log(f"   进度: {i+1}/{len(ext_ids)}")
            _t.sleep(0.05)  # 避免频率限制

        elapsed = _t.time() - t0
        _log(f"✅ {uid}: {len(contacts)} 人 ({elapsed:.1f}s)")
        return contacts

class WeComWindow:
    # 参考坐标基于 900x650 的窗口大小
    REF_W = 900
    REF_H = 650

    def __init__(self):
        self.hwnd = None; self.w = 0; self.h = 0
        print(f"  [WC] WeComWindow 实例已创建")

    def find(self):
        wins = []
        win32gui.EnumWindows(lambda h, _: wins.append(h) if win32gui.GetClassName(h)=="WeWorkWindow" else None, None)
        if not wins:
            print(f"  [WC] ❌ 未找到企业微信窗口 (WeWorkWindow)")
            return False
        self.hwnd = wins[0]
        if win32gui.IsIconic(self.hwnd):
            print(f"  [WC] 窗口已最小化, 正在恢复...")
            win32gui.ShowWindow(self.hwnd, win32con.SW_RESTORE); time.sleep(0.5)
        cr = win32gui.GetClientRect(self.hwnd)
        self.w, self.h = cr[2], cr[3]
        rect = win32gui.GetWindowRect(self.hwnd)
        print(f"  [WC] ✅ 企微窗口: hwnd={self.hwnd}, 客户区={self.w}x{self.h}, 位置=({rect[0]},{rect[1]})-({rect[2]},{rect[3]})")
        return True

    def _jitter(self, val, radius=3):
        """坐标抖动: ±radius 随机偏移"""
        return val + random.randint(-radius, radius)

    def _human_delay(self, base=0.3, variance=0.15):
        """人类反应延时: 高斯分布"""
        delay = max(0.05, random.gauss(base, variance))
        time.sleep(delay)

    def _client_to_screen(self, hwnd, cx, cy):
        """客户区坐标 → 屏幕绝对坐标"""
        import ctypes
        import ctypes.wintypes
        pt = ctypes.wintypes.POINT(cx, cy)
        ctypes.windll.user32.ClientToScreen(hwnd, ctypes.byref(pt))
        return pt.x, pt.y

    def silent_click(self, hwnd, cx, cy):
        """后台点击 (SendMessage) + 人类节奏模拟:
        - SendMessage 确保后台窗口也能点击 (不需要前台焦点)
        - 坐标抖动 ±2px
        - 每步随机延时 (高斯分布)
        - 完整 Chromium hit-test 序列
        """
        cx = self._jitter(cx, 2)
        cy = self._jitter(cy, 2)
        lp = (cy << 16) | (cx & 0xFFFF)

        # Step 1: WM_MOUSEACTIVATE
        win32gui.SendMessage(hwnd, win32con.WM_MOUSEACTIVATE, hwnd,
                             (win32con.WM_LBUTTONDOWN << 16) | 1)
        time.sleep(random.uniform(0.02, 0.06))

        # Step 2: WM_SETCURSOR
        win32gui.SendMessage(hwnd, 0x0020, hwnd,
                             (win32con.WM_LBUTTONDOWN << 16) | 1)
        time.sleep(random.uniform(0.02, 0.06))

        # Step 3: WM_MOUSEMOVE (hover 预热)
        win32gui.SendMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lp)
        time.sleep(random.uniform(0.12, 0.22))

        # Step 4: WM_LBUTTONDOWN
        win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp)
        time.sleep(random.uniform(0.05, 0.12))

        # Step 5: WM_LBUTTONUP
        win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lp)
        self._human_delay(0.25, 0.10)

    def silent_type(self, hwnd, text):
        """后台文字输入 (WM_IME_CHAR) + 人类打字节奏"""
        for ch in text:
            win32gui.SendMessage(hwnd, 0x0286, ord(ch), 0)
            time.sleep(max(0.02, random.gauss(0.07, 0.025)))
        time.sleep(1.0)

    def silent_backspace(self, hwnd, count=20):
        for _ in range(count):
            win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_BACK, 0)
            win32gui.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_BACK, 0)
        time.sleep(0.3)

    def _clear_input(self, hwnd):
        """强力清空输入框: Ctrl+A→Delete + End→多次Backspace"""
        # 方法1: Ctrl+A 全选 → Delete
        win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
        time.sleep(random.uniform(0.02, 0.05))
        win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, 0x41, 0)  # A
        time.sleep(random.uniform(0.02, 0.04))
        win32gui.SendMessage(hwnd, win32con.WM_KEYUP, 0x41, 0)
        time.sleep(random.uniform(0.01, 0.03))
        win32gui.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
        time.sleep(random.uniform(0.05, 0.10))
        win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_DELETE, 0)
        time.sleep(random.uniform(0.02, 0.05))
        win32gui.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_DELETE, 0)
        time.sleep(0.1)
        
        # 方法2: End键到末尾 → 暴力 Backspace 30 次
        win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_END, 0)
        win32gui.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_END, 0)
        time.sleep(0.05)
        for _ in range(30):
            win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_BACK, 0)
            win32gui.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_BACK, 0)
            time.sleep(0.02)

    def _force_foreground(self, hwnd):
        """强制窗口到最前台, 绕过 Windows 前台锁限制。"""
        import ctypes
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        
        # 获取当前前台窗口的线程 ID
        fg_hwnd = user32.GetForegroundWindow()
        fg_tid = user32.GetWindowThreadProcessId(fg_hwnd, None)
        my_tid = kernel32.GetCurrentThreadId()
        
        # 附加到前台线程, 获取前台权限
        if fg_tid != my_tid:
            user32.AttachThreadInput(my_tid, fg_tid, True)
        
        # 模拟 Alt 键释放, 解除 SetForegroundWindow 的限制
        user32.keybd_event(0x12, 0, 0x0002, 0)  # ALT up
        
        # TOPMOST 置顶
        user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 0x0040 | 0x0001 | 0x0002)
        
        # 激活并前台化
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        user32.SetForegroundWindow(hwnd)
        user32.BringWindowToTop(hwnd)
        
        # 解除线程附加
        if fg_tid != my_tid:
            user32.AttachThreadInput(my_tid, fg_tid, False)
        
        time.sleep(0.8)

    def real_click(self, hwnd, cx, cy):
        """前台真实鼠标点击: 用于 Chromium overlay (如群管理面板)。
        SendMessage 无法点击 overlay 元素, 必须用真实鼠标事件。
        """
        import ctypes, ctypes.wintypes
        pt = ctypes.wintypes.POINT(cx, cy)
        ctypes.windll.user32.ClientToScreen(hwnd, ctypes.byref(pt))
        
        # 强制窗口到最前台
        self._force_foreground(hwnd)
        
        # 移动鼠标并点击
        ctypes.windll.user32.SetCursorPos(pt.x, pt.y)
        time.sleep(0.15)
        ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)  # LEFTDOWN
        time.sleep(random.uniform(0.06, 0.12))
        ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)  # LEFTUP
        time.sleep(0.5)
        
        # 恢复非 TOPMOST
        ctypes.windll.user32.SetWindowPos(
            hwnd, -2, 0, 0, 0, 0, 0x0001 | 0x0002)

    def click_main(self, cx, cy, desc=""):
        """Click on the main WeCom window (coordinates in 900x650 reference)"""
        rx = int(cx * self.w / self.REF_W)
        ry = int(cy * self.h / self.REF_H)
        self.silent_click(self.hwnd, rx, ry)

    def find_popup(self):
        pops = []
        win32gui.EnumWindows(lambda h, _: pops.append(h) if win32gui.IsWindowVisible(h) and win32gui.GetClassName(h) == "weWorkSelectUser" else None, None)
        return pops[0] if pops else None

    def create_grp(self, cust, fixed, log):
        """OCR 增强版建群流程"""
        t_start = time.time()

        # 初始化 OCR 引擎 (使用启动时预加载的模块)
        log(f"    [1/7] 初始化 OCR 引擎...")
        if _WeComOCR is None:
            log(f"    [1/7] ❌ OCR 模块未加载, 无法建群")
            return False
        _wc = _WeComOCR()
        _wc.hwnd = self.hwnd
        _wc.pid = self._get_pid()
        _wc.width, _wc.height = win32gui.GetClientRect(self.hwnd)[2], win32gui.GetClientRect(self.hwnd)[3]
        log(f"    [1/7] OCR 就绪 ({time.time()-t_start:.1f}s, 窗口={_wc.width}x{_wc.height})")

        # Step 1: 点击消息 tab (OCR 定位, 侧栏按钮 y 坐标不随窗口缩放)
        log(f"    [2/8] 点击消息 Tab...")
        msg_clicked = False
        try:
            main_items, _ = _wc.ocr_scan(self.hwnd, "main_tab")
            log(f"    [2/8] OCR: {len(main_items)} 项")
            msg_btn = _wc.ocr_find(main_items, "消息")
            if msg_btn:
                log(f"    [2/8] OCR 找到「消息」: ({msg_btn['cx']},{msg_btn['cy']})")
                self.silent_click(self.hwnd, msg_btn['cx'], msg_btn['cy'])
                msg_clicked = True
        except Exception as e:
            log(f"    [2/8] OCR 失败: {e}")
        if not msg_clicked:
            # 回退: 侧栏「消息」在 x≈30, y≈106 (固定值, 不缩放)
            self.silent_click(self.hwnd, 30, 106)
            log(f"    [2/8] 坐标回退 (30,106)")
        self._human_delay(0.8, 0.3)

        # Step 2: 点击 [+] 按钮
        # 智谱 OCR 能将 + 图标识别为 "十", 直接定位
        log(f"    [3/8] 点击 + 按钮...")
        plus_btn = _wc.ocr_find(main_items, "十") or _wc.ocr_find(main_items, "+")
        if plus_btn and plus_btn['cy'] < 60:
            # 确保是顶部搜索栏旁边的 +, 不是其他位置的 "十"
            log(f"    [3/8] OCR 定位 + 按钮: ({plus_btn['cx']},{plus_btn['cy']})")
            self.silent_click(self.hwnd, plus_btn['cx'], plus_btn['cy'])
        else:
            # 回退: 用面板右端推算
            panel_right = 0
            for it in main_items:
                if it['cy'] > 50 and it['cy'] < self.h * 0.5 and it['x2'] < self.w * 0.4:
                    if it['x2'] > panel_right:
                        panel_right = it['x2']
            search_btn = _wc.ocr_find(main_items, "搜索") or _wc.ocr_find(main_items, "搜")
            plus_y = search_btn['cy'] if search_btn else 37
            plus_x = (panel_right - 15) if panel_right > 100 else int(self.w * 0.22)
            log(f"    [3/8] 坐标回退 + ({plus_x},{plus_y})")
            self.silent_click(self.hwnd, plus_x, plus_y)

        self._human_delay(1.5, 0.5)

        # Step 3: 检测弹窗
        popup = self.find_popup()
        if not popup:
            log(f"    [3/8] ⚠️ 弹窗未出现, 重试...")
            # 重试: 用回退坐标
            panel_w = int(self.w * 0.31)
            self.silent_click(self.hwnd, panel_w - 25, 37)
            self._human_delay(2.0, 0.5)
            popup = self.find_popup()
        if not popup:
            self._human_delay(2.0, 0.5)
            popup = self.find_popup()
        if not popup:
            log(f"    [3/8] ❌ 建群弹窗未打开!")
            return False

        pcr = win32gui.GetClientRect(popup)
        log(f"    [3/8] ✅ 弹窗已打开 ({pcr[2]}x{pcr[3]})")

        # Step 4: 逐个搜索并选中成员
        all_members = [cust] + fixed
        log(f"    [4/8] 添加 {len(all_members)} 名成员")
        
        for i, m in enumerate(all_members, 1):
            log(f"    [4/8] ({i}/{len(all_members)}) 搜索: {m}")

            if i > 1:
                self._human_delay(0.5, 0.3)

            # 先清空搜索框
            self.silent_click(popup, 160, 40)
            self._human_delay(0.3, 0.1)
            self._clear_input(popup)
            self._human_delay(0.3, 0.1)

            # 输入搜索
            self.silent_type(popup, m)
            self._human_delay(1.5, 0.5)

            # 勾选第一个搜索结果
            self.silent_click(popup, 25, 95)
            self._human_delay(0.5, 0.2)
            
            # 勾选后清空搜索框
            self.silent_click(popup, 160, 40)
            self._human_delay(0.2, 0.1)
            self._clear_input(popup)
            self._human_delay(0.3, 0.1)
            
            log(f"    [4/8] ({i}/{len(all_members)}) ✅ 已勾选 {m}")

        # Step 5: 清空搜索框 (已在每次勾选后清空, 这里二次确认)
        log(f"    [5/8] 清空搜索框...")
        self.silent_click(popup, 160, 40)
        self._human_delay(0.15, 0.05)
        self._clear_input(popup)
        self._human_delay(0.3, 0.1)

        # Step 6: 点击「完成」按钮 (多策略 + 验证弹窗关闭)
        log(f"    [6/8] 点击完成按钮...")
        
        # 策略1: OCR 定位
        clicked = False
        try:
            popup_items, _ = _wc.ocr_scan(popup, "grp_confirm")
            log(f"    [6/8] OCR 扫描到 {len(popup_items)} 项: {[it['text'] for it in popup_items[:8]]}")
            done_btn = _wc.ocr_find(popup_items, "完成")
            if done_btn:
                log(f"    [6/8] ✅ OCR 找到「完成」: ({done_btn['cx']},{done_btn['cy']})")
                self.silent_click(popup, done_btn['cx'], done_btn['cy'])
                clicked = True
        except Exception as ocr_err:
            log(f"    [6/8] OCR 不可用: {ocr_err}")

        # 策略2: 坐标比例
        if not clicked:
            pr = win32gui.GetClientRect(popup)
            pw, ph = pr[2], pr[3]
            done_x = int(pw * 0.63)
            done_y = int(ph * 0.93)
            log(f"    [6/8] 坐标点击: ({done_x},{done_y}) [弹窗={pw}x{ph}]")
            self.silent_click(popup, done_x, done_y)

        time.sleep(2.0)

        # 验证弹窗是否关闭, 如果没关则重试
        for retry in range(3):
            if not win32gui.IsWindow(popup) or not win32gui.IsWindowVisible(popup):
                log(f"    [6/8] ✅ 弹窗已关闭")
                break
            log(f"    [6/8] ⚠️ 弹窗仍在, 重试 {retry+1}/3...")
            if retry == 0:
                # 重试: 用 Enter 键确认
                win32gui.SendMessage(popup, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                win32gui.SendMessage(popup, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
            elif retry == 1:
                # 再次点击完成按钮 (调整坐标: 更靠左一些)
                pr = win32gui.GetClientRect(popup)
                pw, ph = pr[2], pr[3]
                self.silent_click(popup, int(pw * 0.50), int(ph * 0.93))
            else:
                # 最后手段: Escape 关闭
                win32gui.SendMessage(popup, win32con.WM_KEYDOWN, win32con.VK_ESCAPE, 0)
                win32gui.SendMessage(popup, win32con.WM_KEYUP, win32con.VK_ESCAPE, 0)
            time.sleep(2.0)

        # 随机延时 (降低风控风险)
        delay = random.uniform(1.5, 3.0)
        log(f"    [6/8] 等待 {delay:.1f}s (防风控)...")
        time.sleep(delay)

        # Step 7: 关闭可能残留的弹窗
        log(f"    [7/8] 检查并关闭残留弹窗...")
        for _ in range(3):
            remaining = self.find_popup()
            if not remaining:
                break
            log(f"    [7/8] 发现残留弹窗 {remaining}, 关闭...")
            win32gui.SendMessage(remaining, win32con.WM_KEYDOWN, win32con.VK_ESCAPE, 0)
            win32gui.SendMessage(remaining, win32con.WM_KEYUP, win32con.VK_ESCAPE, 0)
            time.sleep(1.0)

        # Step 8: 设置群管理 - 禁止互加好友 (OCR 定位)
        log(f"    [8/8] 设置群管理 (禁止互加好友)...")
        try:
            self._set_group_privacy(log, _wc)
        except Exception as priv_err:
            log(f"    [8/8] ⚠️ 群管理设置未完成: {priv_err}")

        # 最终清理: 关闭所有残留弹窗
        for _ in range(3):
            p = self.find_popup()
            if not p:
                break
            win32gui.SendMessage(p, win32con.WM_KEYDOWN, win32con.VK_ESCAPE, 0)
            win32gui.SendMessage(p, win32con.WM_KEYUP, win32con.VK_ESCAPE, 0)
            time.sleep(0.5)

        elapsed = time.time() - t_start
        log(f"    ✅ 建群流程完成, 总耗时 {elapsed:.1f}s")
        return True

    def _set_group_privacy(self, log, _wc):
        """群管理设置: OCR 优先定位, 坐标回退。滚轮用 SendMessage 确保后台可用。
        
        流程:
        1. 点击 ··· → 展开群设置面板
        2. 面板内向下滚动 → 找到「群管理」
        3. 点击「群管理」→ 打开弹窗
        4. 在弹窗中找到「禁止互加好友」→ 点击开关
        """
        ocr_ok = _wc is not None
        
        # Step 1: 点击 ··· 展开群设置面板
        # ··· 按钮在窗口右上角, 只扫描右上区域
        dots_btn = None
        if ocr_ok:
            try:
                right_top, _ = _wc.ocr_scan_region(
                    self.hwnd, int(self.w * 0.6), 0, self.w, int(self.h * 0.08), "dots_btn")
                log(f"       -> 右上区域 OCR: {len(right_top)} 项")
                dots_btn = _wc.ocr_find(right_top, "···") or _wc.ocr_find(right_top, "...")
            except Exception:
                pass
        
        if dots_btn:
            log(f"       -> OCR 找到 ··· ({dots_btn['cx']},{dots_btn['cy']})")
            self.silent_click(self.hwnd, dots_btn['cx'], dots_btn['cy'])
        else:
            # 回退坐标: ··· 按钮在群聊区域右上角
            rx = int(self.w * 0.945)
            ry = int(self.h * 0.04)
            log(f"       -> 坐标点击 ··· ({rx},{ry})")
            self.silent_click(self.hwnd, rx, ry)
        self._human_delay(2.0, 0.5)
        
        # 关闭意外弹窗 (如果误点了添加成员)
        accidental = self.find_popup()
        if accidental:
            log(f"       -> ⚠️ 意外弹窗, 关闭...")
            win32gui.SendMessage(accidental, win32con.WM_KEYDOWN, win32con.VK_ESCAPE, 0)
            win32gui.SendMessage(accidental, win32con.WM_KEYUP, win32con.VK_ESCAPE, 0)
            self._human_delay(1.0, 0.3)
        
        # Step 2: 面板内向下滚动 (SendMessage 后台滚轮)
        scroll_x = int(self.w * 0.87)
        scroll_y = int(self.h * 0.50)
        lp_scroll = (scroll_y << 16) | (scroll_x & 0xFFFF)
        log(f"       -> 面板滚动 ({scroll_x},{scroll_y})")
        for _ in range(5):
            wp = ((-120) & 0xFFFFFFFF) << 16  # WHEEL_DELTA=-120 向下
            win32gui.SendMessage(self.hwnd, win32con.WM_MOUSEWHEEL, wp, lp_scroll)
            time.sleep(random.uniform(0.25, 0.4))
        self._human_delay(0.5, 0.2)
        
        # Step 3: OCR 找「群管理」, 找不到则坐标回退
        mgmt_btn = None
        if ocr_ok:
            try:
                # 群管理在右侧面板, 只扫描右半部分
                panel_items, _ = _wc.ocr_scan_region(
                    self.hwnd, int(self.w * 0.6), int(self.h * 0.3), self.w, self.h, "group_panel")
                log(f"       -> 面板区域 OCR: {len(panel_items)} 项: {[it['text'] for it in panel_items[:12]]}")
                mgmt_btn = _wc.ocr_find(panel_items, "群管理")
                if not mgmt_btn:
                    for item in panel_items:
                        if "群管理" in item.get("text", ""):
                            mgmt_btn = item
                            break
            except Exception as e:
                log(f"       -> OCR 扫描失败: {e}")
        
        if mgmt_btn:
            log(f"       -> ✅ OCR 找到「群管理」({mgmt_btn['cx']},{mgmt_btn['cy']})")
            self.silent_click(self.hwnd, mgmt_btn['cx'], mgmt_btn['cy'])
        else:
            # 坐标回退: 群管理通常在滚动后面板中下部
            fx = int(self.w * 0.82)
            fy = int(self.h * 0.70)
            log(f"       -> 坐标点击群管理 ({fx},{fy})")
            self.silent_click(self.hwnd, fx, fy)
        self._human_delay(2.0, 0.5)
        
        # Step 4: 群管理是 Chromium CSS overlay — PrintWindow 截不到!
        # 必须用前台截图 (BitBlt 屏幕 DC) 才能捕获 overlay 内容
        self._human_delay(1.0, 0.3)
        
        toggle_btn = None
        mgmt_items = []
        if ocr_ok:
            try:
                # 前台截图: 能捕获 overlay
                mgmt_items, _ = _wc.ocr_scan_foreground(self.hwnd, "group_mgmt_fg")
                texts = [it['text'] for it in mgmt_items]
                log(f"       -> 群管理面板 OCR (前台): {texts[:25]}")
                
                # 找「禁止互相添加为联系人」— checkbox 标签
                # 注意: 过滤掉长说明文字, checkbox 标签通常 < 15 字
                for kw in ["禁止互相添加为联系人", "禁止互相添加", "互相添加为联系人",
                           "禁止互加", "互加好友"]:
                    toggle_btn = _wc.ocr_find(mgmt_items, kw)
                    # 确保不是说明文字 (太长的排除)
                    if toggle_btn and len(toggle_btn['text']) > 20:
                        toggle_btn = None
                    if toggle_btn:
                        break
                    for item in mgmt_items:
                        if kw in item.get('text', '') and len(item['text']) <= 20:
                            toggle_btn = item
                            break
                    if toggle_btn:
                        break
                
                # 面板可能需要滚动
                if not toggle_btn:
                    log(f"       -> 未找到, 在面板内滚动...")
                    scroll_x = int(self.w * 0.45)
                    scroll_y = int(self.h * 0.50)
                    lps = (scroll_y << 16) | (scroll_x & 0xFFFF)
                    for _ in range(3):
                        wp = ((-120) & 0xFFFFFFFF) << 16
                        win32gui.SendMessage(self.hwnd, win32con.WM_MOUSEWHEEL, wp, lps)
                        time.sleep(0.3)
                    time.sleep(1.0)
                    mgmt_items, _ = _wc.ocr_scan_foreground(self.hwnd, "group_mgmt_fg_scroll")
                    texts = [it['text'] for it in mgmt_items]
                    log(f"       -> 滚动后 OCR: {texts[:25]}")
                    for kw in ["禁止互相添加为联系人", "禁止互相添加", "互相添加为联系人",
                               "禁止互加", "互加好友"]:
                        toggle_btn = _wc.ocr_find(mgmt_items, kw)
                        if toggle_btn and len(toggle_btn['text']) > 20:
                            toggle_btn = None
                        if not toggle_btn:
                            for item in mgmt_items:
                                if kw in item.get('text', '') and len(item['text']) <= 20:
                                    toggle_btn = item
                                    break
                        if toggle_btn:
                            break
            except Exception as e:
                log(f"       -> OCR 扫描失败: {e}")
        
        if toggle_btn:
            txt = toggle_btn['text']
            # OCR 可能把标签和 checkbox 文字连在一起, 如:
            # "成员加好友权限□禁止互相添加为联系人"
            # checkbox ☐ 在 "禁止" 前面, 按比例计算
            pos = -1
            for kw in ["禁止互相", "禁止互加", "互加好友"]:
                idx = txt.find(kw)
                if idx > 0:
                    pos = idx
                    break
            
            if pos > 0 and len(txt) > 12:
                text_width = toggle_btn['x2'] - toggle_btn['x1']
                ratio = pos / len(txt)
                check_x = toggle_btn['x1'] + int(text_width * ratio) - 5
            else:
                check_x = toggle_btn['x1'] - 12
            
            check_y = toggle_btn['cy']
            log(f"       -> ✅ 找到「{txt}」-> 真实点击勾选框 ({check_x},{check_y})")
            # 必须用 real_click (前台鼠标), SendMessage 无法点击 Chromium overlay
            self.real_click(self.hwnd, check_x, check_y)
            self._human_delay(1.0, 0.3)
            log(f"       -> ✅ 禁止互相添加为联系人 已勾选")
        else:
            log(f"       -> ⚠️ 未找到禁止互相添加选项")
        
        # 关闭群管理面板 (× 按钮也在 overlay 上, 用 real_click)
        try:
            close_btn = _wc.ocr_find(mgmt_items, "×") if mgmt_items else None
            if close_btn and close_btn['cy'] < self.h * 0.15:
                self.real_click(self.hwnd, close_btn['cx'], close_btn['cy'])
            else:
                # 点击面板外区域关闭
                self.real_click(self.hwnd, int(self.w * 0.15), int(self.h * 0.5))
        except Exception:
            self.real_click(self.hwnd, int(self.w * 0.15), int(self.h * 0.5))
        self._human_delay(0.5, 0.2)

    def _get_pid(self):
        """Get the PID of the WeCom process."""
        try:
            _, pid = win32process.GetWindowThreadProcessId(self.hwnd)
            return pid
        except Exception:
            return None

    def screenshot_main(self, label="debug"):
        """Take a PrintWindow screenshot of main window for debugging."""
        try:
            from ctypes import windll
            cr = win32gui.GetClientRect(self.hwnd)
            w, h = cr[2], cr[3]
            if w <= 0 or h <= 0: return None
            hwnd_dc = win32gui.GetWindowDC(self.hwnd)
            import win32ui
            mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
            save_dc = mfc_dc.CreateCompatibleDC()
            bmp = win32ui.CreateBitmap()
            bmp.CreateCompatibleBitmap(mfc_dc, w, h)
            save_dc.SelectObject(bmp)
            windll.user32.PrintWindow(self.hwnd, save_dc.GetSafeHdc(), 1)
            bmpinfo = bmp.GetInfo()
            bmpstr = bmp.GetBitmapBits(True)
            from PIL import Image
            img = Image.frombuffer('RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                                   bmpstr, 'raw', 'BGRX', 0, 1)
            fpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"debug_{label}.png")
            img.save(fpath)
            win32gui.DeleteObject(bmp.GetHandle())
            save_dc.DeleteDC()
            mfc_dc.DeleteDC()
            win32gui.ReleaseDC(self.hwnd, hwnd_dc)
            return fpath
        except Exception:
            return None




class WeComAutoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WeCom Auto Agent - 双模式建群中控")
        self.root.geometry("850x700")
        self.root.configure(bg="#f4f4f9")
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TFrame', background='#ffffff')
        style.configure('TLabelFrame', background='#ffffff', padding=10)
        style.configure('TLabel', background='#ffffff', font=('Microsoft YaHei', 10))
        style.configure('TButton', font=('Microsoft YaHei', 10, 'bold'), padding=5)
        
        self.cfg = ConfigManager()
        self.api = APIClient(self.cfg.config)
        self.running = False
        self.thread = None
        self.members = []
        self.member_names = []
        self.member_map = {}
        self.current_contacts = []

        self.build_ui()
        self.log("初始化完成，正在连接 API 获取名单...")
        threading.Thread(target=self.init_data, daemon=True).start()

    def log(self, msg):
        def _append():
            ts = datetime.datetime.now().strftime('%H:%M:%S')
            self.log_area.config(state=tk.NORMAL)
            self.log_area.insert(tk.END, f"[{ts}] {msg}\n")
            self.log_area.see(tk.END)
            self.log_area.config(state=tk.DISABLED)
        self.root.after(0, _append)

    def init_data(self):
        try:
            m = self.api.get_members()
            if m:
                self.members = m
                self.member_names = [x['name'] for x in m if x.get('status') == 1]
                self.member_map = {x['name']: x['userid'] for x in m}
                self.log(f"✅ 成功获取 {len(self.member_names)} 名在职员工数据。")
                self.root.after(0, self.update_comboboxes)
            else:
                self.log("❌ 获取不到员工列表，请验证网络/Token...")
        except BaseException as e:
            self.log(f"网络异常: {e}")

    def update_comboboxes(self):
        for combo in [self.combo_target, self.combo_f1, self.combo_f2, self.combo_f3]:
            combo['values'] = self.member_names

        def set_val(combo, search_val):
            if not search_val: return
            # Try matching by checking if search_val is the exact name, or if search_val is a userid.
            for m in self.members:
                if m['userid'] == search_val or m['name'] == search_val:
                    combo.set(m['name'])
                    return
            combo.set(search_val)
                    
        set_val(self.combo_target, self.cfg.config.get("target_userid"))
        fm = self.cfg.config.get("fixed_members", [])
        if len(fm) > 0: set_val(self.combo_f1, fm[0])
        if len(fm) > 1: set_val(self.combo_f2, fm[1])
        if len(fm) > 2: set_val(self.combo_f3, fm[2])

    def save_settings(self):
        # API 配置
        self.cfg.config["corp_id"] = self.entry_corp_id.get().strip()
        self.cfg.config["secret"] = self.entry_secret.get().strip()
        self.cfg.config["zhipu_api_key"] = self.entry_zhipu_key.get().strip()
        self.cfg.config["group_owner"] = self.entry_group_owner.get().strip()
        # 建群配置
        target_name = self.combo_target.get()
        tgt_uid = self.member_map.get(target_name, target_name)
        f1 = self.combo_f1.get()
        f2 = self.combo_f2.get()
        f3 = self.combo_f3.get()
        self.cfg.config["target_userid"] = tgt_uid
        self.cfg.config["fixed_members"] = [x for x in [f1, f2, f3] if x]
        self.cfg.save_config()
        # 重新初始化 API 客户端
        self.api = APIClient(self.cfg.config)
        self.log(f"✅ 配置已保存！主理人={tgt_uid}，团队={'、'.join(self.cfg.config['fixed_members'])}")

    def build_ui(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ============== API 配置区 ==============
        frame_api = ttk.LabelFrame(self.root, text=" 🔑 API 密钥配置 ")
        frame_api.pack(fill=tk.X, padx=10, pady=3, before=notebook)

        api_row1 = ttk.Frame(frame_api)
        api_row1.pack(fill=tk.X, pady=2)
        ttk.Label(api_row1, text="企业ID:", width=12).pack(side=tk.LEFT)
        self.entry_corp_id = ttk.Entry(api_row1, width=40)
        self.entry_corp_id.pack(side=tk.LEFT, padx=5)
        self.entry_corp_id.insert(0, self.cfg.config.get("corp_id", ""))
        ttk.Label(api_row1, text="应用Secret:", width=12).pack(side=tk.LEFT, padx=(15,0))
        self.entry_secret = ttk.Entry(api_row1, width=40, show="*")
        self.entry_secret.pack(side=tk.LEFT, padx=5)
        self.entry_secret.insert(0, self.cfg.config.get("secret", ""))

        api_row2 = ttk.Frame(frame_api)
        api_row2.pack(fill=tk.X, pady=2)
        ttk.Label(api_row2, text="智谱OCR Key:", width=12).pack(side=tk.LEFT)
        self.entry_zhipu_key = ttk.Entry(api_row2, width=40, show="*")
        self.entry_zhipu_key.pack(side=tk.LEFT, padx=5)
        self.entry_zhipu_key.insert(0, self.cfg.config.get("zhipu_api_key", ""))
        ttk.Label(api_row2, text="群主姓名:", width=12).pack(side=tk.LEFT, padx=(15,0))
        self.entry_group_owner = ttk.Entry(api_row2, width=40)
        self.entry_group_owner.pack(side=tk.LEFT, padx=5)
        self.entry_group_owner.insert(0, self.cfg.config.get("group_owner", ""))

        # ============== 建群规则配置 ==============
        frame_top = ttk.LabelFrame(self.root, text=" 👥 建群规则配置 ")
        frame_top.pack(fill=tk.X, padx=10, pady=3, before=notebook)
        
        row1 = ttk.Frame(frame_top)
        row1.pack(fill=tk.X, pady=3)
        ttk.Label(row1, text="目标主理人（扫谁的客户）:").pack(side=tk.LEFT)
        self.combo_target = ttk.Combobox(row1, width=35, state="readonly")
        self.combo_target.pack(side=tk.LEFT, padx=15)
        
        row2 = ttk.Frame(frame_top)
        row2.pack(fill=tk.X, pady=3)
        ttk.Label(row2, text="固定跟随成员（最多3人）:").pack(side=tk.LEFT)
        self.combo_f1 = ttk.Combobox(row2, width=15, state="readonly"); self.combo_f1.pack(side=tk.LEFT, padx=5)
        self.combo_f2 = ttk.Combobox(row2, width=15, state="readonly"); self.combo_f2.pack(side=tk.LEFT, padx=5)
        self.combo_f3 = ttk.Combobox(row2, width=15, state="readonly"); self.combo_f3.pack(side=tk.LEFT, padx=5)
        ttk.Button(row2, text="💾 保存全部配置", command=self.save_settings).pack(side=tk.RIGHT, padx=10)

        # =============== TAB 1: 测试/手动确认模式 ===============
        tab_manual = ttk.Frame(notebook)
        notebook.add(tab_manual, text=" 🧪 步骤一：测试操作 / 手动指派建群模式 ")

        # =============== TAB 2: 正式/挂机自动模式 ===============
        tab_auto = ttk.Frame(notebook)
        notebook.add(tab_auto, text=" 🚀 步骤二：正式挂机 / 全自动巡检模式 ")
        self.btn_start = tk.Button(tab_auto, text="▶ 全自动挂机巡检建群", command=self.toggle_agent, bg="#0d6efd", fg="#fff", height=2, font=('Microsoft YaHei', 12))
        self.btn_start.pack(fill=tk.X, padx=20, pady=20)
        
        mf1 = ttk.Frame(tab_manual)
        mf1.pack(fill=tk.X, pady=10, padx=10)
        ttk.Button(mf1, text="🔄 读取该主理人的所有外部联系人", command=self.fetch_manual_contacts).pack(side=tk.LEFT, fill=tk.X, expand=True)

        cols = ('uid', 'name', 'status')
        self.tree = ttk.Treeview(tab_manual, columns=cols, show='headings', height=10)
        self.tree.heading('uid', text='微信外部联系人ID')
        self.tree.heading('name', text='客户昵称')
        self.tree.heading('status', text='本地是否已拉网')
        self.tree.column('uid', width=200)
        self.tree.column('name', width=150)
        self.tree.column('status', width=100)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        ttk.Button(tab_manual, text="⚡ 仅将选中的客户拉入新群", command=self.execute_manual).pack(fill=tk.X, padx=10, pady=10)

        # =============== 日志区 ===============
        log_frame = ttk.LabelFrame(self.root, text=" 运行日志 ")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, font=("Consolas", 10), bg="#212529", fg="#20c997", state=tk.DISABLED)
        self.log_area.pack(fill=tk.BOTH, expand=True)

    def fetch_manual_contacts(self):
        self.save_settings()
        tgt = self.cfg.config.get("target_userid")
        if not tgt:
            messagebox.showerror("错误", "请先在上方选择目标主理人！")
            return
        
        self.log(f"📡 正在拉取 {tgt} 的外部联系人列表...")
        self.log(f"   (后端逐个查询联系人详情, 联系人多时需30-60秒)")
        def _task():
            try:
                t0 = time.time()
                self.log(f"   ⏳ 请求已发送, 等待后端响应...")
                self.current_contacts = self.api.get_contacts(tgt, log_fn=self.log)
                elapsed = time.time() - t0
                def _update_ui():
                    for i in self.tree.get_children(): self.tree.delete(i)
                    processed = 0
                    for c in self.current_contacts:
                        uid = c.get('external_userid', '')
                        name = c.get('name', '')
                        is_done = self.cfg.is_processed(uid)
                        status = "已处理" if is_done else "未处理"
                        if is_done: processed += 1
                        self.tree.insert('', tk.END, values=(uid, name, status))
                    total = len(self.current_contacts)
                    self.log(f"✅ 拉取完毕: 共 {total} 人, 已处理 {processed}, 待处理 {total - processed} ({elapsed:.1f}s)")
                    if total == 0:
                        self.log(f"   💡 返回 0 人? 该员工可能不在企微「客户联系」权限范围内")
                self.root.after(0, _update_ui)
            except Exception as e:
                self.log(f"❌ 获取失败: {type(e).__name__}: {e}")
                import traceback; traceback.print_exc()
        threading.Thread(target=_task, daemon=True).start()

    def execute_manual(self):
        # 防止重复点击
        if hasattr(self, '_grp_running') and self._grp_running:
            messagebox.showwarning("提示", "建群任务正在执行中，请等待完成！")
            return

        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请在列表中选中至少 1 名要建群的客户！")
            return
            
        self.save_settings()
        fm = self.cfg.config.get("fixed_members", [])
        if not fm:
            messagebox.showwarning("提示", "请至少配置 1 名内部固定客服！")
            return

        # 收集所有要建群的客户
        tasks = []
        for item in selected:
            vals = self.tree.item(item, "values")
            tasks.append((vals[0], vals[1]))

        # 串行执行（一个线程处理所有选中客户，防止并发冲突）
        self._grp_running = True
        def _run_all():
            try:
                for uid, name in tasks:
                    self.log(f"⚡ 开始为手动选中的【{name}】建立独立群聊...")
                    self._run_single_group(uid, name, fm)
                    time.sleep(2.0)  # 每个群之间间隔
            finally:
                self._grp_running = False
                self.root.after(0, self._refresh_tree_status)
        threading.Thread(target=_run_all, daemon=True).start()

    def _refresh_tree_status(self):
        for item in self.tree.get_children():
            vals = self.tree.item(item, "values")
            uid = vals[0]
            if self.cfg.is_processed(uid):
                self.tree.item(item, values=(uid, vals[1], "已处理"))

    def _run_single_group(self, uid, name, fm):
        self.log(f"🔧 准备建群: 客户={name}, 固定成员={fm}")
        wc = WeComWindow()
        if not wc.find():
            self.log("❌ 找不到企微客户端窗口！请确认企业微信已打开且未最小化。")
            return
        self.log(f"🏗️ 开始建群流程: 客户=【{name}】, 窗口={wc.w}x{wc.h}")
        try:
            if wc.create_grp(name, fm, self.log):
                self.log(f"✅ 【{name}】(uid={uid[:20]}...) 建群成功！")
                self.cfg.mark_processed(uid)
            else:
                self.log(f"❌ 【{name}】建群失败, 可能弹窗未打开或成员未找到。")
        except Exception as e:
            self.log(f"❌ 【{name}】建群异常: {type(e).__name__}: {e}")
            import traceback; traceback.print_exc()
        self.root.after(0, self._refresh_tree_status)

    def toggle_agent(self):
        if not self.running:
            self.save_settings()
            if not self.combo_target.get():
                messagebox.showerror("", "请选择主理人")
                return
            self.running = True
            self.btn_start.config(text="⏹ 停止 Agent", bg="#dc3545")
            self.log("🚀 全自动巡检已启动。")
            threading.Thread(target=self.agent_loop, daemon=True).start()
        else:
            self.running = False
            self.btn_start.config(text="▶ 全自动挂机巡检建群", bg="#0d6efd")
            self.log("🛑 正在退出全自动巡检...")

    def agent_loop(self):
        api = APIClient(self.cfg.config)
        wc = WeComWindow()
        cycle = 0
        while self.running:
            cycle += 1
            try:
                tgt = self.cfg.config.get("target_userid")
                fm = self.cfg.config.get("fixed_members", [])
                interval = self.cfg.config.get("check_interval", 60)
                self.log(f"🔄 [#{cycle}] 全自动核查 (主理人={tgt}, 固定成员={fm})")

                self.log(f"   📡 获取联系人列表...")
                cs = api.get_contacts(tgt, log_fn=self.log)
                self.log(f"   📡 获取群聊列表...")
                gs = api.get_groups()
                g_names = [g.get("name","").lower() for g in gs]
                self.log(f"   📊 联系人={len(cs)}, 群聊={len(gs)}, 已处理={len(self.cfg.state.get('processed_customers',[]))}")
                
                new_c = []
                skipped = 0
                for c in cs:
                    uid, n = c.get("external_userid"), c.get("name", "")
                    if not uid or not n: continue
                    if self.cfg.is_processed(uid):
                        skipped += 1; continue
                    if any(n.lower() in gn for gn in g_names):
                        self.log(f"   ⚠️ 跳过已有群的客户【{n}】")
                        self.cfg.mark_processed(uid); continue
                    new_c.append((uid, n))
                
                self.log(f"   📊 过滤结果: 新客户={len(new_c)}, 已跳过={skipped}")
                
                for i, (uid, n) in enumerate(new_c, 1):
                    if not self.running: break
                    self.log(f"   🏗️ [{i}/{len(new_c)}] 为【{n}】建群...")
                    if not wc.find():
                        self.log("   ❌ 找不到企微窗口, 中止本轮。")
                        break
                    if wc.create_grp(n, fm, self.log):
                        self.log(f"   ✅ [{i}/{len(new_c)}]【{n}】自动建群成功！")
                        self.cfg.mark_processed(uid)
                    else:
                        self.log(f"   ❌ [{i}/{len(new_c)}]【{n}】自动建群失败。")
                    time.sleep(2)

                self.log(f"   💤 本轮结束, {interval}s 后开始下一轮...")
            except Exception as e:
                self.log(f"❌ 后台异常: {type(e).__name__}: {e}")
                import traceback; traceback.print_exc()
            for _ in range(self.cfg.config.get("check_interval", 60)):
                if not self.running: break
                time.sleep(1)

if __name__ == "__main__":
    t = tk.Tk()
    app = WeComAutoApp(t)
    t.mainloop()
