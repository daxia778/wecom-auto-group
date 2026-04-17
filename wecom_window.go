package main

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"image"
	"image/png"
	"math/rand"
	"strings"
	"syscall"
	"time"
	"unsafe"
)

var (
	user32   = syscall.NewLazyDLL("user32.dll")
	gdi32    = syscall.NewLazyDLL("gdi32.dll")
	kernel32 = syscall.NewLazyDLL("kernel32.dll")

	procEnumWindows          = user32.NewProc("EnumWindows")
	procGetClassName         = user32.NewProc("GetClassNameW")
	procGetWindowText        = user32.NewProc("GetWindowTextW")
	procIsWindowVisible      = user32.NewProc("IsWindowVisible")
	procIsIconic             = user32.NewProc("IsIconic")
	procIsWindow             = user32.NewProc("IsWindow")
	procShowWindow           = user32.NewProc("ShowWindow")
	procGetClientRect        = user32.NewProc("GetClientRect")
	procGetWindowRect        = user32.NewProc("GetWindowRect")
	procGetDC                = user32.NewProc("GetDC")
	procReleaseDC            = user32.NewProc("ReleaseDC")
	procSendMessage          = user32.NewProc("SendMessageW")
	procPostMessage          = user32.NewProc("PostMessageW")
	procPrintWindow          = user32.NewProc("PrintWindow")
	procSetWindowPos         = user32.NewProc("SetWindowPos")
	procGetWindowThreadPID   = user32.NewProc("GetWindowThreadProcessId")
	procSetProcessDPIAware   = user32.NewProc("SetProcessDPIAware")
	procClientToScreen       = user32.NewProc("ClientToScreen")
	procSetCursorPos         = user32.NewProc("SetCursorPos")
	procMouseEvent           = user32.NewProc("mouse_event")
	procGetForegroundWindow  = user32.NewProc("GetForegroundWindow")
	procSetForegroundWindow  = user32.NewProc("SetForegroundWindow")
	procAttachThreadInput    = user32.NewProc("AttachThreadInput")
	procKeyboardEvent        = user32.NewProc("keybd_event")
	procBringWindowToTop     = user32.NewProc("BringWindowToTop")

	procCreateCompatibleDC     = gdi32.NewProc("CreateCompatibleDC")
	procCreateCompatibleBitmap = gdi32.NewProc("CreateCompatibleBitmap")
	procSelectObject           = gdi32.NewProc("SelectObject")
	procDeleteObject           = gdi32.NewProc("DeleteObject")
	procDeleteDC               = gdi32.NewProc("DeleteDC")
	procGetDIBits              = gdi32.NewProc("GetDIBits")
	procBitBlt                 = gdi32.NewProc("BitBlt")

	procGetCurrentThreadId = kernel32.NewProc("GetCurrentThreadId")
)

const (
	WM_MOUSEACTIVATE = 0x0021
	WM_SETCURSOR     = 0x0020
	WM_MOUSEMOVE     = 0x0200
	WM_LBUTTONDOWN   = 0x0201
	WM_LBUTTONUP     = 0x0202
	WM_LBUTTONDBLCLK = 0x0203
	WM_MOUSEWHEEL    = 0x020A
	WM_KEYDOWN       = 0x0100
	WM_KEYUP         = 0x0101
	WM_IME_CHAR      = 0x0286
	MK_LBUTTON       = 0x0001
	VK_BACK          = 0x08
	VK_RETURN        = 0x0D
	VK_ESCAPE        = 0x1B
	VK_DELETE        = 0x2E
	VK_END           = 0x23
	VK_CONTROL       = 0x11
	VK_A             = 0x41
	SW_RESTORE       = 9
	HWND_BOTTOM      = 1
	HWND_TOPMOST     = -1
	HWND_NOTOPMOST   = -2
	SWP_NOMOVE       = 0x0002
	SWP_NOSIZE       = 0x0001
	SWP_NOACTIVATE   = 0x0010
	SWP_SHOWWINDOW   = 0x0040
	PW_CLIENTONLY     = 1
	BI_RGB            = 0
	SRCCOPY           = 0x00CC0020
	MOUSEEVENTF_LEFTDOWN = 0x0002
	MOUSEEVENTF_LEFTUP   = 0x0004
)

type RECT struct {
	Left, Top, Right, Bottom int32
}

type POINT struct {
	X, Y int32
}

type BITMAPINFOHEADER struct {
	BiSize          uint32
	BiWidth         int32
	BiHeight        int32
	BiPlanes        uint16
	BiBitCount      uint16
	BiCompression   uint32
	BiSizeImage     uint32
	BiXPelsPerMeter int32
	BiYPelsPerMeter int32
	BiClrUsed       uint32
	BiClrImportant  uint32
}

// WeComWindow 企微窗口操作 (移植自 Python wecom_auto.py)
type WeComWindow struct {
	Hwnd   syscall.Handle
	Pid    uint32
	Width  int
	Height int
}

func makeLParam(x, y int) uintptr {
	return uintptr((y << 16) | (x & 0xFFFF))
}

// jitter 坐标随机偏移 ±radius (模拟人类操作)
func jitter(val, radius int) int {
	return val + rand.Intn(radius*2+1) - radius
}

// humanDelay 人类反应延时 (高斯分布近似)
func humanDelay(baseMs, varianceMs int) {
	delay := baseMs + rand.Intn(varianceMs*2+1) - varianceMs
	if delay < 50 {
		delay = 50
	}
	time.Sleep(time.Duration(delay) * time.Millisecond)
}

// FindWeComWindow 查找企微主窗口
func FindWeComWindow() (*WeComWindow, error) {
	procSetProcessDPIAware.Call()

	var found *WeComWindow
	cb := syscall.NewCallback(func(hwnd syscall.Handle, lParam uintptr) uintptr {
		className := make([]uint16, 256)
		procGetClassName.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&className[0])), 256)
		cls := syscall.UTF16ToString(className)

		if cls == "WeWorkWindow" {
			visible, _, _ := procIsWindowVisible.Call(uintptr(hwnd))
			if visible == 0 {
				return 1 // continue
			}

			var pid uint32
			procGetWindowThreadPID.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&pid)))

			// 检查是否最小化
			iconic, _, _ := procIsIconic.Call(uintptr(hwnd))
			if iconic != 0 {
				procShowWindow.Call(uintptr(hwnd), SW_RESTORE)
				time.Sleep(800 * time.Millisecond)
			}

			var cr RECT
			procGetClientRect.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&cr)))

			found = &WeComWindow{
				Hwnd:   hwnd,
				Pid:    pid,
				Width:  int(cr.Right),
				Height: int(cr.Bottom),
			}
			return 0 // stop
		}
		return 1 // continue
	})

	procEnumWindows.Call(cb, 0)
	if found == nil {
		return nil, fmt.Errorf("找不到企微窗口，请确认企业微信已打开")
	}
	return found, nil
}

// SinkToBottom 将企微窗口压到最底层
func (w *WeComWindow) SinkToBottom() {
	procSetWindowPos.Call(uintptr(w.Hwnd), HWND_BOTTOM, 0, 0, 0, 0,
		SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE)
}

// EnsureNotMinimized 确保窗口不是最小化
func (w *WeComWindow) EnsureNotMinimized() {
	iconic, _, _ := procIsIconic.Call(uintptr(w.Hwnd))
	if iconic != 0 {
		procShowWindow.Call(uintptr(w.Hwnd), SW_RESTORE)
		time.Sleep(500 * time.Millisecond)
		w.SinkToBottom()
	}
}

// ━━━ 后台点击 / 输入 ━━━

// fullClickSeq 完整5步后台点击序列 (适配 Chromium hit-testing)
// 移植自 Python: _full_click_seq()
func (w *WeComWindow) fullClickSeq(hwnd syscall.Handle, x, y int) {
	lp := makeLParam(x, y)
	h := uintptr(hwnd)

	// Step 1: WM_MOUSEACTIVATE
	procSendMessage.Call(h, WM_MOUSEACTIVATE, h, uintptr((WM_LBUTTONDOWN<<16)|1))
	time.Sleep(time.Duration(20+rand.Intn(40)) * time.Millisecond)
	// Step 2: WM_SETCURSOR
	procSendMessage.Call(h, WM_SETCURSOR, h, uintptr((WM_LBUTTONDOWN<<16)|1))
	time.Sleep(time.Duration(20+rand.Intn(40)) * time.Millisecond)
	// Step 3: WM_MOUSEMOVE (hover 预热 — Chromium 关键!)
	procSendMessage.Call(h, WM_MOUSEMOVE, 0, lp)
	time.Sleep(time.Duration(120+rand.Intn(100)) * time.Millisecond)
	// Step 4: WM_LBUTTONDOWN
	procSendMessage.Call(h, WM_LBUTTONDOWN, MK_LBUTTON, lp)
	time.Sleep(time.Duration(50+rand.Intn(70)) * time.Millisecond)
	// Step 5: WM_LBUTTONUP
	procSendMessage.Call(h, WM_LBUTTONUP, 0, lp)
	humanDelay(250, 100)
}

// Click 后台点击主窗口 (带坐标抖动)
func (w *WeComWindow) Click(x, y int) {
	w.EnsureNotMinimized()
	x = jitter(x, 2)
	y = jitter(y, 2)
	w.fullClickSeq(w.Hwnd, x, y)
}

// TypeText 后台输入文字 (WM_IME_CHAR, 支持中文)
func (w *WeComWindow) TypeText(text string) {
	for _, ch := range text {
		procPostMessage.Call(uintptr(w.Hwnd), WM_IME_CHAR, uintptr(ch), 0)
		time.Sleep(time.Duration(40+rand.Intn(60)) * time.Millisecond)
	}
	time.Sleep(200 * time.Millisecond)
}

// SendKey 后台发送按键
func (w *WeComWindow) SendKey(vk int) {
	procPostMessage.Call(uintptr(w.Hwnd), WM_KEYDOWN, uintptr(vk), 0)
	time.Sleep(50 * time.Millisecond)
	procPostMessage.Call(uintptr(w.Hwnd), WM_KEYUP, uintptr(vk), 0)
	humanDelay(250, 80)
}

// ClearInput 强力清空输入框 (Ctrl+A → Delete → End → 30×Backspace)
// 移植自 Python: _clear_input()
func (w *WeComWindow) ClearInput(hwnd syscall.Handle) {
	h := uintptr(hwnd)
	// 方法1: Ctrl+A 全选 → Delete
	procSendMessage.Call(h, WM_KEYDOWN, VK_CONTROL, 0)
	time.Sleep(30 * time.Millisecond)
	procSendMessage.Call(h, WM_KEYDOWN, VK_A, 0)
	time.Sleep(30 * time.Millisecond)
	procSendMessage.Call(h, WM_KEYUP, VK_A, 0)
	time.Sleep(20 * time.Millisecond)
	procSendMessage.Call(h, WM_KEYUP, VK_CONTROL, 0)
	time.Sleep(80 * time.Millisecond)
	procSendMessage.Call(h, WM_KEYDOWN, VK_DELETE, 0)
	time.Sleep(30 * time.Millisecond)
	procSendMessage.Call(h, WM_KEYUP, VK_DELETE, 0)
	time.Sleep(100 * time.Millisecond)

	// 方法2: End 键 → 暴力 Backspace 30 次
	procSendMessage.Call(h, WM_KEYDOWN, VK_END, 0)
	procSendMessage.Call(h, WM_KEYUP, VK_END, 0)
	time.Sleep(50 * time.Millisecond)
	for i := 0; i < 30; i++ {
		procSendMessage.Call(h, WM_KEYDOWN, VK_BACK, 0)
		procSendMessage.Call(h, WM_KEYUP, VK_BACK, 0)
		time.Sleep(20 * time.Millisecond)
	}
}

// SendWheel 后台发送鼠标滚轮 (WM_MOUSEWHEEL)
// delta: 正=向上, 负=向下, 通常 ±120
func (w *WeComWindow) SendWheel(x, y, delta int) {
	lp := makeLParam(x, y)
	// wParam 高16位=滚动量, 低16位=按键状态
	wp := uintptr((delta & 0xFFFF) << 16)
	procSendMessage.Call(uintptr(w.Hwnd), WM_MOUSEWHEEL, wp, uintptr(lp))
}

// ━━━ 弹窗操作 ━━━

// findPopup 查找弹窗句柄
func (w *WeComWindow) findPopup(className string) syscall.Handle {
	var found syscall.Handle
	cb := syscall.NewCallback(func(hwnd syscall.Handle, _ uintptr) uintptr {
		visible, _, _ := procIsWindowVisible.Call(uintptr(hwnd))
		if visible == 0 {
			return 1
		}
		cls := make([]uint16, 256)
		procGetClassName.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&cls[0])), 256)
		if syscall.UTF16ToString(cls) == className {
			found = hwnd
			return 0
		}
		return 1
	})
	procEnumWindows.Call(cb, 0)
	return found
}

// isPopupVisible 检查弹窗是否仍然可见
func (w *WeComWindow) isPopupVisible(className string) bool {
	phwnd := w.findPopup(className)
	if phwnd == 0 {
		return false
	}
	isWin, _, _ := procIsWindow.Call(uintptr(phwnd))
	vis, _, _ := procIsWindowVisible.Call(uintptr(phwnd))
	return isWin != 0 && vis != 0
}

// ClickPopup 点击弹窗
func (w *WeComWindow) ClickPopup(popupClass string, x, y int) {
	phwnd := w.findPopup(popupClass)
	if phwnd != 0 {
		x = jitter(x, 2)
		y = jitter(y, 2)
		w.fullClickSeq(phwnd, x, y)
	}
}

// TypeToPopup 向弹窗输入文字
func (w *WeComWindow) TypeToPopup(popupClass, text string) {
	phwnd := w.findPopup(popupClass)
	if phwnd == 0 {
		return
	}
	for _, ch := range text {
		procPostMessage.Call(uintptr(phwnd), WM_IME_CHAR, uintptr(ch), 0)
		time.Sleep(time.Duration(40+rand.Intn(50)) * time.Millisecond)
	}
	time.Sleep(200 * time.Millisecond)
}

// SendKeyToPopup 向弹窗发送按键
func (w *WeComWindow) SendKeyToPopup(popupClass string, vk int) {
	phwnd := w.findPopup(popupClass)
	if phwnd == 0 {
		return
	}
	procPostMessage.Call(uintptr(phwnd), WM_KEYDOWN, uintptr(vk), 0)
	time.Sleep(50 * time.Millisecond)
	procPostMessage.Call(uintptr(phwnd), WM_KEYUP, uintptr(vk), 0)
	humanDelay(250, 80)
}

// ClearPopupInput 清空弹窗搜索框
func (w *WeComWindow) ClearPopupInput(popupClass string) {
	phwnd := w.findPopup(popupClass)
	if phwnd == 0 {
		return
	}
	w.ClearInput(phwnd)
}

// ClosePopup 用 ESC 关闭弹窗 (注意: ESC 对主窗口会最小化!)
func (w *WeComWindow) ClosePopup(popupClass string) {
	phwnd := w.findPopup(popupClass)
	if phwnd == 0 {
		return
	}
	procSendMessage.Call(uintptr(phwnd), WM_KEYDOWN, VK_ESCAPE, 0)
	procSendMessage.Call(uintptr(phwnd), WM_KEYUP, VK_ESCAPE, 0)
	time.Sleep(500 * time.Millisecond)
}

// ━━━ 截图 (PrintWindow 后台) ━━━

// screenshotHwnd 对任意窗口 HWND 做 PrintWindow 截图
func (w *WeComWindow) screenshotHwnd(hwnd syscall.Handle) (image.Image, []byte, error) {
	var cr RECT
	procGetClientRect.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&cr)))
	width, height := int(cr.Right), int(cr.Bottom)
	if width <= 0 || height <= 0 {
		return nil, nil, fmt.Errorf("窗口尺寸无效: %dx%d", width, height)
	}

	hdc, _, _ := procGetDC.Call(uintptr(hwnd))
	memDC, _, _ := procCreateCompatibleDC.Call(hdc)
	hBitmap, _, _ := procCreateCompatibleBitmap.Call(hdc, uintptr(width), uintptr(height))
	procSelectObject.Call(memDC, hBitmap)
	procPrintWindow.Call(uintptr(hwnd), memDC, PW_CLIENTONLY)

	bmi := BITMAPINFOHEADER{
		BiSize:        uint32(unsafe.Sizeof(BITMAPINFOHEADER{})),
		BiWidth:       int32(width),
		BiHeight:      -int32(height), // 负号 = top-down
		BiPlanes:      1,
		BiBitCount:    32,
		BiCompression: BI_RGB,
	}

	pixelData := make([]byte, width*height*4)
	procGetDIBits.Call(memDC, hBitmap, 0, uintptr(height),
		uintptr(unsafe.Pointer(&pixelData[0])),
		uintptr(unsafe.Pointer(&bmi)), 0)

	procDeleteObject.Call(hBitmap)
	procDeleteDC.Call(memDC)
	procReleaseDC.Call(uintptr(hwnd), hdc)

	// BGRA → RGBA
	img := image.NewNRGBA(image.Rect(0, 0, width, height))
	for y := 0; y < height; y++ {
		for x := 0; x < width; x++ {
			i := (y*width + x) * 4
			if i+3 < len(pixelData) {
				img.Pix[(y*width+x)*4+0] = pixelData[i+2] // R
				img.Pix[(y*width+x)*4+1] = pixelData[i+1] // G
				img.Pix[(y*width+x)*4+2] = pixelData[i+0] // B
				img.Pix[(y*width+x)*4+3] = 255             // A
			}
		}
	}

	var pngBuf bytes.Buffer
	png.Encode(&pngBuf, img)
	return img, pngBuf.Bytes(), nil
}

// Screenshot 后台截图主窗口 (PrintWindow)
func (w *WeComWindow) Screenshot() (image.Image, []byte, error) {
	w.EnsureNotMinimized()
	time.Sleep(500 * time.Millisecond)
	return w.screenshotHwnd(w.Hwnd)
}

// ScreenshotPopup 后台截图弹窗
func (w *WeComWindow) ScreenshotPopup(popupClass string) (image.Image, []byte, error) {
	phwnd := w.findPopup(popupClass)
	if phwnd == 0 {
		return nil, nil, fmt.Errorf("弹窗 [%s] 未找到", popupClass)
	}
	return w.screenshotHwnd(phwnd)
}

// ScreenshotForeground 前台截图 (BitBlt 从屏幕DC, 能捕获 Chromium overlay)
// 移植自 Python: _screenshot_foreground()
func (w *WeComWindow) ScreenshotForeground() (image.Image, []byte, error) {
	// 将窗口强制前台
	w.forceToForeground(w.Hwnd)
	time.Sleep(800 * time.Millisecond)

	var cr RECT
	procGetClientRect.Call(uintptr(w.Hwnd), uintptr(unsafe.Pointer(&cr)))
	width, height := int(cr.Right), int(cr.Bottom)
	if width <= 0 || height <= 0 {
		return nil, nil, fmt.Errorf("窗口尺寸无效")
	}

	// 获取客户区在屏幕上的位置
	var pt POINT
	procClientToScreen.Call(uintptr(w.Hwnd), uintptr(unsafe.Pointer(&pt)))

	// 从屏幕 DC 截取
	screenDC, _, _ := procGetDC.Call(0) // 屏幕 DC
	memDC, _, _ := procCreateCompatibleDC.Call(screenDC)
	hBitmap, _, _ := procCreateCompatibleBitmap.Call(screenDC, uintptr(width), uintptr(height))
	procSelectObject.Call(memDC, hBitmap)
	procBitBlt.Call(memDC, 0, 0, uintptr(width), uintptr(height),
		screenDC, uintptr(pt.X), uintptr(pt.Y), SRCCOPY)

	bmi := BITMAPINFOHEADER{
		BiSize:        uint32(unsafe.Sizeof(BITMAPINFOHEADER{})),
		BiWidth:       int32(width),
		BiHeight:      -int32(height),
		BiPlanes:      1,
		BiBitCount:    32,
		BiCompression: BI_RGB,
	}

	pixelData := make([]byte, width*height*4)
	procGetDIBits.Call(memDC, hBitmap, 0, uintptr(height),
		uintptr(unsafe.Pointer(&pixelData[0])),
		uintptr(unsafe.Pointer(&bmi)), 0)

	procDeleteObject.Call(hBitmap)
	procDeleteDC.Call(memDC)
	procReleaseDC.Call(0, screenDC)

	// 取消 TOPMOST
	procSetWindowPos.Call(uintptr(w.Hwnd), ^uintptr(1), 0, 0, 0, 0, // HWND_NOTOPMOST = -2
		SWP_NOMOVE|SWP_NOSIZE)

	// BGRA → RGBA
	img := image.NewNRGBA(image.Rect(0, 0, width, height))
	for y := 0; y < height; y++ {
		for x := 0; x < width; x++ {
			i := (y*width + x) * 4
			if i+3 < len(pixelData) {
				img.Pix[(y*width+x)*4+0] = pixelData[i+2]
				img.Pix[(y*width+x)*4+1] = pixelData[i+1]
				img.Pix[(y*width+x)*4+2] = pixelData[i+0]
				img.Pix[(y*width+x)*4+3] = 255
			}
		}
	}

	var pngBuf bytes.Buffer
	png.Encode(&pngBuf, img)
	return img, pngBuf.Bytes(), nil
}

// forceToForeground 强制窗口到最前台 (绕过 Windows 前台锁)
// 移植自 Python: _force_foreground()
func (w *WeComWindow) forceToForeground(hwnd syscall.Handle) {
	fgHwnd, _, _ := procGetForegroundWindow.Call()
	fgTid, _, _ := procGetWindowThreadPID.Call(fgHwnd, 0)
	myTid, _, _ := procGetCurrentThreadId.Call()

	if fgTid != myTid {
		procAttachThreadInput.Call(myTid, fgTid, 1) // TRUE
	}

	// 模拟 Alt 键释放, 解除 SetForegroundWindow 限制
	procKeyboardEvent.Call(0x12, 0, 0x0002, 0) // ALT up

	// TOPMOST 置顶
	procSetWindowPos.Call(uintptr(hwnd), ^uintptr(0), 0, 0, 0, 0, // HWND_TOPMOST = -1
		SWP_NOMOVE|SWP_NOSIZE|SWP_SHOWWINDOW)

	procShowWindow.Call(uintptr(hwnd), SW_RESTORE)
	procSetForegroundWindow.Call(uintptr(hwnd))
	procBringWindowToTop.Call(uintptr(hwnd))

	if fgTid != myTid {
		procAttachThreadInput.Call(myTid, fgTid, 0) // FALSE
	}
}

// RealClick 前台真实鼠标点击 (用于 Chromium overlay)
// SendMessage 无法穿透 overlay, 必须用 SetCursorPos + mouse_event
// 移植自 Python: real_click()
func (w *WeComWindow) RealClick(x, y int) {
	var pt POINT
	pt.X = int32(x)
	pt.Y = int32(y)
	procClientToScreen.Call(uintptr(w.Hwnd), uintptr(unsafe.Pointer(&pt)))

	w.forceToForeground(w.Hwnd)
	time.Sleep(150 * time.Millisecond)

	procSetCursorPos.Call(uintptr(pt.X), uintptr(pt.Y))
	time.Sleep(150 * time.Millisecond)
	procMouseEvent.Call(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
	time.Sleep(time.Duration(60+rand.Intn(60)) * time.Millisecond)
	procMouseEvent.Call(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
	time.Sleep(500 * time.Millisecond)

	// 取消 TOPMOST
	procSetWindowPos.Call(uintptr(w.Hwnd), ^uintptr(1), 0, 0, 0, 0, // HWND_NOTOPMOST = -2
		SWP_NOMOVE|SWP_NOSIZE)
}

// ━━━ OCR 集成 ━━━

// OCRScan 截图主窗口 + OCR
func (w *WeComWindow) OCRScan() ([]OCRItem, error) {
	_, pngData, err := w.Screenshot()
	if err != nil {
		return nil, err
	}
	return ZhipuOCR(pngData)
}

// OCRScanPopup 截图弹窗 + OCR
func (w *WeComWindow) OCRScanPopup(popupClass string) ([]OCRItem, error) {
	_, pngData, err := w.ScreenshotPopup(popupClass)
	if err != nil {
		return nil, err
	}
	return ZhipuOCR(pngData)
}

// OCRScanForeground 前台截图 + OCR (能捕获 Chromium overlay)
func (w *WeComWindow) OCRScanForeground() ([]OCRItem, error) {
	_, pngData, err := w.ScreenshotForeground()
	if err != nil {
		return nil, err
	}
	return ZhipuOCR(pngData)
}

// OCRClickText OCR 定位文字并点击主窗口
func (w *WeComWindow) OCRClickText(keyword string) (bool, []OCRItem, error) {
	items, err := w.OCRScan()
	if err != nil {
		return false, nil, err
	}
	match := FindOCRText(items, keyword)
	if match != nil {
		w.Click(match.CX, match.CY)
		return true, items, nil
	}
	return false, items, nil
}

// OCRClickTextInPopup 在弹窗内 OCR 定位文字并点击
func (w *WeComWindow) OCRClickTextInPopup(popupClass, keyword string) (bool, []OCRItem, error) {
	items, err := w.OCRScanPopup(popupClass)
	if err != nil {
		return false, nil, err
	}
	match := FindOCRText(items, keyword)
	if match != nil {
		w.ClickPopup(popupClass, match.CX, match.CY)
		return true, items, nil
	}
	return false, items, nil
}

// WaitForPopup 等待弹窗出现 (轮询)
func (w *WeComWindow) WaitForPopup(popupClass string, timeoutSec int) bool {
	deadline := time.Now().Add(time.Duration(timeoutSec) * time.Second)
	for time.Now().Before(deadline) {
		if w.findPopup(popupClass) != 0 {
			return true
		}
		time.Sleep(300 * time.Millisecond)
	}
	return false
}

// WaitForPopupClosed 等待弹窗关闭 (轮询)
func (w *WeComWindow) WaitForPopupClosed(popupClass string, timeoutSec int) bool {
	deadline := time.Now().Add(time.Duration(timeoutSec) * time.Second)
	for time.Now().Before(deadline) {
		if !w.isPopupVisible(popupClass) {
			return true
		}
		time.Sleep(300 * time.Millisecond)
	}
	return false
}

// popupClientSize 获取弹窗客户区大小
func (w *WeComWindow) popupClientSize(popupClass string) (int, int) {
	phwnd := w.findPopup(popupClass)
	if phwnd == 0 {
		return 0, 0
	}
	var cr RECT
	procGetClientRect.Call(uintptr(phwnd), uintptr(unsafe.Pointer(&cr)))
	return int(cr.Right), int(cr.Bottom)
}

// OCRScanRegion 区域 OCR: 先截全图裁剪到指定区域, 再识别
// 坐标自动转换回窗口坐标系 (移植自 Python: ocr_scan_region)
func (w *WeComWindow) OCRScanRegion(x1, y1, x2, y2 int) ([]OCRItem, error) {
	img, pngData, err := w.Screenshot()
	if err != nil {
		return nil, err
	}
	_ = pngData

	bounds := img.Bounds()
	iw, ih := bounds.Dx(), bounds.Dy()

	// 裁剪边界检查
	if x1 < 0 {
		x1 = 0
	}
	if y1 < 0 {
		y1 = 0
	}
	if x2 > iw {
		x2 = iw
	}
	if y2 > ih {
		y2 = ih
	}
	if x2 <= x1 || y2 <= y1 {
		return nil, fmt.Errorf("区域无效: (%d,%d)-(%d,%d)", x1, y1, x2, y2)
	}

	// 裁剪图片
	type subImager interface {
		SubImage(r image.Rectangle) image.Image
	}
	si, ok := img.(subImager)
	if !ok {
		return nil, fmt.Errorf("图片不支持裁剪")
	}
	cropped := si.SubImage(image.Rect(x1, y1, x2, y2))

	// 编码裁剪后的图片
	var buf bytes.Buffer
	if err := png.Encode(&buf, cropped); err != nil {
		return nil, fmt.Errorf("裁剪图片编码失败: %w", err)
	}

	// OCR 识别
	items, err := ZhipuOCR(buf.Bytes())
	if err != nil {
		return nil, err
	}

	// 坐标转换回窗口坐标系
	for i := range items {
		items[i].CX += x1
		items[i].CY += y1
		items[i].X1 += x1
		items[i].Y1 += y1
		items[i].X2 += x1
		items[i].Y2 += y1
	}
	return items, nil
}

// OCRScanForegroundRegion 前台区域 OCR (能捕获 Chromium overlay)
func (w *WeComWindow) OCRScanForegroundRegion(x1, y1, x2, y2 int) ([]OCRItem, error) {
	img, _, err := w.ScreenshotForeground()
	if err != nil {
		return nil, err
	}

	bounds := img.Bounds()
	iw, ih := bounds.Dx(), bounds.Dy()
	if x1 < 0 {
		x1 = 0
	}
	if y1 < 0 {
		y1 = 0
	}
	if x2 > iw {
		x2 = iw
	}
	if y2 > ih {
		y2 = ih
	}
	if x2 <= x1 || y2 <= y1 {
		return nil, fmt.Errorf("区域无效")
	}

	type subImager interface {
		SubImage(r image.Rectangle) image.Image
	}
	si, ok := img.(subImager)
	if !ok {
		return nil, fmt.Errorf("图片不支持裁剪")
	}
	cropped := si.SubImage(image.Rect(x1, y1, x2, y2))

	var buf bytes.Buffer
	if err := png.Encode(&buf, cropped); err != nil {
		return nil, err
	}

	items, err := ZhipuOCR(buf.Bytes())
	if err != nil {
		return nil, err
	}

	for i := range items {
		items[i].CX += x1
		items[i].CY += y1
		items[i].X1 += x1
		items[i].Y1 += y1
		items[i].X2 += x1
		items[i].Y2 += y1
	}
	return items, nil
}

// IsCheckboxChecked 检测复选框是否已勾选 — 蓝色像素采样
// 移植自 Python: _is_checked()
// WeCom 勾选蓝色 RGB ≈ (76,149,243), B > 200 && B > R + 50
func IsCheckboxChecked(img image.Image, checkX, checkY int) bool {
	if img == nil {
		return false
	}
	bounds := img.Bounds()
	blueCount := 0

	// 在 checkbox 区域采样 3x3 格点
	for _, dx := range []int{-5, 0, 5} {
		for _, dy := range []int{-3, 0, 3} {
			px := checkX + dx
			py := checkY + dy
			if px >= bounds.Min.X && px < bounds.Max.X && py >= bounds.Min.Y && py < bounds.Max.Y {
				r, _, b, _ := img.At(px, py).RGBA()
				// RGBA returns 16-bit, shift to 8-bit
				r8 := r >> 8
				b8 := b >> 8
				if b8 > 200 && b8 > r8+50 {
					blueCount++
				}
			}
		}
	}
	return blueCount >= 3 // 至少 3 个蓝色像素
}

// 兼容性: 确保 binary/strings 包被引用
var _ = binary.LittleEndian
var _ = strings.Contains
