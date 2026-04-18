package main

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"image"
	"image/png"
	"math/rand"
	"strings"
	"sync"
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
	procGetWindowDC          = user32.NewProc("GetWindowDC")
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
	procGetCursorPos         = user32.NewProc("GetCursorPos")
	procSendInput            = user32.NewProc("SendInput")
	procInvalidateRect       = user32.NewProc("InvalidateRect")
	procUpdateWindow         = user32.NewProc("UpdateWindow")

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
	PW_CLIENTONLY             = 1
	PW_RENDERFULLCONTENT      = 2
	BI_RGB            = 0
	SRCCOPY           = 0x00CC0020
	MOUSEEVENTF_LEFTDOWN = 0x0002
	MOUSEEVENTF_LEFTUP   = 0x0004
	MOUSEEVENTF_MOVE     = 0x0001
	MOUSEEVENTF_ABSOLUTE = 0x8000
	MOUSEEVENTF_WHEEL    = 0x0800
	INPUT_MOUSE          = 0
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

// humanDelay 人类反应延时 (高斯分布近似, 增加随机性降低风控)
func humanDelay(baseMs, varianceMs int) {
	// 用多个随机数求和模拟高斯分布 (中心极限定理)
	sum := 0
	for i := 0; i < 3; i++ {
		sum += rand.Intn(varianceMs*2 + 1)
	}
	delay := baseMs + sum/3 - varianceMs
	// 偶尔加入一个“走神”延迟 (5% 概率, 模拟人类走神/切换注意力)
	if rand.Intn(100) < 5 {
		delay += 500 + rand.Intn(1500)
	}
	if delay < 80 {
		delay = 80
	}
	time.Sleep(time.Duration(delay) * time.Millisecond)
}

// MOUSE_INPUT SendInput 的鼠标输入结构 (硬件级伪装)
type MOUSE_INPUT struct {
	InputType uint32
	Dx        int32
	Dy        int32
	MouseData uint32
	DwFlags   uint32
	Time      uint32
	ExtraInfo uintptr
	_pad      [8]byte // 对齐填充 (INPUT union 大小)
}

// sendInputClick 用 SendInput API 发送鼠标点击 (替代 mouse_event)
// SendInput 走硬件输入队列, 比 mouse_event 更难被反自动化检测
func sendInputClick(screenX, screenY int) {
	// 将屏幕坐标转为绝对坐标 (0-65535 归一化)
	smW, _, _ := user32.NewProc("GetSystemMetrics").Call(0) // SM_CXSCREEN
	smH, _, _ := user32.NewProc("GetSystemMetrics").Call(1) // SM_CYSCREEN
	absX := int32(float64(screenX) * 65535.0 / float64(smW))
	absY := int32(float64(screenY) * 65535.0 / float64(smH))

	// 鼠标移动到目标位置
	moveInput := MOUSE_INPUT{
		InputType: INPUT_MOUSE,
		Dx:        absX,
		Dy:        absY,
		DwFlags:   MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE,
	}
	procSendInput.Call(1, uintptr(unsafe.Pointer(&moveInput)), unsafe.Sizeof(moveInput))
	time.Sleep(time.Duration(30+rand.Intn(50)) * time.Millisecond)

	// 鼠标按下
	downInput := MOUSE_INPUT{
		InputType: INPUT_MOUSE,
		Dx:        absX,
		Dy:        absY,
		DwFlags:   MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_ABSOLUTE,
	}
	procSendInput.Call(1, uintptr(unsafe.Pointer(&downInput)), unsafe.Sizeof(downInput))
	time.Sleep(time.Duration(50+rand.Intn(80)) * time.Millisecond)

	// 鼠标抬起
	upInput := MOUSE_INPUT{
		InputType: INPUT_MOUSE,
		Dx:        absX,
		Dy:        absY,
		DwFlags:   MOUSEEVENTF_LEFTUP | MOUSEEVENTF_ABSOLUTE,
	}
	procSendInput.Call(1, uintptr(unsafe.Pointer(&upInput)), unsafe.Sizeof(upInput))
}

// simulateMouseMove 模拟自然鼠标移动路径 (不是瞬移, 而是带曲线的多步移动)
// 从当前位置到 (targetX, targetY), 分 5-8 步完成, 带微小的随机偏移
func simulateMouseMove(targetX, targetY int) {
	var cur POINT
	procGetCursorPos.Call(uintptr(unsafe.Pointer(&cur)))
	startX, startY := float64(cur.X), float64(cur.Y)
	endX, endY := float64(targetX), float64(targetY)

	// 随机步数 5-8
	steps := 5 + rand.Intn(4)
	for i := 1; i <= steps; i++ {
		t := float64(i) / float64(steps)
		// ease-in-out 插值 (先快后慢, 更像人类)
		smooth := t * t * (3 - 2*t)
		curX := startX + (endX-startX)*smooth + float64(rand.Intn(5)-2)
		curY := startY + (endY-startY)*smooth + float64(rand.Intn(5)-2)

		procSetCursorPos.Call(uintptr(int(curX)), uintptr(int(curY)))
		time.Sleep(time.Duration(15+rand.Intn(25)) * time.Millisecond)
	}
	// 最终精确到达目标
	procSetCursorPos.Call(uintptr(targetX), uintptr(targetY))
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
				// 恢复但不激活, 避免抢 WAG Engine 焦点
				procShowWindow.Call(uintptr(hwnd), 4) // SW_SHOWNOACTIVATE = 4
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

// EnsureNotMinimized 确保窗口不是最小化 (不激活窗口, 不影响前台)
func (w *WeComWindow) EnsureNotMinimized() {
	iconic, _, _ := procIsIconic.Call(uintptr(w.Hwnd))
	if iconic != 0 {
		// 使用 SW_SHOWNOACTIVATE 恢复窗口但不激活, 避免抢夺焦点
		procShowWindow.Call(uintptr(w.Hwnd), 4) // SW_SHOWNOACTIVATE = 4
		time.Sleep(500 * time.Millisecond)
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

// ClearInput 强力清空输入框
// 策略: 三击全选 + Ctrl+A + Delete + End + 50×Backspace
// 移植自 Python: _clear_input()
func (w *WeComWindow) ClearInput(hwnd syscall.Handle) {
	h := uintptr(hwnd)

	// 方法1: 三击全选 (WM_LBUTTONDBLCLK 双击选词, 然后 Ctrl+A 全选)
	// 先 Ctrl+A 全选
	procSendMessage.Call(h, WM_KEYDOWN, VK_CONTROL, 0)
	time.Sleep(30 * time.Millisecond)
	procSendMessage.Call(h, WM_KEYDOWN, VK_A, 0)
	time.Sleep(30 * time.Millisecond)
	procSendMessage.Call(h, WM_KEYUP, VK_A, 0)
	time.Sleep(20 * time.Millisecond)
	procSendMessage.Call(h, WM_KEYUP, VK_CONTROL, 0)
	time.Sleep(80 * time.Millisecond)

	// Delete 删除选中内容
	procSendMessage.Call(h, WM_KEYDOWN, VK_DELETE, 0)
	time.Sleep(30 * time.Millisecond)
	procSendMessage.Call(h, WM_KEYUP, VK_DELETE, 0)
	time.Sleep(100 * time.Millisecond)

	// 方法2: 再来一次 Ctrl+A + Backspace (双保险)
	procSendMessage.Call(h, WM_KEYDOWN, VK_CONTROL, 0)
	time.Sleep(20 * time.Millisecond)
	procSendMessage.Call(h, WM_KEYDOWN, VK_A, 0)
	time.Sleep(20 * time.Millisecond)
	procSendMessage.Call(h, WM_KEYUP, VK_A, 0)
	time.Sleep(20 * time.Millisecond)
	procSendMessage.Call(h, WM_KEYUP, VK_CONTROL, 0)
	time.Sleep(50 * time.Millisecond)
	procSendMessage.Call(h, WM_KEYDOWN, VK_BACK, 0)
	procSendMessage.Call(h, WM_KEYUP, VK_BACK, 0)
	time.Sleep(100 * time.Millisecond)

	// 方法3: End 键 → 暴力 Backspace 50 次 (绝对清除)
	procSendMessage.Call(h, WM_KEYDOWN, VK_END, 0)
	procSendMessage.Call(h, WM_KEYUP, VK_END, 0)
	time.Sleep(50 * time.Millisecond)
	for i := 0; i < 50; i++ {
		procSendMessage.Call(h, WM_KEYDOWN, VK_BACK, 0)
		procSendMessage.Call(h, WM_KEYUP, VK_BACK, 0)
		time.Sleep(15 * time.Millisecond)
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
// 企微主窗口 (Chromium): PW_CLIENTONLY 返回黑屏, 必须用 flag=0
// 其他原生窗口 (群管理弹窗等): PW_CLIENTONLY 正常
// 策略: 先 flag=0, 全黑则回退 PW_CLIENTONLY
func (w *WeComWindow) screenshotHwnd(hwnd syscall.Handle) (image.Image, []byte, error) {
	var cr RECT
	procGetClientRect.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&cr)))
	width, height := int(cr.Right), int(cr.Bottom)
	if width <= 0 || height <= 0 {
		return nil, nil, fmt.Errorf("窗口尺寸无效: %dx%d", width, height)
	}

	// 尝试两种 flag: flag=0 对 Chromium 有效, PW_CLIENTONLY 对原生窗口有效
	for _, flag := range []uintptr{0, PW_CLIENTONLY} {
		hdc, _, _ := procGetDC.Call(uintptr(hwnd))
		memDC, _, _ := procCreateCompatibleDC.Call(hdc)
		hBitmap, _, _ := procCreateCompatibleBitmap.Call(hdc, uintptr(width), uintptr(height))
		procSelectObject.Call(memDC, hBitmap)
		procPrintWindow.Call(uintptr(hwnd), memDC, flag)

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
		procReleaseDC.Call(uintptr(hwnd), hdc)

		// 检查是否全黑
		isBlack := true
		step := width * height / 20
		if step < 1 {
			step = 1
		}
		for i := 0; i < len(pixelData)-4; i += step * 4 {
			if pixelData[i] > 5 || pixelData[i+1] > 5 || pixelData[i+2] > 5 {
				isBlack = false
				break
			}
		}
		if isBlack {
			continue
		}

		// 截图有效! BGRA → RGBA
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

	return nil, nil, fmt.Errorf("所有截图方式均返回黑屏")
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
// 操作前保存当前前台窗口, 操作后恢复, 不影响用户窗口状态
// 包含重试逻辑: 如果截图全黑 (BitBlt 时窗口未就绪) 最多重试 3 次
func (w *WeComWindow) ScreenshotForeground() (image.Image, []byte, error) {
	// 保存当前前台窗口 (操作后恢复)
	prevFg, _, _ := procGetForegroundWindow.Call()

	// 将 WeCom 窗口强制前台
	w.forceToForeground(w.Hwnd)

	// 强制窗口重绘 (确保 Chromium overlay 等内容已渲染)
	procInvalidateRect.Call(uintptr(w.Hwnd), 0, 1) // TRUE = erase background
	procUpdateWindow.Call(uintptr(w.Hwnd))
	time.Sleep(1200 * time.Millisecond) // 等待渲染完成 (原 800ms 不够稳定)

	var cr RECT
	procGetClientRect.Call(uintptr(w.Hwnd), uintptr(unsafe.Pointer(&cr)))
	width, height := int(cr.Right), int(cr.Bottom)
	if width <= 0 || height <= 0 {
		w.restoreForeground(prevFg)
		return nil, nil, fmt.Errorf("窗口尺寸无效")
	}

	// 获取客户区在屏幕上的位置
	var pt POINT
	procClientToScreen.Call(uintptr(w.Hwnd), uintptr(unsafe.Pointer(&pt)))

	var img *image.NRGBA
	var pixelData []byte

	// 重试最多 3 次 (BitBlt 在窗口刚前台化时可能返回全黑)
	for attempt := 0; attempt < 3; attempt++ {
		if attempt > 0 {
			// 重试前再次确保前台 + 等待渲染
			w.forceToForeground(w.Hwnd)
			procInvalidateRect.Call(uintptr(w.Hwnd), 0, 1)
			procUpdateWindow.Call(uintptr(w.Hwnd))
			time.Sleep(800 * time.Millisecond)
		}

		// 从屏幕 DC 截取
		screenDC, _, _ := procGetDC.Call(0)
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

		pixelData = make([]byte, width*height*4)
		procGetDIBits.Call(memDC, hBitmap, 0, uintptr(height),
			uintptr(unsafe.Pointer(&pixelData[0])),
			uintptr(unsafe.Pointer(&bmi)), 0)

		procDeleteObject.Call(hBitmap)
		procDeleteDC.Call(memDC)
		procReleaseDC.Call(0, screenDC)

		// BGRA → RGBA
		img = image.NewNRGBA(image.Rect(0, 0, width, height))
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

		// 检查是否全黑: 采样几个点
		isBlack := true
		samplePoints := [][2]int{{width / 4, height / 4}, {width / 2, height / 2}, {width * 3 / 4, height * 3 / 4}}
		for _, sp := range samplePoints {
			r, g, b, _ := img.At(sp[0], sp[1]).RGBA()
			if r > 0x500 || g > 0x500 || b > 0x500 {
				isBlack = false
				break
			}
		}
		if !isBlack {
			break // 截图有效, 跳出重试
		}
		// 全黑, 继续重试
	}

	// 取消 TOPMOST + 恢复之前的前台窗口
	procSetWindowPos.Call(uintptr(w.Hwnd), ^uintptr(1), 0, 0, 0, 0,
		SWP_NOMOVE|SWP_NOSIZE)
	w.restoreForeground(prevFg)

	var pngBuf bytes.Buffer
	png.Encode(&pngBuf, img)
	return img, pngBuf.Bytes(), nil
}

// forceToForeground 强制窗口到最前台 (绕过 Windows 前台锁)
func (w *WeComWindow) forceToForeground(hwnd syscall.Handle) {
	fgHwnd, _, _ := procGetForegroundWindow.Call()
	fgTid, _, _ := procGetWindowThreadPID.Call(fgHwnd, 0)
	myTid, _, _ := procGetCurrentThreadId.Call()

	if fgTid != myTid {
		procAttachThreadInput.Call(myTid, fgTid, 1) // TRUE
	}

	// 模拟 Alt 键释放, 解除 SetForegroundWindow 限制
	procKeyboardEvent.Call(0x12, 0, 0x0002, 0) // ALT up

	// TOPMOST 置顶 (临时, 操作后会取消)
	procSetWindowPos.Call(uintptr(hwnd), ^uintptr(0), 0, 0, 0, 0, // HWND_TOPMOST = -1
		SWP_NOMOVE|SWP_NOSIZE|SWP_SHOWWINDOW)

	procShowWindow.Call(uintptr(hwnd), SW_RESTORE)
	procSetForegroundWindow.Call(uintptr(hwnd))
	procBringWindowToTop.Call(uintptr(hwnd))

	if fgTid != myTid {
		procAttachThreadInput.Call(myTid, fgTid, 0) // FALSE
	}
}

// restoreForeground 恢复之前的前台窗口 (不影响用户窗口状态)
func (w *WeComWindow) restoreForeground(prevFg uintptr) {
	if prevFg != 0 && prevFg != uintptr(w.Hwnd) {
		procSetForegroundWindow.Call(prevFg)
	}
}

// foregroundMu 全局前台操作互斥锁
var foregroundMu sync.Mutex

// RealClick 前台真实鼠标点击 (用于 Chromium overlay)
// 操作前保存前台窗口, 操作后恢复, 不影响用户窗口状态
func (w *WeComWindow) RealClick(x, y int) {
	// 保存当前前台窗口
	prevFg, _, _ := procGetForegroundWindow.Call()

	var pt POINT
	pt.X = int32(x)
	pt.Y = int32(y)
	procClientToScreen.Call(uintptr(w.Hwnd), uintptr(unsafe.Pointer(&pt)))

	w.forceToForeground(w.Hwnd)
	time.Sleep(time.Duration(120+rand.Intn(80)) * time.Millisecond)

	// 自然鼠标移动 (多步带曲线, 不是瞬移)
	simulateMouseMove(int(pt.X), int(pt.Y))
	time.Sleep(time.Duration(80+rand.Intn(80)) * time.Millisecond)

	// 用 SendInput 发送点击 (硬件级)
	sendInputClick(int(pt.X), int(pt.Y))
	time.Sleep(time.Duration(300+rand.Intn(300)) * time.Millisecond)

	// 取消 TOPMOST + 恢复之前的前台窗口
	procSetWindowPos.Call(uintptr(w.Hwnd), ^uintptr(1), 0, 0, 0, 0, // HWND_NOTOPMOST = -2
		SWP_NOMOVE|SWP_NOSIZE)
	w.restoreForeground(prevFg)
}

// SafeRealClick 安全前台点击 (加锁 + 保存/恢复鼠标位置)
// 防止: 1) 多个前台操作并发冲突  2) 用户鼠标被劫持后丢失
func (w *WeComWindow) SafeRealClick(x, y int) {
	foregroundMu.Lock()
	defer foregroundMu.Unlock()

	// 保存当前鼠标位置
	var savedPos POINT
	procGetCursorPos.Call(uintptr(unsafe.Pointer(&savedPos)))

	// 执行点击
	w.RealClick(x, y)

	// 恢复鼠标位置 (减少对用户的干扰)
	procSetCursorPos.Call(uintptr(savedPos.X), uintptr(savedPos.Y))
}

// SafeRealScroll 安全前台鼠标滚轮 (加锁 + 保存/恢复鼠标位置)
// clientX, clientY: 窗口客户区坐标 (滚轮作用位置)
// delta: 正=向上滚, 负=向下滚 (±120 = 1格, ±360 = 3格)
func (w *WeComWindow) SafeRealScroll(clientX, clientY, delta int) {
	foregroundMu.Lock()
	defer foregroundMu.Unlock()

	// 保存当前鼠标位置
	var savedPos POINT
	procGetCursorPos.Call(uintptr(unsafe.Pointer(&savedPos)))

	// 保存当前前台窗口
	prevFg, _, _ := procGetForegroundWindow.Call()

	// 客户区 → 屏幕坐标
	var pt POINT
	pt.X = int32(clientX)
	pt.Y = int32(clientY)
	procClientToScreen.Call(uintptr(w.Hwnd), uintptr(unsafe.Pointer(&pt)))

	// 前台化 + 移动鼠标
	w.forceToForeground(w.Hwnd)
	time.Sleep(80 * time.Millisecond)
	simulateMouseMove(int(pt.X), int(pt.Y))
	time.Sleep(80 * time.Millisecond)

	// 发送真实滚轮事件 (mouse_event 硬件级)
	d32 := uint32(int32(delta))
	procMouseEvent.Call(MOUSEEVENTF_WHEEL, 0, 0, uintptr(d32), 0)
	time.Sleep(150 * time.Millisecond)

	// 取消 TOPMOST + 恢复
	procSetWindowPos.Call(uintptr(w.Hwnd), ^uintptr(1), 0, 0, 0, 0,
		SWP_NOMOVE|SWP_NOSIZE)
	w.restoreForeground(prevFg)
	procSetCursorPos.Call(uintptr(savedPos.X), uintptr(savedPos.Y))
}

// SafeScreenshotForeground 安全前台截图 (加锁保护)
func (w *WeComWindow) SafeScreenshotForeground() (image.Image, []byte, error) {
	foregroundMu.Lock()
	defer foregroundMu.Unlock()

	// 保存当前鼠标位置
	var savedPos POINT
	procGetCursorPos.Call(uintptr(unsafe.Pointer(&savedPos)))

	img, data, err := w.ScreenshotForeground()

	// 恢复鼠标位置
	procSetCursorPos.Call(uintptr(savedPos.X), uintptr(savedPos.Y))

	return img, data, err
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
