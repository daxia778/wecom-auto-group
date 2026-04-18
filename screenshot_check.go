package main

import (
	"fmt"
	"os"
	"path/filepath"
	"syscall"
	"time"
	"unsafe"
)

// ================================================================
// 截图能力深度测试: 找出为什么 PrintWindow 对主窗口返回黑屏
// 测试所有可能的截图方式 + DPI 设置
// 用法: WeComAutoGroup.exe --screenshot-test
// ================================================================

func TestScreenshotCapabilities() {
	fmt.Println("╔═════════════════════════════════════════════╗")
	fmt.Println("║    截图能力深度测试                          ║")
	fmt.Println("║    测试所有截图方式 + DPI                    ║")
	fmt.Println("╚═════════════════════════════════════════════╝")
	fmt.Println()

	debugDir, _ := filepath.Abs("debug_screenshot")
	os.MkdirAll(debugDir, 0755)
	log := func(s string) { fmt.Printf("[%s] %s\n", time.Now().Format("15:04:05"), s) }

	wc, err := FindWeComWindow()
	if err != nil {
		fmt.Printf("❌ %v\n", err)
		return
	}
	fmt.Printf("✅ 窗口: %d×%d (HWND=0x%X PID=%d)\n\n", wc.Width, wc.Height, wc.Hwnd, wc.Pid)

	// 检查 DPI 信息
	log("═══ DPI 信息 ═══")
	shcore := syscall.NewLazyDLL("shcore.dll")
	procGetDpiForMonitor := shcore.NewProc("GetDpiForMonitor")
	procMonitorFromWindow := user32.NewProc("MonitorFromWindow")
	monitor, _, _ := procMonitorFromWindow.Call(uintptr(wc.Hwnd), 2) // MONITOR_DEFAULTTONEAREST
	var dpiX, dpiY uint32
	r, _, _ := procGetDpiForMonitor.Call(monitor, 0, uintptr(unsafe.Pointer(&dpiX)), uintptr(unsafe.Pointer(&dpiY)))
	if r == 0 {
		log(fmt.Sprintf("  显示器 DPI: %d×%d (缩放: %d%%)", dpiX, dpiY, dpiX*100/96))
	}

	// 检查窗口 DPI Awareness
	procGetDpiAwarenessContext := user32.NewProc("GetWindowDpiAwarenessContext")
	procGetAwareFromCtx := user32.NewProc("GetAwarenessFromDpiAwarenessContext")
	ctx, _, e1 := procGetDpiAwarenessContext.Call(uintptr(wc.Hwnd))
	if e1 == nil || ctx != 0 {
		awareness, _, _ := procGetAwareFromCtx.Call(ctx)
		names := map[uintptr]string{0: "UNAWARE", 1: "SYSTEM_AWARE", 2: "PER_MONITOR_AWARE"}
		log(fmt.Sprintf("  企微 DPI Awareness: %s (%d)", names[awareness], awareness))
	}

	// ═══ 测试各种 PrintWindow flag ═══
	log("")
	log("═══ PrintWindow 各种 flag 测试 ═══")

	flags := []struct {
		flag int
		name string
	}{
		{0, "flag=0 (无flag)"},
		{1, "flag=1 (PW_CLIENTONLY)"},
		{2, "flag=2 (PW_RENDERFULLCONTENT)"},
		{3, "flag=3 (CLIENTONLY|RENDERFULLCONTENT)"},
	}

	for _, f := range flags {
		_, data := screenshotHwndWithFlag(wc.Hwnd, uintptr(f.flag))
		fname := fmt.Sprintf("pw_flag%d.png", f.flag)
		isBlack := len(data) < 10000
		status := "⚠️ 黑"
		if !isBlack {
			status = "✅ OK"
		}
		os.WriteFile(filepath.Join(debugDir, fname), data, 0644)
		log(fmt.Sprintf("  %-45s %s (%d bytes)", f.name, status, len(data)))
	}

	// ═══ 测试: 先 forceToForeground 再 PrintWindow ═══
	log("")
	log("═══ forceToForeground + PrintWindow ═══")
	log("  把企微窗口前台化后再 PrintWindow...")

	prevFg, _, _ := procGetForegroundWindow.Call()
	wc.forceToForeground(wc.Hwnd)
	time.Sleep(1500 * time.Millisecond)

	for _, f := range flags {
		_, data := screenshotHwndWithFlag(wc.Hwnd, uintptr(f.flag))
		fname := fmt.Sprintf("pw_fg_flag%d.png", f.flag)
		isBlack := len(data) < 10000
		status := "⚠️ 黑"
		if !isBlack {
			status = "✅ OK"
		}
		os.WriteFile(filepath.Join(debugDir, fname), data, 0644)
		log(fmt.Sprintf("  %-45s %s (%d bytes)", f.name, status, len(data)))
	}

	// 恢复
	procSetWindowPos.Call(uintptr(wc.Hwnd), ^uintptr(1), 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE)
	if prevFg != 0 {
		procSetForegroundWindow.Call(prevFg)
	}

	// ═══ 测试: BitBlt 从屏幕 DC (ScreenshotForeground) ═══
	log("")
	log("═══ BitBlt 前台截图 (ScreenshotForeground) ═══")
	_, fgData, fgErr := wc.ScreenshotForeground()
	if fgErr != nil {
		log(fmt.Sprintf("  ❌ 失败: %v", fgErr))
	} else {
		os.WriteFile(filepath.Join(debugDir, "bitblt_foreground.png"), fgData, 0644)
		isBlack := len(fgData) < 10000
		status := "⚠️ 黑"
		if !isBlack {
			status = "✅ OK"
		}
		log(fmt.Sprintf("  BitBlt 前台截图: %s (%d bytes)", status, len(fgData)))
	}

	// ═══ 测试: 对 PerryShadowWnd 截图 ═══
	log("")
	log("═══ 同进程子窗口截图 ═══")
	cbChild := syscall.NewCallback(func(hwnd syscall.Handle, _ uintptr) uintptr {
		var pid uint32
		procGetWindowThreadPID.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&pid)))
		if pid != wc.Pid {
			return 1
		}
		visible, _, _ := procIsWindowVisible.Call(uintptr(hwnd))
		if visible == 0 {
			return 1
		}
		var cr RECT
		procGetClientRect.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&cr)))
		if int(cr.Right) < 100 {
			return 1
		}
		cls := getClassName(hwnd)

		// 每个窗口测 flag=0 和 flag=3
		for _, flag := range []uintptr{0, 3} {
			_, data := screenshotHwndWithFlag(hwnd, flag)
			isBlack := len(data) < 10000
			status := "⚠️ 黑"
			if !isBlack {
				fname := fmt.Sprintf("child_%s_flag%d_%dx%d.png",
					safeFileName(cls), flag, cr.Right, cr.Bottom)
				os.WriteFile(filepath.Join(debugDir, fname), data, 0644)
				status = "✅ OK"
			}
			log(fmt.Sprintf("  %-25s flag=%d %4d×%-4d %s (%d bytes)",
				cls, flag, cr.Right, cr.Bottom, status, len(data)))
		}
		return 1
	})
	procEnumWindows.Call(cbChild, 0)

	// ═══ 汇总 ═══
	log("")
	log("══════════════════════════════════════")
	entries, _ := os.ReadDir(debugDir)
	okCount := 0
	for _, e := range entries {
		if !e.IsDir() {
			info, _ := e.Info()
			if info != nil {
				if info.Size() > 10000 {
					okCount++
				}
				log(fmt.Sprintf("   📸 %-38s  (%.0f KB)", e.Name(), float64(info.Size())/1024))
			}
		}
	}
	log(fmt.Sprintf("\n  有效截图: %d/%d", okCount, len(entries)))
}
