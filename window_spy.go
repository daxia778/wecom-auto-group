package main

import (
	"bytes"
	"fmt"
	"image"
	"image/png"
	"os"
	"strings"
	"syscall"
	"unsafe"
)

// ================================================================
// 窗口探测工具: 深度枚举所有 WeCom 相关窗口
// 用法: WeComAutoGroup.exe --spy-windows
// 功能: 枚举主窗口的所有子窗口/同进程窗口, 找到群管理 overlay
// ================================================================

var (
	procEnumChildWindows = user32.NewProc("EnumChildWindows")
	procGetParent        = user32.NewProc("GetParent")
	procGetAncestor      = user32.NewProc("GetAncestor")
	procIsWindowEnabled  = user32.NewProc("IsWindowEnabled")
)

func TestSpyWindows() {
	fmt.Println("╔═════════════════════════════════════════════╗")
	fmt.Println("║    企微窗口深度探测工具                      ║")
	fmt.Println("║    枚举所有子窗口 + 同进程窗口               ║")
	fmt.Println("╚═════════════════════════════════════════════╝")
	fmt.Println()

	wc, err := FindWeComWindow()
	if err != nil {
		fmt.Printf("❌ %v\n", err)
		return
	}
	fmt.Printf("✅ 主窗口: %d×%d (HWND=0x%X PID=%d)\n\n", wc.Width, wc.Height, wc.Hwnd, wc.Pid)

	// ═══ 1. 枚举主窗口的所有子窗口 (递归) ═══
	fmt.Println("════════════════════════════════════════")
	fmt.Println("  1. 主窗口的子窗口 (EnumChildWindows)")
	fmt.Println("════════════════════════════════════════")

	childCount := 0
	cbChild := syscall.NewCallback(func(hwnd syscall.Handle, _ uintptr) uintptr {
		childCount++
		printWindowInfo(hwnd, "  ", childCount)
		return 1
	})
	procEnumChildWindows.Call(uintptr(wc.Hwnd), cbChild, 0)
	if childCount == 0 {
		fmt.Println("  (无子窗口)")
	}
	fmt.Printf("\n  共 %d 个子窗口\n\n", childCount)

	// ═══ 2. 枚举所有同进程的顶级窗口 ═══
	fmt.Println("════════════════════════════════════════")
	fmt.Println("  2. 同进程 (PID=%d) 的所有顶级窗口", wc.Pid)
	fmt.Println("════════════════════════════════════════")

	topCount := 0
	cbTop := syscall.NewCallback(func(hwnd syscall.Handle, _ uintptr) uintptr {
		var pid uint32
		procGetWindowThreadPID.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&pid)))
		if pid == wc.Pid {
			topCount++
			printWindowInfo(hwnd, "  ", topCount)
		}
		return 1
	})
	procEnumWindows.Call(cbTop, 0)
	fmt.Printf("\n  共 %d 个同进程顶级窗口\n\n", topCount)

	// ═══ 3. 枚举所有可见窗口 (任何进程, 类名含特殊关键词) ═══
	fmt.Println("════════════════════════════════════════")
	fmt.Println("  3. 所有可见窗口 (类名含 Chrome/Chromium/CefBrowser/WeWork/WeCom)")
	fmt.Println("════════════════════════════════════════")

	specialCount := 0
	cbSpecial := syscall.NewCallback(func(hwnd syscall.Handle, _ uintptr) uintptr {
		visible, _, _ := procIsWindowVisible.Call(uintptr(hwnd))
		if visible == 0 {
			return 1
		}
		cls := getClassName(hwnd)
		clsLow := strings.ToLower(cls)
		if strings.Contains(clsLow, "chrome") ||
			strings.Contains(clsLow, "cef") ||
			strings.Contains(clsLow, "wecom") ||
			strings.Contains(clsLow, "wework") ||
			strings.Contains(clsLow, "wechat") ||
			strings.Contains(clsLow, "wed") ||
			strings.Contains(clsLow, "tencent") {
			specialCount++
			printWindowInfo(hwnd, "  ", specialCount)
		}
		return 1
	})
	procEnumWindows.Call(cbSpecial, 0)
	fmt.Printf("\n  共 %d 个特殊类名窗口\n\n", specialCount)

	// ═══ 4. 尝试对子窗口做截图 ═══
	fmt.Println("════════════════════════════════════════")
	fmt.Println("  4. 尝试对同进程窗口截图 (PrintWindow)")
	fmt.Println("════════════════════════════════════════")

	debugDir := "debug_spy"
	os.MkdirAll(debugDir, 0755)

	shotCount := 0
	cbShot := syscall.NewCallback(func(hwnd syscall.Handle, _ uintptr) uintptr {
		var pid uint32
		procGetWindowThreadPID.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&pid)))
		if pid != wc.Pid {
			return 1
		}

		var cr RECT
		procGetClientRect.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&cr)))
		w, h := int(cr.Right), int(cr.Bottom)
		if w < 50 || h < 50 {
			return 1 // 太小, 跳过
		}

		cls := getClassName(hwnd)
		shotCount++
		name := fmt.Sprintf("win_%02d_%s_%dx%d", shotCount, safeFileName(cls), w, h)

		// 用 PrintWindow 截图 (flag=3 包含 PW_RENDERFULLCONTENT)
		_, pngData, err := wc.screenshotHwnd(hwnd)
		if err == nil && len(pngData) > 5000 {
			path := fmt.Sprintf("%s/%s.png", debugDir, name)
			os.WriteFile(path, pngData, 0644)
			fmt.Printf("  📸 [%d] %-25s %4d×%-4d → %s (%.0f KB)\n",
				shotCount, cls, w, h, path, float64(len(pngData))/1024)
		} else {
			fmt.Printf("  ⚠️ [%d] %-25s %4d×%-4d → 截图失败或全黑\n",
				shotCount, cls, w, h)
		}

		// 也试 flag=3 (PW_CLIENTONLY | PW_RENDERFULLCONTENT)
		_, pngData2 := screenshotHwndWithFlag(hwnd, 3)
		if len(pngData2) > 5000 {
			path2 := fmt.Sprintf("%s/%s_flag3.png", debugDir, name)
			os.WriteFile(path2, pngData2, 0644)
			fmt.Printf("  📸 [%d] (flag=3)              %4d×%-4d → %s (%.0f KB)\n",
				shotCount, w, h, path2, float64(len(pngData2))/1024)
		}

		return 1
	})
	procEnumWindows.Call(cbShot, 0)

	// 同样枚举子窗口
	cbChildShot := syscall.NewCallback(func(hwnd syscall.Handle, _ uintptr) uintptr {
		var cr RECT
		procGetClientRect.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&cr)))
		w, h := int(cr.Right), int(cr.Bottom)
		if w < 50 || h < 50 {
			return 1
		}

		cls := getClassName(hwnd)
		shotCount++
		name := fmt.Sprintf("child_%02d_%s_%dx%d", shotCount, safeFileName(cls), w, h)

		_, pngData, err := wc.screenshotHwnd(hwnd)
		if err == nil && len(pngData) > 5000 {
			path := fmt.Sprintf("%s/%s.png", debugDir, name)
			os.WriteFile(path, pngData, 0644)
			fmt.Printf("  📸 [%d] (child) %-20s %4d×%-4d → %s (%.0f KB)\n",
				shotCount, cls, w, h, path, float64(len(pngData))/1024)
		}

		_, pngData2 := screenshotHwndWithFlag(hwnd, 3)
		if len(pngData2) > 5000 {
			path2 := fmt.Sprintf("%s/%s_flag3.png", debugDir, name)
			os.WriteFile(path2, pngData2, 0644)
			fmt.Printf("  📸 [%d] (child,flag=3) %-15s %4d×%-4d → %s (%.0f KB)\n",
				shotCount, cls, w, h, path2, float64(len(pngData2))/1024)
		}

		return 1
	})
	procEnumChildWindows.Call(uintptr(wc.Hwnd), cbChildShot, 0)

	fmt.Printf("\n  截图保存到: %s/\n", debugDir)
	entries, _ := os.ReadDir(debugDir)
	fmt.Printf("  共 %d 张截图\n", len(entries))

	fmt.Println("\n════════════════════════════════════════")
	fmt.Println("  探测完成!")
	fmt.Println("════════════════════════════════════════")
}

// printWindowInfo 打印单个窗口的详细信息
func printWindowInfo(hwnd syscall.Handle, indent string, idx int) {
	cls := getClassName(hwnd)
	ttl := getWindowTitle(hwnd)

	var cr RECT
	procGetClientRect.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&cr)))

	visible, _, _ := procIsWindowVisible.Call(uintptr(hwnd))
	enabled, _, _ := procIsWindowEnabled.Call(uintptr(hwnd))
	parent, _, _ := procGetParent.Call(uintptr(hwnd))

	visStr := "隐"
	if visible != 0 {
		visStr = "显"
	}
	enStr := ""
	if enabled == 0 {
		enStr = " [禁用]"
	}

	titleStr := ""
	if len(ttl) > 0 {
		if len([]rune(ttl)) > 30 {
			titleStr = " \"" + string([]rune(ttl)[:30]) + "...\""
		} else {
			titleStr = " \"" + ttl + "\""
		}
	}

	fmt.Printf("%s[%2d] %-30s %4d×%-4d  %s%s  HWND=0x%X  Parent=0x%X%s\n",
		indent, idx, cls, cr.Right, cr.Bottom, visStr, enStr, hwnd, parent, titleStr)
}

// getClassName 获取窗口类名
func getClassName(hwnd syscall.Handle) string {
	buf := make([]uint16, 256)
	procGetClassName.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&buf[0])), 256)
	return syscall.UTF16ToString(buf)
}

// getWindowTitle 获取窗口标题
func getWindowTitle(hwnd syscall.Handle) string {
	buf := make([]uint16, 256)
	procGetWindowText.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&buf[0])), 256)
	return syscall.UTF16ToString(buf)
}

// safeFileName 把类名转为安全的文件名
func safeFileName(s string) string {
	s = strings.ReplaceAll(s, ".", "_")
	s = strings.ReplaceAll(s, " ", "_")
	s = strings.ReplaceAll(s, "/", "_")
	s = strings.ReplaceAll(s, "\\", "_")
	if len(s) > 25 {
		s = s[:25]
	}
	return s
}

// screenshotHwndWithFlag 用指定 flag 做 PrintWindow 截图
// flag=1 PW_CLIENTONLY, flag=2 PW_RENDERFULLCONTENT, flag=3 两者合并
func screenshotHwndWithFlag(hwnd syscall.Handle, flag uintptr) ([]byte, []byte) {
	var cr RECT
	procGetClientRect.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&cr)))
	width, height := int(cr.Right), int(cr.Bottom)
	if width <= 0 || height <= 0 {
		return nil, nil
	}

	hdc, _, _ := procGetDC.Call(uintptr(hwnd))
	memDC, _, _ := procCreateCompatibleDC.Call(hdc)
	hBitmap, _, _ := procCreateCompatibleBitmap.Call(hdc, uintptr(width), uintptr(height))
	procSelectObject.Call(memDC, hBitmap)

	// PrintWindow with specified flag
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

	// BGRA → RGBA + PNG
	img := makeNRGBA(pixelData, width, height)
	var pngBuf bytes.Buffer
	png.Encode(&pngBuf, img)
	return pixelData, pngBuf.Bytes()
}

// makeNRGBA 从 BGRA 像素数据构建 NRGBA 图像
func makeNRGBA(pixelData []byte, width, height int) *image.NRGBA {
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
	return img
}
