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
// 后台点击验证工具: 测试 SendMessage 对企微主窗口是否有效
// 用法: WeComAutoGroup.exe --click-test
// ================================================================

func TestBackendClick() {
	fmt.Println("╔═════════════════════════════════════════════╗")
	fmt.Println("║    后台点击验证工具                          ║")
	fmt.Println("║    验证 SendMessage + PrintWindow 能力       ║")
	fmt.Println("╚═════════════════════════════════════════════╝")
	fmt.Println()

	wc, err := FindWeComWindow()
	if err != nil {
		fmt.Printf("❌ %v\n", err)
		return
	}
	fmt.Printf("✅ 窗口: %d×%d (HWND=0x%X PID=%d)\n\n", wc.Width, wc.Height, wc.Hwnd, wc.Pid)

	debugDir, _ := filepath.Abs("debug_click")
	os.MkdirAll(debugDir, 0755)
	log := func(s string) { fmt.Printf("[%s] %s\n", time.Now().Format("15:04:05"), s) }

	// ════ Test 1: PrintWindow 主窗口 (flag=1 vs flag=3) ════
	log("═══ Test 1: PrintWindow 主窗口截图 ═══")

	// flag=1 (PW_CLIENTONLY)
	_, png1 := screenshotHwndWithFlag(wc.Hwnd, 1)
	if len(png1) > 5000 {
		os.WriteFile(filepath.Join(debugDir, "main_flag1.png"), png1, 0644)
		log(fmt.Sprintf("  flag=1: %d bytes ✅", len(png1)))
	} else {
		log(fmt.Sprintf("  flag=1: %d bytes ⚠️ (可能全黑)", len(png1)))
	}

	// flag=2 (PW_RENDERFULLCONTENT)
	_, png2 := screenshotHwndWithFlag(wc.Hwnd, 2)
	if len(png2) > 5000 {
		os.WriteFile(filepath.Join(debugDir, "main_flag2.png"), png2, 0644)
		log(fmt.Sprintf("  flag=2: %d bytes ✅", len(png2)))
	} else {
		log(fmt.Sprintf("  flag=2: %d bytes ⚠️ (可能全黑)", len(png2)))
	}

	// flag=3 (PW_CLIENTONLY | PW_RENDERFULLCONTENT)
	_, png3 := screenshotHwndWithFlag(wc.Hwnd, 3)
	if len(png3) > 5000 {
		os.WriteFile(filepath.Join(debugDir, "main_flag3.png"), png3, 0644)
		log(fmt.Sprintf("  flag=3: %d bytes ✅", len(png3)))
	} else {
		log(fmt.Sprintf("  flag=3: %d bytes ⚠️ (可能全黑)", len(png3)))
	}

	// ════ Test 2: 对 TitleBarWindow 截图 (同进程子窗口) ════
	log("")
	log("═══ Test 2: 同进程可见窗口截图 ═══")
	cb := syscall.NewCallback(func(hwnd syscall.Handle, _ uintptr) uintptr {
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
		if int(cr.Right) < 100 || int(cr.Bottom) < 100 {
			return 1
		}

		cls := getClassName(hwnd)
		_, data := screenshotHwndWithFlag(hwnd, 3)
		status := "⚠️ 黑"
		if len(data) > 5000 {
			fname := fmt.Sprintf("%s_%dx%d.png", safeFileName(cls), cr.Right, cr.Bottom)
			os.WriteFile(filepath.Join(debugDir, fname), data, 0644)
			status = "✅ OK"
		}
		log(fmt.Sprintf("  %-30s %4d×%-4d  %s (%d bytes)",
			cls, cr.Right, cr.Bottom, status, len(data)))
		return 1
	})
	procEnumWindows.Call(cb, 0)

	// ════ Test 3: SendMessage 后台点击 + 按钮 ════
	log("")
	log("═══ Test 3: SendMessage 后台点击 ═══")
	log("  👆 后台点击 + 按钮 (坐标回退)...")

	scaleX := float64(wc.Width) / 1046.0
	plusX := int(283 * scaleX)
	plusY := int(27 * (float64(wc.Height) / 705.0))
	log(fmt.Sprintf("  坐标: (%d, %d)", plusX, plusY))

	wc.Click(plusX, plusY)
	time.Sleep(2000 * time.Millisecond)

	// 检查弹窗是否出现
	popupHwnd := wc.findPopup("weWorkSelectUser")
	if popupHwnd != 0 {
		var pcr RECT
		procGetClientRect.Call(uintptr(popupHwnd), uintptr(unsafe.Pointer(&pcr)))
		log(fmt.Sprintf("  ✅ 弹窗已出现! HWND=0x%X 尺寸=%d×%d", popupHwnd, pcr.Right, pcr.Bottom))

		// 尝试截图弹窗
		_, popupPng, pErr := wc.screenshotHwnd(popupHwnd)
		if pErr == nil && len(popupPng) > 5000 {
			os.WriteFile(filepath.Join(debugDir, "popup_weWorkSelectUser.png"), popupPng, 0644)
			log(fmt.Sprintf("  📸 弹窗截图成功! (%d bytes)", len(popupPng)))
		} else {
			log("  ⚠️ 弹窗截图失败或全黑")
		}

		// 关闭弹窗
		wc.ClosePopup("weWorkSelectUser")
		log("  ↩️ 弹窗已关闭")
	} else {
		log("  ❌ 弹窗未出现! SendMessage 点击可能无效")
		log("  重试: SafeRealClick (前台点击)...")
		wc.SafeRealClick(plusX, plusY)
		time.Sleep(2000 * time.Millisecond)
		popupHwnd = wc.findPopup("weWorkSelectUser")
		if popupHwnd != 0 {
			log("  ✅ SafeRealClick 成功! 弹窗出现了")
			wc.ClosePopup("weWorkSelectUser")
		} else {
			log("  ❌ SafeRealClick 也失败了")
		}
	}

	// ════ Test 4: 汇总 ════
	log("")
	log("══════════════════════════════════════")
	log("✅ 测试完成! 截图:")
	entries, _ := os.ReadDir(debugDir)
	for _, e := range entries {
		if !e.IsDir() {
			info, _ := e.Info()
			if info != nil {
				log(fmt.Sprintf("   📸 %-35s  (%.0f KB)", e.Name(), float64(info.Size())/1024))
			}
		}
	}
}
