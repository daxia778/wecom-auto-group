package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"syscall"
	"time"
	"unsafe"
)

// ================================================================
// 隐私设置完整流程测试 v4
// 用法: WeComAutoGroup.exe --privacy-test
//
// 目标: 尽可能静默地完成 ··· → 群管理 → 勾选checkbox → 关闭
//
// 流程分析:
//   Phase A: 打开群管理窗口 (主窗口操作)
//     - 需要: 点击 ···, 点击 群管理
//     - 方式: SendMessage 后台点击 → 不行再 SafeRealClick
//     - 截图: PrintWindow(flag=0) 后台 → 不行再 BitBlt 前台
//     - 注意: 这一步可能需要短暂前台操作
//
//   Phase B: 操作群管理窗口 (ExternalConversationManagerWindow)
//     - 完全后台! PrintWindow + SendMessage + WM_CLOSE
//     - 100% 不抢鼠标, 不抢焦点, 不影响用户操作
//
// 窗口要求: 企微窗口保持在桌面上 (不最小化即可, 可被遮挡)
// ================================================================

const groupMgmtClass = "ExternalConversationManagerWindow"

func TestPrivacyFlow() {
	fmt.Println("╔═════════════════════════════════════════════╗")
	fmt.Println("║    隐私设置完整流程 v4                       ║")
	fmt.Println("║    Phase A: 尽可能静默打开群管理              ║")
	fmt.Println("║    Phase B: 100% 后台操作 checkbox           ║")
	fmt.Println("╚═════════════════════════════════════════════╝")
	fmt.Println()

	debugDir, _ := filepath.Abs("debug_privacy")
	os.MkdirAll(debugDir, 0755)
	log := func(s string) { fmt.Printf("[%s] %s\n", time.Now().Format("15:04:05"), s) }

	for i := 3; i > 0; i-- {
		fmt.Printf("\r⏳ %d 秒后开始 (请确保企微窗口在桌面上可见)...", i)
		time.Sleep(time.Second)
	}
	fmt.Println("\r🚀 开始!                                        ")
	fmt.Println()

	wc, err := FindWeComWindow()
	if err != nil {
		fmt.Printf("❌ %v\n", err)
		return
	}
	log(fmt.Sprintf("✅ 窗口: %d×%d (HWND=0x%X)", wc.Width, wc.Height, wc.Hwnd))

	// ════════════════════════════════════════════════════
	//  Phase A: 打开群管理窗口
	// ════════════════════════════════════════════════════
	log("")
	log("━━━ Phase A: 打开群管理窗口 ━━━")

	mgmtHwnd := privFindGroupMgmtWindow(wc.Pid)
	if mgmtHwnd != 0 {
		log(fmt.Sprintf("🎯 群管理窗口已打开! HWND=0x%X → 跳到 Phase B", mgmtHwnd))
	} else {
		// 需要打开群管理: ··· → 群管理
		mgmtHwnd = phaseA_OpenGroupMgmt(wc, debugDir, log)
		if mgmtHwnd == 0 {
			log("❌ 无法打开群管理窗口, 测试终止")
			return
		}
	}

	// ════════════════════════════════════════════════════
	//  Phase B: 100% 后台操作 checkbox
	// ════════════════════════════════════════════════════
	log("")
	log("━━━ Phase B: 后台操作 checkbox (100% 静默) ━━━")
	phaseB_ToggleCheckbox(wc, mgmtHwnd, debugDir, log)

	// ════════════════════════════════════════════════════
	//  Phase C: 清理
	// ════════════════════════════════════════════════════
	log("")
	log("━━━ Phase C: 清理 ━━━")

	// 关闭群管理窗口
	log("  发送 WM_CLOSE...")
	const WM_CLOSE = 0x0010
	procSendMessage.Call(uintptr(mgmtHwnd), WM_CLOSE, 0, 0)
	time.Sleep(800 * time.Millisecond)

	if privFindGroupMgmtWindow(wc.Pid) == 0 {
		log("  ✅ 群管理窗口已关闭")
	} else {
		log("  ⚠️ 窗口未关闭, 重试...")
		procSendMessage.Call(uintptr(mgmtHwnd), WM_CLOSE, 0, 0)
	}

	// 收起聊天信息面板 (点击聊天区空白)
	wc.Click(int(float64(wc.Width)*0.45), int(float64(wc.Height)*0.5))
	time.Sleep(500 * time.Millisecond)
	log("  ✅ 面板已收起")

	// 汇总
	log("")
	log("══════════════════════════════════════")
	log("✅ 完成! 截图文件:")
	entries, _ := os.ReadDir(debugDir)
	for _, e := range entries {
		if !e.IsDir() {
			info, _ := e.Info()
			if info != nil {
				log(fmt.Sprintf("   📸 %-35s (%.0f KB)", e.Name(), float64(info.Size())/1024))
			}
		}
	}
}

// ════════════════════════════════════════════════════════
//  Phase A: 打开群管理窗口 (尽可能静默)
// ════════════════════════════════════════════════════════

func phaseA_OpenGroupMgmt(wc *WeComWindow, debugDir string, log func(string)) syscall.Handle {
	scaleX := float64(wc.Width) / refW
	scaleY := float64(wc.Height) / refH

	// ──── A1: 先尝试 OCR 定位 (后台截图) ────
	log("A1: 后台截图主窗口 (尝试 PrintWindow)...")

	var dotsX, dotsY int
	var mgmtX, mgmtY int
	ocrOK := false

	_, mainPng, mainErr := wc.Screenshot()
	if mainErr == nil && len(mainPng) > 10000 {
		privSavePng(mainPng, debugDir, "a1_main_window")
		items, ocrErr := ZhipuOCR(mainPng)
		if ocrErr == nil && len(items) > 0 {
			log(fmt.Sprintf("  ✅ 后台截图 + OCR 成功! (%d 项)", len(items)))
			ocrOK = true

			// 检查聊天信息面板是否已打开
			panelOpen := false
			var mgmtItem *OCRItem
			for idx := range items {
				if strings.Contains(items[idx].Text, "群管理") &&
					items[idx].CX > wc.Width/2 {
					mgmtItem = &items[idx]
					panelOpen = true
				}
			}

			if mgmtItem != nil {
				// 面板已打开且找到群管理
				mgmtX, mgmtY = mgmtItem.CX, mgmtItem.CY
				log(fmt.Sprintf("  📍 面板已打开, 找到「群管理」@(%d,%d)", mgmtX, mgmtY))
			} else if !panelOpen {
				// 需要先点 ···
				dotsX = wc.Width - 50
				dotsY = 47
				// OCR 找标题栏, 精确定位 ···
				headerY := privFindHeaderY(items, wc.Width, wc.Height)
				if headerY > 0 {
					dotsY = headerY
				}
				log(fmt.Sprintf("  📍 OCR定位 ··· @(%d,%d)", dotsX, dotsY))
			}
		}
	}

	if !ocrOK {
		log("  ⚠️ 后台截图失败, 用坐标估算")
		dotsX = wc.Width - int(50*scaleX)
		dotsY = int(47 * scaleY)
	}

	// ──── A2: 点击 ··· 打开聊天信息面板 ────
	if mgmtX == 0 { // 群管理还没找到, 需要先点 ···
		log(fmt.Sprintf("A2: 点击 ··· @(%d,%d) [SendMessage 后台]", dotsX, dotsY))
		wc.Click(dotsX, dotsY)
		time.Sleep(1500 * time.Millisecond)

		// 再次截图看面板是否打开了
		_, panelPng, panelErr := wc.Screenshot()
		if panelErr == nil && len(panelPng) > 10000 {
			privSavePng(panelPng, debugDir, "a2_panel")
			items2, err2 := ZhipuOCR(panelPng)
			if err2 == nil {
				for idx := range items2 {
					if strings.Contains(items2[idx].Text, "群管理") &&
						items2[idx].CX > wc.Width/2 {
						mgmtX, mgmtY = items2[idx].CX, items2[idx].CY
						log(fmt.Sprintf("  📍 OCR找到「群管理」@(%d,%d)", mgmtX, mgmtY))
						break
					}
				}
			}
		}

		if mgmtX == 0 {
			// 坐标回退
			mgmtX = int(float64(wc.Width) * 0.82)
			mgmtY = int(float64(wc.Height) * 0.39)
			log(fmt.Sprintf("  ⚠️ OCR 未找到, 坐标回退 (%d,%d)", mgmtX, mgmtY))
		}
	}

	// ──── A3: 点击「群管理」────
	log(fmt.Sprintf("A3: 点击「群管理」@(%d,%d) [SendMessage 后台]", mgmtX, mgmtY))
	wc.Click(mgmtX, mgmtY)

	// 等待 ExternalConversationManagerWindow 出现
	log("  ⏳ 等待群管理窗口出现...")
	var mgmtHwnd syscall.Handle
	for wait := 0; wait < 10; wait++ {
		time.Sleep(500 * time.Millisecond)
		mgmtHwnd = privFindGroupMgmtWindow(wc.Pid)
		if mgmtHwnd != 0 {
			log(fmt.Sprintf("  ✅ 群管理窗口出现! HWND=0x%X (等待 %dms)", mgmtHwnd, (wait+1)*500))
			return mgmtHwnd
		}
	}

	// ──── A4: SendMessage 没用? 用 SafeRealClick 重试 ────
	log("  ⚠️ SendMessage 点击无效, 改用 SafeRealClick (短暂抢鼠标)...")

	// 重新点击 ··· (SafeRealClick)
	log(fmt.Sprintf("A4: SafeRealClick ··· @(%d,%d)", dotsX, dotsY))
	wc.SafeRealClick(dotsX, dotsY)
	time.Sleep(1500 * time.Millisecond)

	// 截图 + OCR 找群管理
	_, fgPng, fgErr := wc.SafeScreenshotForeground()
	if fgErr == nil && len(fgPng) > 10000 {
		privSavePng(fgPng, debugDir, "a4_foreground")
		items3, err3 := ZhipuOCR(fgPng)
		if err3 == nil {
			for idx := range items3 {
				if strings.Contains(items3[idx].Text, "群管理") &&
					items3[idx].CX > wc.Width/2 {
					mgmtX, mgmtY = items3[idx].CX, items3[idx].CY
					log(fmt.Sprintf("  📍 前台OCR找到「群管理」@(%d,%d)", mgmtX, mgmtY))
					break
				}
			}
		}
	}

	log(fmt.Sprintf("A4: SafeRealClick 群管理 @(%d,%d)", mgmtX, mgmtY))
	wc.SafeRealClick(mgmtX, mgmtY)

	for wait := 0; wait < 10; wait++ {
		time.Sleep(500 * time.Millisecond)
		mgmtHwnd = privFindGroupMgmtWindow(wc.Pid)
		if mgmtHwnd != 0 {
			log(fmt.Sprintf("  ✅ 群管理窗口出现! HWND=0x%X", mgmtHwnd))
			return mgmtHwnd
		}
	}

	log("  ❌ 所有方式均失败, 群管理窗口未出现")
	return 0
}

// ════════════════════════════════════════════════════════
//  Phase B: 后台操作 checkbox (100% 静默)
// ════════════════════════════════════════════════════════

func phaseB_ToggleCheckbox(wc *WeComWindow, mgmtHwnd syscall.Handle, debugDir string, log func(string)) {
	// B1: PrintWindow 后台截图群管理窗口
	log("B1: PrintWindow 后台截图群管理窗口...")
	mgmtImg, mgmtPng, mgmtErr := wc.screenshotHwnd(mgmtHwnd)
	if mgmtErr != nil {
		log(fmt.Sprintf("  ❌ 截图失败: %v", mgmtErr))
		return
	}
	privSavePng(mgmtPng, debugDir, "b1_group_mgmt")
	log(fmt.Sprintf("  ✅ 截图成功! (%.0f KB) — 纯后台, 不抢焦点", float64(len(mgmtPng))/1024))

	// B2: OCR 找 checkbox
	log("B2: OCR 定位 checkbox...")
	items, ocrErr := ZhipuOCR(mgmtPng)
	if ocrErr != nil {
		log(fmt.Sprintf("  ❌ OCR 失败: %v", ocrErr))
		return
	}

	target := privFindCheckboxText(items)
	if target == nil {
		log("  ❌ 未找到「禁止互相添加为联系人」")
		for _, it := range items {
			log(fmt.Sprintf("    \"%s\" @(%d,%d)", it.Text, it.CX, it.CY))
		}
		return
	}

	checkX := target.X1 - 20
	checkY := target.CY
	log(fmt.Sprintf("  📍 「%s」文字@(%d,%d), checkbox@(%d,%d)",
		privTrunc(target.Text, 12), target.CX, target.CY, checkX, checkY))

	// B3: 检查当前状态
	log("B3: 检查勾选状态...")
	alreadyChecked := false
	if mgmtImg != nil {
		alreadyChecked = IsCheckboxChecked(mgmtImg, checkX, checkY)
	}

	if alreadyChecked {
		log("  ✅ 已勾选! 无需操作")
		return
	}
	log("  ⬜ 未勾选, 开始后台点击...")

	// B4: SendMessage 后台点击 (多偏移尝试)
	log("B4: SendMessage 后台点击 checkbox...")
	offsets := []int{0, -5, -10, 5, -15, 10, -25}
	verified := false

	for attempt, dx := range offsets {
		tryX := checkX + dx
		log(fmt.Sprintf("  🖱️ [%d/%d] 后台点击 (%d,%d) — 不动鼠标!",
			attempt+1, len(offsets), tryX, checkY))

		privClickOnWindow(mgmtHwnd, tryX, checkY)
		time.Sleep(800 * time.Millisecond)

		// 后台截图验证
		verifyImg, verifyPng, verifyErr := wc.screenshotHwnd(mgmtHwnd)
		if verifyErr != nil {
			continue
		}
		privSavePng(verifyPng, debugDir, fmt.Sprintf("b4_verify_%d", attempt+1))

		if verifyImg != nil && IsCheckboxChecked(verifyImg, tryX, checkY) {
			log("  ✅ 验证通过! checkbox 已勾选! (纯后台完成)")
			verified = true
			break
		}
		log("  ⚠️ 像素验证未通过, 尝试下个偏移...")
	}

	if !verified {
		log("  ⚠️ 多次尝试后未确认勾选")
		// 最终截图保存, 让用户手动确认
		_, finalPng, _ := wc.screenshotHwnd(mgmtHwnd)
		if finalPng != nil {
			privSavePng(finalPng, debugDir, "b4_final_check")
			log("  📸 已保存最终截图, 请检查 b4_final_check.png")
		}
	}
}

// ════════════════════════════════════════════════════════
//  工具函数
// ════════════════════════════════════════════════════════

func privFindGroupMgmtWindow(pid uint32) syscall.Handle {
	var found syscall.Handle
	cb := syscall.NewCallback(func(hwnd syscall.Handle, _ uintptr) uintptr {
		var winPid uint32
		procGetWindowThreadPID.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&winPid)))
		if winPid != pid {
			return 1
		}
		visible, _, _ := procIsWindowVisible.Call(uintptr(hwnd))
		if visible == 0 {
			return 1
		}
		className := make([]uint16, 256)
		procGetClassName.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&className[0])), 256)
		if syscall.UTF16ToString(className) == groupMgmtClass {
			found = hwnd
			return 0
		}
		return 1
	})
	procEnumWindows.Call(cb, 0)
	return found
}

func privClickOnWindow(hwnd syscall.Handle, x, y int) {
	lParam := uintptr(y<<16 | x)
	procSendMessage.Call(uintptr(hwnd), WM_LBUTTONDOWN, MK_LBUTTON, lParam)
	time.Sleep(80 * time.Millisecond)
	procSendMessage.Call(uintptr(hwnd), WM_LBUTTONUP, 0, lParam)
}

func privSavePng(data []byte, dir, name string) {
	if data == nil || len(data) == 0 {
		return
	}
	os.WriteFile(filepath.Join(dir, name+".png"), data, 0644)
}

func privPrintOCR(items []OCRItem, log func(string), maxItems int) {
	for i, it := range items {
		if i >= maxItems {
			log(fmt.Sprintf("    ... 还有 %d 项", len(items)-maxItems))
			break
		}
		log(fmt.Sprintf("    [%2d] %-22s @(%3d,%3d)", i, privTrunc(it.Text, 20), it.CX, it.CY))
	}
}

func privFindHeaderY(items []OCRItem, wWidth, wHeight int) int {
	maxY := int(float64(wHeight) * 0.10)
	minX := int(float64(wWidth) * 0.25)
	bestY := 0
	for _, it := range items {
		if it.CX < minX || it.CY > maxY || it.CY < 20 {
			continue
		}
		txt := strings.TrimSpace(it.Text)
		if txt == "搜索" || txt == "十" || txt == "+" || txt == "×" || txt == "x" {
			continue
		}
		if bestY == 0 || it.CY < bestY {
			bestY = it.CY
		}
	}
	return bestY
}

func privFindCheckboxText(items []OCRItem) *OCRItem {
	keywords := []string{
		"禁止互相添加为联系人",
		"禁止互相添加",
		"互相添加为联系人",
		"添加为联系人",
	}
	for _, kw := range keywords {
		for idx := range items {
			if strings.Contains(items[idx].Text, kw) {
				return &items[idx]
			}
		}
	}
	return nil
}

func privTrunc(s string, n int) string {
	r := []rune(s)
	if len(r) > n {
		return string(r[:n]) + "…"
	}
	return s
}

func privListWindows() {
	cb := syscall.NewCallback(func(hwnd syscall.Handle, _ uintptr) uintptr {
		visible, _, _ := procIsWindowVisible.Call(uintptr(hwnd))
		if visible == 0 {
			return 1
		}
		className := make([]uint16, 256)
		procGetClassName.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&className[0])), 256)
		cls := syscall.UTF16ToString(className)
		clsLow := strings.ToLower(cls)
		if !strings.Contains(clsLow, "we") && !strings.Contains(clsLow, "external") &&
			!strings.Contains(clsLow, "perry") {
			return 1
		}
		title := make([]uint16, 256)
		procGetWindowText.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&title[0])), 256)
		var cr RECT
		procGetClientRect.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&cr)))
		fmt.Printf("   %-30s %4d×%-4d  HWND=0x%X  %s\n", cls, cr.Right, cr.Bottom, hwnd, syscall.UTF16ToString(title))
		return 1
	})
	procEnumWindows.Call(cb, 0)
}

func privDoCleanup(wc *WeComWindow, fgItems []OCRItem, log func(string)) {
	wc.Click(int(float64(wc.Width)*0.45), int(float64(wc.Height)*0.5))
	time.Sleep(500 * time.Millisecond)
}

func privCapture(wc *WeComWindow, dir, name string, log func(string)) (interface{}, []OCRItem) {
	return nil, nil
}

func privScreenshotOnly(wc *WeComWindow, dir, name string) interface{} {
	return nil
}
