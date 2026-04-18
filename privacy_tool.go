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
//  Phase A: 打开群管理窗口 (v7)
//
//  核心发现 (来自交互测试):
//    - ··· 按钮必须用 SendMessage (bgclick) 点击, 坐标 w*0.945, h*0.04
//    - SafeRealClick 点击 ··· 只会 toggle 成员面板, 不是聊天信息面板!
//    - 「群管理」在聊天信息面板中, 也用 SendMessage 点击 (OCR定位)
//    - 群管理打开后会创建 ExternalConversationManagerWindow 子窗口
//    - Phase B 可以 100% 后台操作该子窗口
//
//  流程:
//    1. 拉高窗口 (确保面板完整显示)
//    2. SendMessage 点击 ··· (w*0.945, h*0.04)
//    3. 前台截图 + OCR 找「群管理」坐标
//    4. SendMessage 点击「群管理」
//    5. 等待 ExternalConversationManagerWindow 出现
// ════════════════════════════════════════════════════════

func phaseA_OpenGroupMgmt(wc *WeComWindow, debugDir string, log func(string)) syscall.Handle {

	// ──── Step 1: 记录原窗口尺寸, 拉高窗口 ────
	log("A1: 拉高窗口以显示完整面板...")
	smH, _, _ := user32.NewProc("GetSystemMetrics").Call(1) // SM_CYSCREEN
	screenH := int(smH)

	var wr RECT
	procGetWindowRect.Call(uintptr(wc.Hwnd), uintptr(unsafe.Pointer(&wr)))
	origLeft := int(wr.Left)
	origTop := int(wr.Top)
	origW := int(wr.Right - wr.Left)
	origH := int(wr.Bottom - wr.Top)

	if origH < screenH-100 {
		newH := screenH - 40
		newTop := 20
		log(fmt.Sprintf("  窗口从 %d×%d 拉高到 %d×%d", origW, origH, origW, newH))
		procSetWindowPos.Call(uintptr(wc.Hwnd), 0,
			uintptr(origLeft), uintptr(newTop), uintptr(origW), uintptr(newH),
			SWP_NOACTIVATE)
		time.Sleep(800 * time.Millisecond)

		// 更新窗口尺寸
		var cr RECT
		procGetClientRect.Call(uintptr(wc.Hwnd), uintptr(unsafe.Pointer(&cr)))
		wc.Width = int(cr.Right)
		wc.Height = int(cr.Bottom)
		log(fmt.Sprintf("  新客户区: %d×%d", wc.Width, wc.Height))
	}

	// defer: 无论成功失败都恢复窗口尺寸
	defer func() {
		if origH < screenH-100 {
			procSetWindowPos.Call(uintptr(wc.Hwnd), 0,
				uintptr(origLeft), uintptr(origTop), uintptr(origW), uintptr(origH),
				SWP_NOACTIVATE)
			time.Sleep(300 * time.Millisecond)
			var cr RECT
			procGetClientRect.Call(uintptr(wc.Hwnd), uintptr(unsafe.Pointer(&cr)))
			wc.Width = int(cr.Right)
			wc.Height = int(cr.Bottom)
		}
	}()

	// ──── Step 2: SendMessage 点击 ··· 打开聊天信息面板 ────
	// Python 版坐标: w*0.945, h*0.04. 交互测试验证有效
	dotsX := int(float64(wc.Width) * 0.945)
	dotsY := int(float64(wc.Height) * 0.04)
	log(fmt.Sprintf("A2: SendMessage 点击 ··· @(%d,%d)...", dotsX, dotsY))
	wc.Click(dotsX, dotsY)
	humanDelay(2000, 500)

	// ──── Step 3: 前台截图 + OCR 找「群管理」 ────
	for attempt := 0; attempt < 3; attempt++ {
		log(fmt.Sprintf("A3: 截图寻找「群管理」(尝试 %d/3)...", attempt+1))

		_, panelPng, panelErr := wc.SafeScreenshotForeground()
		if panelErr != nil || len(panelPng) <= 10000 {
			// 回退到后台截图
			_, panelPng, panelErr = wc.Screenshot()
		}
		if panelErr != nil || len(panelPng) <= 10000 {
			log("  ⚠️ 截图失败")
			humanDelay(1000, 200)
			continue
		}
		privSavePng(panelPng, debugDir, fmt.Sprintf("a3_panel_%d", attempt+1))

		panelItems, ocrErr := ZhipuOCR(panelPng)
		if ocrErr != nil {
			log(fmt.Sprintf("  ⚠️ OCR 失败: %v", ocrErr))
			continue
		}

		// 查找聊天信息面板中的「群管理」
		panelX := wc.Width / 2
		var mgmtItem *OCRItem
		for idx := range panelItems {
			if panelItems[idx].CX > panelX && strings.Contains(panelItems[idx].Text, "群管理") {
				// 排除 "群管理员" 等误匹配
				if panelItems[idx].Text == "群管理" || panelItems[idx].Text == "群管理 >" {
					mgmtItem = &panelItems[idx]
					break
				}
			}
		}

		if mgmtItem != nil {
			log(fmt.Sprintf("  ✅ 找到「%s」@(%d,%d) — SendMessage 点击", mgmtItem.Text, mgmtItem.CX, mgmtItem.CY))
			wc.Click(mgmtItem.CX, mgmtItem.CY)
			humanDelay(2000, 500)

			hwnd := phaseA_WaitForMgmtWindow(wc, log)
			if hwnd != 0 {
				return hwnd
			}

			// SendMessage 没打开窗口, 尝试 SafeRealClick
			log("  ⚠️ SendMessage 未打开群管理窗口, 尝试 SafeRealClick...")
			wc.SafeRealClick(mgmtItem.CX, mgmtItem.CY)
			humanDelay(2000, 500)
			hwnd = phaseA_WaitForMgmtWindow(wc, log)
			if hwnd != 0 {
				return hwnd
			}
			log("  ⚠️ SafeRealClick 也未打开, 重新打开聊天信息面板...")
			// 面板状态可能变了, 重新点 ···
			wc.Click(dotsX, dotsY)
			humanDelay(2000, 500)
		} else {
			// 检查是否看到聊天信息面板
			var panelTexts []string
			for _, it := range panelItems {
				if it.CX > panelX {
					panelTexts = append(panelTexts, fmt.Sprintf("「%s」@%d,%d", it.Text, it.CX, it.CY))
				}
			}
			log(fmt.Sprintf("  面板内容: %v", panelTexts))

			if attempt < 2 {
				// 再次点击 ··· 尝试切换
				log("  重新点击 ···...")
				wc.Click(dotsX, dotsY)
				humanDelay(2000, 500)
			}
		}
	}

	log("  ❌ 未找到「群管理」")
	wc.Click(int(float64(wc.Width)*0.45), int(float64(wc.Height)*0.50))
	return 0
}

// phaseA_Screenshot 截图 + OCR (前台优先, 后台回退)
func phaseA_Screenshot(wc *WeComWindow, debugDir string, log func(string)) []OCRItem {
	_, fgPng, fgErr := wc.SafeScreenshotForeground()
	if fgErr == nil && len(fgPng) > 10000 {
		privSavePng(fgPng, debugDir, "screenshot")
		items, ocrErr := ZhipuOCR(fgPng)
		if ocrErr == nil && len(items) > 0 {
			log(fmt.Sprintf("  ✅ 截图+OCR 成功 (%d 项)", len(items)))
			return items
		}
	}
	_, bgPng, bgErr := wc.Screenshot()
	if bgErr == nil && len(bgPng) > 10000 {
		items, ocrErr := ZhipuOCR(bgPng)
		if ocrErr == nil {
			log(fmt.Sprintf("  ✅ 后台截图+OCR 成功 (%d 项)", len(items)))
			return items
		}
	}
	log("  ❌ 截图失败")
	return nil
}

// phaseA_FindDots 定位 ··· 按钮位置 (OCR + 标题栏估算)
func phaseA_FindDots(items []OCRItem, wc *WeComWindow) (int, int) {
	dotsX := wc.Width - 45
	dotsY := 46

	// OCR 直接找 ···
	for _, it := range items {
		txt := strings.TrimSpace(it.Text)
		if it.CY < int(float64(wc.Height)*0.10) && it.CX > wc.Width/2 {
			if txt == "···" || txt == "..." || txt == "…" || txt == ".." || txt == "·.." {
				return it.CX, it.CY
			}
		}
	}

	// 标题栏 y 坐标估算
	headerY := privFindHeaderY(items, wc.Width, wc.Height)
	if headerY > 0 {
		dotsY = headerY
	}
	return dotsX, dotsY
}

// phaseA_WaitForMgmtWindow 等待群管理窗口出现 (轮询 5s)
func phaseA_WaitForMgmtWindow(wc *WeComWindow, log func(string)) syscall.Handle {
	log("  ⏳ 等待群管理窗口出现...")
	for wait := 0; wait < 10; wait++ {
		humanDelay(500, 100)
		hwnd := privFindGroupMgmtWindow(wc.Pid)
		if hwnd != 0 {
			log(fmt.Sprintf("  ✅ 群管理窗口出现! HWND=0x%X (%dms)", hwnd, (wait+1)*500))
			return hwnd
		}
	}
	log("  ❌ 群管理窗口 5s 超时未出现")
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
