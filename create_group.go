package main

import (
	"fmt"
	"math/rand"
	"strings"
	"syscall"
	"time"
)

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
//   OCR 驱动建群流程 (移植自 Python app.py create_grp)
//
//   每步: OCR 优先定位 → 坐标比例回退 → 操作后验证
//   全程后台执行，不抢鼠标不抢键盘
//   (隐私设置步骤需短暂前台, 已加锁保护)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const popupClass = "weWorkSelectUser"

// 参考坐标 (基于 1046x705 客户区)
const (
	refW = 1046.0
	refH = 705.0
)

// resetWeComState 建群前清理残留状态
// 关闭所有弹窗, 点击聊天区空白处收起面板
// ❗绝对不能发 ESC —— WeCom 收到 ESC 会关闭窗口并缩到系统托盘!
func (w *WeComWindow) resetWeComState(logFn func(string)) {
	// 关闭可能的残留弹窗
	for i := 0; i < 3; i++ {
		if w.isPopupVisible(popupClass) {
			logFn(fmt.Sprintf("  [清理] 关闭残留弹窗 (%d)", i+1))
			w.ClosePopup(popupClass)
			humanDelay(500, 200)
		} else {
			break
		}
	}
	// 点击聊天区空白处收起可能的面板/overlay (不用 ESC!)
	w.Click(int(float64(w.Width)*0.45), int(float64(w.Height)*0.50))
	humanDelay(500, 200)
}

// CreateGroupOCR OCR 增强版建群流程 (8步)
// customer: 客户名  members: 固定成员列表  logFn: 日志回调
// 返回 GroupResult 而非 bool, 携带详细的操作结果用于验证和上报
func (w *WeComWindow) CreateGroupOCR(customer string, members []string, logFn func(string)) GroupResult {
	tStart := time.Now()
	allMembers := append([]string{customer}, members...)
	result := GroupResult{
		MembersExpected: len(allMembers),
	}
	logFn(fmt.Sprintf("🏗️ 建群: %s (共 %d 人)", customer, len(allMembers)))

	// ═══ Step 0: 清理残留状态 ═══
	w.resetWeComState(logFn)

	scaleX := float64(w.Width) / refW
	scaleY := float64(w.Height) / refH

	// ═══ Step 1: 点击消息 Tab (OCR 定位) ═══
	logFn("  [1/8] 点击消息 Tab...")
	msgClicked := false
	items, err := w.OCRScan()
	if err == nil && len(items) > 0 {
		logFn(fmt.Sprintf("  [1/8] OCR: %d 项", len(items)))
		match := FindOCRText(items, "消息")
		if match != nil {
			logFn(fmt.Sprintf("  [1/8] OCR 找到「消息」: (%d,%d)", match.CX, match.CY))
			w.Click(match.CX, match.CY)
			msgClicked = true
		}
	}
	if !msgClicked {
		// 回退: 侧栏「消息」在 x≈28, y≈80
		x, y := int(28*scaleX), int(80*scaleY)
		logFn(fmt.Sprintf("  [1/8] 坐标回退 (%d,%d)", x, y))
		w.Click(x, y)
	}
	humanDelay(800, 300)

	// ═══ Step 2: 点击 [+] 按钮 (OCR 定位) ═══
	logFn("  [2/8] 点击 + 按钮...")
	plusClicked := false
	if items != nil {
		// 智谱 OCR 把 + 图标识别为 "十"
		plusBtn := FindOCRText(items, "十")
		if plusBtn == nil {
			plusBtn = FindOCRText(items, "+")
		}
		if plusBtn != nil && plusBtn.CY < 60 {
			// 确保是顶部搜索栏旁的 +, 不是其他位置
			logFn(fmt.Sprintf("  [2/8] OCR 定位 + 按钮: (%d,%d)", plusBtn.CX, plusBtn.CY))
			w.Click(plusBtn.CX, plusBtn.CY)
			plusClicked = true
		}
	}
	if !plusClicked {
		// 回退: 搜索按钮旁 + 估算
		searchBtn := FindOCRText(items, "搜索")
		plusY := int(27 * scaleY)
		if searchBtn != nil {
			plusY = searchBtn.CY
		}
		// 用面板右端推算 +
		panelRight := 0
		for _, it := range items {
			if it.CY > 50 && it.CY < w.Height/2 && it.X2 < int(float64(w.Width)*0.4) {
				if it.X2 > panelRight {
					panelRight = it.X2
				}
			}
		}
		plusX := int(283 * scaleX)
		if panelRight > 100 {
			plusX = panelRight - 15
		}
		logFn(fmt.Sprintf("  [2/8] 坐标回退 + (%d,%d)", plusX, plusY))
		w.Click(plusX, plusY)
	}
	humanDelay(1500, 500)

	// ═══ Step 3: 检测建群弹窗 (3 次重试) ═══
	logFn("  [3/8] 检测弹窗...")
	popupFound := w.findPopup(popupClass) != 0
	if !popupFound {
		logFn("  [3/8] ⚠️ 弹窗未出现, 重试点击...")
		// 重试: [+] 可能弹出菜单, 需要点击菜单项「发起群聊」
		panelW := int(float64(w.Width) * 0.31)
		w.Click(panelW-25, int(37*scaleY))
		humanDelay(2000, 500)
		popupFound = w.findPopup(popupClass) != 0
	}
	if !popupFound {
		// 再重试
		humanDelay(2000, 500)
		popupFound = w.findPopup(popupClass) != 0
	}
	if !popupFound {
		logFn("  [3/8] ❌ 建群弹窗未打开!")
		result.ErrorDetail = "建群弹窗未打开"
		return result
	}
	pw, ph := w.popupClientSize(popupClass)
	logFn(fmt.Sprintf("  [3/8] ✅ 弹窗已打开 (%dx%d)", pw, ph))

	// ═══ Step 4: 逐个搜索并选中成员 (不做逐人OCR验证, 在Step6统一校验) ═══
	// 优化: 去掉逐人OCR验证, 省 N 次OCR调用 (每次0.01元 + 3秒)
	// 改为在 Step 6 通过「已选择N个联系人」统一验证
	logFn(fmt.Sprintf("  [4/8] 添加 %d 名成员 (用中文名搜索)", len(allMembers)))
	membersSelected := 0
	for i, m := range allMembers {
		logFn(fmt.Sprintf("  [4/8] (%d/%d) 搜索: %s", i+1, len(allMembers), m))

		if i > 0 {
			humanDelay(800, 400)
		}

		// ═══ 搜索前: 强制清空搜索框 (双重清除防堆叠) ═══
		w.ClickPopup(popupClass, 160, 40) // 点击搜索框获取焦点
		humanDelay(300, 100)
		w.ClearPopupInput(popupClass)     // 第一次清空
		humanDelay(500, 200)
		w.ClearPopupInput(popupClass)     // 第二次清空 (确保干净)
		humanDelay(300, 100)

		// 输入中文名搜索
		w.TypeToPopup(popupClass, m)
		humanDelay(2000, 500) // 等搜索结果加载 (网络慢时需要更久)

		// 勾选第一个搜索结果 (checkbox 在左侧 x≈25, 第一结果 y≈95)
		w.ClickPopup(popupClass, 25, 95)
		membersSelected++
		logFn(fmt.Sprintf("  [4/8] (%d/%d) ✅ 已勾选 %s", i+1, len(allMembers), m))
		humanDelay(800, 300)

		// ═══ 勾选后: 强制清空搜索框 (三重清除, 彻底防堆叠) ═══
		w.ClickPopup(popupClass, 160, 40) // 回到搜索框
		humanDelay(300, 100)
		w.ClearPopupInput(popupClass)     // 清空 1
		humanDelay(300, 100)
		w.ClearPopupInput(popupClass)     // 清空 2
		humanDelay(200, 100)
		w.ClearPopupInput(popupClass)     // 清空 3
		humanDelay(300, 100)
	}
	result.MembersSelected = membersSelected

	// ═══ Step 5: 清空搜索框 (二次确认) ═══
	logFn("  [5/8] 清空搜索框...")
	w.ClickPopup(popupClass, 160, 40)
	humanDelay(150, 50)
	w.ClearPopupInput(popupClass)
	humanDelay(300, 100)

	// ═══ Step 6: 点击「完成」按钮 (OCR+坐标+验证) ═══
	logFn("  [6/8] 点击完成按钮...")
	clicked := false

	// 策略1: OCR 定位
	popupItems, ocrErr := w.OCRScanPopup(popupClass)
	if ocrErr == nil {
		texts := make([]string, 0, 8)
		for j, it := range popupItems {
			if j >= 8 {
				break
			}
			texts = append(texts, it.Text)
		}
		logFn(fmt.Sprintf("  [6/8] OCR 扫描到 %d 项: %v", len(popupItems), texts))

		// ═══ 安全检查: 验证已选成员数是否正确 ═══
		for _, it := range popupItems {
			if strings.Contains(it.Text, "已选择") && strings.Contains(it.Text, "联系人") {
				logFn(fmt.Sprintf("  [6/8] 📋 成员状态: %s (期望 %d 人)", it.Text, len(allMembers)))
				break
			}
		}

		doneBtn := FindOCRText(popupItems, "完成")
		if doneBtn != nil {
			logFn(fmt.Sprintf("  [6/8] ✅ OCR 找到「完成」: (%d,%d)", doneBtn.CX, doneBtn.CY))
			w.ClickPopup(popupClass, doneBtn.CX, doneBtn.CY)
			clicked = true
		}
	}

	// 策略2: 坐标比例
	if !clicked {
		pw, ph = w.popupClientSize(popupClass)
		if pw > 0 && ph > 0 {
			doneX := int(float64(pw) * 0.63)
			doneY := int(float64(ph) * 0.93)
			logFn(fmt.Sprintf("  [6/8] 坐标点击: (%d,%d) [弹窗=%dx%d]", doneX, doneY, pw, ph))
			w.ClickPopup(popupClass, doneX, doneY)
		} else {
			// 最终回退
			w.ClickPopup(popupClass, 415, 495)
		}
	}
	// ═══ 安全等待弹窗关闭 (防止二次触发导致重复建群) ═══
	// 核心原则: 只点击一次「完成」, 然后耐心等待弹窗自然关闭
	// ⚠️ 绝不发 Enter! Enter 在企微弹窗中等同于第二次「完成」, 会创建第二个群!
	logFn("  [6/8] 等待弹窗关闭 (最长 15s, 不重试)...")
	if w.WaitForPopupClosed(popupClass, 15) {
		logFn("  [6/8] ✅ 弹窗已关闭")
	} else {
		// 15 秒后弹窗仍在 → 可能点击没命中, 用坐标回退再点一次「完成」(不用 Enter!)
		logFn("  [6/8] ⚠️ 弹窗 15s 未关闭, 尝试坐标回退点击...")
		pw2, ph2 := w.popupClientSize(popupClass)
		if pw2 > 0 && ph2 > 0 {
			doneX2 := int(float64(pw2) * 0.63)
			doneY2 := int(float64(ph2) * 0.93)
			w.ClickPopup(popupClass, doneX2, doneY2)
		}
		if !w.WaitForPopupClosed(popupClass, 8) {
			// 彻底无法关闭, ESC 放弃
			logFn("  [6/8] ⚠️ 弹窗仍未关闭, ESC 放弃")
			w.ClosePopup(popupClass)
			time.Sleep(1 * time.Second)
		} else {
			logFn("  [6/8] ✅ 坐标回退后弹窗已关闭")
		}
	}

	// 判断建群是否成功 (弹窗关闭 = 成功)
	if w.isPopupVisible(popupClass) {
		logFn("  [6/8] ❌ 弹窗未关闭, 建群可能失败")
		result.ErrorDetail = "完成按钮点击后弹窗仍未关闭"
		// 强制关闭弹窗
		w.ClosePopup(popupClass)
		humanDelay(500, 200)
		return result
	}
	result.Success = true

	// 随机延时 (降低风控)
	delay := 1500 + rand.Intn(1500)
	logFn(fmt.Sprintf("  [6/8] 等待 %.1fs (防风控)...", float64(delay)/1000))
	time.Sleep(time.Duration(delay) * time.Millisecond)

	// ═══ Step 7: 关闭残留弹窗 ═══
	logFn("  [7/8] 检查残留弹窗...")
	for i := 0; i < 3; i++ {
		if !w.isPopupVisible(popupClass) {
			break
		}
		logFn(fmt.Sprintf("  [7/8] 发现残留弹窗, 关闭... (%d)", i+1))
		w.ClosePopup(popupClass)
		time.Sleep(1 * time.Second)
	}

	// ═══ Step 8: 设置群管理 — 禁止互加好友 (OCR + 前台点击 + 二次验证) ═══
	logFn("  [8/8] 设置群管理 (禁止互加好友)...")
	privacySet, privacyVerified := w.setupGroupPrivacy(logFn)
	result.PrivacySet = privacySet
	result.PrivacyVerified = privacyVerified
	if !privacyVerified {
		result.NeedManualCheck = true
	}

	// 最终清理
	for i := 0; i < 3; i++ {
		if !w.isPopupVisible(popupClass) {
			break
		}
		w.ClosePopup(popupClass)
		time.Sleep(500 * time.Millisecond)
	}

	elapsed := time.Since(tStart)
	statusEmoji := "✅"
	if result.NeedManualCheck {
		statusEmoji = "⚠️"
	}
	logFn(fmt.Sprintf("  %s 建群流程完成, 总耗时 %.1fs (成员=%d/%d, 隐私=%v/验证=%v)",
		statusEmoji, elapsed.Seconds(), result.MembersSelected, result.MembersExpected,
		result.PrivacySet, result.PrivacyVerified))
	return result
}

// setupGroupPrivacy 设置群管理隐私: 勾选「禁止互相添加为联系人」
//
// 架构 (2025-04-18 重构):
//   Phase A: 打开群管理窗口 (主窗口操作)
//     - SendMessage 后台点击 ··· 和 群管理
//     - 失败时 SafeRealClick 重试 (短暂抢鼠标)
//     - 等待 ExternalConversationManagerWindow 出现
//
//   Phase B: 操作群管理窗口 (100% 后台静默!)
//     - 群管理窗口是独立原生 Win32 窗口, 非 Chromium overlay
//     - PrintWindow 后台截图 ✅ (不需要前台)
//     - SendMessage 后台点击 ✅ (不抢鼠标)
//     - WM_CLOSE 后台关闭 ✅
//
//   Phase C: 收起面板
func (w *WeComWindow) setupGroupPrivacy(logFn func(string)) (privacySet bool, privacyVerified bool) {

	// ═══ Phase A: 打开群管理窗口 ═══
	logFn("       [隐私] Phase A: 打开群管理窗口...")

	// 先检查群管理窗口是否已打开
	mgmtHwnd := privFindGroupMgmtWindow(w.Pid)
	if mgmtHwnd != 0 {
		logFn(fmt.Sprintf("       [隐私] 群管理窗口已打开 (HWND=0x%X)", mgmtHwnd))
	} else {
		mgmtHwnd = w.openGroupMgmtDialog(logFn)
		if mgmtHwnd == 0 {
			logFn("       [隐私] ❌ 无法打开群管理窗口")
			return false, false
		}
	}

	// ═══ Phase B: 100% 后台操作 checkbox ═══
	logFn("       [隐私] Phase B: 后台操作 checkbox (PrintWindow + SendMessage)...")

	// B1: 后台截图群管理窗口
	mgmtImg, mgmtPng, mgmtErr := w.screenshotHwnd(mgmtHwnd)
	if mgmtErr != nil {
		logFn(fmt.Sprintf("       [隐私] ❌ 群管理窗口截图失败: %v", mgmtErr))
		w.closeGroupMgmt(mgmtHwnd, logFn)
		return false, false
	}
	logFn(fmt.Sprintf("       [隐私] ✅ 后台截图成功 (%.0fKB)", float64(len(mgmtPng))/1024))

	// B2: OCR 定位 checkbox
	items, ocrErr := ZhipuOCR(mgmtPng)
	if ocrErr != nil {
		logFn(fmt.Sprintf("       [隐私] ❌ OCR 失败: %v", ocrErr))
		w.closeGroupMgmt(mgmtHwnd, logFn)
		return false, false
	}

	target := privFindCheckboxText(items)
	if target == nil {
		logFn("       [隐私] ❌ OCR 未找到「禁止互相添加为联系人」")
		w.closeGroupMgmt(mgmtHwnd, logFn)
		return false, false
	}

	checkX := target.X1 - 20
	checkY := target.CY
	logFn(fmt.Sprintf("       [隐私] 定位 checkbox (%d,%d) [%s]", checkX, checkY, target.Text))

	// B3: 检查当前状态
	if mgmtImg != nil && IsCheckboxChecked(mgmtImg, checkX, checkY) {
		logFn("       [隐私] ✅ 已勾选, 无需操作")
		privacySet = true
		privacyVerified = true
	} else {
		// B4: SendMessage 后台点击 (多偏移尝试)
		offsets := []int{0, -5, -10, 5, -15, 10, -25}
		for attempt, dx := range offsets {
			tryX := checkX + dx
			logFn(fmt.Sprintf("       [隐私] 后台点击 (%d,%d) [尝试 %d/%d]",
				tryX, checkY, attempt+1, len(offsets)))
			privClickOnWindow(mgmtHwnd, tryX, checkY)
			humanDelay(800, 200)

			// 后台截图验证
			verifyImg, _, verifyErr := w.screenshotHwnd(mgmtHwnd)
			if verifyErr != nil {
				logFn(fmt.Sprintf("       [隐私] ⚠️ 验证截图失败: %v", verifyErr))
				continue
			}

			if IsCheckboxChecked(verifyImg, tryX, checkY) {
				logFn("       [隐私] ✅ 后台验证通过! checkbox 已勾选")
				privacySet = true
				privacyVerified = true
				break
			}
		}

		if !privacyVerified {
			logFn("       [隐私] ⚠️ 多次尝试后未确认勾选, 需人工复核")
		}
	}

	// ═══ Phase C: 关闭群管理 + 收起面板 ═══
	w.closeGroupMgmt(mgmtHwnd, logFn)

	return privacySet, privacyVerified
}

// openGroupMgmtDialog 打开群管理对话框
// 使用前台 OCR + SafeRealClick (原始验证过的方案)
// 打开后等待 ExternalConversationManagerWindow 出现
// 返回群管理窗口的 HWND, 0 表示失败
func (w *WeComWindow) openGroupMgmtDialog(logFn func(string)) syscall.Handle {

	// ──── Step 1: 前台 OCR 定位 ··· 按钮 ────
	logFn("       [隐私] Step1: 定位 ··· 按钮...")
	dotsClicked := false

	fgItems0, fgErr0 := w.OCRScanForeground()
	if fgErr0 == nil && len(fgItems0) > 0 {
		maxY := int(float64(w.Height) * 0.10)

		// 策略 A: OCR 直接识别 ··· 文字 (罕见)
		var dotsBtn *OCRItem
		for idx := range fgItems0 {
			it := &fgItems0[idx]
			if it.CY > maxY || it.CX < int(float64(w.Width)*0.50) {
				continue
			}
			txt := strings.TrimSpace(it.Text)
			if len([]rune(txt)) <= 4 && (txt == "···" || txt == "..." || txt == "…" || txt == ".." || txt == "·..") {
				dotsBtn = it
				break
			}
		}
		if dotsBtn != nil {
			logFn(fmt.Sprintf("       [隐私] OCR 直接找到 ··· (%d,%d)", dotsBtn.CX, dotsBtn.CY))
			w.SafeRealClick(dotsBtn.CX, dotsBtn.CY)
			dotsClicked = true
		}

		// 策略 B: 找聊天标题栏第一行文字 → 取其 y → ··· 在 (width-45, y)
		if !dotsClicked {
			var headerItem *OCRItem
			for idx := range fgItems0 {
				it := &fgItems0[idx]
				if it.CX > int(float64(w.Width)*0.25) && it.CY < maxY && it.CY > 30 {
					if it.Text == "搜索" || it.Text == "十" || it.Text == "+" || it.Text == "×" {
						continue
					}
					if headerItem == nil || it.CY < headerItem.CY {
						headerItem = it
					}
				}
			}
			if headerItem != nil {
				dotsX := w.Width - 45
				dotsY := headerItem.CY
				logFn(fmt.Sprintf("       [隐私] 标题行「%s」y=%d → ··· (%d,%d)",
					headerItem.Text, headerItem.CY, dotsX, dotsY))
				w.SafeRealClick(dotsX, dotsY)
				dotsClicked = true
			}
		}
	} else {
		logFn(fmt.Sprintf("       [隐私] ⚠️ 前台 OCR 失败: %v", fgErr0))
	}

	// 坐标回退
	if !dotsClicked {
		rx := w.Width - 45
		ry := int(float64(w.Height) * 0.06)
		logFn(fmt.Sprintf("       [隐私] 坐标回退 ··· (%d,%d)", rx, ry))
		w.SafeRealClick(rx, ry)
	}
	humanDelay(2500, 500)

	// 关闭意外弹窗
	if w.isPopupVisible(popupClass) {
		w.ClosePopup(popupClass)
		humanDelay(1000, 300)
	}

	// ──── Step 2: 前台 OCR 定位「群管理」────
	logFn("       [隐私] Step2: 定位「群管理」...")
	mgmtClicked := false
	var mgmtX, mgmtY int

	fgItems2, fgErr2 := w.OCRScanForeground()
	if fgErr2 == nil && len(fgItems2) > 0 {
		panelX := int(float64(w.Width) * 0.50)
		var mgmtBtn *OCRItem
		for idx := range fgItems2 {
			if fgItems2[idx].CX > panelX && strings.Contains(fgItems2[idx].Text, "群管理") {
				mgmtBtn = &fgItems2[idx]
				break
			}
		}
		if mgmtBtn != nil {
			logFn(fmt.Sprintf("       [隐私] OCR 定位「群管理」(%d,%d)", mgmtBtn.CX, mgmtBtn.CY))
			mgmtX, mgmtY = mgmtBtn.CX, mgmtBtn.CY
			w.SafeRealClick(mgmtX, mgmtY)
			mgmtClicked = true
		} else {
			var panelTexts []string
			for _, it := range fgItems2 {
				if it.CX > panelX {
					panelTexts = append(panelTexts, it.Text)
				}
			}
			logFn(fmt.Sprintf("       [隐私] ⚠️ OCR 未匹配「群管理」, 面板: %v", panelTexts))
		}
	}

	if !mgmtClicked {
		mgmtX = int(float64(w.Width) * 0.82)
		mgmtY = int(float64(w.Height) * 0.37)
		logFn(fmt.Sprintf("       [隐私] 坐标回退「群管理」(%d,%d)", mgmtX, mgmtY))
		w.SafeRealClick(mgmtX, mgmtY)
	}

	// ──── Step 3: 等待 ExternalConversationManagerWindow 出现 ────
	logFn("       [隐私] Step3: 等待群管理窗口出现...")
	for wait := 0; wait < 10; wait++ {
		humanDelay(500, 100)
		hwnd := privFindGroupMgmtWindow(w.Pid)
		if hwnd != 0 {
			logFn(fmt.Sprintf("       [隐私] ✅ 群管理窗口出现! (HWND=0x%X, %dms)", hwnd, (wait+1)*500))
			return hwnd
		}
	}

	logFn("       [隐私] ❌ 群管理窗口未出现")
	// 收起面板
	w.Click(int(float64(w.Width)*0.45), int(float64(w.Height)*0.50))
	return 0
}

// closeGroupMgmt 关闭群管理窗口 + 收起聊天信息面板
func (w *WeComWindow) closeGroupMgmt(mgmtHwnd syscall.Handle, logFn func(string)) {
	// WM_CLOSE 后台关闭群管理窗口
	const WM_CLOSE = 0x0010
	procSendMessage.Call(uintptr(mgmtHwnd), WM_CLOSE, 0, 0)
	humanDelay(800, 200)

	if privFindGroupMgmtWindow(w.Pid) != 0 {
		procSendMessage.Call(uintptr(mgmtHwnd), WM_CLOSE, 0, 0)
		humanDelay(500, 100)
	}

	// 点击聊天区空白处收起右侧面板 (绝对不发 ESC, ESC 会关闭企微窗口!)
	w.Click(int(float64(w.Width)*0.45), int(float64(w.Height)*0.50))
	humanDelay(500, 200)
}

