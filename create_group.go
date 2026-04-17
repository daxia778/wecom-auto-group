package main

import (
	"bytes"
	"fmt"
	"image/png"
	"math/rand"
	"strings"
	"time"
)

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
//   OCR 驱动建群流程 (移植自 Python app.py create_grp)
//
//   每步: OCR 优先定位 → 坐标比例回退 → 操作后验证
//   全程后台执行，不抢鼠标不抢键盘
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const popupClass = "weWorkSelectUser"

// 参考坐标 (基于 1046x705 客户区)
const (
	refW = 1046.0
	refH = 705.0
)

// CreateGroupOCR OCR 增强版建群流程 (8步)
// customer: 客户名  members: 固定成员列表  logFn: 日志回调
func (w *WeComWindow) CreateGroupOCR(customer string, members []string, logFn func(string)) bool {
	tStart := time.Now()
	allMembers := append([]string{customer}, members...)
	logFn(fmt.Sprintf("🏗️ 建群: %s (共 %d 人)", customer, len(allMembers)))

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
		return false
	}
	pw, ph := w.popupClientSize(popupClass)
	logFn(fmt.Sprintf("  [3/8] ✅ 弹窗已打开 (%dx%d)", pw, ph))

	// ═══ Step 4: 逐个搜索并选中成员 ═══
	logFn(fmt.Sprintf("  [4/8] 添加 %d 名成员", len(allMembers)))
	for i, m := range allMembers {
		logFn(fmt.Sprintf("  [4/8] (%d/%d) 搜索: %s", i+1, len(allMembers), m))

		if i > 0 {
			humanDelay(500, 300)
		}

		// 先清空搜索框
		w.ClickPopup(popupClass, 160, 40)
		humanDelay(300, 100)
		w.ClearPopupInput(popupClass)
		humanDelay(300, 100)

		// 输入搜索
		w.TypeToPopup(popupClass, m)
		humanDelay(1500, 500)

		// 勾选第一个搜索结果 (checkbox 在左侧 x≈25, 第一结果 y≈95)
		w.ClickPopup(popupClass, 25, 95)
		humanDelay(500, 200)

		// 勾选后清空搜索框
		w.ClickPopup(popupClass, 160, 40)
		humanDelay(200, 100)
		w.ClearPopupInput(popupClass)
		humanDelay(300, 100)

		logFn(fmt.Sprintf("  [4/8] (%d/%d) ✅ 已勾选 %s", i+1, len(allMembers), m))
	}

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
	time.Sleep(2 * time.Second)

	// 验证弹窗是否关闭, 如果没关则重试
	for retry := 0; retry < 3; retry++ {
		if !w.isPopupVisible(popupClass) {
			logFn("  [6/8] ✅ 弹窗已关闭")
			break
		}
		logFn(fmt.Sprintf("  [6/8] ⚠️ 弹窗仍在, 重试 %d/3...", retry+1))
		switch retry {
		case 0:
			// Enter 键确认
			w.SendKeyToPopup(popupClass, VK_RETURN)
		case 1:
			// 再点完成 (调整坐标)
			pw, ph = w.popupClientSize(popupClass)
			if pw > 0 && ph > 0 {
				w.ClickPopup(popupClass, int(float64(pw)*0.50), int(float64(ph)*0.93))
			}
		case 2:
			// ESC 关闭
			w.ClosePopup(popupClass)
		}
		time.Sleep(2 * time.Second)
	}

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

	// ═══ Step 8: 设置群管理 — 禁止互加好友 (OCR + 前台点击) ═══
	logFn("  [8/8] 设置群管理 (禁止互加好友)...")
	w.setupGroupPrivacy(logFn)

	// 最终清理
	for i := 0; i < 3; i++ {
		if !w.isPopupVisible(popupClass) {
			break
		}
		w.ClosePopup(popupClass)
		time.Sleep(500 * time.Millisecond)
	}

	elapsed := time.Since(tStart)
	logFn(fmt.Sprintf("  ✅ 建群流程完成, 总耗时 %.1fs", elapsed.Seconds()))
	return true
}

// setupGroupPrivacy 设置群管理: 禁止互加好友
//
// 实际 UI 流程 (基于截图验证):
//   点 ··· → 右侧直接显示「群管理」(无需滚动!)
//   → 点「群管理」→ 弹出 Chromium overlay 模态框
//   → 里面有 checkbox「禁止互相添加为联系人」
//   → 勾选 → 关闭 overlay → ESC 收起面板
//
// 技术要点:
//   overlay 是 Chromium CSS 渲染, PrintWindow 截不到 → 用 BitBlt
//   SendMessage 点不到 → 用 mouse_event (RealClick)
func (w *WeComWindow) setupGroupPrivacy(logFn func(string)) {

	// ═══ Step 1: 点击 ··· 打开右侧面板 ═══
	logFn("       [隐私] Step1: 点击 ··· ...")
	dotsClicked := false
	items, err := w.OCRScan()
	if err == nil {
		// ··· 在聊天区右上角 (右半部分, 上方 15%)
		dotsBtn := FindOCRTextInRegion(items, "···",
			int(float64(w.Width)*0.55), 0, w.Width, int(float64(w.Height)*0.15))
		if dotsBtn == nil {
			dotsBtn = FindOCRTextInRegion(items, "...",
				int(float64(w.Width)*0.55), 0, w.Width, int(float64(w.Height)*0.15))
		}
		if dotsBtn != nil {
			logFn(fmt.Sprintf("       [隐私] OCR ··· (%d,%d)", dotsBtn.CX, dotsBtn.CY))
			w.Click(dotsBtn.CX, dotsBtn.CY)
			dotsClicked = true
		}
	}
	if !dotsClicked {
		// 坐标回退: 距右边 30px, 高度约 5.5%
		rx := w.Width - 30
		ry := int(float64(w.Height) * 0.055)
		logFn(fmt.Sprintf("       [隐私] 坐标回退 ··· (%d,%d)", rx, ry))
		w.Click(rx, ry)
	}
	humanDelay(2000, 500)

	// 关闭意外弹窗
	if w.isPopupVisible(popupClass) {
		w.ClosePopup(popupClass)
		humanDelay(1000, 300)
	}

	// ═══ Step 2: 点击「群管理」(直接可见, 不需要滚动) ═══
	logFn("       [隐私] Step2: 点击「群管理」...")
	mgmtClicked := false

	// 重新 OCR (面板已打开, 内容变了)
	items, err = w.OCRScan()
	if err == nil {
		// 群管理在右侧面板, x > 55% 窗口宽度
		panelX := int(float64(w.Width) * 0.55)
		mgmtBtn := FindOCRTextInRegion(items, "群管理", panelX, 0, w.Width, w.Height)
		if mgmtBtn == nil {
			for idx := range items {
				if strings.Contains(items[idx].Text, "群管理") && items[idx].CX > panelX {
					mgmtBtn = &items[idx]
					break
				}
			}
		}
		if mgmtBtn != nil {
			logFn(fmt.Sprintf("       [隐私] OCR「群管理」(%d,%d)", mgmtBtn.CX, mgmtBtn.CY))
			w.Click(mgmtBtn.CX, mgmtBtn.CY)
			mgmtClicked = true
		} else {
			// 打印面板内 OCR 内容辅助调试
			var panelTexts []string
			for _, it := range items {
				if it.CX > panelX {
					panelTexts = append(panelTexts, it.Text)
				}
			}
			logFn(fmt.Sprintf("       [隐私] 面板 OCR: %v", panelTexts))
		}
	}
	if !mgmtClicked {
		// 坐标回退: 从截图看「群管理」大约在面板列表第4项
		fx := int(float64(w.Width) * 0.78)
		fy := int(float64(w.Height) * 0.45)
		logFn(fmt.Sprintf("       [隐私] 坐标回退「群管理」(%d,%d)", fx, fy))
		w.Click(fx, fy)
	}

	// ═══ Step 3: 等 overlay 加载 → 前台截图 → 找 checkbox ═══
	logFn("       [隐私] Step3: 等待群管理弹窗...")
	humanDelay(2500, 500) // Chromium overlay 需要时间渲染

	// 前台截图 (BitBlt) — overlay 是 Chromium 渲染, PrintWindow 抓不到
	fgImg, _, fgErr := w.ScreenshotForeground()
	if fgErr != nil {
		logFn(fmt.Sprintf("       [隐私] ❌ 前台截图失败: %v", fgErr))
		w.SendKey(VK_ESCAPE)
		return
	}

	var buf bytes.Buffer
	png.Encode(&buf, fgImg)
	fgItems, ocrErr := ZhipuOCR(buf.Bytes())
	if ocrErr != nil {
		logFn(fmt.Sprintf("       [隐私] ❌ OCR 失败: %v", ocrErr))
		w.SendKey(VK_ESCAPE)
		return
	}

	// 打印 OCR 结果
	var allTexts []string
	for _, it := range fgItems {
		allTexts = append(allTexts, it.Text)
	}
	logFn(fmt.Sprintf("       [隐私] 弹窗 OCR(%d项): %v", len(fgItems), allTexts))

	// 搜索「禁止互相添加为联系人」
	keywords := []string{
		"禁止互相添加为联系人", "禁止互相添加",
		"互相添加为联系人", "禁止互加", "互加好友",
		"添加为联系人",
	}

	var toggleBtn *OCRItem
	for _, kw := range keywords {
		// 精确匹配
		match := FindOCRText(fgItems, kw)
		if match != nil {
			toggleBtn = match
			break
		}
		// 子串匹配 (最长不超过 30 字)
		for idx := range fgItems {
			if strings.Contains(fgItems[idx].Text, kw) && len([]rune(fgItems[idx].Text)) <= 30 {
				toggleBtn = &fgItems[idx]
				break
			}
		}
		if toggleBtn != nil {
			break
		}
	}

	if toggleBtn == nil {
		logFn("       [隐私] ❌ 未找到「禁止互相添加为联系人」")
		// 关闭 overlay
		w.RealClick(int(float64(w.Width)*0.15), int(float64(w.Height)*0.5))
		humanDelay(500, 200)
		w.SendKey(VK_ESCAPE)
		return
	}

	// ═══ Step 4: 计算 checkbox 坐标 → 检查状态 → 点击 ═══
	txt := toggleBtn.Text
	// checkbox 在文字左侧
	checkX := toggleBtn.X1 - 15
	checkY := toggleBtn.CY

	// 如果 OCR 把 checkbox 标签连在文字里, 按比例计算真实 checkbox 位置
	txtRunes := []rune(txt)
	for _, kw := range []string{"禁止互相", "禁止互加", "互加好友"} {
		kwRunes := []rune(kw)
		for j := 0; j+len(kwRunes) <= len(txtRunes); j++ {
			if string(txtRunes[j:j+len(kwRunes)]) == kw && j > 0 {
				textWidth := toggleBtn.X2 - toggleBtn.X1
				ratio := float64(j) / float64(len(txtRunes))
				checkX = toggleBtn.X1 + int(float64(textWidth)*ratio) - 5
				break
			}
		}
	}

	// 像素采样检查是否已勾选 (蓝色 = 已勾选)
	if IsCheckboxChecked(fgImg, checkX, checkY) {
		logFn("       [隐私] ✅ 已勾选, 无需操作")
	} else {
		logFn(fmt.Sprintf("       [隐私] 点击勾选 (%d,%d) [%s]", checkX, checkY, txt))
		w.RealClick(checkX, checkY)
		humanDelay(1000, 300)
		logFn("       [隐私] ✅ 已勾选「禁止互相添加为联系人」")
	}

	// ═══ 关闭: overlay × 按钮 → ESC 收起面板 ═══
	closed := false
	// 从截图看 × 按钮在 overlay 右上角
	closeBtn := FindOCRText(fgItems, "×")
	if closeBtn == nil {
		closeBtn = FindOCRText(fgItems, "x")
	}
	if closeBtn != nil {
		logFn(fmt.Sprintf("       [隐私] 关闭 × (%d,%d)", closeBtn.CX, closeBtn.CY))
		w.RealClick(closeBtn.CX, closeBtn.CY)
		closed = true
	}
	if !closed {
		// 点击 overlay 外部区域
		w.RealClick(int(float64(w.Width)*0.15), int(float64(w.Height)*0.5))
	}
	humanDelay(800, 200)

	// ESC 收起右侧面板
	w.SendKey(VK_ESCAPE)
	humanDelay(500, 200)
}



