package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"
)

// TestWindowOCR 诊断企微窗口 + OCR 识别
// 用法: WeComAutoGroup.exe --diag
func TestWindowOCR() {
	fmt.Println("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
	fmt.Println("  企微窗口 + OCR 诊断 (自动化前置检查)")
	fmt.Println("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

	// ═══ Step 1: 查找企微窗口 ═══
	fmt.Println("\n[1/4] 查找企微窗口...")
	wc, err := FindWeComWindow()
	if err != nil {
		fmt.Printf("  ❌ %s\n", err)
		fmt.Println("\n  💡 请确认:")
		fmt.Println("     1. 企业微信已打开")
		fmt.Println("     2. 主窗口未最小化到托盘")
		return
	}
	fmt.Printf("  ✅ 找到企微窗口! HWND=0x%X PID=%d 尺寸=%dx%d\n",
		wc.Hwnd, wc.Pid, wc.Width, wc.Height)

	// ═══ Step 2: 后台截图 ═══
	fmt.Println("\n[2/4] 后台截图...")
	wc.SinkToBottom() // 压到底层，不影响用户操作
	img, pngData, err := wc.Screenshot()
	if err != nil {
		fmt.Printf("  ❌ 截图失败: %s\n", err)
		return
	}
	bounds := img.Bounds()
	fmt.Printf("  ✅ 截图成功! 尺寸=%dx%d 文件大小=%.1fKB\n",
		bounds.Dx(), bounds.Dy(), float64(len(pngData))/1024)

	// 保存截图到本地
	exe, _ := os.Executable()
	screenshotPath := filepath.Join(filepath.Dir(exe), "diag_screenshot.png")
	if err := os.WriteFile(screenshotPath, pngData, 0644); err == nil {
		fmt.Printf("  📸 截图已保存: %s\n", screenshotPath)
	}

	// ═══ Step 3: OCR 识别 ═══
	fmt.Println("\n[3/4] OCR 识别 (智谱 AI)...")
	startTime := time.Now()
	items, err := ZhipuOCR(pngData)
	elapsed := time.Since(startTime)
	if err != nil {
		fmt.Printf("  ❌ OCR 失败: %s\n", err)
		fmt.Println("\n  💡 检查:")
		fmt.Println("     1. ZhipuAPIKey 是否有效")
		fmt.Println("     2. 网络是否可达 open.bigmodel.cn")
		return
	}
	fmt.Printf("  ✅ OCR 识别完成! 耗时=%.1fs 识别到 %d 个文字块\n",
		elapsed.Seconds(), len(items))

	// 打印所有识别到的文字
	fmt.Println("\n  ┌───┬──────────────────────────────┬─────────────┬──────┐")
	fmt.Println("  │ # │ Text                         │ Position    │ Conf │")
	fmt.Println("  ├───┼──────────────────────────────┼─────────────┼──────┤")
	for i, item := range items {
		text := item.Text
		textRunes := []rune(text)
		if len(textRunes) > 26 {
			text = string(textRunes[:26]) + "..."
		}
		fmt.Printf("  │%2d │ %-28s │ (%4d,%4d) │ %.2f │\n",
			i+1, text, item.CX, item.CY, item.Conf)
	}
	fmt.Println("  └───┴──────────────────────────────┴─────────────┴──────┘")

	// ═══ Step 4: 关键元素检测 ═══
	fmt.Println("\n[4/4] 检测企微界面关键元素...")

	keywords := []string{
		"消息", "通讯录", "工作台", "日程",
		"搜索", "发起群聊", "建群",
		"+", "确定", "取消",
	}

	found := 0
	for _, kw := range keywords {
		match := FindOCRText(items, kw)
		if match != nil {
			fmt.Printf("  ✅ [%s] → 坐标(%d, %d) 文字=\"%s\"\n",
				kw, match.CX, match.CY, match.Text)
			found++
		}
	}
	if found == 0 {
		fmt.Println("  ⚠️ 未找到企微标准界面元素，可能不在消息页面")
	}

	// 输出所有可点击的文字元素
	fmt.Println("\n  📋 所有可交互文字:")
	for _, item := range items {
		if len(strings.TrimSpace(item.Text)) > 0 {
			fmt.Printf("    「%s」 → Click(%d, %d)\n", item.Text, item.CX, item.CY)
		}
	}

	fmt.Println("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
	fmt.Printf("  ✅ 诊断完成! 窗口=%dx%d OCR=%d项 关键元素=%d个\n",
		wc.Width, wc.Height, len(items), found)
	fmt.Println("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
}
