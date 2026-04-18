package main

import (
	"bufio"
	"fmt"
	"os"
	"strconv"
	"strings"
	"time"
	"unsafe"
)

// InteractivePrivacyTest 交互式隐私设置测试
// 每一步截图 → 等待指令 → 执行 → 再截图
// 指令格式:
//   click X Y       - 前台点击 (SafeRealClick)
//   bgclick X Y     - 后台点击 (SendMessage)
//   scroll X Y D    - 滚轮 (SafeRealScroll, D 负=向下)
//   screenshot      - 仅截图不操作
//   sleep MS        - 等待毫秒
//   done            - 结束
func InteractivePrivacyTest() {
	fmt.Println("╔═══════════════════════════════════╗")
	fmt.Println("║  交互式隐私设置测试 v1            ║")
	fmt.Println("║  截图 → 等待指令 → 执行 → 重复   ║")
	fmt.Println("╚═══════════════════════════════════╝")
	fmt.Println()

	// 倒计时
	for i := 3; i > 0; i-- {
		fmt.Printf("\r⏳ %d 秒后开始...", i)
		time.Sleep(time.Second)
	}
	fmt.Println("\r🚀 开始!           ")

	// 查找窗口
	wc, err := FindWeComWindow()
	if err != nil || wc == nil {
		fmt.Printf("❌ 未找到企微窗口: %v\n", err)
		return
	}
	fmt.Printf("✅ 窗口: %d×%d (HWND=0x%X)\n\n", wc.Width, wc.Height, wc.Hwnd)

	// 确保 debug 目录
	debugDir := "debug_interactive"
	os.MkdirAll(debugDir, 0755)

	step := 0
	scanner := bufio.NewScanner(os.Stdin)

	// 初始截图
	step++
	takeScreenshot(wc, debugDir, step)

	for {
		fmt.Printf("\n[Step %d] 输入指令 (click/bgclick/scroll/screenshot/sleep/done): ", step)
		if !scanner.Scan() {
			break
		}
		line := strings.TrimSpace(scanner.Text())
		if line == "" {
			continue
		}

		parts := strings.Fields(line)
		cmd := strings.ToLower(parts[0])

		switch cmd {
		case "done", "quit", "exit":
			fmt.Println("✅ 测试结束")
			return

		case "click":
			if len(parts) < 3 {
				fmt.Println("用法: click X Y")
				continue
			}
			x, _ := strconv.Atoi(parts[1])
			y, _ := strconv.Atoi(parts[2])
			fmt.Printf("  → SafeRealClick(%d, %d)...\n", x, y)
			wc.SafeRealClick(x, y)
			time.Sleep(1500 * time.Millisecond)
			step++
			takeScreenshot(wc, debugDir, step)

		case "bgclick":
			if len(parts) < 3 {
				fmt.Println("用法: bgclick X Y")
				continue
			}
			x, _ := strconv.Atoi(parts[1])
			y, _ := strconv.Atoi(parts[2])
			fmt.Printf("  → SendMessage Click(%d, %d)...\n", x, y)
			wc.Click(x, y)
			time.Sleep(1500 * time.Millisecond)
			step++
			takeScreenshot(wc, debugDir, step)

		case "scroll":
			if len(parts) < 4 {
				fmt.Println("用法: scroll X Y delta (负=下)")
				continue
			}
			x, _ := strconv.Atoi(parts[1])
			y, _ := strconv.Atoi(parts[2])
			d, _ := strconv.Atoi(parts[3])
			fmt.Printf("  → SafeRealScroll(%d, %d, %d)...\n", x, y, d)
			wc.SafeRealScroll(x, y, d)
			time.Sleep(800 * time.Millisecond)
			step++
			takeScreenshot(wc, debugDir, step)

		case "screenshot", "ss":
			step++
			takeScreenshot(wc, debugDir, step)

		case "sleep":
			ms := 1000
			if len(parts) >= 2 {
				ms, _ = strconv.Atoi(parts[1])
			}
			fmt.Printf("  → 等待 %dms...\n", ms)
			time.Sleep(time.Duration(ms) * time.Millisecond)

		case "resize":
			if len(parts) < 3 {
				fmt.Println("用法: resize W H")
				continue
			}
			w, _ := strconv.Atoi(parts[1])
			h, _ := strconv.Atoi(parts[2])
			fmt.Printf("  → 调整窗口大小到 %d×%d...\n", w, h)
			procSetWindowPos.Call(uintptr(wc.Hwnd), 0,
				0, 0, uintptr(w), uintptr(h),
				SWP_NOMOVE|SWP_NOACTIVATE)
			time.Sleep(500 * time.Millisecond)
			// 更新尺寸
			var cr RECT
			procGetClientRect.Call(uintptr(wc.Hwnd), uintptr(unsafe.Pointer(&cr)))
			wc.Width = int(cr.Right)
			wc.Height = int(cr.Bottom)
			fmt.Printf("  新客户区: %d×%d\n", wc.Width, wc.Height)
			step++
			takeScreenshot(wc, debugDir, step)

		case "pyclick":
			// 完全模仿 Python 的 real_click: SetCursorPos + mouse_event
			if len(parts) < 3 {
				fmt.Println("用法: pyclick X Y")
				continue
			}
			x, _ := strconv.Atoi(parts[1])
			y, _ := strconv.Atoi(parts[2])
			fmt.Printf("  → Python风格 real_click(%d, %d)...\n", x, y)
			pyRealClick(wc, x, y)
			time.Sleep(1500 * time.Millisecond)
			step++
			takeScreenshot(wc, debugDir, step)

		case "sweep":
			// 系统性扫描 y 坐标, 查找群管理的精确位置
			if len(parts) < 5 {
				fmt.Println("用法: sweep X yStart yEnd yStep")
				continue
			}
			sx, _ := strconv.Atoi(parts[1])
			yStart, _ := strconv.Atoi(parts[2])
			yEnd, _ := strconv.Atoi(parts[3])
			yStep, _ := strconv.Atoi(parts[4])
			if yStep <= 0 {
				yStep = 20
			}
			fmt.Printf("  → 扫描 x=%d, y=%d..%d (步长 %d)\n", sx, yStart, yEnd, yStep)
			for sy := yStart; sy <= yEnd; sy += yStep {
				fmt.Printf("    → bgclick(%d, %d)...\n", sx, sy)
				wc.Click(sx, sy)
				time.Sleep(800 * time.Millisecond)
				step++
				takeScreenshot(wc, debugDir, step)
			}

		default:
			fmt.Printf("  ❓ 未知指令: %s\n", cmd)
			fmt.Println("  可用: click/bgclick/pyclick/scroll/screenshot/sleep/resize/sweep/done")
		}
	}
}

func takeScreenshot(wc *WeComWindow, debugDir string, step int) {
	filename := fmt.Sprintf("%s/step_%02d.png", debugDir, step)
	_, fgPng, err := wc.SafeScreenshotForeground()
	if err != nil || len(fgPng) <= 10000 {
		// 回退到后台截图
		_, fgPng, err = wc.Screenshot()
	}
	if err != nil || len(fgPng) <= 10000 {
		fmt.Printf("  ❌ 截图失败: %v\n", err)
		return
	}
	os.WriteFile(filename, fgPng, 0644)
	fmt.Printf("  📷 已保存: %s (%d bytes)\n", filename, len(fgPng))
}

// pyRealClick 完全模仿 Python 的 real_click 方法:
// 1. ClientToScreen 坐标转换
// 2. _force_foreground 强制前台
// 3. SetCursorPos 移动鼠标
// 4. mouse_event(LEFTDOWN) + mouse_event(LEFTUP)  ← 不用 SendInput!
// 关键区别: Python 用 mouse_event 不带 ABSOLUTE 标志, 在当前光标位置点击
func pyRealClick(wc *WeComWindow, cx, cy int) {
	foregroundMu.Lock()
	defer foregroundMu.Unlock()

	// 保存当前鼠标位置
	var savedPos POINT
	procGetCursorPos.Call(uintptr(unsafe.Pointer(&savedPos)))

	// 客户区坐标 → 屏幕坐标
	var pt POINT
	pt.X = int32(cx)
	pt.Y = int32(cy)
	procClientToScreen.Call(uintptr(wc.Hwnd), uintptr(unsafe.Pointer(&pt)))
	fmt.Printf("    屏幕坐标: (%d, %d)\n", pt.X, pt.Y)

	// 强制前台 (和 Python 的 _force_foreground 完全一致)
	wc.forceToForeground(wc.Hwnd)
	time.Sleep(150 * time.Millisecond)

	// SetCursorPos 移动鼠标到目标位置
	procSetCursorPos.Call(uintptr(pt.X), uintptr(pt.Y))
	time.Sleep(150 * time.Millisecond)

	// mouse_event LEFTDOWN (不用 ABSOLUTE, 在当前光标位置点击)
	procMouseEvent.Call(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
	time.Sleep(80 * time.Millisecond)

	// mouse_event LEFTUP
	procMouseEvent.Call(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
	time.Sleep(500 * time.Millisecond)

	// 取消 TOPMOST
	procSetWindowPos.Call(uintptr(wc.Hwnd), ^uintptr(1), 0, 0, 0, 0,
		SWP_NOMOVE|SWP_NOSIZE)

	// 恢复鼠标位置
	procSetCursorPos.Call(uintptr(savedPos.X), uintptr(savedPos.Y))
}

