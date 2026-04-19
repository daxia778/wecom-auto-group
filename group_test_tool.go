package main

import (
	"fmt"
	"time"
)

// TestGroupCreation 完整建群+隐私设置测试
// 使用测试账号「暴走的手枪腿」建群, 然后自动设置群隐私
func TestGroupCreation() {
	fmt.Println("╔═══════════════════════════════════════════════╗")
	fmt.Println("║    完整建群 + 隐私设置测试                      ║")
	fmt.Println("║    测试账号: 暴走的手枪腿                        ║")
	fmt.Println("╚═══════════════════════════════════════════════╝")
	fmt.Println()

	for i := 3; i > 0; i-- {
		fmt.Printf("\r⏳ %d 秒后开始 (请确保企微窗口可见)...", i)
		time.Sleep(time.Second)
	}
	fmt.Println("\r🚀 开始!                                        ")
	fmt.Println()

	logFn := func(s string) {
		fmt.Printf("[%s] %s\n", time.Now().Format("15:04:05"), s)
	}

	wc, err := FindWeComWindow()
	if err != nil {
		logFn(fmt.Sprintf("❌ %v", err))
		return
	}
	logFn(fmt.Sprintf("✅ 窗口: %d×%d (HWND=0x%X)", wc.Width, wc.Height, wc.Hwnd))

	// 建群: 客户=暴走的手枪腿, 固定成员=吴天宇(夏虫), 吴泽华
	customer := "暴走的手枪腿"
	members := []string{"吴天宇(夏虫)", "吴泽华"}

	logFn(fmt.Sprintf("📋 客户: %s", customer))
	logFn(fmt.Sprintf("📋 固定成员: %v", members))
	logFn("")

	result := wc.CreateGroupOCR(customer, members, logFn)

	logFn("")
	logFn("══════════════════════════════════════")
	logFn(fmt.Sprintf("建群结果:"))
	logFn(fmt.Sprintf("  成功: %v", result.Success))
	logFn(fmt.Sprintf("  成员: %d/%d", result.MembersSelected, result.MembersExpected))
	logFn(fmt.Sprintf("  隐私设置: %v", result.PrivacySet))
	logFn(fmt.Sprintf("  隐私验证: %v", result.PrivacyVerified))
	if result.ErrorDetail != "" {
		logFn(fmt.Sprintf("  错误: %s", result.ErrorDetail))
	}
	if result.NeedManualCheck {
		logFn("  ⚠️ 需要人工复核")
	}
}
