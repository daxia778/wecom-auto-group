package main

// ━━━ 内置配置 (编译进 EXE) ━━━
const (
	WeComCorpID     = "wwdb2f088115fa0fff"
	WeComCorpSecret = "fdaIml1ODRNKyZFPNF04kZz7zn0Mfv8yqW78Fr7zYh0"
	ZhipuAPIKey     = "83af4cc457f74cbca091fa6ff6c604e1.uayvCfEilIDKagSf"
	ZhipuOCRURL     = "https://open.bigmodel.cn/api/paas/v4/files/ocr"
	WeComAPIBase    = "https://qyapi.weixin.qq.com/cgi-bin"

	// ━━━ 服务器 API 中转 (绕过企微 IP 白名单) ━━━
	// 桌面端 → 服务器(118.31.56.141) → 企微API
	ServerAPIBase     = "https://zhiyuanshijue.ltd/api/v1"
	ServerAdminUser   = "admin"           // 管理员用户名
	ServerAdminPass   = "admin888"        // 管理员密码
)
