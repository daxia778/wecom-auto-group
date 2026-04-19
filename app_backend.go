package main

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"sync"
	"time"

	wailsRuntime "github.com/wailsapp/wails/v2/pkg/runtime"
)

// APIClient 统一 API 接口 (支持直连和服务器中转两种模式)
type APIClient interface {
	GetMembers() ([]Member, error)
	GetContacts(userid string) ([]Contact, error)
	GetGroups() ([]GroupChat, error)
	GetFollowUserList() ([]string, error)
}

// App 主应用逻辑 (绑定到前端)
type App struct {
	ctx           context.Context
	api           APIClient  // 统一接口: ServerAPI 或 WeComAPI
	serverAPI     *ServerAPI // 服务器 API 客户端
	running       bool
	stopCh        chan struct{}
	logs          []string
	logMu         sync.Mutex
	logFile       *os.File   // 持久化日志文件
	stateFile     string
	state         AppState
	cachedMembers []Member      // 启动时预加载的员工列表
	startupDone   chan struct{} // 关闭后表示 startup 已完成
}

// AppState 本地持久化状态
type AppState struct {
	// ProcessedCustomers: uid → unix timestamp (改为 map, O(1) 查询)
	// 兼容旧版: 如果 JSON 存的是 []string, loadState 会自动迁移
	ProcessedCustomers map[string]int64 `json:"processed_customers"`
	TargetUserID       string           `json:"target_userid"`
	FixedMembers       []string         `json:"fixed_members"`
	GroupOwner         string           `json:"group_owner"`
	NeedReviewList     []string         `json:"need_review_list,omitempty"`     // 需人工复核的 uid
	TestCustomerNames  []string         `json:"test_customer_names,omitempty"` // 测试账号: 跳过防重, 每次都建群
	RootMode           bool             `json:"root_mode,omitempty"`           // Root模式: 所有客户跳过防重, 无限建群
	AutoCutoffTime     int64            `json:"auto_cutoff_time,omitempty"`    // 全自动模式起始时间 (unix), 仅处理此时间后添加的客户
}

func NewApp() *App {
	// 状态文件保存到 %APPDATA%/WeComAutoGroup/ (不受 wails build -clean 影响)
	appDataDir := filepath.Join(os.Getenv("APPDATA"), "WeComAutoGroup")
	os.MkdirAll(appDataDir, 0755)
	stateFile := filepath.Join(appDataDir, "local_state.json")

	// 迁移: 如果旧位置 (EXE目录) 有状态文件，复制到新位置
	exe, _ := os.Executable()
	exeDir := filepath.Dir(exe)
	oldStateFile := filepath.Join(exeDir, "local_state.json")
	if _, err := os.Stat(stateFile); os.IsNotExist(err) {
		if oldData, err := os.ReadFile(oldStateFile); err == nil {
			os.WriteFile(stateFile, oldData, 0644)
		}
	}

	// 使用服务器 API 中转 (绕过企微 IP 白名单)
	sapi := NewServerAPI()

	app := &App{
		api:         sapi,
		serverAPI:   sapi,
		stateFile:   stateFile,
		state:       AppState{ProcessedCustomers: make(map[string]int64)},
		startupDone: make(chan struct{}),
	}
	app.loadState()
	app.initLogFile(appDataDir)
	return app
}

// initLogFile 初始化持久化日志文件
func (a *App) initLogFile(dir string) {
	logPath := filepath.Join(dir, "wecom_autogroup.log")
	f, err := os.OpenFile(logPath, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
	if err == nil {
		a.logFile = f
		// 写入启动分隔线
		f.WriteString(fmt.Sprintf("\n━━━ 启动 %s ━━━\n", time.Now().Format("2006-01-02 15:04:05")))
	}
}

func (a *App) startup(ctx context.Context) {
	a.ctx = ctx

	// 启动时自动登录服务器 (同步阻塞, 确保 frontend 加载前完成)
	a.addLog("🔐 正在登录服务器...")
	if err := a.serverAPI.Login(ServerAdminUser, ServerAdminPass); err != nil {
		a.addLog(fmt.Sprintf("⚠️ 服务器登录失败: %s", err))
		a.addLog("  将使用直连模式 (需要 IP 白名单)")
		// 回退到直连模式
		a.api = NewWeComAPI()
	} else {
		a.addLog("✅ 服务器登录成功，使用中转模式")
	}

	// 预加载员工列表 (确保前端打开时下拉框立即有数据)
	a.addLog("📡 获取员工列表...")
	members, err := a.api.GetMembers()
	if err != nil {
		a.addLog(fmt.Sprintf("⚠️ 员工列表获取失败: %s", err))
		a.cachedMembers = []Member{}
	} else {
		// 按 UserID (拼音) A-Z 排序
		sort.Slice(members, func(i, j int) bool {
			return strings.ToLower(members[i].UserID) < strings.ToLower(members[j].UserID)
		})
		a.cachedMembers = members
		a.addLog(fmt.Sprintf("✅ 获取到 %d 名员工 (已按拼音排序)", len(members)))
	}

	a.addLog("✅ 应用启动成功")

	// 自动检测 Root 模式: 如果群主是 root 操作员, 自动开启
	if a.state.GroupOwner != "" && IsRootOperator(a.state.GroupOwner) {
		if !a.state.RootMode {
			a.state.RootMode = true
			a.saveState()
		}
		a.addLog("🔓 Root 操作员【" + a.state.GroupOwner + "】: 所有客户可无限建群")
	}
	if a.state.RootMode {
		a.addLog("📋 当前模式: Root (跳过所有防重检查)")
	}

	// 通知所有等待的前端 RPC 调用: startup 已完成
	close(a.startupDone)
}

// waitStartup 阻塞直到 startup() 完成
func (a *App) waitStartup() {
	<-a.startupDone
}

// ━━━ 前端可调用的方法 ━━━

// GetMembers 获取员工列表 (等待 startup 完成后返回缓存数据)
func (a *App) GetMembers() []Member {
	a.waitStartup()
	if len(a.cachedMembers) > 0 {
		a.addLog(fmt.Sprintf("📋 返回缓存员工 %d 人", len(a.cachedMembers)))
		return a.cachedMembers
	}
	// 缓存为空, 重新查询
	a.addLog("📡 缓存为空, 重新获取员工列表...")
	members, err := a.api.GetMembers()
	if err != nil {
		a.addLog(fmt.Sprintf("❌ 获取员工失败: %s", err))
		return []Member{}
	}
	// 按拼音排序
	sort.Slice(members, func(i, j int) bool {
		return strings.ToLower(members[i].UserID) < strings.ToLower(members[j].UserID)
	})
	a.cachedMembers = members
	a.addLog(fmt.Sprintf("✅ 获取到 %d 名员工", len(members)))
	return members
}

// GetContacts 获取某员工的外部联系人
func (a *App) GetContacts(userid string) []Contact {
	a.addLog(fmt.Sprintf("📡 获取 %s 的外部联系人...", userid))
	contacts, err := a.api.GetContacts(userid)
	if err != nil {
		a.addLog(fmt.Sprintf("❌ 获取联系人失败: %s", err))
		return []Contact{}
	}

	// 标记已处理状态
	// Root模式: 不加任何标记, 所有客户都正常显示
	// 非Root: 测试账号显示🧪, 已处理显示✅
	for i := range contacts {
		if !a.state.RootMode {
			if a.isTestAccount(contacts[i].Name) {
				contacts[i].Name = contacts[i].Name + " 🧪测试"
			} else if a.isProcessed(contacts[i].ExternalUserID) {
				contacts[i].Name = contacts[i].Name + " ✅"
			}
		}
	}

	a.addLog(fmt.Sprintf("✅ 获取到 %d 个外部联系人", len(contacts)))
	return contacts
}

// StartLoadContacts 异步流式加载联系人 (通过 Wails 事件推送给前端)
func (a *App) StartLoadContacts(userid string) {
	go func() {
		// 发送开始事件
		wailsRuntime.EventsEmit(a.ctx, "contacts:loading", map[string]interface{}{
			"userid": userid,
			"status": "loading",
		})
		a.addLog(fmt.Sprintf("📡 开始加载 %s 的外部联系人...", userid))

		contacts, err := a.api.GetContacts(userid)
		if err != nil {
			a.addLog(fmt.Sprintf("❌ 获取联系人失败: %s", err))
			wailsRuntime.EventsEmit(a.ctx, "contacts:error", map[string]interface{}{
				"error": err.Error(),
			})
			return
		}

		// 标记已处理状态 (Root模式不加标记)
		for i := range contacts {
			if !a.state.RootMode {
				if a.isTestAccount(contacts[i].Name) {
					contacts[i].Name = contacts[i].Name + " 🧪测试"
				} else if a.isProcessed(contacts[i].ExternalUserID) {
					contacts[i].Name = contacts[i].Name + " ✅"
				}
			}
		}

		// ═══ 批量查询外部群状态 (注入 GroupCount 到每个联系人) ═══
		if a.serverAPI != nil && a.serverAPI.token != "" {
			a.addLog("   📡 批量查询客户群状态...")
			var allIDs []string
			for _, c := range contacts {
				allIDs = append(allIDs, c.ExternalUserID)
			}
			if len(allIDs) > 0 {
				// 分批查询 (每批 50 个, 避免 URL 过长)
				checkBatch := 50
				groupMap := make(map[string]GroupCheckResult)
				for i := 0; i < len(allIDs); i += checkBatch {
					end := i + checkBatch
					if end > len(allIDs) {
						end = len(allIDs)
					}
					batch := allIDs[i:end]
					results, batchErr := a.serverAPI.CheckCustomerInGroups(batch)
					if batchErr != nil {
						a.addLog(fmt.Sprintf("   ⚠️ 群状态查询失败: %v", batchErr))
						break
					}
					for k, v := range results {
						groupMap[k] = v
					}
				}
				// 注入到联系人
				inGroupCount := 0
				multiGroupCount := 0
				for i := range contacts {
					if r, ok := groupMap[contacts[i].ExternalUserID]; ok {
						contacts[i].GroupCount = r.GroupCount
						if r.GroupCount >= 1 {
							inGroupCount++
						}
						if r.GroupCount > 1 {
							multiGroupCount++
						}
					}
				}
				a.addLog(fmt.Sprintf("   ✅ 群状态: %d 人已在群, %d 人多群, %d 人未建群",
					inGroupCount, multiGroupCount, len(contacts)-inGroupCount))
			}
		}

		total := len(contacts)
		a.addLog(fmt.Sprintf("✅ 获取到 %d 个外部联系人，正在推送...", total))

		// 分批推送 (每批 20 个, 模拟流式输出)
		batchSize := 20
		for i := 0; i < total; i += batchSize {
			end := i + batchSize
			if end > total {
				end = total
			}
			batch := contacts[i:end]

			wailsRuntime.EventsEmit(a.ctx, "contacts:batch", map[string]interface{}{
				"contacts": batch,
				"loaded":   end,
				"total":    total,
			})

			// 小延迟让前端有时间渲染
			if end < total {
				time.Sleep(80 * time.Millisecond)
			}
		}

		// 完成事件
		wailsRuntime.EventsEmit(a.ctx, "contacts:done", map[string]interface{}{
			"total": total,
		})
		a.addLog(fmt.Sprintf("✅ 联系人加载完成 (%d 人)", total))
	}()
}


// GetGroups 获取群聊列表
func (a *App) GetGroups() []GroupChat {
	a.addLog("📡 获取群聊列表...")
	groups, err := a.api.GetGroups()
	if err != nil {
		a.addLog(fmt.Sprintf("❌ 获取群聊失败: %s", err))
		return []GroupChat{}
	}
	a.addLog(fmt.Sprintf("✅ 获取到 %d 个群聊", len(groups)))
	return groups
}

// GetFollowUserList 获取有客户联系权限的员工
func (a *App) GetFollowUserList() []string {
	a.addLog("📡 获取客户联系权限员工...")
	users, err := a.api.GetFollowUserList()
	if err != nil {
		a.addLog(fmt.Sprintf("❌ 获取客户联系员工失败: %s", err))
		return []string{}
	}
	a.addLog(fmt.Sprintf("✅ %d 名员工有客户联系权限", len(users)))
	return users
}

// SaveSettings 保存建群设置
func (a *App) SaveSettings(targetUID string, members []string, groupOwner string) {
	a.state.TargetUserID = targetUID
	a.state.FixedMembers = members
	a.state.GroupOwner = groupOwner
	a.saveState()
	a.addLog(fmt.Sprintf("💾 配置已保存: 主理人=%s, 成员=%s", targetUID, strings.Join(members, "、")))
}

// SaveTestAccounts 设置测试账号 (测试账号跳过防重检查, 每次巡检都会建群)
func (a *App) SaveTestAccounts(names []string) {
	a.state.TestCustomerNames = names
	a.saveState()
	a.addLog(fmt.Sprintf("🧪 测试账号已设置: %v", names))
}

// GetTestAccounts 获取当前测试账号列表
func (a *App) GetTestAccounts() []string {
	return a.state.TestCustomerNames
}

// GetSettings 获取当前设置
func (a *App) GetSettings() AppState {
	return a.state
}

// CreateGroupForCustomer 为指定客户建群 (GUI 自动化)
// 注: 带外部客户的群聊只能通过 GUI 创建, 企微 API 不支持外部群
// 返回 GroupResult (包含详细的成功/失败/隐私设置状态)
func (a *App) CreateGroupForCustomer(customerName string, customerUID string) GroupResult {
	// 清除前端显示标签 (GetContacts 会追加 ✅ 或 🧪测试, 搜索时不能带)
	customerName = strings.TrimSpace(customerName)
	customerName = strings.TrimSuffix(customerName, " ✅")
	customerName = strings.TrimSuffix(customerName, " 🧪测试")
	customerName = strings.TrimSuffix(customerName, "✅")
	customerName = strings.TrimSuffix(customerName, "🧪测试")
	customerName = strings.TrimSpace(customerName)

	wc, err := FindWeComWindow()
	if err != nil {
		result := GroupResult{ErrorDetail: err.Error()}
		a.addLog(fmt.Sprintf("❌ %s", err))
		return result
	}

	a.addLog(fmt.Sprintf("🏗️ 开始建群: 客户=%s, 窗口=%dx%d", customerName, wc.Width, wc.Height))

	// ═══ 关键: 将 FixedMembers (UserID) 转为中文名 ═══
	// WeCom 弹窗搜索只认中文名, 不能用 UserID (如 "WuZeHua")
	memberNames := a.resolveToNames(a.state.FixedMembers)
	a.addLog(fmt.Sprintf("   📋 成员名称: %v", memberNames))

	result := wc.CreateGroupOCR(customerName, memberNames, a.addLog)

	// 构建操作报告
	report := NewReport(customerName, customerUID, "create_group")
	report.Success = result.Success
	report.PrivacyOK = result.PrivacyVerified
	report.MembersOK = result.MembersSelected == result.MembersExpected
	report.ErrorDetail = result.ErrorDetail
	report.NeedReview = result.NeedManualCheck

	if result.Success {
		isTest := a.isTestAccount(customerName)
		if result.PrivacyVerified {
			// 完全成功: 建群 OK + 隐私设置已验证
			if !isTest {
				a.markProcessed(customerUID) // 测试账号不标记, 下次还能建
			}
			tag := ""
			if isTest {
				tag = " (测试账号)"
			}
			msg := fmt.Sprintf("✅ 【%s】建群成功！(隐私设置已验证)%s", customerName, tag)
			a.addLog(msg)
		} else {
			// 部分成功: 建群 OK 但隐私设置未能验证
			if !isTest {
				a.markProcessed(customerUID)
			}
			a.markNeedReview(customerUID)
			msg := fmt.Sprintf("⚠️ 【%s】建群成功, 但隐私设置需人工复核!", customerName)
			a.addLog(msg)
			a.reportAlert("warning", fmt.Sprintf("客户【%s】建群成功但「禁止互加好友」未能验证, 请人工检查", customerName))
		}
		// 触发服务端缓存刷新 (异步, 不阻塞)
		if a.serverAPI != nil && a.serverAPI.token != "" {
			go func() {
				if err := a.serverAPI.SyncCustomerGroups(); err != nil {
					a.addLog(fmt.Sprintf("   ⚠️ 缓存刷新失败: %v", err))
				} else {
					a.addLog("   📡 服务端客户群缓存已刷新")
				}
			}()
		}
	} else {
		msg := fmt.Sprintf("❌ 【%s】建群失败: %s", customerName, result.ErrorDetail)
		a.addLog(msg)
	}

	// 异步上报操作日志
	a.reportOperation(report)

	return result
}

// StartAutoAgent 启动全自动巡检
func (a *App) StartAutoAgent() string {
	if a.running {
		return "⚠️ 巡检已在运行中"
	}
	if a.state.TargetUserID == "" {
		return "❌ 请先选择目标主理人"
	}

	a.running = true
	a.stopCh = make(chan struct{})
	go a.agentLoop()
	a.addLog("🚀 全自动巡检已启动")
	return "🚀 已启动"
}

// StopAutoAgent 停止巡检
func (a *App) StopAutoAgent() string {
	if !a.running {
		return "⚠️ 巡检未在运行"
	}
	a.running = false
	close(a.stopCh)
	a.addLog("🛑 全自动巡检已停止")
	return "🛑 已停止"
}

// IsAgentRunning 检查巡检状态
func (a *App) IsAgentRunning() bool {
	return a.running
}

// GetLogs 获取最新日志
func (a *App) GetLogs() []string {
	a.logMu.Lock()
	defer a.logMu.Unlock()
	// 返回最后 100 条
	if len(a.logs) > 100 {
		return a.logs[len(a.logs)-100:]
	}
	return a.logs
}

// TestConnection 测试 API 连接 (使用 startup 预加载的缓存数据)
func (a *App) TestConnection() string {
	a.waitStartup()
	count := len(a.cachedMembers)
	if count > 0 {
		msg := fmt.Sprintf("✅ 连接成功! 获取到 %d 名员工", count)
		a.addLog(msg)
		return msg
	}
	// 缓存为空则认为连接有问题
	a.addLog("⚠️ 员工列表为空, 尝试重新获取...")
	members, err := a.api.GetMembers()
	if err != nil {
		msg := fmt.Sprintf("❌ 连接失败: %s", err)
		a.addLog(msg)
		return msg
	}
	a.cachedMembers = members
	msg := fmt.Sprintf("✅ 连接成功! 获取到 %d 名员工", len(members))
	a.addLog(msg)
	return msg
}

// ━━━ 内部方法 ━━━

func (a *App) addLog(msg string) {
	a.logMu.Lock()
	defer a.logMu.Unlock()
	ts := time.Now().Format("15:04:05")
	line := fmt.Sprintf("[%s] %s", ts, msg)
	a.logs = append(a.logs, line)

	// 持久化到日志文件
	if a.logFile != nil {
		a.logFile.WriteString(line + "\n")
	}
}

// reportOperation 异步上报操作日志到服务器 (不阻塞主流程)
func (a *App) reportOperation(report OperationReport) {
	a.addLog(fmt.Sprintf("📊 上报: %s [%s] success=%v privacy=%v review=%v",
		report.Action, report.CustomerName, report.Success, report.PrivacyOK, report.NeedReview))

	if a.serverAPI != nil && a.serverAPI.token != "" {
		go func() {
			if err := a.serverAPI.ReportOperation(report); err != nil {
				a.addLog(fmt.Sprintf("   ⚠️ 日志上报失败: %v (已记录到本地)", err))
			}
		}()
	}
}

// reportAlert 发送告警通知 (异步, 不阻塞主流程)
func (a *App) reportAlert(level, message string) {
	a.addLog(fmt.Sprintf("🚨 告警[%s]: %s", level, message))
	if a.serverAPI != nil && a.serverAPI.token != "" {
		go func() {
			if err := a.serverAPI.SendAlert(level, message); err != nil {
				a.addLog(fmt.Sprintf("   ⚠️ 告警发送失败: %v (已记录到本地)", err))
			}
		}()
	}
}

func (a *App) agentLoop() {
	cycle := 0
	for a.running {
		cycle++
		a.addLog(fmt.Sprintf("🔄 [#%d] 开始巡检 (主理人=%s, 成员=%v, cutoff=%s)...",
			cycle, a.state.TargetUserID, a.state.FixedMembers,
			time.Unix(a.getAutoCutoffTime(), 0).Format("01-02 15:04")))

		contacts, err := a.api.GetContacts(a.state.TargetUserID)
		if err != nil {
			a.addLog(fmt.Sprintf("❌ 获取联系人失败: %s", err))
		} else {
			a.addLog(fmt.Sprintf("   📊 联系人=%d, 已处理=%d",
				len(contacts), len(a.state.ProcessedCustomers)))

			// ═══ 第 1 层: API 精确检查 — 外部联系人是否已在客户群 ═══
			inGroupMap := make(map[string]GroupCheckResult)
			if a.serverAPI != nil && a.serverAPI.token != "" {
				// 收集未处理联系人的 external_userid
				var uncheckedIDs []string
				for _, c := range contacts {
					if !a.isProcessed(c.ExternalUserID) {
						uncheckedIDs = append(uncheckedIDs, c.ExternalUserID)
					}
				}
				if len(uncheckedIDs) > 0 {
					a.addLog(fmt.Sprintf("   📡 API 检查 %d 个联系人是否已在群...", len(uncheckedIDs)))
					checkResult, checkErr := a.serverAPI.CheckCustomerInGroups(uncheckedIDs)
					if checkErr == nil {
						inGroupMap = checkResult
						inCount := 0
						for _, v := range checkResult {
							if v.InGroup {
								inCount++
							}
						}
						a.addLog(fmt.Sprintf("   📡 API 结果: %d 已在群, %d 不在群",
							inCount, len(uncheckedIDs)-inCount))
					} else {
						a.addLog(fmt.Sprintf("   ⚠️ API 检查失败: %v, 回退群名匹配", checkErr))
					}
				}
			}

			// ═══ 第 2 层备用: 群名匹配 (仅当 API 不可用时) ═══
			var groupNames []string
			if len(inGroupMap) == 0 {
				groups, _ := a.api.GetGroups()
				groupNames = make([]string, len(groups))
				for i, g := range groups {
					groupNames[i] = strings.ToLower(g.Name)
				}
				a.addLog(fmt.Sprintf("   📋 群名匹配模式: %d 个群", len(groups)))
			}

			newCount := 0
			skipped := 0
			consecutiveFails := 0
			const maxConsecutiveFails = 3
			inCycleProcessed := make(map[string]bool) // 轮内去重: 同一轮同一个人只建一次群
			customerFailCount := make(map[string]int)  // 单客户累计失败次数 (防无限重试)

			for _, c := range contacts {
				if !a.running {
					break
				}

				// ═══ 连续失败熔断检查 ═══
				if consecutiveFails >= maxConsecutiveFails {
					msg := fmt.Sprintf("🚨 连续 %d 次建群失败, 暂停巡检!", maxConsecutiveFails)
					a.addLog(msg)
					a.reportAlert("error", msg)
					a.running = false
					break
				}

				// ═══ 轮内去重: 同一轮巡检中同一客户只处理一次 (测试账号不受限) ═══
				if inCycleProcessed[c.ExternalUserID] && !a.isTestAccount(c.Name) {
					skipped++
					continue
				}

				// ═══ 时间过滤: 仅处理 cutoff 时间后添加的新客户 (测试账号跳过) ═══
				cutoff := a.getAutoCutoffTime()
				if c.AddTime > 0 && c.AddTime < cutoff && !a.isTestAccount(c.Name) {
					skipped++
					continue
				}
				if c.AddTime == 0 && !a.isTestAccount(c.Name) {
					// AddTime 未知 (服务端未返回), 跳过以确保安全
					a.addLog(fmt.Sprintf("   ⚠️ 跳过【%s】(无添加时间, 无法确认是否新客户)", c.Name))
					skipped++
					continue
				}

				// 第 3 层: 本地已处理记录 (测试账号跳过此检查)
				if a.isProcessed(c.ExternalUserID) && !a.isTestAccount(c.Name) {
					skipped++
					continue
				}

				// 第 1 层: API 精确检查 (测试账号跳过此检查)
				if gResult, ok := inGroupMap[c.ExternalUserID]; ok && gResult.InGroup && !a.isTestAccount(c.Name) {
					a.addLog(fmt.Sprintf("   ✅ 跳过【%s】(API: 已在 %d 个客户群)", c.Name, gResult.GroupCount))
					a.markProcessed(c.ExternalUserID)
					skipped++
					continue
				}

				// 第 2 层: 群名匹配 (仅当 API 结果不包含该用户时)
				if len(groupNames) > 0 {
					hasGroup := false
					for _, gn := range groupNames {
						if strings.Contains(gn, strings.ToLower(c.Name)) {
							hasGroup = true
							break
						}
					}
					if hasGroup && !a.isTestAccount(c.Name) {
						a.addLog(fmt.Sprintf("   ⚠️ 跳过【%s】(群名匹配)", c.Name))
						a.markProcessed(c.ExternalUserID)
						skipped++
						continue
					}
				}

				newCount++
				if a.isTestAccount(c.Name) {
					a.addLog(fmt.Sprintf("   🧪 [%d] 测试账号【%s】建群...", newCount, c.Name))
				} else {
					a.addLog(fmt.Sprintf("   🏗️ [%d] 为【%s】建群...", newCount, c.Name))
				}

				// ═══ 防崩溃: 先标记已处理, 建群失败再回滚 ═══
				// 避免: 建群成功 → 程序崩溃 → markProcessed 未执行 → 下轮重复建群
				if !a.isTestAccount(c.Name) {
					a.markProcessed(c.ExternalUserID)
				}

				// GUI 自动化建群
				result := a.CreateGroupForCustomer(c.Name, c.ExternalUserID)
				inCycleProcessed[c.ExternalUserID] = true // 轮内标记, 防止同一轮重复建群
				if result.Success {
					consecutiveFails = 0 // 成功则重置熔断计数
					delete(customerFailCount, c.ExternalUserID) // 成功清除失败记录
				} else {
					consecutiveFails++
					customerFailCount[c.ExternalUserID]++
					// 建群失败: 回滚 processed 标记 (允许下轮重试)
					if !a.isTestAccount(c.Name) {
						a.unmarkProcessed(c.ExternalUserID)
					}
					// 单客户累计失败 ≥3 次: 标记需人工复核, 不再自动重试
					if customerFailCount[c.ExternalUserID] >= 3 {
						a.addLog(fmt.Sprintf("   🚫 【%s】累计 3 次失败, 标记人工复核", c.Name))
						a.markProcessed(c.ExternalUserID)
						a.markNeedReview(c.ExternalUserID)
					}
					a.addLog(fmt.Sprintf("   ⚠️ 连续失败计数: %d/%d", consecutiveFails, maxConsecutiveFails))
				}
				time.Sleep(2 * time.Second)
			}
			a.addLog(fmt.Sprintf("   📊 本轮结果: 新建=%d, 跳过=%d", newCount, skipped))
		}

		// 等待下一轮
		a.addLog("   💤 本轮结束, 60s 后开始下一轮...")
		select {
		case <-a.stopCh:
			return
		case <-time.After(60 * time.Second):
		}
	}
}

// ━━━ 状态管理 (map 版, 支持从旧 slice 格式迁移) ━━━

func (a *App) isProcessed(uid string) bool {
	_, exists := a.state.ProcessedCustomers[uid]
	return exists
}

func (a *App) markProcessed(uid string) {
	if !a.isProcessed(uid) {
		a.state.ProcessedCustomers[uid] = time.Now().Unix()
		a.saveState()
	}
}

// unmarkProcessed 回滚已处理标记 (建群失败时使用)
func (a *App) unmarkProcessed(uid string) {
	if a.isProcessed(uid) {
		delete(a.state.ProcessedCustomers, uid)
		a.saveState()
	}
}

func (a *App) markNeedReview(uid string) {
	for _, existing := range a.state.NeedReviewList {
		if existing == uid {
			return
		}
	}
	a.state.NeedReviewList = append(a.state.NeedReviewList, uid)
	a.saveState()
}

// getAutoCutoffTime 获取全自动模式的起始时间 (仅处理此时间后的新客户)
// 默认: 2026-04-20 00:00:00 北京时间
func (a *App) getAutoCutoffTime() int64 {
	if a.state.AutoCutoffTime > 0 {
		return a.state.AutoCutoffTime
	}
	// 默认 cutoff: 2026-04-20 00:00:00 +08:00
	return time.Date(2026, 4, 20, 0, 0, 0, 0, time.FixedZone("CST", 8*3600)).Unix()
}

// SetAutoCutoffTime 设置全自动模式起始时间 (前端可调用)
func (a *App) SetAutoCutoffTime(unixTime int64) {
	a.state.AutoCutoffTime = unixTime
	a.saveState()
	t := time.Unix(unixTime, 0).Format("2006-01-02 15:04:05")
	a.addLog(fmt.Sprintf("⏰ 全自动起始时间已设置: %s", t))
}

// GetAutoCutoffTime 获取当前 cutoff 时间 (前端可调用)
func (a *App) GetAutoCutoffTime() int64 {
	return a.getAutoCutoffTime()
}

// isTestAccount 检查客户名是否是测试账号 (测试账号跳过所有防重检查)
// Root模式下所有客户都视为测试账号
func (a *App) isTestAccount(customerName string) bool {
	// Root模式: 全部跳过防重
	if a.state.RootMode {
		return true
	}
	for _, testName := range a.state.TestCustomerNames {
		if strings.Contains(customerName, testName) || strings.Contains(testName, customerName) {
			return true
		}
	}
	return false
}

// Root 操作员列表 (这些账号可以启用 root 模式)
var rootOperators = []string{"刘浩东"}

// IsRootOperator 检查是否是 root 操作员
func IsRootOperator(name string) bool {
	for _, op := range rootOperators {
		if name == op {
			return true
		}
	}
	return false
}

// SetRootMode 设置 root 模式
func (a *App) SetRootMode(enabled bool) {
	a.state.RootMode = enabled
	a.saveState()
	if enabled {
		a.addLog("🔓 Root 模式已开启: 所有客户跳过防重, 可无限建群")
	} else {
		a.addLog("🔒 Root 模式已关闭: 恢复正常防重检查")
	}
}

// GetRootMode 获取 root 模式状态
func (a *App) GetRootMode() bool {
	return a.state.RootMode
}

func (a *App) loadState() {
	data, err := os.ReadFile(a.stateFile)
	if err != nil {
		return
	}

	// 先尝试新格式 (map)
	if err := json.Unmarshal(data, &a.state); err != nil {
		return
	}

	// ═══ 旧格式迁移: 如果 processed_customers 是 []string 而不是 map ═══
	// JSON 解码时 map[string]int64 遇到 ["str1","str2"] 会失败
	// 此时需要手动检测并迁移
	if a.state.ProcessedCustomers == nil {
		a.state.ProcessedCustomers = make(map[string]int64)

		// 重新解析检查旧格式
		var rawState struct {
			ProcessedCustomers json.RawMessage `json:"processed_customers"`
			TargetUserID       string          `json:"target_userid"`
			FixedMembers       []string        `json:"fixed_members"`
			GroupOwner         string          `json:"group_owner"`
		}
		if json.Unmarshal(data, &rawState) == nil && len(rawState.ProcessedCustomers) > 0 {
			// 尝试解析为 []string (旧格式)
			var oldList []string
			if json.Unmarshal(rawState.ProcessedCustomers, &oldList) == nil && len(oldList) > 0 {
				now := time.Now().Unix()
				for _, uid := range oldList {
					a.state.ProcessedCustomers[uid] = now
				}
				a.state.TargetUserID = rawState.TargetUserID
				a.state.FixedMembers = rawState.FixedMembers
				a.state.GroupOwner = rawState.GroupOwner
				// 立即保存新格式
				a.saveState()
			}
		}
	}
}

func (a *App) saveState() {
	data, _ := json.MarshalIndent(a.state, "", "  ")
	os.WriteFile(a.stateFile, data, 0644)
}

// resolveToNames 将 UserID 列表转为中文名列表
// FixedMembers 存的是 UserID (如 "WuZeHua"), WeCom 弹窗搜索需要中文名 (如 "吴泽华")
func (a *App) resolveToNames(userIDs []string) []string {
	names := make([]string, 0, len(userIDs))
	for _, uid := range userIDs {
		resolved := uid // 默认: 如果找不到就原样返回
		for _, m := range a.cachedMembers {
			if strings.EqualFold(m.UserID, uid) {
				resolved = m.Name
				break
			}
		}
		names = append(names, resolved)
	}
	return names
}
