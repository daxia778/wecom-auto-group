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
	stateFile     string
	state         AppState
	cachedMembers []Member      // 启动时预加载的员工列表
	startupDone   chan struct{} // 关闭后表示 startup 已完成
}

// AppState 本地持久化状态
type AppState struct {
	ProcessedCustomers []string `json:"processed_customers"`
	TargetUserID       string   `json:"target_userid"`
	FixedMembers       []string `json:"fixed_members"`
	GroupOwner         string   `json:"group_owner"`
}

func NewApp() *App {
	exe, _ := os.Executable()
	dir := filepath.Dir(exe)
	stateFile := filepath.Join(dir, "local_state.json")

	// 使用服务器 API 中转 (绕过企微 IP 白名单)
	sapi := NewServerAPI()

	app := &App{
		api:         sapi,
		serverAPI:   sapi,
		stateFile:   stateFile,
		state:       AppState{},
		startupDone: make(chan struct{}),
	}
	app.loadState()
	return app
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
	for i := range contacts {
		if a.isProcessed(contacts[i].ExternalUserID) {
			contacts[i].Name = contacts[i].Name + " ✅"
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

		// 标记已处理状态
		for i := range contacts {
			if a.isProcessed(contacts[i].ExternalUserID) {
				contacts[i].Name = contacts[i].Name + " ✅"
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

// GetSettings 获取当前设置
func (a *App) GetSettings() AppState {
	return a.state
}

// CreateGroupForCustomer 为指定客户建群 (GUI 自动化)
// 注: 带外部客户的群聊只能通过 GUI 创建, 企微 API 不支持外部群
func (a *App) CreateGroupForCustomer(customerName string, customerUID string) string {
	wc, err := FindWeComWindow()
	if err != nil {
		msg := fmt.Sprintf("❌ %s", err)
		a.addLog(msg)
		return msg
	}

	a.addLog(fmt.Sprintf("🏗️ 开始建群: 客户=%s, 窗口=%dx%d", customerName, wc.Width, wc.Height))
	wc.SinkToBottom()

	success := wc.CreateGroupOCR(customerName, a.state.FixedMembers, a.addLog)
	if success {
		a.markProcessed(customerUID)
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
		msg := fmt.Sprintf("✅ 【%s】建群成功！", customerName)
		a.addLog(msg)
		return msg
	}
	msg := fmt.Sprintf("❌ 【%s】建群失败", customerName)
	a.addLog(msg)
	return msg
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
	a.logs = append(a.logs, fmt.Sprintf("[%s] %s", ts, msg))
}

func (a *App) agentLoop() {
	cycle := 0
	for a.running {
		cycle++
		a.addLog(fmt.Sprintf("🔄 [#%d] 开始巡检 (主理人=%s, 成员=%v)...",
			cycle, a.state.TargetUserID, a.state.FixedMembers))

		contacts, err := a.api.GetContacts(a.state.TargetUserID)
		if err != nil {
			a.addLog(fmt.Sprintf("❌ 获取联系人失败: %s", err))
		} else {
			a.addLog(fmt.Sprintf("   📊 联系人=%d, 已处理=%d",
				len(contacts), len(a.state.ProcessedCustomers)))

			// ═══ 第 1 层: API 精确检查 — 外部联系人是否已在客户群 ═══
			inGroupMap := make(map[string]bool)
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
							if v {
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
			for _, c := range contacts {
				if !a.running {
					break
				}
				// 第 3 层: 本地已处理记录
				if a.isProcessed(c.ExternalUserID) {
					skipped++
					continue
				}

				// 第 1 层: API 精确检查
				if inGroup, ok := inGroupMap[c.ExternalUserID]; ok && inGroup {
					a.addLog(fmt.Sprintf("   ✅ 跳过【%s】(API: 已在客户群)", c.Name))
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
					if hasGroup {
						a.addLog(fmt.Sprintf("   ⚠️ 跳过【%s】(群名匹配)", c.Name))
						a.markProcessed(c.ExternalUserID)
						skipped++
						continue
					}
				}

				newCount++
				a.addLog(fmt.Sprintf("   🏗️ [%d] 为【%s】建群...", newCount, c.Name))
				// GUI 自动化建群
				a.CreateGroupForCustomer(c.Name, c.ExternalUserID)
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

func (a *App) isProcessed(uid string) bool {
	for _, p := range a.state.ProcessedCustomers {
		if p == uid {
			return true
		}
	}
	return false
}

func (a *App) markProcessed(uid string) {
	if !a.isProcessed(uid) {
		a.state.ProcessedCustomers = append(a.state.ProcessedCustomers, uid)
		a.saveState()
	}
}

func (a *App) loadState() {
	data, err := os.ReadFile(a.stateFile)
	if err != nil {
		return
	}
	json.Unmarshal(data, &a.state)
}

func (a *App) saveState() {
	data, _ := json.MarshalIndent(a.state, "", "  ")
	os.WriteFile(a.stateFile, data, 0644)
}
