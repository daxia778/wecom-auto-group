package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"strings"
	"sync"
	"time"
)

// ServerAPI 通过服务器 API 中转调用企微接口
// 解决 IP 白名单限制: 桌面端 → 服务器(118.31.56.141) → 企微API
type ServerAPI struct {
	baseURL string // https://zhiyuanshijue.ltd/api/v1
	token   string // JWT token
	mu      sync.Mutex
	client  *http.Client
}

func NewServerAPI() *ServerAPI {
	return &ServerAPI{
		baseURL: ServerAPIBase,
		client:  &http.Client{Timeout: 120 * time.Second},
	}
}

// Login 通过管理员账号登录获取 JWT (含重试)
func (s *ServerAPI) Login(username, password string) error {
	s.mu.Lock()
	defer s.mu.Unlock()

	payload := map[string]string{
		"username": username,
		"password": password,
	}
	body, _ := json.Marshal(payload)

	var lastErr error
	for attempt := 1; attempt <= 3; attempt++ {
		resp, err := s.client.Post(
			s.baseURL+"/auth/admin_login",
			"application/json",
			bytes.NewReader(body),
		)
		if err != nil {
			lastErr = fmt.Errorf("登录请求失败(第%d次): %w", attempt, err)
			if attempt < 3 {
				time.Sleep(time.Duration(attempt) * 2 * time.Second)
			}
			continue
		}
		respBody, _ := io.ReadAll(resp.Body)
		resp.Body.Close()

		var result struct {
			Token string `json:"token"`
			Error string `json:"error"`
		}
		if err := json.Unmarshal(respBody, &result); err != nil {
			lastErr = fmt.Errorf("解析登录响应失败: %w", err)
			continue
		}
		if result.Token == "" {
			return fmt.Errorf("登录失败: %s", result.Error)
		}

		s.token = result.Token
		return nil
	}
	return lastErr
}

// SetToken 直接设置 JWT token
func (s *ServerAPI) SetToken(token string) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.token = token
}

// apiGet 带 JWT 的 GET 请求 (含重试)
func (s *ServerAPI) apiGet(path string, params map[string]string) ([]byte, error) {
	apiURL := s.baseURL + path
	if len(params) > 0 {
		values := url.Values{}
		for k, v := range params {
			values.Set(k, v)
		}
		apiURL += "?" + values.Encode()
	}

	var lastErr error
	for attempt := 1; attempt <= 3; attempt++ {
		req, err := http.NewRequest("GET", apiURL, nil)
		if err != nil {
			return nil, err
		}
		req.Header.Set("Authorization", "Bearer "+s.token)

		resp, err := s.client.Do(req)
		if err != nil {
			lastErr = fmt.Errorf("请求 %s 失败(第%d次): %w", path, attempt, err)
			if attempt < 3 {
				time.Sleep(time.Duration(attempt) * 2 * time.Second)
			}
			continue
		}
		data, err := io.ReadAll(resp.Body)
		resp.Body.Close()
		if err != nil {
			lastErr = fmt.Errorf("读取响应失败: %w", err)
			continue
		}
		return data, nil
	}
	return nil, lastErr
}

// GetMembers 通过服务器获取员工列表
// 服务器响应: {"data": [{"userid":"xxx","name":"xxx","department":"1",...}]}
func (s *ServerAPI) GetMembers() ([]Member, error) {
	data, err := s.apiGet("/admin/wecom/members", nil)
	if err != nil {
		return nil, err
	}

	var result struct {
		Data []struct {
			UserID     string `json:"userid"`
			Name       string `json:"name"`
			Status     int    `json:"status"`
			IsEmployee bool   `json:"is_employee"`
		} `json:"data"`
	}
	if err := json.Unmarshal(data, &result); err != nil {
		return nil, fmt.Errorf("解析员工列表失败: %w (原始: %s)", err, string(data[:min(len(data), 200)]))
	}

	members := make([]Member, 0, len(result.Data))
	for _, u := range result.Data {
		status := 1
		if u.Status != 0 {
			status = u.Status
		}
		members = append(members, Member{
			UserID: u.UserID,
			Name:   u.Name,
			Status: status,
		})
	}
	return members, nil
}

// GetContacts 通过服务器获取外部联系人
// 服务器: GET /admin/transfer/external-contacts?userid=xxx
// 响应: {"contacts": [{external_userid, name, type, corp_name, ...}], "total": N}
func (s *ServerAPI) GetContacts(userid string) ([]Contact, error) {
	data, err := s.apiGet("/admin/transfer/external-contacts", map[string]string{
		"userid": userid,
	})
	if err != nil {
		return nil, err
	}

	// 检查错误响应
	var errResp struct {
		Code    string `json:"code"`
		Message string `json:"message"`
	}
	if json.Unmarshal(data, &errResp) == nil && errResp.Code != "" {
		return nil, fmt.Errorf("服务器错误: %s - %s", errResp.Code, errResp.Message)
	}

	// 服务器返回: {"contacts": [...], "total": N}
	var result struct {
		Contacts []struct {
			ExternalUserID string `json:"external_userid"`
			Name           string `json:"name"`
			Type           int    `json:"type"`
			CorpName       string `json:"corp_name"`
			Avatar         string `json:"avatar"`
			Gender         int    `json:"gender"`
			AddTime        int64  `json:"add_time"`
		} `json:"contacts"`
		Total int `json:"total"`
	}
	if err := json.Unmarshal(data, &result); err != nil {
		return nil, fmt.Errorf("解析联系人失败: %w (原始: %s)", err, string(data[:min(len(data), 300)]))
	}

	contacts := make([]Contact, len(result.Contacts))
	for i, c := range result.Contacts {
		contacts[i] = Contact{
			ExternalUserID: c.ExternalUserID,
			Name:           c.Name,
			Type:           c.Type,
			CorpName:       c.CorpName,
			AddTime:        c.AddTime,
		}
	}
	return contacts, nil
}

// GetGroups 通过服务器获取群聊列表
// 服务器响应: {"data": [{"id":2,"chat_id":"xxx","name":"xxx","owner_id":"xxx",...}]}
func (s *ServerAPI) GetGroups() ([]GroupChat, error) {
	data, err := s.apiGet("/admin/wecom/groups", nil)
	if err != nil {
		return nil, err
	}

	var result struct {
		Data []struct {
			ChatID    string `json:"chat_id"`
			Name      string `json:"name"`
			OwnerID   string `json:"owner_id"`
			MemberIDs string `json:"member_ids"`
			Status    string `json:"status"`
		} `json:"data"`
	}
	if err := json.Unmarshal(data, &result); err != nil {
		return nil, fmt.Errorf("解析群聊列表失败: %w (原始: %s)", err, string(data[:min(len(data), 200)]))
	}

	groups := make([]GroupChat, 0, len(result.Data))
	for _, g := range result.Data {
		// member_ids 是逗号分隔的字符串
		memberCount := 0
		if g.MemberIDs != "" {
			for _, c := range g.MemberIDs {
				if c == ',' {
					memberCount++
				}
			}
			memberCount++ // 逗号数 + 1
		}
		groups = append(groups, GroupChat{
			ChatID:      g.ChatID,
			Name:        g.Name,
			Owner:       g.OwnerID,
			MemberCount: memberCount,
		})
	}
	return groups, nil
}

// GetFollowUserList 获取有客户联系权限的员工
func (s *ServerAPI) GetFollowUserList() ([]string, error) {
	members, err := s.GetMembers()
	if err != nil {
		return nil, err
	}
	userIDs := make([]string, len(members))
	for i, m := range members {
		userIDs[i] = m.UserID
	}
	return userIDs, nil
}

// GroupCheckResult 单个客户的群检查结果
type GroupCheckResult struct {
	InGroup    bool
	GroupCount int
}

// CheckCustomerInGroups 批量检查外部联系人是否已在客户群中
// 调用: GET /admin/wecom/customer-groups/members-check?external_userids=id1,id2,...
// 服务端有内存缓存, O(1) 查询, 瞬间返回
func (s *ServerAPI) CheckCustomerInGroups(externalUserIDs []string) (map[string]GroupCheckResult, error) {
	if len(externalUserIDs) == 0 {
		return map[string]GroupCheckResult{}, nil
	}

	// 逗号拼接 (最多 50 个一批)
	idStr := strings.Join(externalUserIDs, ",")
	params := map[string]string{"external_userids": idStr}

	data, err := s.apiGet("/admin/wecom/customer-groups/members-check", params)
	if err != nil {
		return nil, fmt.Errorf("客户群检查失败: %w", err)
	}

	var result struct {
		Results []struct {
			ExternalUserID string `json:"external_user_id"`
			InGroup        bool   `json:"in_group"`
			GroupCount     int    `json:"group_count"`
		} `json:"results"`
		InGroup    int `json:"in_group"`
		NotInGroup int `json:"not_in_group"`
	}
	if err := json.Unmarshal(data, &result); err != nil {
		return nil, fmt.Errorf("解析客户群检查结果失败: %w", err)
	}

	// 构建 map: external_userid → GroupCheckResult
	m := make(map[string]GroupCheckResult, len(result.Results))
	for _, r := range result.Results {
		m[r.ExternalUserID] = GroupCheckResult{
			InGroup:    r.InGroup,
			GroupCount: r.GroupCount,
		}
	}
	return m, nil
}

// SyncCustomerGroups 触发服务端客户群缓存刷新
// 建群成功后调用, 让服务端立即感知新群, 避免 30 分钟刷新窗口内的重复建群
func (s *ServerAPI) SyncCustomerGroups() error {
	data, err := s.apiGet("/admin/wecom/customer-groups/sync", nil)
	if err != nil {
		return fmt.Errorf("触发同步失败: %w", err)
	}
	_ = data
	return nil
}

// apiPost 带 JWT 的 POST 请求 (含重试)
func (s *ServerAPI) apiPost(path string, payload interface{}) ([]byte, error) {
	body, err := json.Marshal(payload)
	if err != nil {
		return nil, fmt.Errorf("JSON 编码失败: %w", err)
	}

	var lastErr error
	for attempt := 1; attempt <= 3; attempt++ {
		req, err := http.NewRequest("POST", s.baseURL+path, bytes.NewReader(body))
		if err != nil {
			return nil, err
		}
		req.Header.Set("Authorization", "Bearer "+s.token)
		req.Header.Set("Content-Type", "application/json")

		resp, err := s.client.Do(req)
		if err != nil {
			lastErr = fmt.Errorf("请求 %s 失败(第%d次): %w", path, attempt, err)
			if attempt < 3 {
				time.Sleep(time.Duration(attempt) * 2 * time.Second)
			}
			continue
		}
		data, err := io.ReadAll(resp.Body)
		resp.Body.Close()
		if err != nil {
			lastErr = fmt.Errorf("读取响应失败: %w", err)
			continue
		}

		// 检查 HTTP 状态码
		if resp.StatusCode >= 400 {
			lastErr = fmt.Errorf("HTTP %d: %s", resp.StatusCode, string(data[:min(len(data), 200)]))
			if attempt < 3 {
				time.Sleep(time.Duration(attempt) * 2 * time.Second)
			}
			continue
		}
		return data, nil
	}
	return nil, lastErr
}

// CreateGroupChat 通过服务器 API 创建群聊
// 调用: POST /admin/wecom/groups/create
// 返回: chat_id (群聊ID)
func (s *ServerAPI) CreateGroupChat(name, owner string, memberUserIDs []string) (string, error) {
	payload := map[string]interface{}{
		"name":    name,
		"owner":   owner,
		"userids": memberUserIDs,
	}

	data, err := s.apiPost("/admin/wecom/groups/create", payload)
	if err != nil {
		return "", fmt.Errorf("创建群聊失败: %w", err)
	}

	var result struct {
		ChatID  string `json:"chat_id"`
		Error   string `json:"error"`
		Message string `json:"message"`
	}
	if err := json.Unmarshal(data, &result); err != nil {
		return "", fmt.Errorf("解析群聊响应失败: %w (原始: %s)", err, string(data[:min(len(data), 200)]))
	}
	if result.Error != "" {
		return "", fmt.Errorf("服务器: %s", result.Error)
	}
	if result.ChatID == "" {
		// 可能在 data 字段里
		var result2 struct {
			Data struct {
				ChatID string `json:"chat_id"`
			} `json:"data"`
		}
		json.Unmarshal(data, &result2)
		if result2.Data.ChatID != "" {
			return result2.Data.ChatID, nil
		}
		return "", fmt.Errorf("服务器未返回 chat_id (响应: %s)", string(data[:min(len(data), 200)]))
	}
	return result.ChatID, nil
}

// ReportOperation 上报操作日志到服务器
// 服务端接口: POST /admin/wecom/operation-log
// 如果服务端未实现此接口, 会返回 404 错误但不影响主流程
func (s *ServerAPI) ReportOperation(report OperationReport) error {
	_, err := s.apiPost("/admin/wecom/operation-log", report)
	if err != nil {
		return fmt.Errorf("上报操作日志失败: %w", err)
	}
	return nil
}

// SendAlert 发送告警通知 (调用服务器 webhook 推送到企微群)
// 服务端接口: POST /admin/wecom/alert
// level: "warning" / "error" / "critical"
func (s *ServerAPI) SendAlert(level, message string) error {
	payload := map[string]string{
		"level":   level,
		"message": message,
		"source":  "WeComAutoGroup",
	}
	_, err := s.apiPost("/admin/wecom/alert", payload)
	if err != nil {
		return fmt.Errorf("发送告警失败: %w", err)
	}
	return nil
}

