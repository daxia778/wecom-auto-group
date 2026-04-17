package main

import (
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"strings"
	"sync"
	"time"
)

// WeComAPI 企微官方 API 客户端
type WeComAPI struct {
	corpID      string
	corpSecret  string
	accessToken string
	tokenExpiry time.Time
	mu          sync.Mutex
	client      *http.Client
}

// Member 企微员工
type Member struct {
	UserID string `json:"userid"`
	Name   string `json:"name"`
	Status int    `json:"status"`
}

// Contact 外部联系人
type Contact struct {
	ExternalUserID string `json:"external_userid"`
	Name           string `json:"name"`
	Type           int    `json:"type"`
	CorpName       string `json:"corp_name"`
}

// GroupChat 客户群
type GroupChat struct {
	ChatID      string `json:"chat_id"`
	Name        string `json:"name"`
	Owner       string `json:"owner"`
	MemberCount int    `json:"member_count"`
}

func NewWeComAPI() *WeComAPI {
	return &WeComAPI{
		corpID:     WeComCorpID,
		corpSecret: WeComCorpSecret,
		client:     &http.Client{Timeout: 30 * time.Second},
	}
}

// GetAccessToken 获取/刷新 access_token (自动缓存)
func (w *WeComAPI) GetAccessToken() (string, error) {
	w.mu.Lock()
	defer w.mu.Unlock()

	if w.accessToken != "" && time.Now().Before(w.tokenExpiry) {
		return w.accessToken, nil
	}

	apiURL := fmt.Sprintf("%s/gettoken?corpid=%s&corpsecret=%s",
		WeComAPIBase, w.corpID, w.corpSecret)

	resp, err := w.client.Get(apiURL)
	if err != nil {
		return "", fmt.Errorf("请求token失败: %w", err)
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	var result struct {
		ErrCode     int    `json:"errcode"`
		ErrMsg      string `json:"errmsg"`
		AccessToken string `json:"access_token"`
		ExpiresIn   int    `json:"expires_in"`
	}
	if err := json.Unmarshal(body, &result); err != nil {
		return "", fmt.Errorf("解析token响应失败: %w", err)
	}
	if result.ErrCode != 0 {
		return "", fmt.Errorf("获取token失败: %s", result.ErrMsg)
	}

	w.accessToken = result.AccessToken
	w.tokenExpiry = time.Now().Add(time.Duration(result.ExpiresIn-300) * time.Second)
	return w.accessToken, nil
}

// apiGet GET 请求
func (w *WeComAPI) apiGet(path string, params map[string]string) ([]byte, error) {
	token, err := w.GetAccessToken()
	if err != nil {
		return nil, err
	}

	apiURL := fmt.Sprintf("%s%s?access_token=%s", WeComAPIBase, path, token)
	for k, v := range params {
		apiURL += "&" + k + "=" + url.QueryEscape(v)
	}

	resp, err := w.client.Get(apiURL)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	return io.ReadAll(resp.Body)
}

// apiPost POST 请求
func (w *WeComAPI) apiPost(path string, body interface{}) ([]byte, error) {
	token, err := w.GetAccessToken()
	if err != nil {
		return nil, err
	}

	apiURL := fmt.Sprintf("%s%s?access_token=%s", WeComAPIBase, path, token)
	jsonBody, _ := json.Marshal(body)

	resp, err := w.client.Post(apiURL, "application/json", strings.NewReader(string(jsonBody)))
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	return io.ReadAll(resp.Body)
}

// GetMembers 获取所有在职员工
func (w *WeComAPI) GetMembers() ([]Member, error) {
	data, err := w.apiGet("/user/simplelist", map[string]string{
		"department_id": "1",
		"fetch_child":   "1",
	})
	if err != nil {
		return nil, err
	}

	var result struct {
		ErrCode  int    `json:"errcode"`
		ErrMsg   string `json:"errmsg"`
		UserList []struct {
			UserID string `json:"userid"`
			Name   string `json:"name"`
			Status int    `json:"status"`
		} `json:"userlist"`
	}
	if err := json.Unmarshal(data, &result); err != nil {
		return nil, fmt.Errorf("解析员工列表失败: %w", err)
	}
	if result.ErrCode != 0 {
		return nil, fmt.Errorf("获取员工列表失败: errcode=%d, %s", result.ErrCode, result.ErrMsg)
	}

	members := make([]Member, len(result.UserList))
	for i, u := range result.UserList {
		members[i] = Member{UserID: u.UserID, Name: u.Name, Status: u.Status}
	}
	return members, nil
}

// GetFollowUserList 获取配置了客户联系功能的成员列表
func (w *WeComAPI) GetFollowUserList() ([]string, error) {
	data, err := w.apiGet("/externalcontact/get_follow_user_list", nil)
	if err != nil {
		return nil, err
	}

	var result struct {
		ErrCode    int      `json:"errcode"`
		ErrMsg     string   `json:"errmsg"`
		FollowUser []string `json:"follow_user"`
	}
	if err := json.Unmarshal(data, &result); err != nil {
		return nil, fmt.Errorf("解析客户联系员工列表失败: %w", err)
	}
	if result.ErrCode != 0 {
		return nil, fmt.Errorf("获取客户联系员工失败: errcode=%d, %s", result.ErrCode, result.ErrMsg)
	}
	return result.FollowUser, nil
}

// GetContacts 获取员工的外部联系人
func (w *WeComAPI) GetContacts(userid string) ([]Contact, error) {
	// Step 1: 获取 external_userid 列表
	data, err := w.apiGet("/externalcontact/list", map[string]string{"userid": userid})
	if err != nil {
		return nil, err
	}

	var listResult struct {
		ErrCode        int      `json:"errcode"`
		ErrMsg         string   `json:"errmsg"`
		ExternalUserID []string `json:"external_userid"`
	}
	if err := json.Unmarshal(data, &listResult); err != nil {
		return nil, fmt.Errorf("解析联系人列表失败: %w", err)
	}
	if listResult.ErrCode != 0 {
		return nil, fmt.Errorf("%s 获取联系人失败: errcode=%d, %s (可能不在客户联系权限范围)",
			userid, listResult.ErrCode, listResult.ErrMsg)
	}
	if len(listResult.ExternalUserID) == 0 {
		return []Contact{}, nil
	}

	// Step 2: 逐个获取详情 (与 Python 参考逻辑一致)
	contacts := make([]Contact, 0, len(listResult.ExternalUserID))
	for i, extID := range listResult.ExternalUserID {
		detail, err := w.apiGet("/externalcontact/get", map[string]string{"external_userid": extID})
		if err != nil {
			continue
		}
		var detailResult struct {
			ErrCode         int `json:"errcode"`
			ExternalContact struct {
				Name     string `json:"name"`
				Type     int    `json:"type"`
				CorpName string `json:"corp_name"`
			} `json:"external_contact"`
		}
		if err := json.Unmarshal(detail, &detailResult); err != nil {
			continue
		}
		if detailResult.ErrCode != 0 {
			continue
		}
		contacts = append(contacts, Contact{
			ExternalUserID: extID,
			Name:           detailResult.ExternalContact.Name,
			Type:           detailResult.ExternalContact.Type,
			CorpName:       detailResult.ExternalContact.CorpName,
		})
		// 每 20 个打印一次进度 (用于调试)
		_ = i
		time.Sleep(50 * time.Millisecond) // 避免频率限制
	}
	return contacts, nil
}

// GetGroups 获取客户群列表
func (w *WeComAPI) GetGroups() ([]GroupChat, error) {
	var allGroups []GroupChat
	cursor := ""

	for {
		body := map[string]interface{}{"status_filter": 0, "limit": 100}
		if cursor != "" {
			body["cursor"] = cursor
		}

		data, err := w.apiPost("/externalcontact/groupchat/list", body)
		if err != nil {
			return allGroups, err
		}

		var result struct {
			ErrCode       int    `json:"errcode"`
			ErrMsg        string `json:"errmsg"`
			GroupChatList []struct {
				ChatID string `json:"chat_id"`
			} `json:"group_chat_list"`
			NextCursor string `json:"next_cursor"`
		}
		if err := json.Unmarshal(data, &result); err != nil {
			return allGroups, fmt.Errorf("解析群聊列表失败: %w", err)
		}
		if result.ErrCode != 0 {
			return allGroups, fmt.Errorf("获取群聊列表失败: errcode=%d, %s", result.ErrCode, result.ErrMsg)
		}

		for _, g := range result.GroupChatList {
			// 获取群详情
			detailData, err := w.apiPost("/externalcontact/groupchat/get", map[string]string{"chat_id": g.ChatID})
			if err != nil {
				allGroups = append(allGroups, GroupChat{ChatID: g.ChatID})
				continue
			}
			var detail struct {
				GroupChat struct {
					Name       string `json:"name"`
					Owner      string `json:"owner"`
					MemberList []struct {
						UserID string `json:"userid"`
					} `json:"member_list"`
				} `json:"group_chat"`
			}
			if err := json.Unmarshal(detailData, &detail); err != nil {
				allGroups = append(allGroups, GroupChat{ChatID: g.ChatID})
				continue
			}
			allGroups = append(allGroups, GroupChat{
				ChatID:      g.ChatID,
				Name:        detail.GroupChat.Name,
				Owner:       detail.GroupChat.Owner,
				MemberCount: len(detail.GroupChat.MemberList),
			})
		}

		cursor = result.NextCursor
		if cursor == "" {
			break
		}
	}
	return allGroups, nil
}
