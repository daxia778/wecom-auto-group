package main

import (
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"strings"
	"time"
)

// TestWeComAPI 测试完整 API 链路 (通过服务器中转)
func TestWeComAPI() {
	fmt.Println("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
	fmt.Println("  WeCom API 集成测试 (服务器中转模式)")
	fmt.Println("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

	client := &http.Client{Timeout: 120 * time.Second}

	// ═══ Step 1: 登录服务器 ═══
	fmt.Printf("\n[1/4] 登录服务器 (%s)...\n", ServerAPIBase)
	loginBody, _ := json.Marshal(map[string]string{
		"username": ServerAdminUser,
		"password": ServerAdminPass,
	})
	resp, err := client.Post(
		ServerAPIBase+"/auth/admin_login",
		"application/json",
		strings.NewReader(string(loginBody)),
	)
	if err != nil {
		fmt.Printf("  ❌ 请求失败: %s\n", err)
		return
	}
	defer resp.Body.Close()
	body, _ := io.ReadAll(resp.Body)

	var loginResult struct {
		Token    string `json:"token"`
		Error    string `json:"error"`
		Employee struct {
			Name string `json:"name"`
			Role string `json:"role"`
		} `json:"employee"`
	}
	json.Unmarshal(body, &loginResult)

	if loginResult.Token == "" {
		fmt.Printf("  ❌ 登录失败: %s\n", loginResult.Error)
		return
	}
	fmt.Printf("  ✅ 登录成功! 用户=%s 角色=%s\n", loginResult.Employee.Name, loginResult.Employee.Role)

	// 使用 ServerAPI
	sapi := NewServerAPI()
	sapi.SetToken(loginResult.Token)

	// ═══ Step 2: 获取员工列表 ═══
	fmt.Println("\n[2/4] 获取企微员工列表...")
	members, err := sapi.GetMembers()
	if err != nil {
		fmt.Printf("  ❌ 失败: %s\n", err)
	} else {
		fmt.Printf("  ✅ 获取到 %d 名员工\n", len(members))
		fmt.Println("  ┌─────────────────┬──────────────────┐")
		fmt.Println("  │ UserID          │ Name             │")
		fmt.Println("  ├─────────────────┼──────────────────┤")
		for i, m := range members {
			if i >= 8 {
				fmt.Printf("  │ ... 还有 %d 人                     │\n", len(members)-8)
				break
			}
			fmt.Printf("  │ %-15s │ %-16s │\n", m.UserID, m.Name)
		}
		fmt.Println("  └─────────────────┴──────────────────┘")
	}

	// ═══ Step 3: 获取外部联系人 ═══
	// 用 WuZeHua (已知有 140+ 客户)
	testUserID := "WuZeHua"
	fmt.Printf("\n[3/4] 获取 %s 的外部联系人 (可能需要 1-2 分钟)...\n", testUserID)
	contacts, err := sapi.GetContacts(testUserID)
	if err != nil {
		fmt.Printf("  ⚠️ %s\n", err)
	} else {
		fmt.Printf("  ✅ 获取到 %d 个外部联系人\n", len(contacts))
		fmt.Println("  ┌──────────────────────┬──────────────────┬──────┐")
		fmt.Println("  │ ExternalUserID       │ Name             │ Type │")
		fmt.Println("  ├──────────────────────┼──────────────────┼──────┤")
		for i, c := range contacts {
			if i >= 8 {
				fmt.Printf("  │ ... 还有 %d 个                                   │\n", len(contacts)-8)
				break
			}
			typeText := "微信"
			if c.Type != 1 {
				typeText = "企业"
			}
			eid := c.ExternalUserID
			if len(eid) > 20 {
				eid = eid[:20]
			}
			fmt.Printf("  │ %-20s │ %-16s │ %-4s │\n", eid, c.Name, typeText)
		}
		fmt.Println("  └──────────────────────┴──────────────────┴──────┘")
	}

	// ═══ Step 4: 获取群聊列表 ═══
	fmt.Println("\n[4/4] 获取群聊列表...")
	groups, err := sapi.GetGroups()
	if err != nil {
		fmt.Printf("  ❌ 失败: %s\n", err)
	} else {
		fmt.Printf("  ✅ 获取到 %d 个群聊\n", len(groups))
		for i, g := range groups {
			if i >= 5 {
				fmt.Printf("    ... 还有 %d 个群\n", len(groups)-5)
				break
			}
			name := g.Name
			if name == "" {
				name = "(未命名)"
			}
			fmt.Printf("    %d. %s (群主=%s, 成员≈%d)\n", i+1, name, g.Owner, g.MemberCount)
		}
	}

	fmt.Println("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
	fmt.Println("  ✅ API 测试完成!")
	fmt.Println("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
}
