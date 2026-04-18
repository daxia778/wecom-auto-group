package main

import "time"

// GroupResult 建群操作的详细结果 (替代原来的 bool 返回值)
type GroupResult struct {
	Success         bool   // 建群流程是否走完 (弹窗关闭)
	PrivacySet      bool   // 隐私 checkbox 是否已勾选
	PrivacyVerified bool   // 是否经过二次截图验证确认
	MembersSelected int    // 实际勾选的成员数
	MembersExpected int    // 期望的成员数
	ErrorDetail     string // 失败原因
	NeedManualCheck bool   // 是否需要人工复核
}

// OperationReport 操作日志 (上报到服务器)
type OperationReport struct {
	Timestamp    string `json:"timestamp"`
	CustomerName string `json:"customer_name"`
	CustomerUID  string `json:"customer_uid"`
	Action       string `json:"action"`       // "create_group" / "privacy_set"
	Success      bool   `json:"success"`
	PrivacyOK    bool   `json:"privacy_ok"`
	MembersOK    bool   `json:"members_ok"`
	ErrorDetail  string `json:"error_detail,omitempty"`
	NeedReview   bool   `json:"need_review"`
}

// NewReport 创建一份操作报告
func NewReport(customerName, customerUID, action string) OperationReport {
	return OperationReport{
		Timestamp:    time.Now().Format("2006-01-02 15:04:05"),
		CustomerName: customerName,
		CustomerUID:  customerUID,
		Action:       action,
	}
}
