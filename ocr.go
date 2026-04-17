package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"mime/multipart"
	"net/http"
	"strings"
	"time"
)

// OCRItem OCR 识别结果
type OCRItem struct {
	Text string  `json:"text"`
	CX   int     `json:"cx"`
	CY   int     `json:"cy"`
	Conf float64 `json:"conf"`
	X1   int     `json:"x1"`
	Y1   int     `json:"y1"`
	X2   int     `json:"x2"`
	Y2   int     `json:"y2"`
}

// ZhipuOCR 智谱 AI OCR 识别 (0.01元/次)
func ZhipuOCR(pngData []byte) ([]OCRItem, error) {
	var buf bytes.Buffer
	writer := multipart.NewWriter(&buf)

	// 表单字段
	writer.WriteField("tool_type", "hand_write")
	writer.WriteField("language_type", "CHN_ENG")
	writer.WriteField("probability", "true")

	// 文件字段
	part, err := writer.CreateFormFile("file", "screenshot.png")
	if err != nil {
		return nil, err
	}
	part.Write(pngData)
	writer.Close()

	req, err := http.NewRequest("POST", ZhipuOCRURL, &buf)
	if err != nil {
		return nil, err
	}
	req.Header.Set("Authorization", "Bearer "+ZhipuAPIKey)
	req.Header.Set("Content-Type", writer.FormDataContentType())

	client := &http.Client{Timeout: 15 * time.Second}
	resp, err := client.Do(req)
	if err != nil {
		return nil, fmt.Errorf("OCR 请求失败: %w", err)
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	var result struct {
		WordsResult []struct {
			Words    string `json:"words"`
			Location struct {
				Left   int `json:"left"`
				Top    int `json:"top"`
				Width  int `json:"width"`
				Height int `json:"height"`
			} `json:"location"`
			Probability struct {
				Average float64 `json:"average"`
			} `json:"probability"`
		} `json:"words_result"`
	}

	if err := json.Unmarshal(body, &result); err != nil {
		return nil, fmt.Errorf("OCR 解析失败: %w", err)
	}

	items := make([]OCRItem, len(result.WordsResult))
	for i, w := range result.WordsResult {
		items[i] = OCRItem{
			Text: w.Words,
			CX:   w.Location.Left + w.Location.Width/2,
			CY:   w.Location.Top + w.Location.Height/2,
			Conf: w.Probability.Average,
			X1:   w.Location.Left,
			Y1:   w.Location.Top,
			X2:   w.Location.Left + w.Location.Width,
			Y2:   w.Location.Top + w.Location.Height,
		}
	}
	return items, nil
}

// FindOCRText 在 OCR 结果中查找匹配项 (多策略模糊匹配)
// 移植自 Python wecom_auto.py ocr_find()
func FindOCRText(items []OCRItem, keyword string) *OCRItem {
	if len(items) == 0 || keyword == "" {
		return nil
	}
	kwRunes := []rune(keyword)

	// 策略1: 精确子串匹配
	for i := range items {
		if strings.Contains(items[i].Text, keyword) {
			return &items[i]
		}
	}

	// 策略2: 首字匹配 + 长度接近 (±3字以内)
	for i := range items {
		tRunes := []rune(items[i].Text)
		if len(tRunes) > 0 && len(kwRunes) > 0 &&
			tRunes[0] == kwRunes[0] &&
			abs(len(tRunes)-len(kwRunes)) <= 3 {
			return &items[i]
		}
	}

	// 策略3: 字符重叠率 >= 50% (keyword 的字有多少出现在 text 中)
	if len(kwRunes) >= 2 {
		var bestItem *OCRItem
		bestRatio := 0.0
		for i := range items {
			t := items[i].Text
			if len(t) < 1 {
				continue
			}
			matchCount := 0
			for _, c := range kwRunes {
				if strings.ContainsRune(t, c) {
					matchCount++
				}
			}
			ratio := float64(matchCount) / float64(len(kwRunes))
			if ratio > bestRatio && ratio >= 0.5 {
				bestRatio = ratio
				bestItem = &items[i]
			}
		}
		if bestItem != nil {
			return bestItem
		}
	}

	return nil
}

// FindOCRTextAll 返回所有匹配项 (用于位置过滤)
func FindOCRTextAll(items []OCRItem, keyword string) []OCRItem {
	var matches []OCRItem
	for _, item := range items {
		if strings.Contains(item.Text, keyword) {
			matches = append(matches, item)
		}
	}
	return matches
}

// FindOCRTextInRegion 在指定矩形区域内搜索 OCR 结果
func FindOCRTextInRegion(items []OCRItem, keyword string, x1, y1, x2, y2 int) *OCRItem {
	var filtered []OCRItem
	for _, item := range items {
		if item.CX >= x1 && item.CX <= x2 && item.CY >= y1 && item.CY <= y2 {
			filtered = append(filtered, item)
		}
	}
	return FindOCRText(filtered, keyword)
}

func abs(x int) int {
	if x < 0 {
		return -x
	}
	return x
}
