# -*- coding: utf-8 -*-
"""快速验证: Python 能否调通企微 API (对比 Go 版本)"""
import urllib.request, json

CORP_ID = "wwdb2f088115fa0fff"
SECRET = "fdaIml1ODRNKyZFPNF04kZz7zn0Mfv8yqW78Fr7zYh0"
BASE = "https://qyapi.weixin.qq.com/cgi-bin"

print("=== Python 企微 API 验证 ===\n")

# Step 1: 获取 token
print("[1] 获取 access_token...")
url1 = f"{BASE}/gettoken?corpid={CORP_ID}&corpsecret={SECRET}"
r1 = json.loads(urllib.request.urlopen(url1, timeout=15).read())
if r1.get("errcode", 0) != 0:
    print(f"  FAIL: {r1}")
    exit(1)
token = r1["access_token"]
print(f"  OK: token={token[:20]}...")

# Step 2: 获取员工列表
print("\n[2] 获取员工列表 (/user/simplelist)...")
url2 = f"{BASE}/user/simplelist?access_token={token}&department_id=1&fetch_child=1"
r2 = json.loads(urllib.request.urlopen(url2, timeout=15).read())
print(f"  errcode={r2.get('errcode')}, members={len(r2.get('userlist', []))}")
if r2.get("errcode", 0) != 0:
    print(f"  errmsg={r2.get('errmsg', '')}")

# Step 3: 获取外部联系人
print("\n[3] 获取外部联系人 (/externalcontact/list)...")
url3 = f"{BASE}/externalcontact/list?access_token={token}&userid=WuZeHua"
r3 = json.loads(urllib.request.urlopen(url3, timeout=15).read())
print(f"  errcode={r3.get('errcode')}, contacts={len(r3.get('external_userid', []))}")
if r3.get("errcode", 0) != 0:
    print(f"  errmsg={r3.get('errmsg', '')}")

print("\n=== 验证完毕 ===")
