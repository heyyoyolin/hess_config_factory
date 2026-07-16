# HESS FortiGate Config 批次生成工具 v0.9

## 1. 專案目的

本工具使用 Streamlit 建立網頁介面，讓使用者上傳：

1. 何嘉仁專案 Excel 總表。
2. FortiGate Config 範本。

系統會逐列讀取 Excel 的分校資料，將 Config 範本內的預留字串替換成各分校的實際設定，批次產生 `.conf` 檔案，最後打包成 ZIP 供下載。

一筆有效 Excel 資料對應一份 FortiGate Config。

---

## 2. 執行環境

建議 Python 3.10 以上。

必要套件：

```text
streamlit
pandas
openpyxl
```

安裝：

```bash
pip install streamlit pandas openpyxl
```

啟動：

```bash
streamlit run HESS_CONFIG_FACTORY_v0.9.py
```

---

## 3. 輸入檔案

### 3.1 Excel

程式固定讀取工作表：

```text
總表
```

必要欄位：

| 欄位 | 用途 |
|---|---|
| `分校代碼` | Hostname、物件名稱、DDNS、SNMP 與輸出檔名 |
| `IP` | 產生 SITE 與 Mobile 網段 |
| `Voice 2 IP` | 產生 Voice2 網段 |
| `HQ_IPSEC_IP(SPOKE)` | 產生 VPN Tunnel 本地與遠端 IP |
| `客戶號碼` | 產生 WAN、LAN1 的 PPPoE 帳號；欄位可空白，但欄名必須存在 |

Excel 的 `客戶密碼`欄位不參與 Config 產生。

### 3.2 Config 範本

支援副檔名：

```text
.conf
.txt
```

程式依序嘗試以下編碼：

1. `utf-8-sig`
2. `utf-8`
3. `cp950`
4. `big5`

若全部無法解碼，整批停止。

---

## 4. Config 密碼規格

WAN 與 LAN1 的密碼固定使用 Config 範本內的：

```text
set password "!QAZ2wsx"
```

程式不讀取 Excel 的 `客戶密碼`，也不替換此固定值。

---

## 5. 分校代碼規格

正式分校代碼必須是恰好四位數字，例如：

```text
1133
3803
2164
```

Excel 若將數字讀成 `3803.0`，程式會轉為 `3803`。

以下格式視為異常，不產出：

```text
66
202301
1108-01
2010(暫定)
```

若相同四碼分校代碼在 Excel 出現兩次以上，所有重複列都不產出，避免 ZIP 內出現同名 Config。

---

## 6. 預留字串與替換規則

### 6.1 `HESSXXXX`

來源：

```text
分校代碼
```

例如分校代碼為 `3803`：

```text
HESSXXXX
```

替換為：

```text
HESS3803
```

此替換套用於 Config 內所有出現位置，包括 Hostname、DDNS、SNMP、Address Object 與其他物件名稱。

### 6.2 `HN_NUMBER`

來源：

```text
客戶號碼
```

規則：

| Excel 原始內容 | 處理結果 |
|---|---|
| `HN78211899` | 使用 `HN78211899` |
| `hn78211899` | 正規化為 `HN78211899` |
| `78211899` | 自動補成 `HN78211899` |
| 空白 | 視為無 HN |
| `xxxxxxxx` | 視為無 HN |
| 其他非 HN、非純數字格式 | 視為無 HN |

有 HN 時：

```text
set username "HN_NUMBER@ip.hinet.net"
set username "HN_NUMBER@hinet.net"
```

會替換為實際帳號。

無 HN 時：

- Config 仍會產出。
- Config 內保留 `HN_NUMBER`，供後續人工補登。
- 檔名加上 `_noHN`。

例如：

```text
1133_noHN.conf
```

### 6.3 `.XXX.YYY.`

來源為 `IP` 第一行。

例如：

```text
10.38.3.x/24
```

取第二、第三段：

```text
.38.3.
```

因此：

```text
10.XXX.YYY.254
```

會變成：

```text
10.38.3.254
```

### 6.4 `.AAA.BBB.`

若 `IP` 有第二行，使用第二行的第二、第三段。

例如：

```text
10.38.3.x/24
10.138.3.x/24
```

得到：

```text
.138.3.
```

若只有一行，使用第一行第二段加 100。

例如：

```text
10.38.3.x/24
```

得到：

```text
.138.3.
```

若加 100 後超過 255：

- 不跨網段。
- 不產出該分校 Config。
- 在 `Execution_Summary.txt` 清楚記錄計算溢位。

### 6.5 `CCC.DDD.EEE.`

來源：

```text
Voice 2 IP
```

例如：

```text
172.24.50.0/24
```

取前三段：

```text
172.24.50.
```

Config 可自行組合 `.0`、`.1`、`.2` 等主機位址。

### 6.6 `KKK.LLL.MMM.NNN`

來源：

```text
HQ_IPSEC_IP(SPOKE)
```

例如：

```text
169.254.254.101
```

完整替換為：

```text
169.254.254.101
```

### 6.7 `KKK.LLL.MMM.NNN+1`

將 SPOKE 最後一段加 1。

例如：

```text
169.254.254.101
```

得到：

```text
169.254.254.102
```

若原始 IP 最後一段為 `255`，加 1 會跨越原本 `/24` 網段：

- 不產出該分校 Config。
- 不產生 `.256`。
- 在 `Execution_Summary.txt` 記錄原始 IP、計算結果及跨網段原因。

---

## 7. IP 欄位格式

`IP`欄位接受一行或兩行，並支援：

```text
10.38.3.x/24
10.38.3.254/24
10.38.3.x
10.38.3.254
```

支援 Windows、Linux 與舊 Mac 換行格式。

檢查規則：

1. 不可空白。
2. 最少一行、最多兩行。
3. 每行必須有四段。
4. 前三段必須是 `0-255` 的數字。
5. 第四段可為 `x` 或 `0-255` 的數字。
6. CIDR 若存在，必須介於 `/0` 至 `/32`。
7. 有兩行時，兩行的第一段與第三段必須一致。
8. 格式不符時，該列不產出。

---

## 8. Voice 2 IP 檢查

以下情況不產出：

- 欄位空白。
- 不是四段 IPv4。
- 含雙句點。
- 含非數字內容。
- 任一段超出 `0-255`。
- CIDR 前綴不合法。

---

## 9. 整批停止條件

以下問題會停止整批作業：

- Excel 無法讀取。
- 找不到 `總表`工作表。
- 缺少任一必要欄位。
- Config 是空檔。
- Config 無法以支援的編碼解碼。
- Config 缺少任一必要預留字串。

必要預留字串：

```text
HESSXXXX
HN_NUMBER
.XXX.YYY.
.AAA.BBB.
CCC.DDD.EEE.
KKK.LLL.MMM.NNN
KKK.LLL.MMM.NNN+1
```

---

## 10. 單筆停止條件

以下問題只停止該分校，不影響其他分校：

- 分校代碼不是四位數字。
- 分校代碼重複。
- IP 缺失或格式錯誤。
- IP 超過兩行。
- `.AAA.BBB.` 加 100 溢位。
- Voice 2 IP 缺失或格式錯誤。
- SPOKE IP 缺失或格式錯誤。
- SPOKE IP +1 跨越原本 `/24`。
- Config 替換後仍殘留必要預留字串。

無 HN 帳號不是停止條件。

---

## 11. 輸出檔名

有合法 HN 帳號：

```text
{分校代碼}.conf
```

例如：

```text
3803.conf
```

沒有合法 HN 帳號：

```text
{分校代碼}_noHN.conf
```

例如：

```text
1133_noHN.conf
```

ZIP 檔名：

```text
HESS_Configs_Batch.zip
```

ZIP 內另包含：

```text
Execution_Summary.txt
```

---

## 12. Execution_Summary.txt

執行摘要使用 UTF-8 BOM，方便 Windows 記事本直接顯示中文。

內容包含：

- 執行時間。
- Excel 工作表名稱。
- Config 解碼方式。
- Excel 總資料列。
- 成功數量。
- 失敗數量。
- 跳過空白列數量。
- Config 預留字串及出現次數。
- 成功檔名清單。
- 每筆成功資料的 HN 狀態。
- 每筆失敗資料的 Excel 列號、分校代碼、錯誤原因及未產出結果。

無 HN 的成功項目會明確標示：

```text
HN：缺失，Config 保留 HN_NUMBER
```

---

## 13. 部署注意事項

GitHub 建議至少包含：

```text
HESS_CONFIG_FACTORY_v0.9.py
REQUIREMENTS_v0.9.md
requirements.txt
```

`requirements.txt` 可使用：

```text
streamlit
pandas
openpyxl
```

部署至 Streamlit Community Cloud 時，Main file path 設為：

```text
HESS_CONFIG_FACTORY_v0.9.py
```

Config 範本與 Excel 由使用者在網頁上傳，不需要提交實際客戶資料至公開 GitHub Repository。

---

## 14. 資安注意事項

本專案會處理 PPPoE 帳號、內部網段與 VPN IP 等資訊。

建議：

- GitHub Repository 設為 Private。
- 不提交正式 Excel。
- 不提交已產生的分校 Config。
- 不提交含正式客戶帳號或設備設定的 ZIP。
- 定期檢查 Git 歷史是否誤提交敏感檔案。
