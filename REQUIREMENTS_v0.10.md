# HESS FortiGate Config 批次生成工具 v0.10

## 1. 專案目的

本工具使用 Streamlit 建立網頁介面。使用者上傳何嘉仁專案 Excel 總表及 FortiGate Config 範本後，系統會依 Excel 每一列的分校資料進行檢查、預覽與 Config 產生。

v0.10 同時支援兩種產出方式：

1. **整批產出**：處理 Excel 全部分校，成功的 Config 打包成 ZIP，失敗資料寫入 `Execution_Summary.txt`。
2. **指定分校產出**：網頁顯示全部分校資料，可搜尋分校代碼、勾選需要的分校，再下載單一 `.conf` 或指定分校 ZIP。

資料即使有缺漏或格式錯誤，也必須顯示於網頁清單中，並明確說明無法產出的原因。

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
streamlit run HESS_CONFIG_FACTORY_v0.10.py
```

---

## 3. Excel 輸入規格

程式固定讀取工作表：

```text
總表
```

必要欄位：

| 欄位 | 用途 |
|---|---|
| `分校代碼` | Hostname、物件名稱、DDNS、SNMP、查詢及輸出檔名 |
| `IP` | SITE 與 Mobile 網段 |
| `Voice 2 IP` | Voice2 網段 |
| `HQ_IPSEC_IP(SPOKE)` | VPN Tunnel 本地及遠端 IP |
| `客戶號碼` | WAN、LAN1 PPPoE 帳號；欄位可空白，但欄名必須存在 |

Excel 的 `客戶密碼`完全不參與 Config 產生。

---

## 4. Config 輸入規格

支援：

```text
.conf
.txt
```

依序嘗試編碼：

1. `utf-8-sig`
2. `utf-8`
3. `cp950`
4. `big5`

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

若缺少任一必要預留字串，整批停止。

---

## 5. 固定密碼規格

Excel 的 `客戶密碼`不讀取、不驗證、不替換。

Config 的 WAN 與 LAN1 密碼固定維持：

```text
set password "!QAZ2wsx"
```

---

## 6. 分校代碼規格

### 6.1 四碼分校代碼

正常四碼直接使用：

```text
3803 -> 3803
2164 -> 2164
```

Excel 若讀成小數：

```text
3803.0 -> 3803
```

### 6.2 兩碼分校代碼自動補零

若分校代碼**恰好只有兩碼且為純數字**，前方補零至四碼：

```text
66 -> 0066
24 -> 0024
```

補零後的四碼會套用於：

- `HESSXXXX` 替換。
- 網頁查詢與顯示。
- Config 檔名。
- 重複分校代碼檢查。

例如 Excel 分校代碼為 `66`，產生：

```text
HESS0066
0066.conf
```

若該分校沒有 HN：

```text
0066_noHN.conf
```

### 6.3 不自動修正的代碼

只有「兩碼」會自動補零。其他不符合四碼的格式仍視為錯誤，例如：

```text
6
123
202301
1108-01
2010(暫定)
```

### 6.4 重複代碼

重複檢查使用**正規化後的分校代碼**。

因此：

```text
66
0066
```

會被視為同一個分校代碼 `0066`，兩筆都不產出，並說明重複的 Excel 列號。

---

## 7. HN_NUMBER 規格

`HN_NUMBER`只放入客戶號碼的數字部分。

| Excel 客戶號碼 | Config 替換值 |
|---|---|
| `HN78211899` | `78211899` |
| `hn78211899` | `78211899` |
| `78211899` | `78211899` |
| 空白 | 無 HN |
| `xxxxxxxx` | 無 HN |
| 其他格式 | 無 HN |

例如：

```text
set username "HN_NUMBER@ip.hinet.net"
set username "HN_NUMBER@hinet.net"
```

Excel 為 `HN78211899` 時產生：

```text
set username "78211899@ip.hinet.net"
set username "78211899@hinet.net"
```

無 HN 不是失敗條件：

- Config 仍產出。
- Config 保留 `HN_NUMBER`。
- 檔名加 `_noHN`。

---

## 8. IP 替換規格

### 8.1 `.XXX.YYY.`

來源為 `IP`第一行的第二、第三段。

```text
10.38.3.x/24 -> .38.3.
```

### 8.2 `.AAA.BBB.`

若 `IP`有第二行，使用第二行的第二、第三段：

```text
10.38.3.x/24
10.138.3.x/24
```

得到：

```text
.138.3.
```

若只有一行，第一行第二段加 100。

```text
10.38.3.x/24 -> .138.3.
```

若加 100 超過 255，不跨網段、不產出，並顯示錯誤。

### 8.3 IP 格式檢查

接受：

```text
10.38.3.x/24
10.38.3.254/24
10.38.3.x
10.38.3.254
```

規則：

- 最少一行、最多兩行。
- 每行四段。
- 前三段必須為 0-255。
- 第四段可為 `x` 或 0-255。
- CIDR 若存在必須為 `/0` 到 `/32`。
- 有兩行時，第一段與第三段必須一致。

---

## 9. Voice 2 IP 規格

來源：

```text
Voice 2 IP
```

例如：

```text
172.24.50.0/24 -> 172.24.50.
```

以下情況不可產出：

- 空白。
- 非四段 IPv4。
- 雙句點。
- 非數字內容。
- 任一段超過 255。
- CIDR 不合法。

---

## 10. VPN SPOKE 規格

`KKK.LLL.MMM.NNN`直接使用 `HQ_IPSEC_IP(SPOKE)`。

`KKK.LLL.MMM.NNN+1`只增加第四段：

```text
169.254.254.101 -> 169.254.254.102
```

若第四段為 255，加一會跨原 `/24`：

- 不產出。
- 不產生 `.256`。
- 網頁與 Execution Summary 都顯示跨網段原因。

---

## 11. 網頁資料預檢功能

Excel 與 Config 驗證完成後，程式先對每一列執行與正式產檔相同的檢查，再顯示完整清單。

清單包含：

| 欄位 | 說明 |
|---|---|
| `選取` | 使用者勾選指定分校 |
| `分校代碼` | 已正規化的代碼；兩碼會顯示成四碼 |
| `單位名稱` | Excel 的單位名稱，若有 |
| `Excel列號` | 方便回頭修改來源資料 |
| `HN狀態` | 顯示數字帳號或無 HN |
| `預計檔名` | 實際產出檔名 |
| `可產出` | `✅` 或 `❌` |
| `檢查結果 / 缺漏原因` | 可產出或失敗的具體原因 |

即使資料有錯誤也不從畫面隱藏，例如：

```text
分校 3104 | ❌ | Voice 2 IP 欄位缺失
分校 2179 | ❌ | 分校代碼重複，出現於 Excel 第 38, 40 列
分校 0066 | ✅ | 可產出
```

---

## 12. 分校代碼搜尋

網頁提供「查詢分校代碼」文字欄位。

使用者可以輸入：

```text
66
0066
3803
```

搜尋使用目前畫面顯示的正規化代碼進行包含比對。

留空時顯示全部資料。

---

## 13. 指定分校 / 單檔產出

### 13.1 勾選

使用者可在資料表的 `選取`欄勾選需要產出的分校。

若勾選的資料有缺漏：

- 網頁立即顯示該分校不能產出的原因。
- 不產生錯誤 Config。

### 13.2 只選一個合法分校

直接提供單一 `.conf`下載，例如：

```text
3803.conf
```

或：

```text
0066_noHN.conf
```

### 13.3 選擇多個合法分校

打包為：

```text
HESS_Selected_Configs.zip
```

ZIP 內包含選取且合法的 Config，以及：

```text
Execution_Summary.txt
```

---

## 14. 整批產出

原本的整批產出功能保留。

檔名：

```text
HESS_Configs_Batch.zip
```

有問題的資料不產出 Config，但會寫入 `Execution_Summary.txt`。

---

## 15. Execution_Summary.txt

使用 UTF-8 BOM，方便 Windows 記事本直接顯示中文。

內容包含：

- 執行時間。
- 產出模式：整批或指定分校。
- Excel 工作表。
- Config 解碼方式。
- 總資料列。
- 成功、失敗、跳過數量。
- 預留字串掃描結果。
- 成功分校、檔名、HN 狀態。
- 失敗分校、Excel 列號、完整原因。

---

## 16. 整批停止條件

以下問題停止整批系統，不進入分校清單：

- Excel 無法讀取。
- 找不到 `總表`。
- 缺少必要 Excel 欄位。
- Config 為空。
- Config 無法解碼。
- Config 缺少必要預留字串。

---

## 17. 單筆停止條件

以下問題只影響該分校：

- 分校代碼不符合四碼，且不是可自動補零的兩碼。
- 正規化後分校代碼重複。
- IP 缺失或格式錯誤。
- IP 超過兩行。
- `.AAA.BBB.` 加 100 溢位。
- Voice 2 IP 缺失或格式錯誤。
- SPOKE IP 缺失或格式錯誤。
- SPOKE +1 跨原 `/24`。
- Config 替換後仍殘留必要預留字串。

**無 HN 不屬於停止條件。**

---

## 18. GitHub / Streamlit 部署

建議 Repository：

```text
HESS_CONFIG_FACTORY_v0.10.py
REQUIREMENTS_v0.10.md
requirements.txt
```

`requirements.txt`：

```text
streamlit
pandas
openpyxl
```

Streamlit Community Cloud Main file path：

```text
HESS_CONFIG_FACTORY_v0.10.py
```

正式 Excel、正式 Config、產出的 ZIP 與客戶資料不應提交至公開 Repository。
