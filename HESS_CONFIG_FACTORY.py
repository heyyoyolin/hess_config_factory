import streamlit as st
import pandas as pd
import ipaddress
import zipfile
import io
import os

# --- 設定網頁標題 ---
st.set_page_config(page_title="HESS FortiGate 自動化配置系統", layout="wide")
st.title("🏫 HESS FortiGate Config 批次生成工具 (v0.8)")

# --- 1. 檔案上傳區 ---
st.header("1. 上傳必要檔案")
col1, col2 = st.columns(2)

with col1:
    uploaded_excel = st.file_uploader("上傳何嘉仁專案 Excel (xlsx)", type=["xlsx"])
with col2:
    uploaded_template = st.file_uploader("上傳 Config 範本 (.conf 或 .txt)", type=["conf", "txt", "conf.txt"])

# --- 輔助函式：IP 驗證 ---
def is_valid_ip(ip_str):
    try:
        clean_ip = ip_str.split('/')[0].strip()
        ipaddress.IPv4Address(clean_ip)
        return True
    except:
        return False

# --- 2. 核心處理邏輯 ---
if uploaded_excel and uploaded_template:
    # 讀取 Excel 總表
    df = pd.read_excel(uploaded_excel, sheet_name='總表')
    
    # 讀取範本內容
    template_content = uploaded_template.read().decode("utf-8")
    
    st.success("檔案上傳成功！準備處理資料...")
    
    if st.button("🚀 開始批次生成並打包下載"):
        zip_buffer = io.BytesIO()
        logs = []
        success_count = 0
        
        # 建立記憶體內的 ZIP 檔
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for index, row in df.iterrows():
                site_id_raw = row.get('分校代碼', '')
                
                # 排除無效的分校代碼
                if pd.isna(site_id_raw) or str(site_id_raw).strip() == '' or site_id_raw == 0:
                    continue
                
                try:
                    # 格式化 Site ID (避免出現 .0)
                    site_id = str(int(float(site_id_raw)))
                    
                    # 讀取欄位資料
                    raw_ip = str(row.get('IP', ''))
                    raw_voice_ip = str(row.get('VoiceIP', ''))
                    raw_spoke_ip = str(row.get('HQ_IPSEC_IP(SPOKE)', ''))
                    
                    if 'nan' in raw_ip.lower() or 'nan' in raw_spoke_ip.lower():
                        raise ValueError("關鍵 IP 欄位資料缺失")

                    # --- [置換邏輯 1] 解析 IP 規則 ---
                    ip_lines = [line.strip() for line in raw_ip.split('\n') if line.strip()]
                    p1 = ip_lines[0].split('.')
                    xxx_yyy = f".{p1[1]}.{p1[2]}."
                    
                    # .AAA.BBB. 邏輯：如果有第二行取第二行，否則第一行首碼+100
                    if len(ip_lines) >= 2:
                        p2 = ip_lines[1].split('.')
                        aaa_bbb = f".{p2[1]}.{p2[2]}."
                    else:
                        new_val = int(p1[1]) + 100
                        if new_val > 255:
                            raise ValueError(f".AAA.BBB. 計算溢位 ({new_val})")
                        aaa_bbb = f".{new_val}.{p1[2]}."
                    
                    # --- [置換邏輯 2] Voice IP (CCC.DDD.EEE.) ---
                    voice_parts = raw_voice_ip.replace('/24','').strip().split('.')
                    ccc_ddd_eee = f"{voice_parts[0]}.{voice_parts[1]}.{voice_parts[2]}."
                    
                    # --- [置換邏輯 3] VPN IP (KKK.LLL.MMM.NNN) ---
                    kkk_lll_mmm_nnn = raw_spoke_ip.strip()
                    spoke_parts = kkk_lll_mmm_nnn.split('.')
                    last_octet = int(spoke_parts[3])
                    kkk_lll_mmm_nnn_plus_1 = f"{spoke_parts[0]}.{spoke_parts[1]}.{spoke_parts[2]}.{last_octet + 1}"
                    
                    # --- 執行字串置換 ---
                    config = template_content
                    config = config.replace("HESSXXXX", f"HESS{site_id}")
                    config = config.replace(".XXX.YYY.", xxx_yyy)
                    config = config.replace(".AAA.BBB.", aaa_bbb)
                    config = config.replace("CCC.DDD.EEE.", ccc_ddd_eee)
                    # 先換 +1 的，避免字串重疊
                    config = config.replace("KKK.LLL.MMM.NNN+1", kkk_lll_mmm_nnn_plus_1)
                    config = config.replace("KKK.LLL.MMM.NNN", kkk_lll_mmm_nnn)
                    
                    # 將生成內容寫入 ZIP (檔名：分校代碼.conf)
                    zip_file.writestr(f"{site_id}.conf", config)
                    success_count += 1

                except Exception as e:
                    logs.append(f"❌ 分校 {site_id_raw} 處理失敗: {str(e)}")

            # 額外加入執行紀錄檔
            summary_info = f"成功生成數量: {success_count}\n錯誤數量: {len(logs)}\n\n" + "\n".join(logs)
            zip_file.writestr("Execution_Summary.txt", summary_info)

        # --- 3. 提供下載 ---
        st.success(f"處理完畢！成功生成 {success_count} 份設定檔。")
        st.download_button(
            label="💾 下載 HESS 設定檔壓縮包 (.zip)",
            data=zip_buffer.getvalue(),
            file_name="HESS_Configs_Batch.zip",
            mime="application/zip"
        )
        
        if logs:
            st.warning("部分資料處理有誤，請查看下載包內的 Execution_Summary.txt")

else:
    st.info("💡 請上傳 Excel 總表與 .conf 範本檔案以開始。")