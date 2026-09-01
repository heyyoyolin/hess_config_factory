import io
import ipaddress
import re
import zipfile
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st


APP_VERSION = "v0.10"
SHEET_NAME = "總表"

REQUIRED_COLUMNS = [
    "分校代碼",
    "IP",
    "Voice 2 IP",
    "HQ_IPSEC_IP(SPOKE)",
    "客戶號碼",
]

REQUIRED_PLACEHOLDERS = [
    "HESSXXXX",
    "HN_NUMBER",
    ".XXX.YYY.",
    ".AAA.BBB.",
    "CCC.DDD.EEE.",
    "KKK.LLL.MMM.NNN",
    "KKK.LLL.MMM.NNN+1",
]

CONFIG_ENCODINGS = ["utf-8-sig", "utf-8", "cp950", "big5"]


st.set_page_config(
    page_title="HESS FortiGate 自動化配置系統",
    layout="wide",
)
st.title(f"🏫 HESS FortiGate Config 批次生成工具 ({APP_VERSION})")


def is_valid_ip(ip_str):
    """保留舊版函式；依需求不變更既有用途。"""
    try:
        clean_ip = ip_str.split("/")[0].strip()
        ipaddress.IPv4Address(clean_ip)
        return True
    except Exception:
        return False


def is_blank(value):
    """判斷 Excel 儲存格是否為空白、NaN 或僅包含空白字元。"""
    return pd.isna(value) or str(value).strip() == ""


def decode_config(file_bytes):
    """依序嘗試常見 Config 編碼，回傳內容與成功使用的編碼。"""
    for encoding in CONFIG_ENCODINGS:
        try:
            return file_bytes.decode(encoding), encoding
        except UnicodeDecodeError:
            continue

    raise ValueError(
        "Config 範本無法解碼。已嘗試："
        + "、".join(CONFIG_ENCODINGS)
        + "。"
    )


def normalize_site_id(value):
    """
    分校代碼規則：
    - 3803、3803.0 -> 3803
    - 兩碼純數字，例如 66 -> 0066
    - 四碼純數字直接使用
    - 其他長度或格式視為異常
    """
    if is_blank(value):
        raise ValueError("分校代碼缺失")

    raw = str(value).strip()

    if re.fullmatch(r"\d+\.0+", raw):
        raw = raw.split(".", 1)[0]

    if re.fullmatch(r"\d{2}", raw):
        return raw.zfill(4)

    if re.fullmatch(r"\d{4}", raw):
        return raw

    raise ValueError(
        "分校代碼必須為四碼數字，或可自動補零的兩碼數字，"
        f"原始內容：{raw}"
    )


def normalize_hn_number(value):
    """
    客戶號碼規則：
    - HN78211899 -> 78211899
    - hn78211899 -> 78211899
    - 78211899   -> 78211899
    - 空白、xxxxxxxx 或其他格式 -> 無 HN

    Config 的 HN_NUMBER 僅替換為數字部分。
    無 HN 時 Config 仍產出，保留 HN_NUMBER，檔名加 _noHN。
    """
    if is_blank(value):
        return None

    raw = str(value).strip().upper().replace(" ", "")

    if re.fullmatch(r"HN\d+", raw):
        return raw[2:]

    if re.fullmatch(r"\d+", raw):
        return raw

    return None


def split_nonempty_lines(value):
    """統一處理 Windows、Linux、舊 Mac 換行並移除空白行。"""
    if is_blank(value):
        return []

    normalized = str(value).replace("\r\n", "\n").replace("\r", "\n")
    return [line.strip() for line in normalized.split("\n") if line.strip()]


def parse_ip_pattern_line(raw_line, field_name, allow_x=True):
    """解析 IPv4/CIDR 或第四段為 x 的樣板。"""
    line = str(raw_line).strip()

    if line.count("/") > 1:
        raise ValueError(f"{field_name} CIDR 格式錯誤：{line}")

    address_part, _, prefix_part = line.partition("/")

    if prefix_part:
        if not prefix_part.isdigit() or not 0 <= int(prefix_part) <= 32:
            raise ValueError(f"{field_name} CIDR 前綴不合法：{line}")

    parts = [part.strip() for part in address_part.split(".")]

    if len(parts) != 4:
        raise ValueError(f"{field_name} 必須為四段 IPv4，原始內容：{line}")

    parsed = []
    for position, part in enumerate(parts, start=1):
        if position == 4 and allow_x and part.lower() == "x":
            parsed.append("x")
            continue

        if not part.isdigit():
            allowed = "或 x" if position == 4 and allow_x else ""
            raise ValueError(
                f"{field_name} 第 {position} 段不是數字{allowed}，"
                f"原始內容：{line}"
            )

        number = int(part)
        if not 0 <= number <= 255:
            raise ValueError(
                f"{field_name} 第 {position} 段超出 0-255，原始內容：{line}"
            )

        parsed.append(number)

    return parsed


def parse_site_ip_field(value):
    """解析 SITE / Mobile 所需網段。"""
    lines = split_nonempty_lines(value)

    if not lines:
        raise ValueError("IP 欄位缺失")

    if len(lines) > 2:
        raise ValueError(
            f"IP 欄位最多只能有兩行，目前有 {len(lines)} 行：{lines}"
        )

    first = parse_ip_pattern_line(lines[0], "IP 第一行", allow_x=True)
    xxx_yyy = f".{first[1]}.{first[2]}."

    if len(lines) == 2:
        second = parse_ip_pattern_line(lines[1], "IP 第二行", allow_x=True)

        if first[0] != second[0] or first[2] != second[2]:
            raise ValueError(
                "IP 第一、二行的第一段與第三段必須一致，"
                f"第一行：{lines[0]}，第二行：{lines[1]}"
            )

        aaa_bbb = f".{second[1]}.{second[2]}."
    else:
        calculated_second_octet = first[1] + 100

        if calculated_second_octet > 255:
            raise ValueError(
                ".AAA.BBB. 計算溢位："
                f"IP 第二段 {first[1]} + 100 = {calculated_second_octet}，"
                "不產出 Config"
            )

        aaa_bbb = f".{calculated_second_octet}.{first[2]}."

    return xxx_yyy, aaa_bbb


def parse_voice_ip(value):
    """Voice 2 IP 必須是合法 IPv4/CIDR，Config 只取前三段。"""
    if is_blank(value):
        raise ValueError("Voice 2 IP 欄位缺失")

    raw = str(value).strip()
    parts = parse_ip_pattern_line(raw, "Voice 2 IP", allow_x=False)
    return f"{parts[0]}.{parts[1]}.{parts[2]}."


def parse_spoke_ip(value):
    """SPOKE +1 僅允許最後一段加一，不允許跨越原 /24。"""
    if is_blank(value):
        raise ValueError("HQ_IPSEC_IP(SPOKE) 欄位缺失")

    raw = str(value).strip()

    if "/" in raw:
        raw = raw.split("/", 1)[0].strip()

    try:
        spoke_ip = ipaddress.IPv4Address(raw)
    except ipaddress.AddressValueError as exc:
        raise ValueError(
            f"HQ_IPSEC_IP(SPOKE) 格式錯誤：{value}"
        ) from exc

    octets = [int(part) for part in str(spoke_ip).split(".")]

    if octets[3] == 255:
        next_ip = spoke_ip + 1
        raise ValueError(
            "VPN IP +1 會跨網段："
            f"原始 IP {spoke_ip}，計算結果 {next_ip}，"
            "依規格不產出 Config"
        )

    octets[3] += 1
    next_ip = ".".join(str(part) for part in octets)

    return str(spoke_ip), next_ip


def count_placeholders(template_content):
    """計算範本預留字串出現次數。"""
    counts = {
        placeholder: template_content.count(placeholder)
        for placeholder in REQUIRED_PLACEHOLDERS
    }
    counts["KKK.LLL.MMM.NNN（獨立）"] = (
        counts["KKK.LLL.MMM.NNN"]
        - counts["KKK.LLL.MMM.NNN+1"]
    )
    return counts


def validate_template(template_content):
    """必要預留字串缺失時停止整批處理。"""
    missing = [
        placeholder
        for placeholder in REQUIRED_PLACEHOLDERS
        if placeholder not in template_content
    ]

    if missing:
        raise ValueError(
            "Config 範本缺少必要預留字串：" + "、".join(missing)
        )


def get_duplicate_site_ids(df):
    """依正規化後的四碼分校代碼檢查重複。"""
    normalized = []

    for value in df["分校代碼"]:
        try:
            normalized.append(normalize_site_id(value))
        except ValueError:
            normalized.append(None)

    series = pd.Series(normalized, index=df.index, dtype="object")
    duplicate_mask = series.notna() & series.duplicated(keep=False)

    duplicate_rows = {}
    for site_id in sorted(series[duplicate_mask].dropna().unique()):
        duplicate_rows[site_id] = [
            int(index) + 2
            for index in series[series == site_id].index
        ]

    return duplicate_rows


def replace_template(
    template_content,
    site_id,
    hn_number,
    xxx_yyy,
    aaa_bbb,
    voice_prefix,
    spoke_ip,
    spoke_ip_plus_1,
):
    """依既定順序執行 Config 全域字串替換。"""
    config = template_content

    replacements = [
        ("HESSXXXX", f"HESS{site_id}"),
        (".XXX.YYY.", xxx_yyy),
        (".AAA.BBB.", aaa_bbb),
        ("CCC.DDD.EEE.", voice_prefix),
        ("KKK.LLL.MMM.NNN+1", spoke_ip_plus_1),
        ("KKK.LLL.MMM.NNN", spoke_ip),
    ]

    if hn_number:
        replacements.append(("HN_NUMBER", hn_number))

    for source, target in replacements:
        config = config.replace(source, target)

    # 無 HN 時允許 HN_NUMBER 保留，供人工補登。
    forbidden_remaining = [
        "HESSXXXX",
        ".XXX.YYY.",
        ".AAA.BBB.",
        "CCC.DDD.EEE.",
        "KKK.LLL.MMM.NNN+1",
        "KKK.LLL.MMM.NNN",
    ]

    if hn_number:
        forbidden_remaining.append("HN_NUMBER")

    remaining = [
        placeholder
        for placeholder in forbidden_remaining
        if placeholder in config
    ]

    if remaining:
        raise ValueError(
            "替換完成後仍殘留必要預留字串：" + "、".join(remaining)
        )

    return config


def evaluate_row(row, index, duplicate_site_ids, template_content=None):
    """
    預先檢查單筆資料。
    回傳網頁清單與實際產檔都能共用的結果，避免畫面與產出規則不一致。
    """
    excel_row = int(index) + 2
    site_id_raw = row.get("分校代碼", "")
    unit_name = "" if is_blank(row.get("單位名稱", "")) else str(row.get("單位名稱", "")).strip()

    result = {
        "source_index": index,
        "excel_row": excel_row,
        "site_id_raw": "" if is_blank(site_id_raw) else str(site_id_raw).strip(),
        "site_id": "",
        "unit_name": unit_name,
        "hn_number": None,
        "hn_status": "",
        "filename": "",
        "can_generate": False,
        "reason": "",
        "config": None,
    }

    if is_blank(site_id_raw) or str(site_id_raw).strip() == "0":
        result["reason"] = "分校代碼缺失或為 0"
        return result

    try:
        site_id = normalize_site_id(site_id_raw)
        result["site_id"] = site_id

        if site_id in duplicate_site_ids:
            rows = ", ".join(str(n) for n in duplicate_site_ids[site_id])
            raise ValueError(f"分校代碼重複，出現於 Excel 第 {rows} 列")

        xxx_yyy, aaa_bbb = parse_site_ip_field(row.get("IP", ""))
        voice_prefix = parse_voice_ip(row.get("Voice 2 IP", ""))
        spoke_ip, spoke_ip_plus_1 = parse_spoke_ip(
            row.get("HQ_IPSEC_IP(SPOKE)", "")
        )
        hn_number = normalize_hn_number(row.get("客戶號碼", ""))

        result["hn_number"] = hn_number
        result["hn_status"] = (
            f"有 HN：{hn_number}"
            if hn_number
            else "無 HN：Config 保留 HN_NUMBER"
        )
        result["filename"] = (
            f"{site_id}.conf" if hn_number else f"{site_id}_noHN.conf"
        )

        if template_content is not None:
            result["config"] = replace_template(
                template_content=template_content,
                site_id=site_id,
                hn_number=hn_number,
                xxx_yyy=xxx_yyy,
                aaa_bbb=aaa_bbb,
                voice_prefix=voice_prefix,
                spoke_ip=spoke_ip,
                spoke_ip_plus_1=spoke_ip_plus_1,
            )

        result["can_generate"] = True
        result["reason"] = "可產出"

    except Exception as exc:
        result["reason"] = str(exc)
        if not result["hn_status"]:
            result["hn_status"] = "未檢查或無法判定"

    return result


def build_summary(
    total_rows,
    success_items,
    error_items,
    skipped_blank_rows,
    encoding_used,
    placeholder_counts,
    mode="整批產出",
):
    """產生 Execution_Summary.txt。"""
    timestamp = datetime.now(ZoneInfo("Asia/Taipei")).strftime(
        "%Y-%m-%d %H:%M:%S %Z"
    )

    lines = [
        "HESS FortiGate Config 產生結果",
        "=" * 42,
        "",
        f"執行時間：{timestamp}",
        f"產出模式：{mode}",
        f"Excel 工作表：{SHEET_NAME}",
        f"Config 解碼方式：{encoding_used}",
        "",
        f"Excel 總資料列：{total_rows}",
        f"成功產出：{len(success_items)}",
        f"失敗：{len(error_items)}",
        f"跳過空白列：{skipped_blank_rows}",
        "",
        "【Config 預留字串掃描】",
    ]

    for placeholder, count in placeholder_counts.items():
        lines.append(f"{placeholder}：{count} 處")

    lines.extend(["", "【成功清單】"])

    if success_items:
        for item in success_items:
            hn_status = (
                f"客戶號碼：{item['hn_number']}"
                if item["hn_number"]
                else "HN：缺失，Config 保留 HN_NUMBER"
            )
            lines.append(
                f"Excel 第 {item['excel_row']} 列｜"
                f"分校 {item['site_id']}｜"
                f"{item['filename']}｜{hn_status}"
            )
    else:
        lines.append("無")

    lines.extend(["", "【失敗清單】"])

    if error_items:
        for item in error_items:
            lines.extend(
                [
                    "",
                    f"Excel 第 {item['excel_row']} 列｜分校 {item['site_id']}",
                    f"錯誤類型：{item['error']}",
                    "處理結果：未產出 Config",
                ]
            )
    else:
        lines.append("無")

    return "\n".join(lines) + "\n"


def build_preview_dataframe(df, duplicate_site_ids, template_content):
    """建立所有分校的網頁檢查清單。"""
    records = []
    results_by_index = {}

    for index, row in df.iterrows():
        result = evaluate_row(
            row=row,
            index=index,
            duplicate_site_ids=duplicate_site_ids,
            template_content=template_content,
        )
        results_by_index[index] = result

        display_site_id = result["site_id"] or result["site_id_raw"] or "（空白）"
        records.append(
            {
                "選取": False,
                "分校代碼": display_site_id,
                "單位名稱": result["unit_name"],
                "Excel列號": result["excel_row"],
                "HN狀態": result["hn_status"],
                "預計檔名": result["filename"],
                "可產出": "✅" if result["can_generate"] else "❌",
                "檢查結果 / 缺漏原因": result["reason"],
                "_source_index": index,
            }
        )

    return pd.DataFrame(records), results_by_index


st.header("1. 上傳必要檔案")
col1, col2 = st.columns(2)

with col1:
    uploaded_excel = st.file_uploader(
        "上傳何嘉仁專案 Excel (xlsx)",
        type=["xlsx"],
    )

with col2:
    uploaded_template = st.file_uploader(
        "上傳 Config 範本 (.conf 或 .txt)",
        type=["conf", "txt"],
    )


if uploaded_excel and uploaded_template:
    try:
        excel_file = pd.ExcelFile(uploaded_excel)

        if SHEET_NAME not in excel_file.sheet_names:
            raise ValueError(
                f"Excel 找不到工作表「{SHEET_NAME}」。"
                f"目前工作表：{', '.join(excel_file.sheet_names)}"
            )

        df = pd.read_excel(excel_file, sheet_name=SHEET_NAME)

        missing_columns = [
            column for column in REQUIRED_COLUMNS if column not in df.columns
        ]

        if missing_columns:
            raise ValueError("Excel 缺少必要欄位：" + "、".join(missing_columns))

        template_bytes = uploaded_template.getvalue()
        if not template_bytes:
            raise ValueError("Config 範本為空檔案")

        template_content, encoding_used = decode_config(template_bytes)
        validate_template(template_content)
        placeholder_counts = count_placeholders(template_content)
        duplicate_site_ids = get_duplicate_site_ids(df)

        st.success(
            f"檔案驗證成功。Excel 工作表：{SHEET_NAME}；"
            f"Config 編碼：{encoding_used}"
        )

        st.subheader("2. Config 預留字串掃描")
        placeholder_df = pd.DataFrame(
            [
                {"預留字串": key, "出現次數": value}
                for key, value in placeholder_counts.items()
            ]
        )
        st.dataframe(placeholder_df, use_container_width=True, hide_index=True)

        if duplicate_site_ids:
            duplicate_text = "；".join(
                f"{site_id}（Excel 第 {', '.join(map(str, rows))} 列）"
                for site_id, rows in duplicate_site_ids.items()
            )
            st.warning(
                "偵測到重複分校代碼，相關列將全部停止產出：" + duplicate_text
            )

        preview_df, row_results = build_preview_dataframe(
            df=df,
            duplicate_site_ids=duplicate_site_ids,
            template_content=template_content,
        )

        valid_count = sum(1 for item in row_results.values() if item["can_generate"])
        invalid_count = len(row_results) - valid_count

        st.subheader("3. 分校資料檢查與指定產出")
        metric1, metric2, metric3 = st.columns(3)
        metric1.metric("總資料列", len(preview_df))
        metric2.metric("可產出", valid_count)
        metric3.metric("有問題 / 缺漏", invalid_count)

        search_text = st.text_input(
            "🔎 查詢分校代碼",
            placeholder="例如：66、0066、3803；留空顯示全部",
        ).strip()

        filtered_preview = preview_df.copy()
        if search_text:
            filtered_preview = filtered_preview[
                filtered_preview["分校代碼"].astype(str).str.contains(
                    search_text,
                    case=False,
                    na=False,
                    regex=False,
                )
            ].copy()

        if filtered_preview.empty:
            st.info("查無符合的分校代碼。")
        else:
            editor_input = filtered_preview.drop(columns=["_source_index"])
            edited_df = st.data_editor(
                editor_input,
                use_container_width=True,
                hide_index=True,
                disabled=[
                    "分校代碼",
                    "單位名稱",
                    "Excel列號",
                    "HN狀態",
                    "預計檔名",
                    "可產出",
                    "檢查結果 / 缺漏原因",
                ],
                column_config={
                    "選取": st.column_config.CheckboxColumn(
                        "選取",
                        help="勾選要指定產出的分校",
                        default=False,
                    ),
                    "檢查結果 / 缺漏原因": st.column_config.TextColumn(
                        "檢查結果 / 缺漏原因",
                        width="large",
                    ),
                },
                key=f"site_selector_{search_text}",
            )

            selected_rows = edited_df[edited_df["選取"]]
            selected_excel_rows = set(selected_rows["Excel列號"].tolist())
            selected_results = [
                result
                for result in row_results.values()
                if result["excel_row"] in selected_excel_rows
            ]
            selected_valid = [r for r in selected_results if r["can_generate"]]
            selected_invalid = [r for r in selected_results if not r["can_generate"]]

            if selected_results:
                st.caption(
                    f"已選取 {len(selected_results)} 筆："
                    f"可產出 {len(selected_valid)} 筆，"
                    f"不可產出 {len(selected_invalid)} 筆。"
                )

            if selected_invalid:
                with st.expander("⚠️ 已勾選但無法產出的分校", expanded=True):
                    for result in selected_invalid:
                        display_id = result["site_id"] or result["site_id_raw"] or "（空白）"
                        st.write(
                            f"- Excel 第 {result['excel_row']} 列｜"
                            f"分校 {display_id}：{result['reason']}"
                        )

            if len(selected_valid) == 1:
                selected = selected_valid[0]
                st.download_button(
                    label=f"💾 下載單一 Config：{selected['filename']}",
                    data=selected["config"].encode("utf-8"),
                    file_name=selected["filename"],
                    mime="text/plain",
                    key="download_single_config",
                )

            elif len(selected_valid) > 1:
                selected_zip = io.BytesIO()
                selected_success_items = []
                selected_error_items = []

                with zipfile.ZipFile(
                    selected_zip,
                    mode="w",
                    compression=zipfile.ZIP_DEFLATED,
                ) as zip_file:
                    for result in selected_valid:
                        zip_file.writestr(
                            result["filename"],
                            result["config"].encode("utf-8"),
                        )
                        selected_success_items.append(
                            {
                                "excel_row": result["excel_row"],
                                "site_id": result["site_id"],
                                "filename": result["filename"],
                                "hn_number": result["hn_number"],
                            }
                        )

                    for result in selected_invalid:
                        selected_error_items.append(
                            {
                                "excel_row": result["excel_row"],
                                "site_id": result["site_id"] or result["site_id_raw"],
                                "error": result["reason"],
                            }
                        )

                    summary = build_summary(
                        total_rows=len(selected_results),
                        success_items=selected_success_items,
                        error_items=selected_error_items,
                        skipped_blank_rows=0,
                        encoding_used=encoding_used,
                        placeholder_counts=placeholder_counts,
                        mode="指定分校產出",
                    )
                    zip_file.writestr(
                        "Execution_Summary.txt",
                        summary.encode("utf-8-sig"),
                    )

                st.download_button(
                    label=f"💾 下載已勾選的 {len(selected_valid)} 份 Config (.zip)",
                    data=selected_zip.getvalue(),
                    file_name="HESS_Selected_Configs.zip",
                    mime="application/zip",
                    key="download_selected_configs",
                )

        st.divider()
        st.subheader("4. 整批產出")
        st.caption(
            "整批模式會處理 Excel 全部資料；有問題的分校不產出，"
            "並寫入 Execution_Summary.txt。"
        )

        if st.button(
            "🚀 開始整批生成並打包下載",
            type="primary",
        ):
            zip_buffer = io.BytesIO()
            success_items = []
            error_items = []
            skipped_blank_rows = 0

            with zipfile.ZipFile(
                zip_buffer,
                mode="w",
                compression=zipfile.ZIP_DEFLATED,
            ) as zip_file:
                for result in row_results.values():
                    if not result["site_id_raw"] or result["site_id_raw"] == "0":
                        skipped_blank_rows += 1
                        continue

                    if result["can_generate"]:
                        zip_file.writestr(
                            result["filename"],
                            result["config"].encode("utf-8"),
                        )
                        success_items.append(
                            {
                                "excel_row": result["excel_row"],
                                "site_id": result["site_id"],
                                "filename": result["filename"],
                                "hn_number": result["hn_number"],
                            }
                        )
                    else:
                        error_items.append(
                            {
                                "excel_row": result["excel_row"],
                                "site_id": result["site_id"] or result["site_id_raw"],
                                "error": result["reason"],
                            }
                        )

                summary_content = build_summary(
                    total_rows=len(df),
                    success_items=success_items,
                    error_items=error_items,
                    skipped_blank_rows=skipped_blank_rows,
                    encoding_used=encoding_used,
                    placeholder_counts=placeholder_counts,
                    mode="整批產出",
                )

                zip_file.writestr(
                    "Execution_Summary.txt",
                    summary_content.encode("utf-8-sig"),
                )

            st.success(
                f"處理完畢：成功 {len(success_items)} 份，"
                f"失敗 {len(error_items)} 筆，"
                f"跳過空白列 {skipped_blank_rows} 筆。"
            )

            st.download_button(
                label="💾 下載 HESS 設定檔壓縮包 (.zip)",
                data=zip_buffer.getvalue(),
                file_name="HESS_Configs_Batch.zip",
                mime="application/zip",
                key="download_batch_configs",
            )

            if error_items:
                st.warning(
                    "部分資料未產出，請查看 ZIP 內的 Execution_Summary.txt。"
                )

    except Exception as exc:
        st.error(f"檔案驗證失敗：{exc}")

else:
    st.info("💡 請上傳 Excel 總表與 Config 範本以開始。")
