import io
import ipaddress
import re
import zipfile
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st


APP_VERSION = "v0.9.1"
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
    errors = []

    for encoding in CONFIG_ENCODINGS:
        try:
            return file_bytes.decode(encoding), encoding
        except UnicodeDecodeError as exc:
            errors.append(f"{encoding}: {exc}")

    raise ValueError(
        "Config 範本無法解碼。已嘗試："
        + "、".join(CONFIG_ENCODINGS)
        + "。"
    )


def normalize_site_id(value):
    """
    將 Excel 內的 3803、3803.0 轉成 3803。
    正式分校代碼必須為恰好四位數字。
    """
    if is_blank(value):
        raise ValueError("分校代碼缺失")

    raw = str(value).strip()

    if re.fullmatch(r"\d+\.0+", raw):
        raw = raw.split(".", 1)[0]

    if not re.fullmatch(r"\d{4}", raw):
        raise ValueError(f"分校代碼必須為四位數字，原始內容：{raw}")

    return raw


def normalize_hn_number(value):
    """
    客戶號碼規則：
    1. HN + 純數字：移除 HN，只保留數字部分。
    2. 純數字：直接使用數字部分。
    3. 空白、xxxxxxxx 或其他格式：視為無 HN 帳號。
       Config 仍產出，但保留 HN_NUMBER，檔名加 _noHN。

    範例：
    - HN78211899 -> 78211899
    - hn78211899 -> 78211899
    - 78211899   -> 78211899
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
    """
    解析 Excel 中的 IPv4/CIDR 或 x 樣板，例如：
    10.38.3.x/24
    10.38.3.254/24
    10.38.3.x
    10.38.3.254
    """
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
            raise ValueError(
                f"{field_name} 第 {position} 段不是數字"
                + ("或 x" if position == 4 and allow_x else "")
                + f"，原始內容：{line}"
            )

        number = int(part)
        if not 0 <= number <= 255:
            raise ValueError(
                f"{field_name} 第 {position} 段超出 0-255，原始內容：{line}"
            )

        parsed.append(number)

    return parsed


def parse_site_ip_field(value):
    """
    IP 欄位規格：
    - 必須有 1 或 2 行。
    - 第一行產生 .XXX.YYY.。
    - 第二行存在時產生 .AAA.BBB.。
    - 第二行不存在時，第一行第二段 +100。
    - 不允許 +100 後超過 255。
    - 兩行的第一段與第三段必須一致，避免不同分校網段誤配。
    """
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
    """
    Voice 2 IP 必須是合法 IPv4/CIDR。
    Config 只取前三段，形成 CCC.DDD.EEE.。
    """
    if is_blank(value):
        raise ValueError("Voice 2 IP 欄位缺失")

    raw = str(value).strip()
    parts = parse_ip_pattern_line(raw, "Voice 2 IP", allow_x=False)
    return f"{parts[0]}.{parts[1]}.{parts[2]}."


def parse_spoke_ip(value):
    """
    SPOKE 必須是合法 IPv4。
    +1 僅允許最後一段加一，不允許跨越原本 /24 網段。
    """
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
    """計算範本預留字串出現次數，避免 +1 與原始 SPOKE 字串重疊誤判。"""
    counts = {}

    for placeholder in REQUIRED_PLACEHOLDERS:
        counts[placeholder] = template_content.count(placeholder)

    counts["KKK.LLL.MMM.NNN（獨立）"] = (
        counts["KKK.LLL.MMM.NNN"]
        - counts["KKK.LLL.MMM.NNN+1"]
    )

    return counts


def validate_template(template_content):
    """必要預留字串缺失時，停止整批處理。"""
    missing = [
        placeholder
        for placeholder in REQUIRED_PLACEHOLDERS
        if placeholder not in template_content
    ]

    if missing:
        raise ValueError(
            "Config 範本缺少必要預留字串："
            + "、".join(missing)
        )


def get_duplicate_site_ids(df):
    """
    只針對可被正規化為四碼的分校代碼檢查重複。
    所有重複列均不產出，避免 ZIP 內產生同名 Config。
    """
    normalized = []

    for value in df["分校代碼"]:
        try:
            normalized.append(normalize_site_id(value))
        except ValueError:
            normalized.append(None)

    series = pd.Series(normalized, index=df.index)
    duplicate_mask = series.notna() & series.duplicated(keep=False)

    duplicate_rows = {}
    for site_id in sorted(series[duplicate_mask].unique()):
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
    """依照既定順序執行全域字串替換。"""
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
            "替換完成後仍殘留必要預留字串："
            + "、".join(remaining)
        )

    return config


def build_summary(
    total_rows,
    success_items,
    error_items,
    skipped_blank_rows,
    encoding_used,
    placeholder_counts,
):
    """產生新版 Execution_Summary.txt。"""
    timestamp = datetime.now(ZoneInfo("Asia/Taipei")).strftime(
        "%Y-%m-%d %H:%M:%S %Z"
    )

    lines = [
        "HESS FortiGate Config 批次產生結果",
        "=" * 42,
        "",
        f"執行時間：{timestamp}",
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
                    f"Excel 第 {item['excel_row']} 列｜"
                    f"分校 {item['site_id']}",
                    f"錯誤類型：{item['error']}",
                    f"處理結果：未產出 Config",
                ]
            )
    else:
        lines.append("無")

    return "\n".join(lines) + "\n"


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
            column
            for column in REQUIRED_COLUMNS
            if column not in df.columns
        ]

        if missing_columns:
            raise ValueError(
                "Excel 缺少必要欄位："
                + "、".join(missing_columns)
            )

        template_bytes = uploaded_template.getvalue()

        if not template_bytes:
            raise ValueError("Config 範本為空檔案")

        template_content, encoding_used = decode_config(template_bytes)
        validate_template(template_content)
        placeholder_counts = count_placeholders(template_content)

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
        st.dataframe(placeholder_df, use_container_width=True)

        duplicate_site_ids = get_duplicate_site_ids(df)

        if duplicate_site_ids:
            duplicate_text = "；".join(
                f"{site_id}（Excel 第 {', '.join(map(str, rows))} 列）"
                for site_id, rows in duplicate_site_ids.items()
            )
            st.warning(
                "偵測到重複分校代碼，相關列將全部停止產出："
                + duplicate_text
            )

        st.subheader("3. 批次產生")

        if st.button(
            "🚀 開始批次生成並打包下載",
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
                for index, row in df.iterrows():
                    excel_row = int(index) + 2
                    site_id_raw = row.get("分校代碼", "")

                    if is_blank(site_id_raw) or str(site_id_raw).strip() == "0":
                        skipped_blank_rows += 1
                        continue

                    site_id_for_log = str(site_id_raw).strip()

                    try:
                        site_id = normalize_site_id(site_id_raw)
                        site_id_for_log = site_id

                        if site_id in duplicate_site_ids:
                            rows = ", ".join(
                                str(row_number)
                                for row_number in duplicate_site_ids[site_id]
                            )
                            raise ValueError(
                                f"分校代碼重複，出現於 Excel 第 {rows} 列"
                            )

                        xxx_yyy, aaa_bbb = parse_site_ip_field(
                            row.get("IP", "")
                        )
                        voice_prefix = parse_voice_ip(
                            row.get("Voice 2 IP", "")
                        )
                        spoke_ip, spoke_ip_plus_1 = parse_spoke_ip(
                            row.get("HQ_IPSEC_IP(SPOKE)", "")
                        )
                        hn_number = normalize_hn_number(
                            row.get("客戶號碼", "")
                        )

                        config = replace_template(
                            template_content=template_content,
                            site_id=site_id,
                            hn_number=hn_number,
                            xxx_yyy=xxx_yyy,
                            aaa_bbb=aaa_bbb,
                            voice_prefix=voice_prefix,
                            spoke_ip=spoke_ip,
                            spoke_ip_plus_1=spoke_ip_plus_1,
                        )

                        filename = (
                            f"{site_id}.conf"
                            if hn_number
                            else f"{site_id}_noHN.conf"
                        )

                        zip_file.writestr(
                            filename,
                            config.encode("utf-8"),
                        )

                        success_items.append(
                            {
                                "excel_row": excel_row,
                                "site_id": site_id,
                                "filename": filename,
                                "hn_number": hn_number,
                            }
                        )

                    except Exception as exc:
                        error_items.append(
                            {
                                "excel_row": excel_row,
                                "site_id": site_id_for_log,
                                "error": str(exc),
                            }
                        )

                summary_content = build_summary(
                    total_rows=len(df),
                    success_items=success_items,
                    error_items=error_items,
                    skipped_blank_rows=skipped_blank_rows,
                    encoding_used=encoding_used,
                    placeholder_counts=placeholder_counts,
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
            )

            if error_items:
                st.warning(
                    "部分資料未產出，請查看 ZIP 內的 "
                    "Execution_Summary.txt。"
                )

    except Exception as exc:
        st.error(f"檔案驗證失敗：{exc}")

else:
    st.info(
        "💡 請上傳 Excel 總表與 Config 範本以開始。"
    )
