#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""批量读取 .eml 邮件文件（仅标准库，无第三方依赖）。

用法示例：
    # 读取一个文件夹（不递归），结果写到 ./eml_output
    python read_eml.py "D:\\mail"

    # 递归 + 导出附件 + 指定输出目录
    python read_eml.py "D:\\mail" --recursive --attachments -o "D:\\out"

    # 指定若干文件
    python read_eml.py a.eml b.eml

    # 强制全量重建（忽略上一次的处理基线）
    python read_eml.py "D:\\mail" --rebuild

输出：
    eml_result.xlsx  每封邮件一行：发件人、收件时间、邮件主题、正文内容、分类
    eml_summary.md   分类汇总：各类别封数统计 + 每封邮件要点
    .eml_index.json  界面索引（单文件 JSON，供工作台读正文/元信息/附件列表）
    会议日程.xlsx     会议类邮件整理成的日程表（9 列）
    attachments/     仅在加 --attachments 时生成，子目录名带收件日期
    .eml_state.json  处理基线（记录已处理的文件指纹），供下次增补时比对

    需要完整字段的清单时加 --full，此时 eml_result.xlsx 输出全部列。

增补行为（默认）：
    输出目录里已有基础清单 + 基线时，只解析新出现的邮件，把结果并入既有清单/汇总/
    会议日程（附件只导出新邮件的），旧邮件不重复解析。
    下列情况自动退回全量重建：首次运行（无基线）、基础清单缺失或列结构变化、
    有邮件被移出文件夹或被改动 —— 此时会先说明原因再全量重做。
    需要强制全量重建时加 --rebuild。

也可作为库调用：
    from read_eml import read_eml_file, read_eml_dir
    msg = read_eml_file("a.eml")
    print(msg["subject"], msg["body_text"])
"""

from __future__ import annotations

import argparse
import glob as globlib
import hashlib
import json
import mimetypes
import os
import re
import shutil
import sys
import zipfile
from datetime import datetime, timedelta, timezone
from email import policy
from email.parser import BytesHeaderParser, BytesParser
from email.utils import getaddresses, parsedate_to_datetime
from html.parser import HTMLParser
from pathlib import Path

try:  # 3.9+ 才有 zoneinfo，缺失时按固定偏移处理
    from zoneinfo import ZoneInfo
except Exception:  # pragma: no cover
    ZoneInfo = None

# ---------------------------------------------------------------- 解析基础

PARSER = BytesParser(policy=policy.default)

_BLOCK_TAGS = {
    "p", "div", "br", "tr", "td", "th", "li", "ul", "ol", "table", "hr",
    "h1", "h2", "h3", "h4", "h5", "h6", "blockquote", "pre",
}


class _HTMLToText(HTMLParser):
    """极简 HTML 转纯文本：去标签、保留块级换行。"""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._skip = 0
        self._parts: list[str] = []

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "title"):
            self._skip += 1
        elif tag in _BLOCK_TAGS:
            self._parts.append("\n")

    def handle_endtag(self, tag):
        if tag in ("script", "style", "title") and self._skip:
            self._skip -= 1
        elif tag in _BLOCK_TAGS:
            self._parts.append("\n")

    def handle_data(self, data):
        if not self._skip:
            self._parts.append(data)

    def text(self) -> str:
        return "".join(self._parts)


def html_to_text(html: str) -> str:
    parser = _HTMLToText()
    try:
        parser.feed(html)
        parser.close()
    except Exception:  # 残缺 HTML 不致命，返回已解析部分
        pass
    raw = parser.text()
    raw = raw.replace("\xa0", " ")
    raw = re.sub(r"[ \t]+", " ", raw)
    raw = re.sub(r"\n\s*\n\s*\n+", "\n\n", raw)
    return raw.strip()


def _decode(part) -> str:
    """按声明的 charset 解码零件，失败则逐级回退。

    逐个字符集尝试严格解码（失败才算不合格），全部失败才用替换方式兜底，
    否则 errors="replace" 会让任何字符集都“解码成功”，回退链形同虚设。
    """
    payload = part.get_payload(decode=True)
    if payload is None:  # 某些 8bit 邮件没有正确编码，退回原始字符串
        payload = part.get_payload()
        if not isinstance(payload, str):
            return ""
        payload = payload.encode("utf-8", "surrogateescape")
    charsets = [part.get_content_charset(), "utf-8", "gb18030", "big5", "latin-1"]
    for charset in charsets:
        if not charset:
            continue
        try:
            return payload.decode(charset)
        except (LookupError, UnicodeDecodeError):
            continue
    return payload.decode("utf-8", "replace")


def _is_attachment(part) -> bool:
    if part.get_content_disposition() == "attachment":
        return True
    # 正文内嵌的图片（签名 logo 等）带 Content-ID、disposition=inline，不算附件
    if (part.get_content_disposition() == "inline"
            and part.get("Content-ID")
            and part.get_content_type().startswith("image/")):
        return False
    return bool(part.get_filename())


def _fix_surrogates(text: str) -> str:
    """修复非法 charset 的头字段留下的孤立代理字符（否则无法写入 UTF-8 文件）。

    email 包在解码失败时用 surrogateescape 保留原始字节，这里还原成字节后
    按常见编码重新解码；仍失败则丢弃不可表示的字符。
    """
    if not any(0xD800 <= ord(char) <= 0xDFFF for char in text):
        return text
    try:
        raw = text.encode("utf-8", "surrogateescape")
    except UnicodeEncodeError:
        # 含 surrogateescape 覆盖不到的代理字符：只丢掉非法字符，不能让整段变空
        return "".join(char for char in text if not 0xD800 <= ord(char) <= 0xDFFF)
    for encoding in ("utf-8", "gb18030", "big5", "shift_jis", "latin-1"):
        try:
            return raw.decode(encoding)
        except (UnicodeDecodeError, LookupError):
            continue
    return "".join(char for char in text if not 0xD800 <= ord(char) <= 0xDFFF)


def _addr_list(header) -> str:
    if header is None:
        return ""
    addresses = getattr(header, "addresses", None)
    if addresses is None:
        return str(header).strip()
    shapes = []
    for address in addresses:
        name, addr = (address.display_name or "").strip(), address.addr_spec
        # 有些发件人把显示名写成了自己的邮箱，去掉重复的 <addr>
        if name and addr and name.lower() == addr.lower():
            name = ""
        shapes.append(f"{name} <{addr}>" if name else addr)
    return ", ".join(shapes)


def _trace_add(trace: list | None, reason: str, text: str) -> None:
    """记录一段被清洗掉的内容（供界面标注“删了什么、为什么删”）。"""
    if trace is None:
        return
    value = (text or "").strip("\n")
    if not value.strip():
        return
    # 同一段提示在 HTML/纯文本里常重复出现，去重避免界面标出一堆重复块
    if any(item["reason"] == reason and item["text"] == value for item in trace):
        return
    trace.append({"reason": reason, "text": value})


def _strip_boilerplate(text: str, trace: list | None = None) -> str:
    """删除正文中的整段免责/环保提示（按其关键词判定，逐段丢弃）。

    这些段落通常是邮件系统的自动附加内容，不含正文信息；整封都是提示时
    返回空字符串，由调用方另行标注。
    """
    markers: list[tuple[str, tuple[str, ...]]] = [
        # 邮件网关附加的外部邮件安全提醒
        ("提醒：这是一封外部邮件", ()),
        ("reminder: this is an external email", ()),
        ("友情提示", ("打印", "节约用纸", "收件人")),
        ("本电子邮件及其附件", ("敏感信息", "收件人")),
        ("电子邮件及其附件", ("敏感信息", "收件人")),
        ("此邮件及附件", ("保密", "收件人")),
        ("免责声明", ()),
        ("please consider the environment", ()),
        ("this email and any attachments", ("confidential", "privileged")),
        ("privileged and confidential", ()),
    ]
    paragraphs = re.split(r"\n\s*\n", text)
    kept = []
    for paragraph in paragraphs:
        if not paragraph.strip():
            continue
        low = paragraph.lower()
        hit = any(keyword.lower() in low and (not hints or any(h in low for h in hints))
                  for keyword, hints in markers)
        if not hit:
            kept.append(paragraph.strip())
            continue
        lines = [line for line in paragraph.splitlines() if line.strip()]
        # 短段落（典型是网关/环保提示）整段丢弃，避免把分段里的提示行残留进正文
        if len(paragraph) < 200 or sum(
                1 for line in lines if any(k.lower() in line.lower() for k, _ in markers)
        ) * 2 >= len(lines):
            _trace_add(trace, "套话/免责声明", paragraph)
            continue
        # 长段落里混着人工内容：只删命中的行，保留其余正文
        rest = [line for line in lines
                if not any(k.lower() in line.lower() for k, _ in markers)]
        dropped = [line for line in lines
                   if any(k.lower() in line.lower() for k, _ in markers)]
        _trace_add(trace, "套话/免责声明", "\n".join(dropped))
        if rest:
            kept.append("\n".join(rest).strip())
    return "\n\n".join(kept)


def _strip_quoted(text: str, trace: list | None = None) -> str:
    """截掉正文里引用的历史邮件（转发头、Origin message、引用块、发件人: 引言等）。

    保留本封邮件自己写的内容；引用他人历史邮件、回复链一律丢弃。
    """
    cut_pairs = [
        r"\n\s*[-—]{3,}\s*Origin message",
        r"\n\s*[-—]{3,}\s*转发的邮件",
        r"\n\s*[-—]{3,}\s*转 发 的 邮 件",
        r"\n\s*[-—]{3,}\s*原始邮件",
        r"\n\s*[-—]{3,}\s*Original Message",
        r"\n\s*[-—]{3,}\s*Forwarded message",
        r"\n\s*[-—]{3,}\s*Forwarded by",
        r"\n\s*[-—]{3,}\s*以下为转发",
        r"(?:^|\n)\s*>+\s*(?:From|\"|“|发件人)",
        r"(?:^|\n)\s*发件人[:：]",
        r"(?:^|\n)\s*发送时间[:：]",
        r"\n\s*在.{0,80}写道[:：]",
        r"(?:^|\n)\s*On .{0,80} wrote[:：]",
    ]
    earliest = len(text)
    for pattern in cut_pairs:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match and match.start() < earliest:
            earliest = match.start()

    # 纯 "> " 引用块：以连续引用行开头、且附近出现邮件头字段（Subject/Date/主题 等）
    header_re = re.compile(r"(?:subject|from|date|to|cc|bcc|主题|发件人|收件人|日期)[:：]",
                           re.IGNORECASE)
    lines = text.splitlines(keepends=True)
    quote_start = None
    for index, line in enumerate(lines):
        stripped = line.lstrip()
        if stripped.startswith((">", "＞")):
            if quote_start is None:
                quote_start = index
            continue
        if quote_start is not None:
            # 只看行首字段：正文里的 "reply to:"、"send to:" 不算邮件头
            if any(header_re.match(candidate.strip()) for candidate in lines[index:index + 8]):
                break  # 引用块后面跟邮件头 → 从引用块起始处截断
            quote_start = None
    if quote_start is not None:
        position = sum(len(item) for item in lines[:quote_start])
        earliest = min(earliest, position)

    if earliest < len(text):
        _trace_add(trace, "历史转发/引用", text[earliest:])
    return text[:earliest].rstrip()


def _own_body(text: str, trace: list | None = None) -> str:
    """本封邮件自己的正文：去掉引用历史邮件与签名档。

    整封都是转发/引用的内容时，不把历史邮件当成正文塞进结果，只标注一句。
    """
    trimmed = _trim_closing(_trim_signature(_strip_quoted(text, trace), trace), trace)
    if trimmed.strip():
        return trimmed
    if text.strip():
        _trace_add(trace, "整封无自写正文", text)
        return "（本封无自写正文，内容为转发/引用的历史邮件）"
    return ""


def _strip_banners(subject: str) -> str:
    """去掉邮件网关在主题前加的外部邮件提醒，例如“[邮件来自外部，注意安全]”。"""
    subject = subject.strip()
    banners = [
        r"^\[[^\]]*邮件来自外部[^\]]*\]\s*",
        r"^【[^】]*外部邮件[^】]*】\s*",
        r"^\[[^\]]*(EXTERNAL|外部邮件|外部发送)[^\]]*\]\s*",
        r"^【[^】]*(外部邮件|外部发送)[^】]*】\s*",
    ]
    changed = True
    while changed:
        changed = False
        for pattern in banners:
            match = re.match(pattern, subject, re.IGNORECASE)
            if match:
                subject = subject[match.end():].strip()
                changed = True
    return subject


_RESERVED_NAMES = {"CON", "PRN", "AUX", "NUL"}


def _safe_name(name: str, fallback: str) -> str:
    """清掉路径分隔符与非法字符，并限制长度（保留扩展名，避免截断后文件名不带后缀）。"""
    name = Path(name.replace("\\", "/")).name.strip() or fallback
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    name = name.rstrip(" .")  # Windows 不允许结尾的点与空格
    if not name:
        name = fallback
    # Windows 保留设备名（con、aux、com1 …）会被当成设备，前缀下划线规避
    base = name.split(".")[0].upper()
    if base in _RESERVED_NAMES or re.fullmatch(r"(COM|LPT)[1-9]", base):
        name = "_" + name
    if len(name) <= 120:
        return name
    suffix = Path(name).suffix
    if len(suffix) >= 20:  # 扩展名本身过长，按无后缀处理，避免重复拼接
        return name[:120]
    return name[:120 - len(suffix)] + suffix


# ---------------------------------------------------------------- 日历邀请

def parse_ics(text: str) -> list[dict]:
    """极简 iCalendar 解析：展开折行，抽取 VEVENT 的常用字段。

    除摘要与时间外，还取出组织者（含 CN 显示名）、地点、参会人、加入链接，
    以及描述原文，供日程表展示与判断会议形式。
    """
    unfolded: list[str] = []
    for raw in text.splitlines():
        if raw[:1] in (" ", "\t") and unfolded:
            unfolded[-1] += raw[1:]
        else:
            unfolded.append(raw)

    events: list[dict] = []
    current: dict | None = None
    in_alarm = False
    for line in unfolded:
        upper = line.upper()
        # 结构行必须先判断：提醒块没闭合（.ics 被截断）时也不能吞掉 END:VEVENT
        if upper.startswith("BEGIN:VEVENT"):
            current = {"attendees": [], "start_tzid": "", "end_tzid": ""}
            in_alarm = False
            continue
        if upper.startswith("END:VEVENT"):
            if current is not None:
                events.append(current)
                current = None
            in_alarm = False
            continue
        if upper.startswith("BEGIN:VALARM"):
            in_alarm = True
            continue
        if upper.startswith("END:VALARM"):
            in_alarm = False
            continue
        if in_alarm:  # 提醒块的 DESCRIPTION 会覆盖会议描述，必须跳过
            continue
        if current is None:
            continue
        pieces = _ics_split_line(line)
        if pieces is None:
            continue
        head, value = pieces
        key = head.split(";")[0].upper()
        value = value.strip()
        params = dict(
            (m.group(1).upper(), m.group(2).strip('"'))
            for m in re.finditer(r';([A-Za-z0-9\-]+)=("[^"]*"|[^;:]*)', head)
        )
        if key == "ATTENDEE":
            current["attendees"].append({
                "name": params.get("CN", ""),
                "email": re.sub(r"(?i)^mailto:", "", value.strip()),
            })
            continue
        current[key] = value
        if key == "ORGANIZER":
            current["organizer_name"] = params.get("CN", "")
        if key in ("DTSTART", "DTEND") and params.get("TZID"):
            current["start_tzid" if key == "DTSTART" else "end_tzid"] = params["TZID"]

    if current is not None:  # .ics 在最后一个事件中途结束：已解析的部分也收下
        events.append(current)

    normalized = []
    for event in events:
        organizer_raw = re.sub(r"(?i)^mailto:", "", event.get("ORGANIZER", "").strip())
        start_text, start_zone = _ics_datetime(event.get("DTSTART", ""),
                                               event.get("start_tzid", ""))
        end_text, end_zone = _ics_datetime(event.get("DTEND", ""),
                                           event.get("end_tzid", "") or start_zone)
        normalized.append({
            "summary": _ics_unescape(event.get("SUMMARY", "")),
            "start": start_text,
            "end": end_text,
            "location": _ics_unescape(event.get("LOCATION", "")),
            "organizer": organizer_raw,
            "organizer_name": event.get("organizer_name", ""),
            "attendees": event.get("attendees", []),
            "description": _ics_unescape(event.get("DESCRIPTION", "")),
            "url": event.get("X-MICROSOFT-SKYPETEAMSMEETINGURL", "") or event.get("URL", ""),
            "uid": event.get("UID", ""),
            "tzid": start_zone or end_zone,
        })
    return normalized


def _ics_split_line(line: str) -> tuple[str, str] | None:
    """按引号外的第一个冒号切成 (参数段, 值段)。

    RFC 5545 允许参数值用双引号包裹并含冒号（如 CN="会议: 讨论"），
    split(":", 1) 会把值截断。
    """
    in_quote = False
    for index, char in enumerate(line):
        if char == '"':
            in_quote = not in_quote
        elif char == ":" and not in_quote:
            return line[:index], line[index + 1:]
    return None


def _ics_unescape(value: str) -> str:
    """还原 iCalendar 文本转义（\\n、\\,、\\;、\\\\），单次扫描避免链式覆盖。"""
    return re.sub(r"\\([\\;,nN])", lambda m: "\n" if m.group(1) in "nN" else m.group(1),
                  value or "")


def _local_zone_label(moment: datetime) -> str:
    name = moment.tzname() or ""
    offset = moment.strftime("%z")
    if offset:
        offset = f"{offset[:3]}:{offset[3:]}"
        return f"本机时间 UTC{offset}" if not name else f"{name}（UTC{offset}）"
    return name or "本机时间"


def _ics_datetime(value: str, tzid: str = "") -> tuple[str, str]:
    """把 20261112T150000Z / 20261112T150000 / 20261112 转成 (可读时间, 展示时区)。

    带 Z 的是 UTC，必须换算成本机时区；带 TZID 的保留邀请里写的墙上时间。
    """
    value = (value or "").strip()
    match = re.match(r"(\d{8})(?:T(\d{6}))?(Z)?", value)
    if not match:
        return "", tzid
    date_part, time_part, utc = match.group(1), match.group(2), match.group(3)
    if not time_part:
        return f"{date_part[:4]}-{date_part[4:6]}-{date_part[6:]}（全天）", tzid
    try:
        moment = datetime.strptime(date_part + "T" + time_part, "%Y%m%dT%H%M%S")
    except ValueError:
        return "", tzid
    if utc:
        moment = moment.replace(tzinfo=timezone.utc).astimezone()
        return moment.strftime("%Y-%m-%d %H:%M"), _local_zone_label(moment)
    return moment.strftime("%Y-%m-%d %H:%M"), tzid


_LINK_PATTERN = re.compile(r"https?://[^\s<>\"'）\)\]]+")

_PLATFORM_KEYWORDS = [
    ("Microsoft Teams", ("teams.microsoft.com", "microsoft teams", "teams meeting", "teams 会议")),
    ("Zoom", ("zoom.us", "zoom meeting", "zoom 会议")),
    ("Webex", ("webex",)),
    ("腾讯会议", ("腾讯会议", "meeting.tencent.com")),
    ("钉钉会议", ("钉钉", "dingtalk")),
    ("飞书会议", ("飞书", "feishu", "larksuite")),
    ("Google Meet", ("meet.google.com",)),
]

# 平台自动附加的入会说明行，不属于人工写的会议安排（命中即说明其后都是入会指引）
_BOILERPLATE_LINES = (
    "microsoft teams meeting", "join:", "meeting id:", "passcode:", "need help?",
    "join on a video conferencing device", "for organisers:", "tenant key:",
    "video id:", "more info", "system reference", "générer", "————",
)

# 邮件网关在正文顶部加的外部邮件安全提醒：只是声明，跳过即可，不能据此截断议程
_REMINDER_LINES = (
    "提醒：这是一封外部邮件", "reminder: this is an external email",
    "请谨慎对待邮件内容、链接和附件", "please be careful with content, links",
)


def _meeting_platform(event: dict) -> str:
    text = " ".join([
        event.get("location") or "", event.get("description") or "", event.get("url") or "",
    ]).lower()
    for name, keys in _PLATFORM_KEYWORDS:
        if any(key in text for key in keys):
            return name
    return ""


def _meeting_form(event: dict) -> str:
    """判断线上/线下/混合，附平台或场地。"""
    platform = _meeting_platform(event)
    location = (event.get("location") or "").strip()
    has_link = bool(event.get("url")) or bool(
        _LINK_PATTERN.search(event.get("description", "") or ""))
    place_is_platform = bool(platform) and platform.lower() in location.lower()
    physical = bool(location) and not place_is_platform

    if platform and physical:
        return f"线上线下结合（{platform}；{location}）"
    if platform:
        return f"线上（{platform}）"
    if has_link:
        return "线上（视频会议，平台未注明）"
    if location:
        return f"线下（{location}）"
    return "未注明"


def _join_info(event: dict) -> str:
    """提取入会链接、会议号与密码。"""
    text = "\n".join([event.get("description") or "", event.get("location") or ""])
    link = ""
    match = re.search(r"(?:Join|加入会议|会议链接|Join meeting)[:：]\s*(https?://[^\s<>]+)",
                      text, re.IGNORECASE)
    if match:
        link = match.group(1)
    elif event.get("url"):
        link = event["url"]
    else:
        found = _LINK_PATTERN.search(text)
        link = found.group(0) if found else ""

    meeting_id = re.search(r"(?:Meeting ID|会议号)[:：]?\s*([0-9][0-9 \-]{6,})", text, re.IGNORECASE)
    passcode = re.search(r"(?:Passcode|Password|入会密码|密码)[:：]\s*([^\s\n，。；;,.、]+)",
                         text, re.IGNORECASE)

    parts = []
    if link:
        parts.append(f"链接：{link}")
    if meeting_id:
        parts.append(f"会议号：{meeting_id.group(1).strip()}")
    if passcode:
        parts.append(f"密码：{passcode.group(1).strip()}")
    return " ｜ ".join(parts)


def _agenda_text(event: dict) -> str:
    """人工撰写的会议说明/议程（剔除平台自动附加的入会指引）。"""
    lines = []
    for line in (event.get("description") or "").splitlines():
        stripped = line.strip()
        if set(stripped) == {"_"} and stripped:
            break  # 平台自动附加的分隔线，其后都是入会指引
        low = stripped.lower()
        if stripped and any(low.startswith(marker) for marker in _REMINDER_LINES):
            continue  # 网关提醒常在正文顶部，跳过该行而不是截断
        if stripped and any(low.startswith(marker) for marker in _BOILERPLATE_LINES):
            break
        lines.append(stripped)
    text = "\n".join(lines).strip()
    return re.sub(r"\n{3,}", "\n\n", text)


def _attendee_text(event: dict) -> str:
    attendees = event.get("attendees") or []
    if not attendees:
        return ""
    names = []
    for person in attendees:
        label = (person.get("name") or person.get("email") or "").strip()
        if label:
            names.append(label)
    if not names:
        return ""
    head = "、".join(names[:5])
    tail = " 等" if len(names) > 5 else ""
    prefix = "收件人 " if event.get("attendee_kind") == "recipients" else ""
    return f"{prefix}{len(names)} 人：{head}{tail}"


# ---------------------------------------------------------------- 正文中的会议通知

# 正文里常见的时区写法 → 展示用的时区名
_BODY_TIMEZONES = [
    ("北京时间", "北京时间/Asia/Shanghai"),
    ("GMT+8", "GMT+8"),
    ("UTC+8", "UTC+8"),
    ("W. Central Africa Standard Time", "W. Central Africa Standard Time"),
    ("UTC+1", "UTC+1"),
    ("GMT+1", "GMT+1"),
    ("UTC", "UTC"),
]

# 正文会议通知里，签名档、祝福语之后的内容不属于会议安排
_SIGNATURE_CUTS = (
    "祝好", "此致", "顺祝", "thanks & best", "thanks and best", "best regards",
    "kind regards", "p save a tree", "mail:", "mobile:", "tel ", "add:",
)

_BODY_MEETING_HINTS = ("会议", "例会", "汇报会", "会商", "meeting", "conference", "workshop")

_BODY_DATE_RE = re.compile(r"(20\d{2})\s*[年\-/.]\s*(\d{1,2})\s*[月\-/.]\s*(\d{1,2})\s*日?")
# 时段词单独成组：下午/晚上 3:00 要换算成 15:00，不能只把词吞掉
_BODY_TIME_RE = re.compile(r"(?P<ap>上午|下午|晚上|中午)?\s*(?P<hour>\d{1,2})\s*[:：]\s*(?P<minute>\d{2})")
_BODY_RANGE_RE = re.compile(
    r"(\d{1,2}\s*[:：]\s*\d{2})\s*[~～\-–—至到]\s*(\d{1,2}\s*[:：]\s*\d{2})")
_BODY_PLACE_RE = re.compile(r"(?:会议地点|地点|会议室)[:：]\s*([^\n，。；;]+)")


def _clock_parts(match) -> tuple[int, int] | None:
    """把 _BODY_TIME_RE 的匹配结果换算成 24 小时制的 (时, 分)。"""
    hour, minute = int(match.group("hour")), int(match.group("minute"))
    ap = match.group("ap") or ""
    if ap in ("下午", "晚上") and hour < 12:
        hour += 12
    if ap == "中午" and hour < 12:
        hour = 12
    if not (0 <= hour <= 23 and 0 <= minute <= 59):
        return None
    return hour, minute


def _parse_people(value: str) -> list[dict]:
    """把头部地址串拆成 [{'name','email'}]，用于会议通知的参会人列。"""
    people = []
    chunks = [chunk for chunk in re.split(r"[;]", value or "") if chunk.strip()]
    for chunk in chunks:
        parsed = [(name.strip().strip('"'), addr.strip())
                  for name, addr in getaddresses([chunk]) if addr.strip()]
        if parsed:
            for name, email in parsed:
                people.append({"name": name, "email": email})
            continue
        text = chunk.strip()
        if text:
            people.append({"name": "", "email": text})
    return people


def _trim_signature(text: str, trace: list | None = None) -> str:
    """截掉签名档与祝福语，只留会议通知本身的内容。"""
    kept = []
    lines = text.splitlines()
    for index, line in enumerate(lines):
        stripped = line.strip()
        if stripped and stripped.lower().startswith(_SIGNATURE_CUTS):
            _trace_add(trace, "签名档", "\n".join(lines[index:]))
            break
        kept.append(line)
    return "\n".join(kept).strip()


# 邮件末尾的礼貌收尾语（单独成行、或与姓名写在同一行）；含寒暄式告别
_CLOSING_RE = re.compile(
    r"^(?:regards|best regards|kind regards|warm regards|thank you|thanks|many thanks|"
    r"best wishes|br|rgds|good night|goodnight|good day|goodbye|bye|cheers|"
    r"yours sincerely|yours faithfully|sincerely|respectfully|"
    r"此致|敬礼|顺颂商祺|顺祝商祺|祝好|商祺|晚安|再见|谢谢|礼！)"
    r"\s*[,，:：!！]?\s*(?P<rest>[^\n]*)$", re.IGNORECASE)

# 正文与收尾语写在同一行末尾时（"Please find attached, Regards, Robert"）
_INLINE_CLOSING_RE = re.compile(
    r"[,，;；。]\s*(?:best regards|kind regards|warm regards|regards|thanks|thank you|"
    r"good night|goodnight|goodbye|cheers|"
    r"此致|敬礼|顺颂商祺|祝好|晚安|再见)\s*[,，:：]?\s*(?:[^,，。！？\n]{0,40})$", re.IGNORECASE)


def _looks_like_name(line: str) -> bool:
    """收尾语后面的署名行：短、无句末标点与数字，最多 5 个词。"""
    text = line.strip()
    return bool(text) and len(text) <= 40 and not re.search(r"[。！？.!?]|\d", text) \
        and len(text.split()) <= 5


def _trim_closing(text: str, trace: list | None = None) -> str:
    """截掉结尾的礼貌收尾语（Regards / 此致 / Thanks …）及其后面的署名。

    先按行找单独成行的收尾语；找不到再看末行末尾是否挂着收尾语
    （如 "Please find attached, Regards, Robert"）。只有收尾语后跟的都是
    署名性质的短行时才截，否则视为正文内容保持原样。
    """
    lines = text.splitlines()

    # 从末行往前：跳过空行与署名短行，找到的第一行若既不是收尾语也不是署名，就不动
    index = len(lines) - 1
    while index >= 0:
        stripped = lines[index].strip()
        if not stripped:
            index -= 1
            continue

        # 形态一：整行就是收尾语（可能带姓名）
        match = _CLOSING_RE.match(stripped)
        if match:
            name = (match.group("rest") or "").strip().strip(".,;:，；。：！!")
            if name and not _looks_like_name(name):
                return text.strip()
            _trace_add(trace, "收尾语/寒暄", "\n".join(lines[index:]))
            return "\n".join(lines[:index]).strip()

        # 形态二：正文和收尾语写在同一行末尾
        inline = _INLINE_CLOSING_RE.search(stripped)
        if inline and inline.start() > 0:
            head = stripped[:inline.start()].rstrip(" ,，;；")
            _trace_add(trace, "收尾语/寒暄", stripped[inline.start():])
            lines[index] = head
            return "\n".join(lines[:index + 1]).strip()

        # 署名行继续往前找；普通正文行则到此为止
        if not _looks_like_name(stripped):
            break
        index -= 1

    return text.strip()


def extract_body_meeting(record: dict) -> dict | None:
    """从正文识别人工口述的会议通知（无日历邀请），返回事件字典；识别不到返回 None。

    只在非引用正文里找，且必须同时出现会议类字样与可解析的「年-月-日」。
    """
    text = _strip_quoted(record.get("body_text", "") or "")
    hint_text = f"{record.get('subject', '')} {text}".lower()
    if not text or not any(hint.lower() in hint_text for hint in _BODY_MEETING_HINTS):
        return None

    # 正文可能先提到别的日期（如上次例会），取第一处其后 60 字内能解析出时分的日期
    match = None
    start_clock = ""
    for candidate in _BODY_DATE_RE.finditer(text):
        window = text[candidate.end():candidate.end() + 60]
        clock = _BODY_TIME_RE.search(window)
        if not clock:
            clock = _BODY_TIME_RE.search(text[max(0, candidate.start() - 40):candidate.start()])
        if not clock:
            continue
        parts = _clock_parts(clock)
        if parts is None:
            continue
        match, start_clock = candidate, f"{parts[0]:02d}:{parts[1]:02d}"
        break
    if match is None:
        # 只有日期没有时分的不算日程（日报里提到日期 + meeting 字样会误判）
        return None
    try:
        year, month, day = int(match.group(1)), int(match.group(2)), int(match.group(3))
        if not (1 <= month <= 12 and 1 <= day <= 31):
            return None
        start_date = datetime(year, month, day)
    except ValueError:
        return None
    window = text[match.end():match.end() + 60]

    end_clock = ""
    span = _BODY_RANGE_RE.search(window)
    if span:
        tail_time = _BODY_TIME_RE.search(span.group(2))
        if tail_time:
            parts = _clock_parts(tail_time)
            if parts and f"{parts[0]:02d}:{parts[1]:02d}" > start_clock:
                end_clock = f"{parts[0]:02d}:{parts[1]:02d}"

    start_text = f"{start_date:%Y-%m-%d} {start_clock}".strip()
    end_text = f"{start_date:%Y-%m-%d} {end_clock}" if end_clock else ""

    zone = ""
    for marker, label in _BODY_TIMEZONES:
        if marker.lower() in text.lower():
            zone = label
            break

    place = ""
    place_match = _BODY_PLACE_RE.search(text)
    if place_match:
        place = place_match.group(1).strip()

    low = text.lower()
    form = ""
    if any(word in low for word in ("视频会议", "线上会议", "网络会议", "腾讯会议", "teams", "zoom", "钉钉会议", "飞书会议")):
        form = "线上（视频会议，平台未注明）"
    elif place:
        form = f"线下（{place}）"

    plain = _trim_closing(_trim_signature(text))
    paragraphs = [p.strip() for p in re.split(r"\n\s*\n", plain) if p.strip()]
    description = "\n\n".join(paragraphs[:2])

    url = ""
    link_match = _LINK_PATTERN.search(text)
    if link_match:
        url = link_match.group(0)

    attendees = _parse_people(record.get("to", "")) + _parse_people(record.get("cc", ""))

    return {
        "summary": (record.get("subject") or "").strip() or (paragraphs[0][:60] if paragraphs else ""),
        "start": start_text,
        "end": end_text,
        "tzid": zone,
        "location": place,
        "form": form,
        "description": description,
        "url": url,
        "organizer": record.get("from_addr", ""),
        "organizer_name": "",
        "attendees": attendees,
        "attendee_kind": "recipients",
        "source": "body",
    }


# ---------------------------------------------------------------- 单封解析

def _unique_target(dir_path: Path, filename: str) -> Path:
    """同名附件不互相覆盖：第二个起加 _1、_2 后缀。"""
    target = dir_path / filename
    index = 1
    while target.exists():
        target = dir_path / f"{Path(filename).stem}_{index}{Path(filename).suffix}"
        index += 1
    return target


def read_eml_file(path: str | Path, save_attachments: bool = False,
                  attach_root: str | Path | None = None,
                  attach_subdir: str | None = None,
                  trace: list | None = None) -> dict:
    """读取单个 .eml，返回字典（字段见模块顶部注释 / 索引文件输出）。"""
    path = Path(path)
    record: dict = {
        "file": str(path),
        "name": path.name,
        "size_bytes": 0,
        "date": "",
        "subject": "",
        "from": "",
        "from_addr": "",
        "to": "",
        "cc": "",
        "reply_to": "",
        "message_id": "",
        "in_reply_to": "",
        "references": "",
        "body_text": "",
        "body_full": "",
        "body_html": "",
        "body_chars": 0,
        "attachments": [],
        "attachment_count": 0,
        "is_multipart": False,
        "ics_events": [],
        "ics_raw": [],
        "body_events": [],
        "parse_error": "",
    }

    try:
        record["size_bytes"] = path.stat().st_size
        with path.open("rb") as handle:
            msg = PARSER.parse(handle)
    except Exception as exc:  # 单封失败不影响整批
        record["parse_error"] = f"{type(exc).__name__}: {exc}"
        return record

    record["is_multipart"] = msg.is_multipart()
    record["subject"] = _strip_banners(str(msg.get("Subject", "") or ""))
    record["from"] = _addr_list(msg.get("From"))
    record["to"] = _addr_list(msg.get("To"))
    record["cc"] = _addr_list(msg.get("Cc"))
    record["reply_to"] = _addr_list(msg.get("Reply-To"))
    record["message_id"] = str(msg.get("Message-ID", "") or "").strip()
    record["in_reply_to"] = str(msg.get("In-Reply-To", "") or "").strip()
    record["references"] = str(msg.get("References", "") or "").strip()

    from_header = msg.get("From")
    if from_header is not None and getattr(from_header, "addresses", None):
        record["from_addr"] = from_header.addresses[0].addr_spec

    date_header = msg.get("Date")
    when = getattr(date_header, "datetime", None)
    if when is None and date_header is not None:
        try:
            when = parsedate_to_datetime(str(date_header))
        except Exception:
            when = None
    if when is not None:
        record["date"] = when.isoformat()

    text_chunks: list[str] = []
    html_chunks: list[str] = []
    attachments: list[dict] = []
    ics_raw: list[str] = []
    ics_events: list[dict] = []

    # 每封邮件的附件单独放一个子目录；重跑时先清空该子目录，保证结果可重复
    target_dir: Path | None = None
    if save_attachments and attach_root is not None:
        folder = attach_subdir or _safe_name(path.stem, "mail")
        target_dir = Path(attach_root) / folder
        if target_dir.is_dir():
            for stale in target_dir.iterdir():
                if stale.is_file():
                    try:
                        stale.unlink()
                    except OSError as exc:  # 被 Excel/Word 占用时删不掉，改名写会堆积，必须提示
                        print(f"警告：无法清理旧附件 {stale}：{exc}", file=sys.stderr)
        # 目录等真正写出附件时再建，没有附件的邮件不留空目录

    for part in msg.walk():
        if part.is_multipart():
            continue
        content_type = part.get_content_type()
        if content_type == "message/rfc822":
            continue

        # 日历邀请（text/calendar）保留原文用于导出 .ics，并解析出会议时间地点
        if content_type == "text/calendar":
            raw_ics = _decode(part)
            ics_raw.append(raw_ics)
            ics_events.extend(parse_ics(raw_ics))
            if part.get_filename() and target_dir is not None:
                # 以附件形式附带的 invite.ics 也要落盘，不能只解析不留文件
                filename = _safe_name(part.get_filename() or "",
                                      f"invite_{len(attachments) + 1}.ics")
                payload = part.get_payload(decode=True) or b""
                entry = {"filename": filename, "content_type": content_type,
                         "size_bytes": len(payload), "saved_path": "", "save_error": ""}
                target_dir.mkdir(parents=True, exist_ok=True)
                target = _unique_target(target_dir, filename)
                try:
                    target.write_bytes(payload)
                    entry["saved_path"] = str(target)
                except OSError as exc:
                    entry["save_error"] = f"{type(exc).__name__}: {exc}"
                    print(f"警告：附件写出失败 {target}：{exc}", file=sys.stderr)
                attachments.append(entry)
            continue

        if _is_attachment(part):
            filename = _safe_name(
                part.get_filename() or "", f"attachment_{len(attachments) + 1}"
            )
            payload = part.get_payload(decode=True) or b""
            # 发送方常把 xlsx/docx 声明成 octet-stream，按扩展名补正便于识别
            if content_type == "application/octet-stream":
                guessed, _ = mimetypes.guess_type(filename)
                if guessed:
                    content_type = guessed
            entry = {
                "filename": filename,
                "content_type": content_type,
                "size_bytes": len(payload),
                "saved_path": "",
                "save_error": "",
            }
            if target_dir is not None:
                target_dir.mkdir(parents=True, exist_ok=True)
                target = _unique_target(target_dir, filename)
                try:
                    target.write_bytes(payload)
                    entry["saved_path"] = str(target)
                except OSError as exc:
                    # 错误不塞进 saved_path：该字段必须是真实路径，否则下游按路径读会拿到不存在的文件
                    entry["save_error"] = f"{type(exc).__name__}: {exc}"
                    print(f"警告：附件写出失败 {target}：{exc}", file=sys.stderr)
            attachments.append(entry)
            continue

        if content_type == "text/plain":
            text_chunks.append(_decode(part))
        elif content_type == "text/html":
            html_chunks.append(_decode(part))

    body_text = "\n\n".join(chunk.strip() for chunk in text_chunks if chunk.strip())
    body_html = "\n".join(html_chunks)
    if not body_text and body_html:
        body_text = html_to_text(body_html)
    # trace 非空时保留清洗前的正文快照，供界面标注“删了哪段、为什么删”
    raw_body = body_text if trace is not None else ""
    body_text = _strip_boilerplate(body_text, trace)
    # 纯文本部分只有「外部邮件提醒」这类声明时，正文在 HTML 里，改从 HTML 取
    if not body_text.strip() and body_html:
        html_text = html_to_text(body_html)
        raw_body = raw_body or (html_text if trace is not None else "")
        body_text = _strip_boilerplate(html_text, trace)
    # 正文只留本封自己的内容：引用/转发来的历史邮件与签名档都不进结果；
    # 完整原文另存 body_full（只留在内存/界面回溯用，不写索引）
    full_body = body_text
    body_text = _own_body(body_text, trace)
    # 正文被清空（只剩安全提醒/免责声明/签名）但有附件时，注明内容在附件里，别留空行
    if not body_text.strip() and attachments:
        names = "、".join(item["filename"] for item in attachments[:3])
        more = " 等" if len(attachments) > 3 else ""
        body_text = f"（本封无正文内容，实质内容见附件：{names}{more}）"

    record["body_full"] = full_body
    if trace is not None:
        # 仅界面按需带上：清洗前原文 + 删除记录，不进索引文件，避免体积膨胀
        record["body_raw"] = raw_body or full_body

    record["body_text"] = body_text
    record["body_html"] = body_html
    record["body_chars"] = len(body_text)
    record["attachments"] = attachments
    record["attachment_count"] = len(attachments)
    record["ics_events"] = ics_events
    record["ics_raw"] = ics_raw
    # 没有日历邀请、但正文里写明时间地点的会议通知，同样进日程表
    if not ics_events:
        body_event = extract_body_meeting(record)
        if body_event:
            record["body_events"] = [_clean_record({"event": body_event})["event"]]
    record = _clean_record(record)
    if (not record["subject"] and not record["from"] and not record["date"]
            and not record["body_text"]):
        record["parse_error"] = "无有效邮件头与正文（文件可能不是标准 eml）"
    return record


def _clean_record(record: dict) -> dict:
    """统一清理所有字符串字段中的孤立代理字符，保证可安全写入文件。"""
    def _walk(value):
        if isinstance(value, str):
            return _fix_surrogates(value)
        if isinstance(value, list):
            return [_walk(item) for item in value]
        if isinstance(value, dict):
            return {key: _walk(item) for key, item in value.items()}
        return value

    for key, value in list(record.items()):
        record[key] = _walk(value)
    record["body_chars"] = len(record.get("body_text", ""))
    return record


# ---------------------------------------------------------------- 批量收集

def read_eml_dir(root: str | Path, recursive: bool = False,
                 pattern: str = "*.eml") -> list[dict]:
    """批量读取目录下的 .eml，返回记录列表（按真实时间倒序，无日期者排后）。"""
    root = Path(root)
    if not root.is_dir():
        raise FileNotFoundError(f"邮件目录不存在: {root}")
    files = sorted(root.rglob(pattern) if recursive else root.glob(pattern))
    records = []
    for path in files:
        try:
            records.append(read_eml_file(path))
        except Exception as exc:  # 单封失败不影响整批
            records.append({"file": str(path), "name": path.name,
                            "parse_error": f"{type(exc).__name__}: {exc}"})
    return sort_records(records)


def _record_sort_key(record: dict) -> tuple[int, float, str]:
    """排序键：有时间的在前且按时间倒序，无时间/时间无法解析的排最后。"""
    stamp = None
    if record.get("date"):
        try:
            stamp = datetime.fromisoformat(record["date"]).timestamp()
        except (ValueError, OSError, OverflowError):
            # Windows 上 1970 年前的 timestamp() 会抛 OSError，不能让整批排序崩掉
            stamp = None
    return (stamp is None, -stamp if stamp is not None else 0.0, record.get("name", ""))


def sort_records(records: list[dict]) -> list[dict]:
    """按邮件时间倒序排列（新邮件在前）；无日期或时间无法解析的排最后。"""
    return sorted(records, key=_record_sort_key)


# ---------------------------------------------------------------- 邮件分类

# 顺序即优先级：先命中者优先，避免“关于……的通知”被归到沟通类
# 会议类只由 classify() 依据日历邀请（ICS）判定：正文里通知开会的邮件属于通知类
CATEGORY_RULES: list[tuple[str, list[str]]] = [
    ("通知类", [
        "通知", "通报", "通告", "公告", "关于开展", "关于梳理", "关于做好",
        "关于报送", "关于上报", "关于召开", "会议纪要", "纪要", "排查", "整改",
        "报送", "上报", "请于", "要求", "请各单位", "各项目部",
        # 会议通知类措辞：只告知时间地点、无日历邀请，仍按通知类归档
        "会议", "召开", "例会", "汇报会", "视频会议", "会议邀请", "邀请函",
        "参会", "出席", "日程", "invitation", "invite", "calendar",
        "meeting", "conference call", "zoom", "teams",
        "notice", "circular", "memo",
    ]),
    ("日报类", [
        "日报", "班报", "值班报", "日报表", "daily", "report", "reporting",
        "reading", "记录表", "今日完成", "明日计划", "progress report",
    ]),
    ("沟通类", [
        "回复", "答复", "请示", "申请", "确认", "协调", "咨询", "反馈",
        "审阅", "沟通", "thanks", "regards", "please confirm", "re:", "fw:",
    ]),
]


def _keyword_hit(keyword: str, haystack: str) -> bool:
    """关键词命中判断：英文按词边界，中文按子串。

    "report" 不能命中 "reporting"，"re:" 不能命中 "score:"，否则分类会被无关词干扰。
    """
    low = keyword.lower()
    if not low:
        return False
    if re.search(r"[a-z]", low) and re.fullmatch(r"[a-z0-9:'.\- ]+", low):
        if low.endswith(":"):
            return bool(re.search(rf"(?<![a-z0-9]){re.escape(low)}", haystack))
        return bool(re.search(rf"(?<![a-z0-9]){re.escape(low)}(?![a-z0-9])", haystack))
    return low in haystack


def classify(record: dict) -> tuple[str, str]:
    """按主题与正文关键词归类，返回 (类别, 判定依据)。

    只有带日历邀请（ICS）的邮件归会议类；正文里通知开会的邮件（没有 ICS）按通知类
    归档。主题命中权重 3 分、正文命中 1 分，取得分最高者；全为 0 归“其他类”。
    """
    if record.get("ics_events"):
        return "会议类", "含日历邀请（ICS）"

    subject = (record.get("subject", "") or "").lower()
    body = (record.get("body_text", "") or "")[:BODY_MATCH_LIMIT].lower()
    best_category, best_score, best_hit = "其他类", 0, ""
    best_hit_part = "正文"

    for category, keywords in CATEGORY_RULES:
        score, subject_hit, body_hit = 0, "", ""
        for keyword in keywords:
            low_keyword = keyword.lower()
            if _keyword_hit(low_keyword, subject):
                score += 3
                if not subject_hit:
                    subject_hit = keyword
            elif _keyword_hit(low_keyword, body):
                score += 1
                if not body_hit:
                    body_hit = keyword
        if score > best_score:
            best_category, best_score = category, score
            best_hit = subject_hit or body_hit
            best_hit_part = "主题" if subject_hit else "正文"

    if best_score == 0:
        return "其他类", "未命中规则"
    return best_category, f"{best_hit_part}关键词「{best_hit}」"


def _key_point(record: dict, limit: int = 150) -> str:
    """取正文开头作为要点，并优先把带期限的句子提出来（不看引用的历史邮件）。"""
    body = _strip_quoted(record.get("body_text", "") or "")
    body = re.sub(r"\s+", " ", body).strip()
    if not body:
        names = [item.get("filename", "未命名附件")
                 for item in (record.get("attachments") or []) if isinstance(item, dict)]
        if names:
            return "（正文为空，内容见附件：" + "、".join(names) + "）"
        return "（无正文）"
    deadline = ""
    match = re.search(r"((?:于|在|请于)?\s*\d{1,2}月\d{1,2}日(?:[（(][^）)]*[）)])?[^。；;\n]{0,20}?前)", body)
    if match:
        deadline = f"【期限】{match.group(1).strip()}"
    head = body[:limit] + ("……" if len(body) > limit else "")
    return f"{deadline}{head}" if not deadline else f"{head}（{deadline}）"


# ---------------------------------------------------------------- 输出

CSV_FIELDS = [
    "file", "date", "from", "from_addr", "to", "cc", "subject",
    "attachment_count", "attachment_names", "size_bytes", "body_chars",
    "body_text", "parse_error",
]

BRIEF_FIELDS = ["发件人", "收件时间", "邮件主题", "正文内容", "分类"]

# ---------------------------------------------------------------- 输出语言
# 只影响表头、分类值、状态等“标签”文字；正文内容与文件名保持原样。
LABELS: dict[str, dict[str, str]] = {
    "zh": {},
    "en": {
        "发件人": "From", "收件时间": "Received", "邮件主题": "Subject",
        "正文内容": "Body", "分类": "Category",
        "file": "File", "date": "Date", "from": "From", "from_addr": "From address",
        "to": "To", "cc": "Cc", "subject": "Subject",
        "attachment_count": "Attachments", "attachment_names": "Attachment names",
        "size_bytes": "Size (bytes)", "body_chars": "Body characters",
        "body_text": "Body text", "parse_error": "Parse error",
        "会议组织者": "Organizer", "会议议题": "Topic", "会议时间": "Time",
        "会议形式": "Format", "会议安排": "Agenda", "会议号/链接": "Meeting ID / link",
        "参会人": "Attendees", "状态": "Status", "来源邮件": "Source email",
        "会议日程表": "Meeting schedule",
    },
}

CATEGORY_EN = {"日报类": "Daily report", "会议类": "Meeting", "通知类": "Notice",
               "沟通类": "General", "其他类": "Other"}

STATUS_EN = {"待确认": "TBD", "全天": "All day", "待参会": "Upcoming",
             "已结束": "Ended", "进行中": "In progress"}

MODE_EN = {"全量重建": "Full rebuild", "增补": "Incremental", "无新增": "No new mail"}


def _label(text: str, lang: str) -> str:
    """表头等标签的显示文本。"""
    return LABELS.get(lang, {}).get(text, text)


def _category_label(value: str, lang: str) -> str:
    return CATEGORY_EN.get(value, value) if lang == "en" else value


def _status_label(value: str, lang: str) -> str:
    return STATUS_EN.get(value, value) if lang == "en" else value

# 分类只看正文前这么多字：长邮件后半段的关键词不参与评分（口径固定，便于解释）
BODY_MATCH_LIMIT = 2000


def _date_text(stamp: str) -> str:
    """邮件头 Date → 本机时区的 "YYYY-MM-DD HH:MM"。"""
    if not stamp:
        return ""
    moment = None
    for attempt in (stamp, stamp.replace("Z", "+00:00")):
        try:
            moment = datetime.fromisoformat(attempt)
            break
        except ValueError:
            continue
    if moment is None:
        try:  # RFC 5322 原始写法（如 "Wed, 12 Nov 2026 15:00:00 +0100"）
            moment = parsedate_to_datetime(stamp)
        except (TypeError, ValueError, IndexError):
            return ""
    if moment.tzinfo is not None:
        moment = moment.astimezone()  # 换算成本机时区，跨时区邮件也按同一把尺子排
    return moment.strftime("%Y-%m-%d %H:%M")


def _received_text(record: dict) -> str:
    """收件时间：取邮件头 Date，统一换算到本机时区显示（与排序口径一致）。"""
    return _date_text(record.get("date", "") or "")


def _header_date_text(path: Path) -> str:
    """只读邮件头取收件时间：给附件目录加日期前缀，不必先解析全文。"""
    try:
        with path.open("rb") as handle:
            headers = BytesHeaderParser().parse(handle)
        return _date_text(headers.get("Date", "") or "")
    except Exception:
        return ""


def _atomic_write(target: Path, write_func) -> Path:
    """先写临时文件再替换；目标被占用（如正在 Excel 中打开）时另存为不冲突的新名字。"""
    tmp = target.with_name(f"{target.name}.{os.getpid()}.tmp")
    write_func(tmp)
    try:
        os.replace(tmp, target)
        return target
    except OSError:
        pass
    # 目标被占用：换一个当前不存在的名字，绝不覆盖用户已有的 _new 文件
    for index in range(100):
        suffix = "" if index == 0 else str(index)
        fallback = target.with_name(f"{target.stem}_new{suffix}{target.suffix}")
        if fallback.exists():
            continue
        try:
            os.replace(tmp, fallback)
            print(f"警告：{target.name} 被占用，结果另存为 {fallback.name}", file=sys.stderr)
            return fallback
        except OSError:
            continue
    print(f"警告：无法替换 {target}，临时文件保留在 {tmp}", file=sys.stderr)
    return tmp


# ---------------------------------------------------------------- xlsx 写出（仅标准库）

_XLSX_CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"""

_XLSX_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

_XLSX_WB_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

# 样式索引：0 默认、1 表头、2 正文（顶端对齐）、3 正文（居中）
_XLSX_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><name val="微软雅黑"/><color rgb="FF1F2329"/></font><font><b/><sz val="11"/><name val="微软雅黑"/><color rgb="FF1F2329"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE1EAFF"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFB8C7D9"/></left><right style="thin"><color rgb="FFB8C7D9"/></right><top style="thin"><color rgb="FFB8C7D9"/></top><bottom style="thin"><color rgb="FFB8C7D9"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="4">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def _xml_text(value) -> str:
    """XML 文本转义 + 去掉 XML 1.0 不允许的控制字符 + 统一换行为 \\n。"""
    text = str(value if value is not None else "")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = "".join(ch for ch in text
                   if ch in "\t\n" or (0x20 <= ord(ch) and ord(ch) != 0xFFFE and ord(ch) != 0xFFFF))
    return (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _col_letter(index: int) -> str:
    letters = ""
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _xlsx_visual_width(text: str) -> float:
    """按列宽单位估算文本宽度：中日韩字符按 2，其余按 1。"""
    return sum(2 if ord(ch) > 0x2E7F else 1 for ch in str(text or ""))


def _xlsx_line_count(text: str, width: float) -> int:
    """估算单元格在给定列宽下要占几行（按 \n 分段后再按宽度折行）。"""
    if not text:
        return 1
    total = 0
    for part in str(text).split("\n"):
        span = _xlsx_visual_width(part)
        total += max(1, -(-span // max(width - 1, 1)))
    return total


def _xlsx_sheet_xml(table: list[list[str]], widths: list[float],
                    title: str | None, center_cols: set[int]) -> str:
    columns = max((len(row) for row in table), default=0)
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
             ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">']
    merged_rows = 0
    if title:
        merged_rows = 1
        parts.append('<sheetViews><sheetView workbookViewId="0">'
                     f'<pane ySplit="{merged_rows + 1}" topLeftCell="A{merged_rows + 2}"'
                     ' activePane="bottomLeft" state="frozen"/>'
                     '</sheetView></sheetViews>')
    else:
        parts.append('<sheetViews><sheetView workbookViewId="0">'
                     '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
                     '</sheetView></sheetViews>')
    parts.append("<cols>")
    for index in range(columns):
        width = widths[index] if index < len(widths) else 30
        parts.append(f'<col min="{index + 1}" max="{index + 1}" width="{width}" customWidth="1"/>')
    parts.append("</cols><sheetData>")

    row_number = 1
    if title:
        height = 28
        parts.append(f'<row r="1" ht="{height}" customHeight="1">'
                     f'<c r="A1" s="1" t="inlineStr">'
                     f'<is><t xml:space="preserve">{_xml_text(title)}</t></is></c>'
                     '</row>')
        row_number = 2

    for offset, row in enumerate(table):
        header = offset == 0
        style = 1 if header else 2
        cells = []
        for col_index in range(columns):
            value = row[col_index] if col_index < len(row) else ""
            ref = f"{_col_letter(col_index + 1)}{row_number}"
            cell_style = style if header or col_index not in center_cols else 3
            cells.append(f'<c r="{ref}" s="{cell_style}" t="inlineStr">'
                         f'<is><t xml:space="preserve">{_xml_text(value)}</t></is></c>')
        if header:
            height = 24
        else:
            lines = max(_xlsx_line_count(row[i] if i < len(row) else "",
                                         widths[i] if i < len(widths) else 30)
                        for i in range(columns)) if columns else 1
            height = min(300, max(20, lines * 15 + 5))
        parts.append(f'<row r="{row_number}" ht="{height}" customHeight="1">'
                     + "".join(cells) + "</row>")
        row_number += 1

    parts.append("</sheetData>")
    last_row = row_number - 1
    if columns:
        # 顺序按 ECMA-376 的 CT_Worksheet 序列：autoFilter 在 mergeCells 之前
        parts.append(f'<autoFilter ref="A{1 + merged_rows}:{_col_letter(columns)}{last_row}"/>')
    if title and columns:
        parts.append(f'<mergeCells count="1"><mergeCell ref="A1:{_col_letter(columns)}1"/></mergeCells>')
    parts.append("</worksheet>")
    return "".join(parts)


def write_xlsx(table: list[list[str]], target: str | Path, sheet_name: str,
               widths: list[float], title: str | None = None,
               center_cols: set[int] | None = None) -> Path:
    """把二维表写成单工作表 xlsx（仅标准库：zipfile + SpreadsheetML）。

    table 的第 0 行是表头；title 非空时在最上方加一行合并标题；center_cols 里的
    列按居中显示（如状态列），其余列顶端对齐；行高按文本折行行数估算。
    """
    target = Path(target)
    center_cols = center_cols or set()
    sheet_xml = _xlsx_sheet_xml(table, widths, title, center_cols)
    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
                ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                f'<sheets><sheet name="{_xml_text(sheet_name)[:31]}" sheetId="1" r:id="rId1"/></sheets>'
                '</workbook>')

    def _write(path: Path) -> None:
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
            package.writestr("[Content_Types].xml", _XLSX_CT)
            package.writestr("_rels/.rels", _XLSX_RELS)
            package.writestr("xl/workbook.xml", workbook)
            package.writestr("xl/_rels/workbook.xml.rels", _XLSX_WB_RELS)
            package.writestr("xl/styles.xml", _XLSX_STYLES)
            package.writestr("xl/worksheets/sheet1.xml", sheet_xml)

    return _atomic_write(target, _write)


BRIEF_WIDTHS = [34, 20, 40, 90, 12]


def _brief_rows(records: list[dict], lang: str = "zh") -> list[list[str]]:
    table = [[_label(field, lang) for field in BRIEF_FIELDS]]
    for record in records:
        table.append([
            record.get("from", ""),
            _received_text(record),
            record.get("subject", ""),
            record.get("body_text") or "",
            _category_label(record.get("category", "其他类"), lang),
        ])
    return table


def write_brief_table(records: list[dict], target: str | Path, lang: str = "zh") -> Path:
    """精简版清单：只保留发件人、收件时间、主题、正文、分类，写成 xlsx。"""
    return write_xlsx(_brief_rows(records, lang), target, "eml_result", BRIEF_WIDTHS)


def write_full_table(records: list[dict], target: str | Path,
                     with_body: bool = True, lang: str = "zh") -> Path:
    """完整字段清单，写成 xlsx（附件列拼成文件名列表）。"""
    fields = CSV_FIELDS if with_body else [f for f in CSV_FIELDS if f != "body_text"]
    widths = [30, 20, 18, 40, 40, 24, 24, 16, 60]
    widths += [36] * max(0, len(fields) - len(widths))
    table = [[_label(field, lang) for field in fields]]
    for record in records:
        row = dict(record)
        row["attachment_names"] = "; ".join(
            item.get("filename", "") for item in (record.get("attachments") or [])
            if isinstance(item, dict))
        table.append([str(row.get(field, "")) for field in fields])
    return write_xlsx(table, target, "eml_result", widths[:len(fields)])


def _prepare_attachment_dirs(files: list[Path], to_parse: list[Path], attach_root,
                             legacy: dict[str, str], subdirs: dict[str, str], log) -> None:
    """附件目录按「收件日期 + 原目录名」命名：先把旧的（无日期）目录改名，再清空本轮重写的目录。

    改名对所有文件都做（只是换名，不丢文件）；清空只对本轮真正要重新解析的邮件做，
    否则增补模式下没重写的旧附件会被清掉。目录只属于一封邮件（同名邮件用内容哈希
    后缀区分），所以清空它在重跑同一封邮件时能避免堆出 xxx_1、xxx_2 的重复副本。
    """
    if attach_root is None:
        return
    root = Path(attach_root)
    for path in files:
        key = _path_key(path)
        base, name = legacy.get(key) or "", subdirs.get(key) or ""
        if not base or not name or base == name:
            continue
        old_dir, new_dir = root / base, root / name
        if old_dir.is_dir() and not new_dir.exists():
            try:
                os.replace(old_dir, new_dir)
            except OSError as exc:
                log(f"警告：无法重命名附件目录 {base}：{exc}")
                continue
    for path in to_parse:
        key = _path_key(path)
        new_dir = root / (subdirs.get(key) or "")
        if not subdirs.get(key) or not new_dir.is_dir():
            continue
        for child in new_dir.iterdir():
            try:
                if child.is_dir():
                    shutil.rmtree(child, ignore_errors=True)
                else:
                    child.unlink()
            except OSError as exc:
                log(f"警告：无法清理旧附件 {child}：{exc}")


def _sync_attachment_paths(records: list[dict], legacy: dict[str, str],
                           subdirs: dict[str, str]) -> None:
    """改名后把上一轮记录里的 saved_path 指到新目录，避免界面定位失效。"""
    for record in records:
        key = _path_key(record.get("file", ""))
        base, name = legacy.get(key), subdirs.get(key)
        if not base or not name or base == name:
            continue
        for item in record.get("attachments") or []:
            if not isinstance(item, dict):
                continue
            saved = item.get("saved_path") or ""
            if not saved:
                continue
            path = Path(saved)
            if path.parent.name == base:
                item["saved_path"] = str(path.parent.parent / name / path.name)


def _index_entry(record: dict) -> dict:
    """写进索引文件的精简记录：只留界面与导出要用的字段。

    不放 body_full（清洗前的整段原文，体积最大且只在回溯时有用）——原文页需要时
    直接从 .eml 重新解析；源文件不在了就退回 body_text。
    """
    attachments = []
    for item in record.get("attachments") or []:
        if isinstance(item, dict):
            attachments.append({"filename": item.get("filename", ""),
                                "content_type": item.get("content_type", ""),
                                "size_bytes": item.get("size_bytes", 0),
                                "saved_path": item.get("saved_path", ""),
                                "save_error": item.get("save_error", "")})
    return {
        "file": record.get("file", ""),
        "name": record.get("name", ""),
        "subject": record.get("subject", ""),
        "from": record.get("from", ""),
        "from_addr": record.get("from_addr", ""),
        "to": record.get("to", ""),
        "cc": record.get("cc", ""),
        "date": record.get("date", ""),
        "message_id": record.get("message_id", ""),
        "category": record.get("category", ""),
        "category_hit": record.get("category_hit", ""),
        "body_text": record.get("body_text", ""),
        "body_chars": record.get("body_chars", 0),
        "attachment_count": record.get("attachment_count", len(attachments)),
        "attachments": attachments,
        "parse_error": record.get("parse_error", ""),
        "ics_events": record.get("ics_events") or [],
        "body_events": record.get("body_events") or [],
    }


def load_index(path: str | Path) -> dict[str, dict]:
    """读回索引里已有的条目（按文件绝对路径索引），增量运行用它保住旧邮件的数据。"""
    path = Path(path)
    if not path.is_file():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    merged = {}
    for item in payload.get("records") or []:
        if isinstance(item, dict) and item.get("file"):
            merged[_path_key(item["file"])] = item
    return merged


INDEX_FILE = ".eml_index.json"


def write_index(records: list[dict], target: str | Path,
                previous: dict[str, dict] | None = None) -> Path:
    """写界面索引：单文件 JSON，按当前记录顺序排列。

    已经有条目、且本轮没有重新解析的邮件，沿用旧条目（基线摘要是精简过的，
    不能拿来重写索引，否则正文/附件路径会被覆盖丢失）。
    """
    target = Path(target)
    previous = previous or {}
    entries = []
    for record in records:
        key = _path_key(record.get("file", ""))
        fresh = _index_entry(record)
        old = previous.get(key) or {}
        for field in ("body_text", "attachments", "body_chars"):
            if not fresh.get(field) and old.get(field):
                fresh[field] = old[field]
        entries.append(fresh)
    payload = {"version": 1,
               "updated": datetime.now().isoformat(timespec="seconds"),
               "records": entries}

    def _write(path: Path) -> None:
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")

    return _atomic_write(target, _write)


def write_ics(records: list[dict], target_dir: str | Path) -> list[Path]:
    """把邮件里收到的日历邀请原样导出为 .ics（双击可由日历客户端导入）。"""
    target_dir = Path(target_dir)
    written: list[Path] = []
    for record in records:
        raws = record.get("ics_raw") or []
        if not raws:
            continue
        target_dir.mkdir(parents=True, exist_ok=True)
        stem = _safe_name(Path(record["name"]).stem, "meeting")
        for index, raw in enumerate(raws, 1):
            base = f"{stem}.ics" if len(raws) == 1 else f"{stem}_{index}.ics"
            path = target_dir / base
            count = 1
            while path.exists():  # 同名邮件不能互相覆盖邀请文件
                path = target_dir / f"{Path(base).stem}_{count}.ics"
                count += 1
            # iCalendar 规定使用 CRLF 换行
            body = "\r\n".join(raw.replace("\r", "").split("\n"))
            path.write_bytes((body.rstrip("\r\n") + "\r\n").encode("utf-8"))
            written.append(path)
    return written


_WINDOWS_TZ_OFFSETS = {
    "w. central africa standard time": 1,
    "utc": 0,
    "gmt standard time": 0,
    "china standard time": 8,
    "singapore standard time": 8,
    "tokyo standard time": 9,
    "eastern standard time": -5,
    "central standard time": -6,
    "mountain standard time": -7,
    "pacific standard time": -8,
    # Windows 上没有 tzdata 时 zoneinfo 查不到库，常用地区按固定偏移兜底
    "asia/shanghai": 8, "asia/hong_kong": 8, "asia/taipei": 8, "asia/singapore": 8,
    "asia/tokyo": 9, "asia/seoul": 9, "asia/dubai": 4, "asia/kolkata": 5.5,
    "asia/bangkok": 7, "asia/jakarta": 7,
    "africa/lagos": 1, "africa/accra": 0, "africa/cairo": 2, "africa/johannesburg": 2,
    "europe/london": 0, "europe/paris": 1, "europe/berlin": 1, "europe/moscow": 3,
    "america/new_york": -5, "america/chicago": -6, "america/denver": -7,
    "america/los_angeles": -8, "america/sao_paulo": -3,
    "australia/sydney": 10, "pacific/auckland": 12,
}


def _event_tzinfo(tzid: str):
    """按邀请里的 tzid 取时区（IANA 名 / UTC±hh:mm / 常见 Windows 时区名）；认不出返回 None。"""
    label = (tzid or "").strip()
    if not label:
        return None
    if ZoneInfo is not None:
        for candidate in re.findall(r"[A-Za-z]+/[A-Za-z_]+", label):
            try:
                return ZoneInfo(candidate)
            except Exception:
                pass
        try:
            return ZoneInfo(label)
        except Exception:
            pass
    match = re.search(r"UTC\s*([+-])\s*(\d{1,2})(?::?(\d{2}))?", label, re.IGNORECASE)
    if match:
        sign = 1 if match.group(1) == "+" else -1
        return timezone(sign * timedelta(hours=int(match.group(2)),
                                          minutes=int(match.group(3) or 0)))
    for token in re.findall(r"[A-Za-z]+/[A-Za-z_]+", label):
        offset = _WINDOWS_TZ_OFFSETS.get(token.lower())
        if offset is not None:
            return timezone(timedelta(hours=offset))
    if "北京" in label or "中国" in label:
        return timezone(timedelta(hours=8))
    offset = _WINDOWS_TZ_OFFSETS.get(label.lower())
    if offset is not None:
        return timezone(timedelta(hours=offset))
    return None


def write_meeting_table(records: list[dict], target: str | Path, lang: str = "zh") -> Path:
    """把会议类邮件整理成一张日程表（xlsx），供排期或导入 Excel 使用。

    字段按使用顺序排列：组织者、议题、时间、形式、安排、会议号/链接、参会人、状态、来源。
    """
    target = Path(target)
    fields = ["会议组织者", "会议议题", "会议时间", "会议形式",
              "会议安排", "会议号/链接", "参会人", "状态", "来源邮件"]
    body_origin = "（正文通知，无日历邀请）" if lang == "zh" \
        else " (from email body, no calendar invite)"
    no_time = "未给出时间" if lang == "zh" else "Not given"
    no_end = "（结束时间未注明）" if lang == "zh" else " (end time not specified)"
    no_link = "（正文未给入会链接）" if lang == "zh" else " (no join link in body)"
    now = datetime.now(timezone.utc)
    local_zone = datetime.now().astimezone().tzinfo
    rows = []
    for record in records:
        events = [(event, "") for event in (record.get("ics_events") or [])]
        events += [(event, body_origin) for event in (record.get("body_events") or [])]
        for event, origin in events:
            start_text = event.get("start") or ""
            end_text = event.get("end") or ""
            status = "待确认"
            start_dt = end_dt = None
            try:
                start_dt = datetime.strptime(start_text, "%Y-%m-%d %H:%M")
                end_dt = datetime.strptime(end_text, "%Y-%m-%d %H:%M") if end_text else start_dt
            except (ValueError, TypeError):
                start_dt = end_dt = None
                if start_text.endswith("（全天）"):
                    status = "全天"
            sort_key = float("-inf")
            if start_dt is not None:
                # 状态要按邀请所在时区判断，否则北京时间 12:30 的会被当成拉各斯时间
                zone = _event_tzinfo(event.get("tzid") or "") or local_zone
                start_cmp = start_dt.replace(tzinfo=zone)
                end_cmp = (end_dt or start_dt).replace(tzinfo=zone)
                sort_key = start_cmp.timestamp()
                if now < start_cmp:
                    status = "待参会"
                elif now > end_cmp:
                    status = "已结束"
                else:
                    status = "进行中"

            if start_text and end_text:
                if start_text[:10] == end_text[:10]:
                    when = f"{start_text}–{end_text[11:]}"
                else:
                    when = f"{start_text}–{end_text}"
            else:
                when = start_text or no_time
            if not end_text and " " in start_text:
                when += no_end
            zone = f"（{event['tzid']}）" if event.get("tzid") else ""

            organizer = event.get("organizer_name") or event.get("organizer") \
                or record.get("from", "") or ""
            if (event.get("organizer") and event.get("organizer_name")
                    and event["organizer_name"].lower() != event["organizer"].lower()):
                organizer = f"{event['organizer_name']} <{event['organizer']}>"

            rows.append({
                "会议组织者": organizer,
                "会议议题": event.get("summary") or record.get("subject", ""),
                "会议时间": f"{when}{zone}",
                "会议形式": event.get("form") or _meeting_form(event),
                "会议安排": _agenda_text(event),
                "会议号/链接": _join_info(event) or (no_link if origin else ""),
                "参会人": _attendee_text(event),
                "状态": _status_label(status, lang),
                "来源邮件": f"{record.get('name', '')}{origin}",
                "_sort": sort_key,
            })
    rows.sort(key=lambda row: row["_sort"], reverse=True)

    table = [[_label(field, lang) for field in fields]]
    table += [[str(row.get(field, "")) for field in fields] for row in rows]
    return write_xlsx(table, target, "Meetings" if lang == "en" else "会议日程",
                      [36, 44, 34, 22, 50, 62, 42, 10, 46],
                      title=_label("会议日程表", lang), center_cols={7})


def write_summary(records: list[dict], target: str | Path, lang: str = "zh") -> Path:
    """生成分类汇总 Markdown：先统计，再按类别列出每封邮件。"""
    target = Path(target)
    groups: dict[str, list[dict]] = {}
    for record in records:
        groups.setdefault(record["category"], []).append(record)

    order = ["会议类", "通知类", "日报类", "沟通类", "其他类"]
    order += [c for c in groups if c not in order]
    en = lang == "en"
    lines = ["# Email summary" if en else "# 邮件分类汇总", ""]
    lines.append(f"{len(records)} emails in {len(groups)} categories."
                 if en else f"共 {len(records)} 封邮件，分为 {len(groups)} 类。")
    lines += ["", "| Category | Count | Share |" if en else "| 分类 | 封数 | 占比 |",
              "| --- | --- | --- |"]
    for category in order:
        items = groups.get(category)
        if not items:
            continue
        pct = f"{len(items) / len(records) * 100:.0f}%" if records else "0%"
        lines.append(f"| {_category_label(category, lang)} | {len(items)} | {pct} |")

    for category in order:
        items = groups.get(category)
        if not items:
            continue
        name = _category_label(category, lang)
        lines += ["", (f"## {name} ({len(items)})" if en else f"## {name}（{len(items)} 封）"), ""]
        for record in items:
            stamp = _received_text(record) or ("no date" if en else "时间未知")
            lines.append(f"- **{record['subject'] or ('(no subject)' if en else '（无主题）')}**")
            if en:
                lines.append(f"  - From: {record['from'] or '(unknown)'} | "
                             f"Received: {stamp} | Matched by: {record['category_hit']}")
            else:
                lines.append(f"  - 发件人：{record['from'] or '（未知）'} ｜ 收件时间：{stamp} "
                             f"｜ 判定依据：{record['category_hit']}")
            for event in record.get("ics_events") or []:
                when = (" to " if en else " 至 ").join(
                    x for x in (event.get("start", ""), event.get("end", "")) if x)
                zone = f" ({event['tzid']})" if (en and event.get("tzid")) \
                    else (f"（{event['tzid']} 时区）" if event.get("tzid") else "")
                place = (f" | Location: {event['location']}" if en
                         else f" ｜ 地点：{event['location']}") if event.get("location") else ""
                label = "Meeting time" if en else "会议时间"
                empty = "(time not given in invite)" if en else "邀请中未给出时间"
                lines.append(f"  - **{label}: {when or empty}**{zone}{place}")
            for event in record.get("body_events") or []:
                when = event.get("start", "") or ("time not given in body" if en
                                                  else "正文未给出时间")
                zone = f" ({event['tzid']})" if (en and event.get("tzid")) \
                    else (f"（{event['tzid']}）" if event.get("tzid") else "")
                place = (f" | Location: {event['location']}" if en
                         else f" ｜ 地点：{event['location']}") if event.get("location") else ""
                tail = " (from email body, no invite; added to schedule)" if en \
                    else "（据邮件正文通知，无日历邀请，已入会议日程）"
                label = "Meeting time" if en else "会议时间"
                lines.append(f"  - **{label}: {when}**{zone}{place}{tail}")
            lines.append(f"  - Key points: {_key_point(record)}" if en
                         else f"  - 要点：{_key_point(record)}")

    def _write(path: Path) -> None:
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    return _atomic_write(target, _write)


def print_summary(records: list[dict], out_dir: Path, log=None, lang: str = "zh") -> None:
    emit = log or (lambda text: print(text))
    en = lang == "en"
    total = len(records)
    failed = sum(1 for r in records if r["parse_error"])
    attached = sum(1 for r in records if r["attachment_count"])
    counts: dict[str, int] = {}
    for record in records:
        counts[record["category"]] = counts.get(record["category"], 0) + 1
    if en:
        emit(f"{total} emails, {failed} failed to parse, {attached} with attachments")
        emit("Categories: " + ", ".join(
            f"{_category_label(k, lang)} {v}" for k, v in counts.items()))
        emit(f"Output: {out_dir}")
    else:
        emit(f"共 {total} 封，解析失败 {failed} 封，含附件 {attached} 封")
        emit("分类：" + "，".join(f"{k} {v} 封" for k, v in counts.items()))
        emit(f"输出目录：{out_dir}")
    for record in records[:10]:
        flag = "!" if record["parse_error"] else " "
        note = f"  [{record['parse_error']}]" if record["parse_error"] else ""
        category = _category_label(record["category"], lang)
        empty = "(no date)" if en else "(无日期)"
        subject = record["subject"] or ("(no subject)" if en else "(无主题)")
        emit(f"{flag} {_received_text(record) or empty:16} | {category:14} | "
             f"{subject[:34]} | {record['name']}{note}")
    if total > 10:
        emit(f"... {total - 10} more in the output files" if en
             else f"…… 其余 {total - 10} 封见输出文件")


# ---------------------------------------------------------------- 增量状态

STATE_FILE = ".eml_state.json"

# 抽取规则版本：正文清洗、分类、日程提取的口径一改就加 1，
# 让旧基线自动失效、下次运行时全量重做，避免新旧口径混在一张表里。
RULES_VERSION = 13

# 增补模式下需要保留的字段：足够重建 CSV / 汇总 / 会议日程，不必重新解析旧邮件
INCREMENTAL_FIELDS = (
    "file", "name", "size_bytes", "date", "subject", "from", "from_addr",
    "to", "cc", "message_id", "body_text", "body_chars", "attachment_count",
    "attachments", "ics_events", "body_events", "parse_error", "category",
    "category_hit",
)


def _normalize_message_id(value: str) -> str:
    """Message-ID 归一化：去掉尖括号、空白、统一小写；无 ID 返回空串。"""
    text = (value or "").strip().strip("<>").strip().lower()
    return text


def dedupe_records(entries: list[tuple[str, str, dict]]) -> tuple[list[dict], list[dict]]:
    """按 Message-ID 与「文件名 + 内容指纹」去重。

    entries 为 (路径, 内容指纹, 记录)，顺序应为时间倒序。判定规则：
    1) Message-ID 相同 → 同一封邮件被保存多次，只留一封（保留先出现的那封）；
    2) 文件名相同且内容指纹相同 → 完全重复，跳过；
    3) 文件名相同但内容不同 → 不是重复，保留（附件目录另起后缀，互不覆盖）。
    返回 (保留的记录, [{path, name, reason}]) 被跳过的记录。
    """
    seen_mid: dict[str, str] = {}
    seen_name: dict[str, set[str]] = {}
    kept: list[dict] = []
    skipped: list[dict] = []

    for path, fingerprint, record in entries:
        name = record.get("name", "") or Path(path).name
        message_id = _normalize_message_id(record.get("message_id", ""))
        reason = ""
        if message_id and message_id in seen_mid:
            reason = f"Message-ID 与「{seen_mid[message_id]}」相同"
        elif fingerprint in seen_name.get(name, set()):
            reason = f"与同名的「{name}」内容完全相同"
        if reason:
            skipped.append({"path": path, "name": name, "reason": reason,
                            "sha256": fingerprint})
            continue
        if message_id:
            seen_mid[message_id] = name
        # 同名可能有多份不同内容，全部记下指纹：第三个相同内容才能被判重
        seen_name.setdefault(name, set()).add(fingerprint)
        kept.append(record)
    return kept, skipped


def _file_fingerprint(path: Path) -> str:
    """按文件内容算指纹，判断这封邮件是否已经处理过、是否被改动。"""
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _record_digest(record: dict) -> dict:
    return {key: record.get(key) for key in INCREMENTAL_FIELDS}


def load_state(out_dir: Path) -> dict | None:
    path = out_dir / STATE_FILE
    if not path.is_file():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        # 损坏的基线先留证，别让下一次 save_state 静默覆盖掉增量信息
        backup = path.with_name(path.name + ".bad")
        try:
            if not backup.exists():
                path.rename(backup)
            print(f"警告：基线文件无法解析（{exc}），已备份为 {backup.name}，本轮按首次运行处理",
                  file=sys.stderr)
        except OSError:
            print(f"警告：基线文件无法解析（{exc}），且备份失败", file=sys.stderr)
        return None


def _path_key(path) -> str:
    """文件的统一键：绝对路径 + 小写，避免同一邮件用相对/绝对路径各算一份。"""
    text = str(path or "").strip()
    if not text:
        return ""
    return os.path.normcase(str(Path(text).resolve()))


def save_state(out_dir: Path, records: list[dict],
               fingerprints: dict[str, str], shape: str,
               skipped: list[dict] | None = None) -> Path:
    payload = {
        "version": 1,
        "rules": RULES_VERSION,
        "shape": shape,
        "updated": datetime.now().isoformat(timespec="seconds"),
        "files": [
            {"path": record.get("file", ""),
             "sha256": fingerprints.get(_path_key(record.get("file", "")), ""),
             "record": _record_digest(record)}
            for record in records
        ],
        # 已判定为重复、没有再进清单的文件：记下来，下次运行不必反复重新解析
        "skipped": skipped or [],
    }

    def _write(path: Path) -> None:
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    return _atomic_write(out_dir / STATE_FILE, _write)


# ---------------------------------------------------------------- 命令行

def collect_inputs(paths: list[str], recursive: bool, pattern: str) -> list[Path]:
    found: list[Path] = []
    for raw in paths:
        item = Path(raw)
        if item.is_dir():
            found.extend(item.rglob(pattern) if recursive else item.glob(pattern))
        elif item.is_file():
            found.append(item)
        elif globlib.has_magic(raw):
            # 通配符：交给 glob 处理（pathlib.glob 不接受绝对路径模式）
            found.extend(Path(match) for match in globlib.glob(raw, recursive=recursive))
        else:
            raise SystemExit(f"路径不存在：{raw}")
    return sorted({path.resolve() for path in found if path.is_file()})


def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        description="批量读取 .eml 邮件文件，导出 xlsx 清单与汇总，可选导出附件。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    ap.add_argument("paths", nargs="+", help="eml 文件、所在目录或通配符")
    ap.add_argument("-o", "--out", default=None, help="输出目录，默认 ./eml_output")
    ap.add_argument("--recursive", action="store_true", help="递归搜索子目录")
    ap.add_argument("--pattern", default="*.eml", help="文件匹配规则，默认 *.eml")
    ap.add_argument("--attachments", action="store_true", help="导出附件到 <out>/attachments")
    ap.add_argument("--no-body", action="store_true", help="--full 模式下 xlsx 不写正文列（正文仍保留在索引里）")
    ap.add_argument("--full", action="store_true",
                    help="输出完整字段（文件、日期、收件人、附件等），默认只输出发件人/主题/正文/分类")
    ap.add_argument("--no-summary", action="store_true", help="不生成分类汇总 eml_summary.md")
    ap.add_argument("--no-calendar", action="store_true", help="不导出邮件里的日历邀请（默认导出到 <out>/calendar）")
    ap.add_argument("--quiet", action="store_true", help="不打印摘要")
    ap.add_argument("--rebuild", action="store_true",
                    help="忽略上次的处理记录，强制全量重建（默认是：有基础文件则只增补新邮件）")
    ap.add_argument("--lang", choices=["zh", "en"], default="zh",
                    help="输出表头与日志语言：zh（默认）或 en（英文表头）")
    return ap


class BatchCancelled(Exception):
    """用户中止批量处理。"""


def run_batch(args, log=None, progress=None, cancelled=None, lang: str | None = None) -> dict:
    """执行一次批量处理，返回结果摘要。

    命令行（main）与可视化界面共用这一条路径，避免两边逻辑分叉。
    log(text) 输出过程信息；progress(阶段, 已完成, 总数) 报告进度；
    cancelled() 返回 True 时尽快中止并抛出 BatchCancelled。
    lang 决定输出表头与日志语言（"zh"/"en"），缺省取 args.lang。
    """
    log = log or (lambda text: print(text))
    lang = (lang or getattr(args, "lang", "zh") or "zh").lower()
    if lang not in LABELS:
        lang = "zh"
    en = lang == "en"

    def _check_cancel() -> None:
        if cancelled is not None and cancelled():
            raise BatchCancelled()

    files = collect_inputs(args.paths, args.recursive, args.pattern)
    if not files:
        log("未找到任何 .eml 文件，请检查路径与 --pattern")
        return {"ok": False, "error": "未找到任何 .eml 文件", "records": [], "mode": "",
                "reason": "", "duplicates": [], "file_count": 0}

    out_dir = (Path(args.out) if args.out else Path.cwd() / "eml_output").resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    # 附件路径要落成绝对路径：界面按这些路径打开附件，相对路径换个工作目录就失效
    attach_root = out_dir / "attachments" if args.attachments else None

    shape = "brief" if not args.full else ("full-no-body" if args.no_body else "full")
    shape = f"{shape}:{lang}"  # 语言切换也算列结构变化，触发全量重写，避免表头中英混排

    if progress is not None:
        progress("读取文件", 0, len(files))
    fingerprints = {_path_key(path): _file_fingerprint(path) for path in files}

    # 判断走增补还是全量：没有基线、列结构变化、基础文件缺失或不匹配、旧邮件被移走/改动，都退回全量
    base_entries: list[tuple[str, str, dict]] = []
    new_paths: list[Path] = list(files)
    mode, reason = "全量重建", "首次生成（无基础文件）"
    state = None if args.rebuild else load_state(out_dir)
    if args.rebuild:
        reason = "指定 --rebuild"
    elif state is not None:
        if state.get("rules") != RULES_VERSION:
            reason = f"抽取规则已更新（v{state.get('rules')} → v{RULES_VERSION}）"
        elif state.get("shape") != shape:
            reason = "输出列结构已变化"
        elif not (out_dir / "eml_result.xlsx").is_file():
            reason = "基础清单缺失或列不匹配"
        else:
            known = {_path_key(item.get("path", "")): item for item in state.get("files", [])}
            # 上一轮判定为重复的文件也算“已知”，否则每轮都会被当成新邮件重算一遍
            known.update({_path_key(item.get("path", "")): {"sha256": item.get("sha256", "")}
                          for item in state.get("skipped", [])})
            moved = [key for key in known if key not in fingerprints]
            changed = [key for key, sha in fingerprints.items()
                       if key in known and known[key].get("sha256") != sha]
            if moved:
                reason = f"有 {len(moved)} 封邮件已移出文件夹"
            elif changed:
                reason = f"有 {len(changed)} 封邮件内容被修改"
            else:
                base_entries = [
                    (item.get("path", ""), item.get("sha256", ""),
                     dict(item.get("record") or {}))
                    for item in state.get("files", [])
                ]
                new_paths = [path for path in files if _path_key(path) not in known]
                mode = "增补" if new_paths else "无新增"
                reason = f"新增 {len(new_paths)} 封"

    if args.full and mode != "全量重建":
        # 基线里只存了精简字段，拿它重写完整清单会把收件人/正文原文等列清空
        mode, reason = "全量重建", "完整清单模式下需重读全部邮件以补齐字段"
        base_entries, new_paths = [], list(files)

    # 同名邮件（文件名相同、内容不同）各占一个附件子目录，避免互相清空；
    # 子目录名按内容哈希固定，新邮件插入不会让旧邮件换目录、清掉别人的附件。
    # 目录名统一带收件日期前缀（2026-11-12 xxx），便于在资源管理器里按时间找。
    legacy_dirs: dict[str, str] = {}
    subdirs: dict[str, str] = {}
    groups: dict[str, list[Path]] = {}
    for path in files:
        groups.setdefault(_safe_name(path.stem, "mail"), []).append(path)
    for stem, group in groups.items():
        for path in group:
            key = _path_key(path)
            base = (stem if len(group) == 1 else f"{stem}_{fingerprints[key][:6]}")
            legacy_dirs[key] = base
            date_text = _header_date_text(path)[:10]
            subdirs[key] = f"{date_text} {base}" if date_text else base

    parsed = []
    _prepare_attachment_dirs(files, new_paths, attach_root, legacy_dirs, subdirs, log)
    if progress is not None:
        progress("解析邮件", 0, len(new_paths))
    for index, path in enumerate(new_paths, 1):
        _check_cancel()
        try:
            parsed.append(read_eml_file(path, save_attachments=args.attachments,
                                        attach_root=attach_root,
                                        attach_subdir=subdirs[_path_key(path)]))
        except Exception as exc:  # 单封失败不影响整批
            parsed.append({"file": str(path), "name": path.name,
                           "parse_error": f"{type(exc).__name__}: {exc}"})
        if progress is not None:
            progress("解析邮件", index, len(new_paths))
    _check_cancel()

    # 按 Message-ID / 同名同内容去重（时间倒序后再去重，保留最新的一份）
    entries = base_entries + [(_path_key(path), fingerprints[_path_key(path)], record)
                              for path, record in zip(new_paths, parsed)]
    ordered = sorted(entries, key=lambda item: _record_sort_key(item[2]))
    records, duplicates = dedupe_records(ordered)
    records = sort_records(records)
    for record in records:
        record["category"], record["category_hit"] = classify(record)

    # 被判定为重复的新邮件：删掉刚导出的附件目录，避免留下孤儿目录
    parsed_keys = {_path_key(path) for path in new_paths}
    kept_keys = {_path_key(record.get("file", "")) for record in records}
    if attach_root is not None:
        for path in new_paths:
            key = _path_key(path)
            if key in parsed_keys and key not in kept_keys:
                stale_dir = Path(attach_root) / subdirs[key]
                if stale_dir.is_dir():
                    for child in stale_dir.iterdir():
                        if child.is_file():
                            try:
                                child.unlink()
                            except OSError as exc:
                                log(f"警告：无法删除残留附件 {child}：{exc}")
                    try:
                        stale_dir.rmdir()
                    except OSError as exc:
                        log(f"警告：无法删除空目录 {stale_dir}：{exc}")
    _sync_attachment_paths(records, legacy_dirs, subdirs)
    kept_new = [record for record in records
                if _path_key(record.get("file", "")) in parsed_keys]

    if args.full:
        table_path = write_full_table(records, out_dir / "eml_result.xlsx",
                                      with_body=not args.no_body, lang=lang)
    else:
        table_path = write_brief_table(records, out_dir / "eml_result.xlsx", lang=lang)
    # 界面索引：单文件 JSON，界面冷启动、元信息页与附件列表都从它读
    index_path = out_dir / INDEX_FILE
    previous_index = load_index(index_path)
    index_path = write_index(records, index_path, previous=previous_index)
    summary_path = out_dir / "eml_summary.md"
    if not args.no_summary:
        summary_path = write_summary(records, summary_path, lang=lang)

    ics_paths = [] if args.no_calendar else write_ics(records, out_dir / "calendar")
    if mode == "全量重建" and not args.no_calendar:
        # 全量重建时清掉这次没有重新导出的旧 .ics，避免会议邮件被移走后留下幽灵日程
        keep = {path.name for path in ics_paths}
        for stale in (out_dir / "calendar").glob("*.ics"):
            if stale.name not in keep:
                try:
                    stale.unlink()
                except OSError as exc:
                    log(f"警告：无法删除旧日程文件 {stale}：{exc}")
    has_events = any(r.get("ics_events") or r.get("body_events") for r in records)
    # 全量重建时即使没有会议也要落一张（只有表头）的日程表，避免留着上一轮的旧行
    meeting_path = (write_meeting_table(records, out_dir / "会议日程.xlsx", lang=lang)
                    if has_events or mode == "全量重建" else None)

    # 判重记录要跨轮保留：本轮没解析任何新邮件时 duplicates 为空，
    # 若直接覆盖，下一轮会把这些重复文件当成新增邮件再算一遍
    previous_skipped = (state or {}).get("skipped", [])
    skipped_state = [dict(item) for item in previous_skipped
                     if _path_key(item.get("path", "")) in fingerprints]
    seen_skipped = {_path_key(item.get("path", "")) for item in skipped_state}
    for item in duplicates:
        key = _path_key(item.get("path", ""))
        if key not in seen_skipped:
            skipped_state.append({"path": item["path"],
                                  "sha256": item.get("sha256")
                                  or fingerprints.get(key, ""),
                                  "name": item["name"],
                                  "reason": item["reason"]})
            seen_skipped.add(key)

    state_path = save_state(out_dir, records, fingerprints, shape, skipped=skipped_state)

    if not args.quiet:
        if en:
            log(f"Mode: {MODE_EN.get(mode, mode)}")
        else:
            log(f"处理方式：{mode}（{reason}）")
        if duplicates:
            log(f"Skipped {len(duplicates)} duplicate emails" if en
                else f"去重：跳过 {len(duplicates)} 封重复邮件")
            for item in duplicates:
                log(f"  - {item['name']}: {item['reason']}")
        print_summary(records, out_dir, log, lang=lang)
        log(f"Table : {table_path}" if en else f"清单  : {table_path}")
        log(f"Index : {index_path}" if en else f"索引  : {index_path}")
        if not args.no_summary:
            log(f"Summary: {summary_path}" if en else f"汇总  : {summary_path}")
        if ics_paths:
            log(f"Calendar: {len(ics_paths)} invite(s) exported to {out_dir / 'calendar'}" if en
                else f"日历  : {len(ics_paths)} 个邀请已导出到 {out_dir / 'calendar'}")
        if meeting_path:
            log(f"Meetings: {meeting_path}" if en else f"日程  : {meeting_path}")
        log(f"Baseline: {state_path}" if en else f"基线  : {state_path}")

    return {"ok": True, "mode": mode, "reason": reason, "records": records,
            "duplicates": duplicates, "file_count": len(files),
            "out_dir": str(out_dir), "table_path": str(table_path),
            "jsonl_path": str(index_path), "index_path": str(index_path),
            "summary_path": str(summary_path) if not args.no_summary else "",
            "meeting_path": str(meeting_path) if meeting_path else "",
            "ics_count": len(ics_paths), "state_path": str(state_path),
            "shape": shape}


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        result = run_batch(args)
    except BatchCancelled:
        return 130
    return 0 if result.get("ok") else 2


if __name__ == "__main__":
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    sys.exit(main())
