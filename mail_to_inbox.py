#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把采货侠 / 拍机堂发到网易邮箱的账单，自动写入 ERP 的「待入库」。

拍机堂：按物品编号去重；成本=支出合计+其他调整分摊；成色默认待定。

只使用 Python 标准库 + openpyxl，不需要服务器。
在 GitHub Actions 里每 5 分钟跑一次；也可以本地 --file 测试解析。
"""

import argparse
import csv
import email
import imaplib
import io
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
import zipfile
from datetime import date, timedelta
from email.header import decode_header
from pathlib import Path

import openpyxl

DEFAULT_KEYWORDS = "采货侠,拍机堂,转转,caihuoxia,paijitang,zhuanzhuan"
COND_MAP = {
    "S": "靓机", "A+": "靓机", "A": "靓机",
    "A1": "小花", "A2": "小花",
    "B+": "小花", "B": "小花", "B+1": "小花", "B+2": "小花", "B1": "小花", "B2": "小花",
    "C+": "小花", "C+1": "小花", "C+2": "小花",
    "C": "大花", "C1": "大花", "C2": "大花", "D": "大花", "D+1": "大花", "D1": "大花",
    "E+1": "小花", "G+": "小花",
    "D+2": "外爆", "D+3": "外爆", "E1": "外爆",
    "D3": "后爆", "F+": "后爆", "H+": "后爆",
    "D2": "后爆", "E": "小破烂", "E+": "小破烂", "E+2": "小破烂", "E2": "小破烂", "F": "小破烂",
    "G": "小破烂", "H": "小破烂", "H1": "小破烂", "H2": "小破烂", "K+": "小破烂", "K": "小破烂",
    "G1": "炸弹", "G2": "炸弹", "I+": "炸弹", "I": "炸弹", "I1": "炸弹", "I2": "炸弹",
    "J+": "炸弹", "J": "炸弹",
}
BRANDS = [
    ("iPhone", ("iphone", "苹果", "apple")),
    ("Redmi", ("redmi", "红米")),
    ("荣耀", ("荣耀", "honor")),
    ("华为", ("华为", "huawei")),
    ("小米", ("小米", "xiaomi")),
    ("OPPO", ("oppo",)),
    ("vivo", ("vivo",)),
    ("iQOO", ("iqoo",)),
    ("一加", ("一加", "oneplus")),
    ("realme", ("realme", "真我")),
    ("三星", ("三星", "samsung")),
    ("魅族", ("魅族", "meizu")),
    ("努比亚", ("努比亚", "nubia")),
    ("摩托罗拉", ("摩托罗拉", "motorola")),
]
COLORS = {
    "冰魄蓝": "蓝", "岩晶白": "白", "暗影": "黑", "墨羽": "黑", "幽芒": "黑",
    "光刃": "白", "碳纤黑": "黑", "先锋灰": "灰", "蝶羽翠": "绿",
    "深空黑": "黑", "银色": "银", "白色": "白", "黑色": "黑", "蓝色": "蓝",
    "绿色": "绿", "紫色": "紫", "金色": "金", "灰色": "灰", "粉色": "粉",
}


def log(*a):
    print(*a, flush=True)


def clean_text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def num(v):
    try:
        return float(str(v).replace(",", "").strip())
    except Exception:
        return 0.0


def money(v):
    return round(num(v) + 1e-9, 2)


def decode_mime(v):
    if not v:
        return ""
    try:
        parts = decode_header(v)
        out = []
        for text, enc in parts:
            if isinstance(text, bytes):
                out.append(text.decode(enc or "utf-8", "ignore"))
            else:
                out.append(text)
        return "".join(out)
    except Exception:
        return str(v)


def pick_brand(model, brand_raw=""):
    text = (clean_text(brand_raw) + " " + clean_text(model)).lower()
    for name, aliases in BRANDS:
        if any(a in text for a in aliases):
            return name
    return clean_text(brand_raw) or "其他"


def norm_model(model, brand):
    m = clean_text(model)
    m = re.sub(r"^\s*(苹果|apple|Apple)\s+", "", m)
    if brand == "iPhone" and not m.lower().startswith("iphone"):
        m = "iPhone " + m
    if brand == "Redmi":
        m = re.sub(r"^\s*红米\s*", "Redmi ", m)
    if brand == "荣耀":
        m = re.sub(r"^\s*华为\s*", "", m)
    return m.strip()


def desc_part(desc, key):
    m = re.search(key + r"\s*[:：]\s*([^|]+)", desc or "")
    return m.group(1).strip() if m else ""


def norm_mem(mem):
    m = clean_text(mem)
    m = m.replace("GB", "G").replace("gb", "G").replace("G+", "+")
    m = re.sub(r"(\d+)\s*G\s*\+\s*(\d+)\s*G", r"\1+\2G", m)
    m = re.sub(r"(\d+)\s*\+\s*(\d+)\s*G", r"\1+\2G", m)
    return m


def base_color(color):
    c = clean_text(color)
    if not c:
        return ""
    if c in COLORS:
        return COLORS[c]
    tail = c[-1]
    return tail if tail in "黑白灰青金蓝绿紫银粉橙红黄棕钛" else c


def cond_of_grade(grade):
    g = clean_text(grade).upper()
    if g in COND_MAP:
        return COND_MAP[g]
    if re.match(r"^(I|J)", g) or re.match(r"^G[12]", g):
        return "炸弹"
    if re.match(r"^D", g):
        return "大花"
    if re.match(r"^(C|B|A1|A2)", g):
        return "小花"
    if re.match(r"^A", g):
        return "靓机"
    return "靓机"


def header_index(row):
    vals = [clean_text(x).replace(" ", "") for x in row]
    joined = "|".join(vals).lower()
    if "订单号" in joined and "imei" in joined:
        return {v: i for i, v in enumerate(vals) if v}
    return None


def cell_by(row, cols, *names):
    for name in names:
        idx = cols.get(name)
        if idx is not None and idx < len(row):
            return clean_text(row[idx])
    # 容忍表头末尾空格 / 大小写
    low = {k.lower(): v for k, v in cols.items()}
    for name in names:
        idx = low.get(name.lower())
        if idx is not None and idx < len(row):
            return clean_text(row[idx])
    return ""


def sheets_from_xlsx(data):
    """返回 [(sheet名, 行列表)]，保留其它工作表给拍机堂的「其他调整明细」用。"""
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    try:
        out = []
        for name in wb.sheetnames:
            ws = wb[name]
            try:
                ws.reset_dimensions()
            except Exception:
                pass
            out.append((name, [[c for c in row] for row in ws.iter_rows(values_only=True)]))
        return out
    finally:
        wb.close()


def rows_from_xlsx(data):
    sheets = sheets_from_xlsx(data)
    target = None
    for n, rows in sheets:
        if "买入" in n or "明细" in n:
            target = rows
            break
    return target if target is not None else (sheets[0][1] if sheets else [])


def phone_name_parts(name):
    """拍机堂物品名称 → (品牌型号, 内存, 颜色)。例：Redmi Note 12T Pro 8G+256G 碳纤黑 国行 全网通"""
    text = clean_text(name)
    text = re.sub(r"^\s+", "", text)
    for _ in range(4):
        text = re.sub(r"\s*(国行|港版|美版|日版|欧版|韩版|国际版|全网通|双卡|单卡|公开版|移动版|联通版|电信版)\s*$", "", text, flags=re.I).strip()
    mem = ""
    m = re.search(r"(\d{1,2})\s*(?:G|GB)?\s*[+＋]\s*(\d{2,4})\s*(?:G|GB)?", text, flags=re.I)
    if m:
        mem = m.group(1) + "+" + m.group(2) + "G"
        before = text[:m.start()].strip()
        after = text[m.end():].strip()
    else:
        before, after = text, ""
    model_base = re.sub(r"\s+", " ", before).strip()
    color = ""
    if after:
        token = after.split()[0]
        color = base_color(token)
    brand = pick_brand(model_base)
    model = norm_model(model_base, brand)
    if mem and not re.search(r"\d+\s*\+\s*\d+\s*G", model, flags=re.I):
        model += " " + mem
    if color:
        model += " " + color
    return brand, re.sub(r"\s+", " ", model).strip(), mem, color


def pjt_other_fees(sheets):
    """拍机堂「其他调整明细」→ {订单号: {'channel': 支付通道费, 'flat': 发货服务费等}}。金额取绝对值。"""
    rows = []
    for name, rs in sheets:
        if "其他调整" in name:
            rows = rs
            break
    if not rows:
        return {}
    head_at = -1
    for i, r in enumerate(rows[:6]):
        joined = "|".join(clean_text(x) for x in r)
        if "费用类型" in joined and "实收" in joined:
            head_at = i
            break
    if head_at < 0:
        return {}
    cols = {clean_text(x).replace(" ", ""): j for j, x in enumerate(rows[head_at]) if clean_text(x)}
    fees = {}
    for r in rows[head_at + 1:]:
        desc = cell_by(r, cols, "费用说明")
        m = re.search(r"订单编号[：:]\s*([A-Za-z0-9]+)", desc)
        if not m:
            continue
        order = m.group(1)
        amount = abs(num(cell_by(r, cols, "实收/实扣金额", "实收/实扣金额总计")))
        if not amount:
            continue
        kind = cell_by(r, cols, "费用类型")
        d = fees.setdefault(order, {"channel": 0.0, "flat": 0.0})
        if "支付通道" in kind:
            d["channel"] += amount
        else:
            d["flat"] += amount
    return fees


def allocate_order_fees(items, fee):
    """同一订单多台：通道费按物品金额比例分，发货服务费等按台平均分；尾差补到最后 1 台。"""
    if not items or not fee:
        return
    total = round(fee.get("channel", 0.0) + fee.get("flat", 0.0), 2)
    if total <= 0:
        return
    weights = [max(0.0, num(x.get("_goods_amt"))) for x in items]
    sw = sum(weights)
    channel = round(fee.get("channel", 0.0), 2)
    flat = round(fee.get("flat", 0.0), 2)
    allocated = 0.0
    for i, x in enumerate(items):
        if sw > 0:
            c = round(channel * weights[i] / sw, 2)
        else:
            c = round(channel / len(items), 2)
        f = round(flat / len(items), 2)
        x["_other_fee"] = round(c + f, 2)
        allocated += x["_other_fee"]
    diff = round(total - allocated, 2)
    if items and diff:
        items[-1]["_other_fee"] = round(items[-1].get("_other_fee", 0.0) + diff, 2)


def rows_from_csv(data):
    txt = data.decode("utf-8-sig", "ignore")
    return list(csv.reader(io.StringIO(txt)))


def parse_bill_bytes(data, filename):
    """返回 (records, stats)。支持采货侠 / 拍机堂；退款、取消、非手机不生成记录。"""
    sheets = []
    if filename.lower().endswith(".csv"):
        rows = rows_from_csv(data)
        sheets = [("csv", rows)]
    else:
        sheets = sheets_from_xlsx(data)
    target = None
    for n, rs in sheets:
        if "买入" in n or "明细" in n:
            target = rs
            break
    if target is None:
        target = sheets[0][1] if sheets else []

    stats = {"total": 0, "added": 0, "refunded": 0, "cancel": 0, "other": 0, "no_cost": 0}
    cols = None
    header_at = -1
    for i, r in enumerate(target[:10]):
        cols = header_index(r)
        if cols:
            header_at = i
            break
    if not cols:
        raise ValueError("账单里没找到表头（订单号 / imei）")

    is_pjt = any(k in cols for k in ("物品编号", "物品名称", "支出合计")) and "物品金额" in cols
    records = []

    if is_pjt:
        raw = []
        for r in target[header_at + 1:]:
            po = cell_by(r, cols, "订单号")
            item_no = cell_by(r, cols, "物品编号")
            if not po and not item_no:
                continue
            stats["total"] += 1
            cancel_flag = re.search(r"取消|关闭", cell_by(r, cols, "交易类型"))
            if cell_by(r, cols, "取消时间") or cancel_flag or num(cell_by(r, cols, "订单取消补偿")) or num(cell_by(r, cols, "订单取消罚款")):
                stats["cancel"] += 1
                continue
            refund_amt = abs(num(cell_by(r, cols, "物品退款")))
            after = " ".join([cell_by(r, cols, "售后状态"), cell_by(r, cols, "售后合计")]).strip()
            if refund_amt > 0 or re.search(r"退款|退货|售后成功|售后完成", after):
                stats["refunded"] += 1
                continue
            name = cell_by(r, cols, "物品名称")
            imei = re.sub(r"[^0-9]", "", cell_by(r, cols, "imei"))
            if not imei and not re.search(r"手机|iPhone|Redmi|OPPO|vivo|荣耀|华为|小米|真我|realme|一加|三星|魅族|iQOO", name, flags=re.I):
                stats["other"] += 1
                continue
            goods = abs(num(cell_by(r, cols, "物品金额")))
            spend = abs(num(cell_by(r, cols, "支出合计")))
            if not spend:
                spend = abs(goods) + abs(num(cell_by(r, cols, "交易服务费")))
            if not spend:
                stats["no_cost"] += 1
            brand, model, mem, color = phone_name_parts(name)
            buy_time = cell_by(r, cols, "付款/退款时间", "付款时间", "下单时间")
            buy_date = buy_time[:10] if buy_time else date.today().isoformat()
            raw.append({
                "_order": po,
                "_goods_amt": goods,
                "poNo": item_no or (po + "-" + (imei[-4:] if imei else str(len(raw) + 1))),
                "imei": imei,
                "brand": brand,
                "model": model,
                "cond": "待定",
                "grade": "",
                "cost": money(spend),
                "buyDate": buy_date,
                "note": ("拍机堂订单 " + po).strip(),
                "source": "拍机堂",
            })
        fee_map = pjt_other_fees(sheets)
        by_order = {}
        for x in raw:
            by_order.setdefault(x["_order"], []).append(x)
        for order, items in by_order.items():
            allocate_order_fees(items, fee_map.get(order, {}))
        for x in raw:
            x["cost"] = money(x["cost"] + x.get("_other_fee", 0.0))
            x.pop("_order", None)
            x.pop("_goods_amt", None)
            x.pop("_other_fee", None)
            records.append(x)
        stats["added"] = len(records)
        return records, stats

    # 采货侠 / 转转：原有字段
    for r in target[header_at + 1:]:
        po = cell_by(r, cols, "订单号")
        if not po:
            continue
        stats["total"] += 1
        if cell_by(r, cols, "取消时间"):
            stats["cancel"] += 1
            continue
        refund_amt = abs(num(cell_by(r, cols, "机器退款金额")))
        after = cell_by(r, cols, "售后状态")
        if refund_amt > 0 or cell_by(r, cols, "退款时间") or re.search(r"退款|售后成功|退货|售后完成", after):
            stats["refunded"] += 1
            continue
        kind = cell_by(r, cols, "品类")
        if kind and kind != "手机":
            stats["other"] += 1
            continue
        imei = re.sub(r"[^0-9]", "", cell_by(r, cols, "imei"))
        brand_raw = cell_by(r, cols, "品牌")
        model_raw = cell_by(r, cols, "机型")
        if not model_raw:
            model_raw = (brand_raw + " " + cell_by(r, cols, "型号")).strip()
        brand = pick_brand(model_raw, brand_raw)
        desc = cell_by(r, cols, "机型描述")
        mem = norm_mem(desc_part(desc, "容量"))
        color_raw = desc_part(desc, "颜色")
        color = base_color(color_raw)
        model = norm_model(model_raw, brand)
        if mem and not re.search(r"\d+\s*\+\s*\d+\s*G", model):
            model += " " + mem
        if color and color not in model.split()[-1:]:
            model += " " + color
        grade = cell_by(r, cols, "成色/等级")
        cond = cond_of_grade(grade)
        paid = abs(num(cell_by(r, cols, "实付金额")))
        goods = abs(num(cell_by(r, cols, "商品金额")))
        cost = money(paid if paid > 0 else goods)
        if not cost:
            stats["no_cost"] += 1
        buy_time = cell_by(r, cols, "下单/竞得时间", "下单时间", "支付时间")
        buy_date = buy_time[:10] if buy_time else date.today().isoformat()
        note = ""
        if grade:
            note += "定级" + grade
        qc = cell_by(r, cols, "质检码")
        if qc:
            note += (" " if note else "") + "质检码" + qc
        if after:
            note += (" " if note else "") + "售后:" + after
        records.append({
            "poNo": po,
            "imei": imei,
            "brand": brand,
            "model": model,
            "cond": cond,
            "grade": grade,
            "cost": cost,
            "buyDate": buy_date,
            "note": note.strip(),
            "source": "采货侠",
        })
    stats["added"] = len(records)
    return records, stats

def rec_hash(o):
    parts = []
    for k in sorted(o.keys()):
        if k == "t":
            continue
        v = o[k]
        if v is None:
            s = ""
        elif isinstance(v, bool):
            s = "true" if v else "false"
        elif isinstance(v, (dict, list)):
            s = json.dumps(v, ensure_ascii=False, separators=(",", ":"))
        else:
            s = str(v)
        parts.append(k + "=" + s)
    return "|".join(parts)


def blank_payload():
    return {
        "devices": [], "sales": [], "purch": [], "returns": [], "inbox": [],
        "seq": {"sj": 1, "xs": 1, "hg": 1}, "meta": {}, "trash": [], "dirty": 0, "rev": 0,
    }


def merge_records(payload, records, source, batch_label):
    payload.setdefault("inbox", [])
    payload.setdefault("devices", [])
    payload.setdefault("meta", {})
    existing_pending = {
        (clean_text(x.get("poNo")), clean_text(x.get("imei")))
        for x in payload["inbox"] if x.get("status") != "已退货"
    }
    existing_devices = {
        (clean_text(x.get("poNo")), clean_text(x.get("imei")))
        for x in payload["devices"]
    }
    ids = {clean_text(x.get("id")) for x in payload["inbox"]}
    now = int(time.time() * 1000)
    batch_id = "b" + str(now) + "_" + os.urandom(2).hex()
    added = 0
    for idx, r in enumerate(records):
        key = (r["poNo"], r["imei"])
        if key in existing_pending or key in existing_devices:
            continue
        base_id = "ib_" + (r["imei"] or r["poNo"])
        rid = base_id
        n = 2
        while rid in ids:
            rid = base_id + "_" + str(n)
            n += 1
        rec = {
            "id": rid,
            "source": r.get("source") or source,
            "model": r["model"],
            "brand": r["brand"],
            "cond": r["cond"],
            "grade": r["grade"],
            "cost": r["cost"],
            "price": 0,
            "imei": r["imei"],
            "poNo": r["poNo"],
            "buyDate": r["buyDate"],
            "note": r["note"],
            "addedAt": now + idx,
            "status": "待入库",
            "batchId": batch_id,
            "batchLabel": (batch_label or source)[:24],
        }
        payload["inbox"].append(rec)
        payload["meta"]["i:" + rid] = {
            "h": rec_hash(rec), "t": now + idx, "imei": r["imei"],
        }
        ids.add(rid)
        existing_pending.add(key)
        added += 1
    if added:
        payload["rev"] = int(payload.get("rev") or 0) + 1
    return added


def supabase_base():
    base = os.environ["SUPABASE_URL"].strip().rstrip("/")
    return re.sub(r"/rest/v1.*$", "", base, flags=re.I)


def supabase_headers():
    key = os.environ["SUPABASE_KEY"].strip()
    return {
        "apikey": key,
        "Authorization": "Bearer " + key,
        "Content-Type": "application/json",
        "Prefer": "resolution=merge-duplicates",
    }


def load_payload(shop_code):
    base = supabase_base()
    url = base + "/rest/v1/shop_erp?shop_code=eq." + urllib.parse.quote(shop_code) + "&select=payload"
    req = urllib.request.Request(url, headers=supabase_headers())
    with urllib.request.urlopen(req, timeout=30) as resp:
        rows = json.loads(resp.read().decode("utf-8"))
    if rows and isinstance(rows[0].get("payload"), dict):
        p = rows[0]["payload"]
        p.setdefault("inbox", [])
        p.setdefault("meta", {})
        p.setdefault("devices", [])
        return p
    return blank_payload()


def save_payload(shop_code, payload):
    base = supabase_base()
    body = json.dumps([{"shop_code": shop_code, "payload": payload}], ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    req = urllib.request.Request(base + "/rest/v1/shop_erp", data=body, headers=supabase_headers(), method="POST")
    with urllib.request.urlopen(req, timeout=30) as resp:
        if resp.status not in (200, 201, 204):
            raise RuntimeError("Supabase HTTP " + str(resp.status))


def imap_host(user):
    if os.environ.get("NETEASE_HOST"):
        return os.environ["NETEASE_HOST"].strip()
    domain = user.split("@")[-1].lower()
    if domain.endswith("126.com"):
        return "imap.126.com"
    if domain.endswith("yeah.net"):
        return "imap.yeah.net"
    return "imap.163.com"


def attachment_files(msg):
    """返回 [(filename, bytes)]；支持 xlsx/csv，以及包含它们的 zip。"""
    out = []
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        filename = decode_mime(part.get_filename())
        payload = part.get_payload(decode=True)
        if not filename or not payload:
            continue
        low = filename.lower()
        if low.endswith((".xlsx", ".csv")):
            out.append((filename, payload))
        elif low.endswith(".zip"):
            try:
                with zipfile.ZipFile(io.BytesIO(payload)) as z:
                    for name in z.namelist():
                        if name.lower().endswith((".xlsx", ".csv")):
                            out.append((name, z.read(name)))
            except Exception:
                continue
    return out


def recent_message_ids(imap, days):
    since = (date.today() - timedelta(days=days)).strftime("%d-%b-%Y")
    status, data = imap.search(None, "SINCE", since)
    if status != "OK":
        return []
    return data[0].split()


def header_text(msg):
    return (decode_mime(msg.get("From")) + " " + decode_mime(msg.get("Subject"))).lower()


def process_mailbox():
    user = os.environ["NETEASE_USER"].strip()
    auth = os.environ["NETEASE_PASS"].strip()
    shop = os.environ["SHOP_CODE"].strip()
    raw_keywords = os.environ.get("BILL_KEYWORDS", "").strip() or DEFAULT_KEYWORDS
    keywords = [x.strip().lower() for x in raw_keywords.split(",") if x.strip()]
    host = imap_host(user)

    imap = imaplib.IMAP4_SSL(host, 993)
    try:
        imap.login(user, auth)
    except Exception as e:
        raise RuntimeError("网易邮箱登录失败：确认已开启 IMAP，且 NETEASE_PASS 填的是授权码不是登录密码") from e
    try:
        imap.select("INBOX")
        payload = load_payload(shop)
        total_added = 0
        for num in recent_message_ids(imap, 14):
            status, data = imap.fetch(num, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])")
            if status != "OK" or not data or not data[0]:
                continue
            head = email.message_from_bytes(data[0][1])
            text = header_text(head)
            if not any(k in text for k in keywords):
                continue
            status, data = imap.fetch(num, "(RFC822)")
            if status != "OK" or not data or not data[0]:
                continue
            msg = email.message_from_bytes(data[0][1])
            subject = decode_mime(msg.get("Subject")) or "账单"
            source = "拍机堂" if ("拍机堂" in text or "paijitang" in text) else "采货侠"
            files = attachment_files(msg)
            if not files:
                continue
            added_this = 0
            for filename, blob in files:
                try:
                    records, st = parse_bill_bytes(blob, filename)
                except Exception as e:
                    log("跳过附件", filename, "解析失败:", str(e)[:120])
                    continue
                added = merge_records(payload, records, source, subject + " · " + Path(filename).stem)
                added_this += added
                log(f"账单 {filename}: 读到 {st['total']} 单，新增 {added} 台，退款 {st['refunded']}，取消 {st['cancel']}，非手机 {st['other']}")
            if added_this:
                total_added += added_this
                save_payload(shop, payload)
            try:
                imap.store(num, "+FLAGS", "\\Seen")
            except Exception:
                pass
        log("本次共新增待入库", total_added, "台")
    finally:
        try:
            imap.logout()
        except Exception:
            pass


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", help="本地账单文件，测试解析用")
    ap.add_argument("--dry-run", action="store_true", help="只解析，不连接邮箱和 Supabase")
    args = ap.parse_args()

    if args.file:
        data = Path(args.file).read_bytes()
        records, st = parse_bill_bytes(data, Path(args.file).name)
        log(json.dumps({"stats": st, "first": records[:1]}, ensure_ascii=False, indent=2))
        return
    required = ["NETEASE_USER", "NETEASE_PASS", "SUPABASE_URL", "SUPABASE_KEY", "SHOP_CODE"]
    missing = [k for k in required if not os.environ.get(k)]
    if missing:
        raise SystemExit("缺少环境变量: " + ", ".join(missing))
    if args.dry_run:
        log("dry-run 只支持 --file，不执行邮箱同步")
        return
    process_mailbox()


if __name__ == "__main__":
    main()
