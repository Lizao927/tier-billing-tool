# -*- coding: utf-8 -*-
"""企业级/商业级/尊享级阶梯计费工具 - Windows GUI"""
import os
import re
import sys
import json
import traceback
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    raise RuntimeError("缺少 openpyxl 依赖。请运行安装依赖.bat。")

APP_NAME = "多等级阶梯计费工具"
BASE_BLACKLIST = [
    "AI Infra智算规划设计与实施服务",
    "AI平台辅助运营服务",
    "AI平台开发与实施服务",
    "AI平台运维与保障服务",
    "AI平台咨询与规划服务",
    "AI使能服务优化与提升服务",
    "aPaaS确定性运维",
    "aPaaS上云与实施服务",
    "aPaaS咨询与规划服务",
    "CBS对话机器人规划设计与实施服务",
    "CDN运维管理服务",
    "Cloud VR培训服务",
    "CloudDC确定性运维服务",
    "CSS专家服务",
    "FRS人脸识别规划设计与实施服务",
    "GES规划设计与实施服务",
    "HCS 盘古大模型集成与开发支持服务",
    "HCSO安全专业服务包周期 模式1",
    "HCSO安全专业服务一次性 模式1",
    "HCSO计算专业服务 模式1",
    "HCSO培训服务 模式1",
    "HCSO企业云化咨询服务 模式1",
    "HCSO数据库专业服务 模式1",
    "HCSO支持计划 模式1",
    "HiLens规划设计与实施服务",
    "iDEE设计与实施服务",
    "iDME实施与支持服务",
    "IPDCenter确定性运维技术支持服务",
    "IPDCenter上云与实施服务",
    "IT辅助运营专业服务_数据库",
    "IVA视频分析平台规划设计与实施服务",
    "IVA视频分析平台开发支持服务",
    "IVA视频分析平台专家服务",
    "Landing Zone设计与实施服务",
    "ModelArts平台规划设计与安装实施服务",
    "ModelArts平台培训服务",
    "NLP自然语言处理规划设计与实施服务",
    "OCR规划设计与实施服务",
    "pEDA确定性运维技术支持服务",
    "pEDA上云与实施服务",
    "ROMA集成专业服务",
    "ROMA设计与实施服务",
    "ROMA运维与管理服务",
    "ROMA专家服务",
    "ROMA咨询服务",
    "ROMA咨询与规划服务",
    "SAP上云专业服务",
    "SIM上云与实施服务",
    "SIS语音交互规划设计与实施服务",
    "WeLink “Athena+”专业服务",
    "WeLink“Search+”专业服务",
    "WeLink交付服务",
    "WeLink开发支持服务",
    "WeLink培训服务",
    "WeLink企业品牌LOGO专享专业服务",
    "WeLink运维保障服务",
    "WeLink运营支撑服务",
    "WeLink专业服务",
    "WeLink咨询服务",
    "Workspace交付实施服务",
    "Workspace运维服务",
    "安全集成开发与实施服务",
    "安全使能规划与设计服务",
    "安全运营与运维服务",
    "城市智能体使能服务",
    "存储安全优化与提升服务",
    "大数据基础设施专家服务",
    "大数据上云与实施服务",
    "大数据优化与提升服务",
    "大数据咨询服务",
    "地理智能体专家服务",
    "低代码运维与管理服务",
    "低代码运维与管理服务-伙伴云",
    "高级人工智能模型需求分析与设计",
    "工程交付规划设计与实施服务",
    "工业互联网培训服务",
    "工业软件云咨询服务",
    "工业智能体使能服务",
    "供暖智能体专家服务",
    "规划设计与实施服务",
    "规划设计与实施服务 模式一",
    "华为八爪鱼自动驾驶专业服务",
    "华为云DevSecOps规划服务",
    "华为云DevSecOps设计与实施服务",
    "华为云DevSecOps运维与管理服务",
    "华为云Workspace咨询规划服务",
    "华为云安全培训服务",
    "华为云成长型企业数字化转型包集成实施服务",
    "华为云工业软件培训服务",
    "华为云基础服务人才培养服务",
    "华为云敏捷与DevOps培训服务",
    "华为云数据库培训服务",
    "华为云通用培训服务",
    "华为云政务类解决方案集成交付服务",
    "华为云智慧园区规划设计与集成实施服务",
    "混合云HCS Online云平台下电支持及上电恢复服务",
    "混合云基础服务集成交付服务",
    "混合云基础架构规划咨询与设计服务",
    "混合云现场技术管理服务",
    "混合云现场项目管理服务",
    "混合云支持计划",
    "混合云驻场运维服务",
    "伙伴云AI平台开发与实施服务",
    "伙伴云aPaaS规划与设计服务",
    "伙伴云CloudDC确定性运维服务",
    "伙伴云ModelArts平台规划设计与安装实施服务",
    "伙伴云ROMA设计与实施服务",
    "伙伴云安全管理检测与响应",
    "伙伴云安全认证测试中心服务",
    "伙伴云大数据优化与提升服务",
    "伙伴云工程交付规划设计与实施服务",
    "伙伴云混合云基础服务交付服务",
    "伙伴云混合云基础架构规划咨询与设计服务",
    "伙伴云鲲鹏测评专家服务",
    "伙伴云鲲鹏移植专家服务",
    "伙伴云上云设计与实施工具支持服务",
    "伙伴云数据库方案设计服务",
    "伙伴云数据库规划设计与实施服务",
    "伙伴云数据库开发支持服务",
    "伙伴云数据库培训服务",
    "伙伴云数据库迁移服务",
    "伙伴云云原生专家服务",
    "伙伴云云灾备交付实施服务",
    "伙伴云云灾备咨询服务",
    "伙伴云智能数据专家服务",
    "交通智能体专家服务",
    "鲲鹏测评专家服务",
    "鲲鹏计算人才培养服务",
    "鲲鹏计算移植调优服务",
    "鲲鹏人才培养专家服务",
    "鲲鹏移植专家服务",
    "流程IT-IDC运维服务",
    "敏捷与DevOps专家服务",
    "企业上云网络规划设计服务",
    "企业云化咨询服务",
    "区块链培训服务",
    "区块链专家服务",
    "确定性运维大数据服务",
    "确定性运维管理服务",
    "确定性运维规划与设计服务",
    "确定性运维支持服务",
    "人工智能云培训服务",
    "人工智能云专家咨询服务",
    "人工智能咨询服务 失效",
    "认证测试中心",
    "商专车自动驾驶专业服务",
    "上云迁移服务",
    "上云设计与实施工具支持服务",
    "视频分析规划设计与实施服务",
    "视频直播运维与管理服务",
    "数据仓库上云与实施服务",
    "数据仓库优化与支持服务",
    "数据仓库专家服务",
    "数据仓库咨询服务",
    "数据仓库咨询与规划服务",
    "数据可控可信服务",
    "数据库方案设计服务",
    "数据库规划设计与实施",
    "数据库开发支持服务",
    "数据库迁移服务",
    "数据库确定性运维服务",
    "数据库专家服务",
    "数据上云设计与实施服务",
    "数据运营专家服务",
    "数据治理服务",
    "数据中心确定性运维",
    "数字化规划治理服务",
    "数字化诊断治理专家服务",
    "水智能体专家服务",
    "图像识别规划设计与实施服务",
    "微服务专家服务",
    "稳定性提升服务",
    "物联网上云与实施服务",
    "物联网上云与实施服务_伙伴云",
    "物联网优化与提升服务",
    "物联网云培训服务",
    "物联网运维与管理服务",
    "物联网专家服务",
    "物联网咨询与规划服务",
    "行业AI大赛&培训专业服务",
    "行业AI辅助运营服务",
    "行业AI集成与开发支持服务",
    "行业AI上云与实施服务",
    "行业AI咨询服务",
    "应用迁移支持服务",
    "应用上云规划服务",
    "应用现代化服务",
    "应用性能调优专家服务",
    "应用中间件设计与实施服务",
    "应用中间件咨询与规划服务",
    "营销数字化规划设计服务",
    "用云专家服务",
    "园区智能体专家服务",
    "云存储规划与设计服务",
    "云存储优化与提升服务",
    "云工厂规划设计服务",
    "云会议保障服务",
    "云会议交付实施服务",
    "云会议支持订阅服务",
    "云基础设施采用规划与设计",
    "云容灾交付实施服务",
    "云容灾咨询服务",
    "云上保障护航服务",
    "云上保障护航服务_Inactive",
    "云视频交付实施服务",
    "云视频咨询规划服务",
    "云通信调研与使能服务",
    "云通信上云与实施服务",
    "云通信运维与管理服务",
    "云应用平台采用规划与设计",
    "云原生专家服务",
    "云运维中心实施支持服务",
    "云运营平台集成服务",
    "云灾备交付实施服务",
    "云灾备咨询服务",
    "云智慧运营规划与设计服务",
    "云助工软使能服务",
    "运维平台集成与实施服务",
    "运维平台培训服务",
    "政企自服务管理规划设计与实施服务",
    "支持计划",
    "知识图谱规划设计与实施服务",
    "智能数据培训服务",
    "智能数据专家服务",
    "云商店",
]
LEVELS = {
    "企业级": {
        "minimum": 55000,
        "tiers": [(550000, .10), (1800000, .07), (3300000, .05), (None, .03)]
    },
    "商业级": {
        "minimum": 28000,
        "tiers": [(280000, .10), (900000, .07), (1650000, .05), (None, .03)]
    },
    "尊享级": {
        "minimum": 300000,
        "tiers": [(8000000, .10), (25000000, .06), (None, .05)]
    },
}
NAVY = "17365D"
BLUE = "D9EAF7"
GREEN = "E2F0D9"
RED = "FCE4D6"
WHITE = "FFFFFF"

def normalize(value):
    s = str(value or "").strip().lower()
    # 忽略空格、全角空格、NBSP及常见分隔符
    return re.sub(r"[\s\u3000\u00a0\-_—（）()\[\]【】,，.。/\\:：]+", "", s)

def is_blacklisted(product_type, keywords):
    p = normalize(product_type)
    if not p:
        return None
    for kw in keywords:
        k = normalize(kw)
        if k and k in p:
            return kw
    return None

def amount(value):
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    for token in ("¥", "￥", "元", ",", "，"):
        s = s.replace(token, "")
    if not s:
        return 0.0
    try:
        return float(s)
    except ValueError:
        raise ValueError(f"官网价存在无法识别的金额：{value!r}")

def progressive_fee(monthly, level):
    remaining = max(float(monthly), 0.0)
    previous = 0.0
    fee = 0.0
    for upper, rate in LEVELS[level]["tiers"]:
        if upper is None:
            portion = max(remaining - previous, 0)
        else:
            portion = max(min(remaining, upper) - previous, 0)
        fee += portion * rate
        if upper is None or remaining <= upper:
            break
        previous = upper
    return fee

def find_col(headers, names):
    cleaned = [re.sub(r"\s+", "", str(x or "")).lower() for x in headers]
    for name in names:
        target = re.sub(r"\s+", "", name).lower()
        if target in cleaned:
            return cleaned.index(target)
    return None

def read_input(path, sheet_name):
    keep_vba = path.lower().endswith(".xlsm")
    wb = load_workbook(path, data_only=True, read_only=True, keep_vba=keep_vba)
    ws = wb[sheet_name]
    rows = ws.iter_rows(values_only=True)
    try:
        headers = list(next(rows))
    except StopIteration:
        wb.close()
        raise ValueError("选择的工作表为空。")
    data = [list(row) for row in rows if any(v is not None and str(v).strip() for v in row)]
    wb.close()
    return headers, data

def create_report(input_path, sheet_name, months, level, keywords, output_path):
    headers, data = read_input(input_path, sheet_name)
    type_idx = find_col(headers, ["产品类型"])
    price_idx = find_col(headers, ["官网价(¥)", "官网价（¥）", "官网价", "官网价(元)", "官网价（元）"])
    if type_idx is None:
        raise ValueError("没有找到“产品类型”列，请检查第一行表头。")
    if price_idx is None:
        raise ValueError("没有找到“官网价(¥)”列，请检查第一行表头。")

    all_total = filtered_total = eligible_total = 0.0
    filtered_count = eligible_count = 0
    output_rows = []
    for row_num, row in enumerate(data, start=2):
        row = list(row) + [None] * max(0, len(headers) - len(row))
        product_type = row[type_idx]
        price = amount(row[price_idx])
        all_total += price
        hit = is_blacklisted(product_type, keywords)
        if hit:
            filtered_total += price
            filtered_count += 1
            included = 0.0
            status = "是"
        else:
            eligible_total += price
            eligible_count += 1
            included = price
            status = "否"
        output_rows.append(row[:len(headers)] + [status, hit or "", included])

    monthly = eligible_total / months
    tier = progressive_fee(monthly, level)
    floor = LEVELS[level]["minimum"]
    monthly_final = max(tier, floor)
    period_final = monthly_final * months
    note = "按保底价计费" if tier <= floor else "阶梯计费金额高于保底价，按阶梯计费金额计费"

    wb = Workbook()
    ws = wb.active
    ws.title = "计费汇总"
    ws["A1"] = "多等级阶梯计费结果"
    ws["A1"].font = Font(name="微软雅黑", size=16, bold=True, color=WHITE)
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws.merge_cells("A1:C1")

    summary = [
        ("项目", "计算结果", "说明"),
        ("产品等级", level, "本次选择的计费等级"),
        ("账单月份数", months, "账单覆盖月份"),
        ("导入明细官网价总额", all_total, "包含黑名单产品"),
        ("黑名单过滤金额", filtered_total, "命中黑名单产品的官网价"),
        ("计入消费官网价总额", eligible_total, "过滤后累计官网价，按折前价"),
        ("月均云服务消费", monthly, f"计入消费官网价总额 ÷ {months}"),
        ("月均阶梯计费金额", tier, "按所选等级分段累进计算"),
        ("对应等级月保底价", floor, "所选产品等级的月度保底价"),
        ("最终月度应收", monthly_final, "月均阶梯费与月保底价取高者"),
        ("本账期最终应收总额", period_final, "最终月度应收 × 账单月份数"),
        ("超出保底价金额（月度）", max(tier-floor, 0), "阶梯计费金额超过月保底价的差额"),
        ("计费备注", note, "自动判断"),
        ("过滤明细行数", filtered_count, "命中黑名单的明细行数"),
        ("有效明细行数", eligible_count, "未命中黑名单的明细行数"),
    ]
    for r, row in enumerate(summary, start=3):
        for c, val in enumerate(row, start=1):
            cell = ws.cell(r, c, val)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if r == 3:
                cell.font = Font(bold=True, color=WHITE)
                cell.fill = PatternFill("solid", fgColor=NAVY)
            else:
                cell.font = Font(name="微软雅黑", size=10, bold=(c == 1))
                if c == 1:
                    cell.fill = PatternFill("solid", fgColor=BLUE)
                if r in (10, 13):
                    cell.fill = PatternFill("solid", fgColor=GREEN)
        ws.row_dimensions[r].height = 28
    for r in range(6, 16):
        ws.cell(r, 2).number_format = '#,##0.00'
    ws["B5"].number_format = "0"
    ws["B17"].number_format = "0"
    ws["B18"].number_format = "0"
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 66
    ws.freeze_panes = "A4"

    detail = wb.create_sheet("消费明细")
    out_headers = headers + ["是否过滤", "过滤命中项", "计入消费官网价(¥)"]
    for c, h in enumerate(out_headers, 1):
        cell = detail.cell(1, c, h)
        cell.font = Font(bold=True, color=WHITE, name="微软雅黑", size=9)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    detail.row_dimensions[1].height = 32
    for r, row in enumerate(output_rows, 2):
        for c, val in enumerate(row, 1):
            cell = detail.cell(r, c, val)
            cell.font = Font(name="微软雅黑", size=9)
            if c == len(headers) + 1:
                cell.fill = PatternFill("solid", fgColor=RED if val == "是" else GREEN)
            if c == len(headers) + 3 or c == price_idx + 1:
                cell.number_format = '#,##0.00'
    for c, h in enumerate(out_headers, 1):
        width = min(max(len(str(h)) + 3, 12), 36)
        detail.column_dimensions[get_column_letter(c)].width = width
    detail.freeze_panes = "A2"
    detail.auto_filter.ref = detail.dimensions

    rules = wb.create_sheet("计费规则")
    tier_descriptions = []
    previous = 0
    for upper, rate in LEVELS[level]["tiers"]:
        if upper is None:
            tier_descriptions.append(f"超过{previous:,.0f}元部分 × {rate:.0%}")
        else:
            tier_descriptions.append(f"{previous:,.0f}至{upper:,.0f}元部分 × {rate:.0%}")
            previous = upper
    rule_lines = [
        ("产品等级", level),
        ("月保底价", floor),
        ("月均消费计算", f"过滤后官网价总额 ÷ {months}个月"),
        ("阶梯计算", "；".join(tier_descriptions)),
        ("最终月度应收", "MAX(月均阶梯计费金额, 月保底价)"),
        ("本账期最终应收", "最终月度应收 × 账单月份数"),
        ("黑名单关键词", "；".join(keywords)),
    ]
    rules["A1"] = "本次计算使用的规则"
    rules["A1"].font = Font(bold=True, color=WHITE, size=14)
    rules["A1"].fill = PatternFill("solid", fgColor=NAVY)
    rules.merge_cells("A1:B1")
    for r, row in enumerate(rule_lines, 3):
        rules.cell(r, 1, row[0]).font = Font(bold=True, color=NAVY)
        rules.cell(r, 1).fill = PatternFill("solid", fgColor=BLUE)
        rules.cell(r, 2, row[1]).alignment = Alignment(wrap_text=True, vertical="top")
        rules.row_dimensions[r].height = 36
    rules.column_dimensions["A"].width = 24
    rules.column_dimensions["B"].width = 110

    wb.save(output_path)
    return {
        "all_total": all_total, "filtered_total": filtered_total,
        "eligible_total": eligible_total, "monthly": monthly,
        "tier": tier, "floor": floor, "monthly_final": monthly_final,
        "period_final": period_final, "filtered_count": filtered_count,
        "eligible_count": eligible_count, "note": note
    }

class BillingGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("900x720")
        self.root.minsize(820, 650)
        self.input_path = ""
        self.sheet_names = []
        self.keywords = BASE_BLACKLIST.copy()
        self.build()

    def build(self):
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except Exception:
            pass
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 18, "bold"), foreground="#17365D")
        style.configure("Run.TButton", font=("Microsoft YaHei UI", 11, "bold"), padding=8)

        main = ttk.Frame(self.root, padding=18)
        main.pack(fill="both", expand=True)
        ttk.Label(main, text="多等级阶梯计费工具", style="Title.TLabel").pack(anchor="w")
        ttk.Label(main, text="导入账单｜维护黑名单｜选择产品等级｜自动计算并导出 Excel",
                  foreground="#555").pack(anchor="w", pady=(4, 14))

        filebox = ttk.LabelFrame(main, text="一、客户消费明细", padding=12)
        filebox.pack(fill="x", pady=5)
        frow = ttk.Frame(filebox); frow.pack(fill="x")
        self.file_label = ttk.Label(frow, text="尚未选择 Excel 文件", wraplength=680)
        self.file_label.pack(side="left", fill="x", expand=True)
        ttk.Button(frow, text="选择账单 Excel", command=self.choose_file).pack(side="right")
        ttk.Label(filebox, text="支持 .xlsx / .xlsm；表头需包含“产品类型”和“官网价(¥)”。",
                  foreground="#666").pack(anchor="w", pady=(6, 0))
        sheetrow = ttk.Frame(filebox); sheetrow.pack(fill="x", pady=(7, 0))
        ttk.Label(sheetrow, text="工作表：").pack(side="left")
        self.sheet_var = tk.StringVar(value="")
        self.sheet_combo = ttk.Combobox(sheetrow, textvariable=self.sheet_var, state="readonly", width=38)
        self.sheet_combo.pack(side="left")

        optbox = ttk.LabelFrame(main, text="二、计费参数", padding=12)
        optbox.pack(fill="x", pady=5)
        row = ttk.Frame(optbox); row.pack(fill="x")
        ttk.Label(row, text="产品等级：").pack(side="left")
        self.level_var = tk.StringVar(value="企业级")
        ttk.Combobox(row, textvariable=self.level_var, values=list(LEVELS), state="readonly", width=12).pack(side="left", padx=(0, 22))
        ttk.Label(row, text="账单月份数：").pack(side="left")
        self.month_var = tk.StringVar(value="1")
        ttk.Spinbox(row, from_=1, to=120, textvariable=self.month_var, width=8).pack(side="left", padx=(0, 10))
        ttk.Label(row, text="例如1—11月账单填写11", foreground="#666").pack(side="left")

        blbox = ttk.LabelFrame(main, text="三、黑名单维护（可新增、删除、修改）", padding=12)
        blbox.pack(fill="both", expand=True, pady=5)
        body = ttk.Frame(blbox); body.pack(fill="both", expand=True)
        self.listbox = tk.Listbox(body, height=10, font=("Microsoft YaHei UI", 10), selectmode=tk.SINGLE)
        self.listbox.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(body, orient="vertical", command=self.listbox.yview)
        scroll.pack(side="left", fill="y")
        self.listbox.config(yscrollcommand=scroll.set)
        side = ttk.Frame(body); side.pack(side="left", fill="y", padx=(12, 0))
        ttk.Label(side, text="关键词：").pack(anchor="w")
        self.keyword_var = tk.StringVar()
        ttk.Entry(side, textvariable=self.keyword_var, width=28).pack(anchor="w", pady=(3, 8))
        ttk.Button(side, text="添加关键词", command=self.add_keyword).pack(fill="x", pady=3)
        ttk.Button(side, text="修改选中项", command=self.edit_keyword).pack(fill="x", pady=3)
        ttk.Button(side, text="删除选中项", command=self.delete_keyword).pack(fill="x", pady=3)
        ttk.Button(side, text="恢复默认黑名单", command=self.reset_keywords).pack(fill="x", pady=(14, 3))
        ttk.Label(blbox, text="匹配方式：忽略大小写和空格；产品类型只要包含黑名单关键词就过滤。例如“支持计划”匹配“支持计划Support Plan”。",
                  foreground="#666", wraplength=820).pack(anchor="w", pady=(8, 0))

        bottom = ttk.Frame(main); bottom.pack(fill="x", pady=(12, 4))
        self.run_btn = ttk.Button(bottom, text="计算并导出 Excel 报告", style="Run.TButton", command=self.run)
        self.run_btn.pack(side="left")
        ttk.Button(bottom, text="保存黑名单配置", command=self.save_keywords).pack(side="left", padx=10)
        ttk.Label(main, text="输出报告包含：计费汇总、消费明细（含过滤命中项）、计费规则。",
                  foreground="#666").pack(anchor="w", pady=(6, 0))
        self.populate_list()

    def populate_list(self):
        self.listbox.delete(0, "end")
        for kw in self.keywords:
            self.listbox.insert("end", kw)

    def add_keyword(self):
        kw = self.keyword_var.get().strip()
        if not kw:
            messagebox.showwarning("提示", "请先输入关键词。"); return
        if kw in self.keywords:
            messagebox.showinfo("提示", "该关键词已经存在。"); return
        self.keywords.append(kw); self.populate_list(); self.keyword_var.set("")

    def edit_keyword(self):
        sel = self.listbox.curselection()
        kw = self.keyword_var.get().strip()
        if not sel:
            messagebox.showwarning("提示", "请先选择要修改的关键词。"); return
        if not kw:
            messagebox.showwarning("提示", "请在关键词输入框填写新内容。"); return
        idx = sel[0]
        if kw in self.keywords and self.keywords[idx] != kw:
            messagebox.showwarning("提示", "该关键词已经存在。"); return
        self.keywords[idx] = kw; self.populate_list(); self.keyword_var.set("")

    def delete_keyword(self):
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请先选择要删除的关键词。"); return
        del self.keywords[sel[0]]; self.populate_list()

    def reset_keywords(self):
        if messagebox.askyesno("确认", "恢复默认黑名单？当前自定义关键词将被替换。"):
            self.keywords = BASE_BLACKLIST.copy(); self.populate_list()

    def config_path(self):
        return os.path.join(os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)),
                            "黑名单配置.json")

    def save_keywords(self):
        try:
            with open(self.config_path(), "w", encoding="utf-8") as f:
                json.dump(self.keywords, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("保存成功", f"黑名单配置已保存：\n{self.config_path()}")
        except Exception as e:
            messagebox.showerror("保存失败", str(e))

    def load_keywords(self):
        try:
            with open(self.config_path(), "r", encoding="utf-8") as f:
                values = json.load(f)
            if isinstance(values, list):
                self.keywords = [str(x).strip() for x in values if str(x).strip()]
        except Exception:
            pass

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="选择客户消费明细",
            filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")]
        )
        if not path:
            return
        self.input_path = path
        self.file_label.config(text=path)
        try:
            wb = load_workbook(path, read_only=True, data_only=True, keep_vba=path.lower().endswith(".xlsm"))
            self.sheet_names = wb.sheetnames
            wb.close()
            self.sheet_combo["values"] = self.sheet_names
            self.sheet_var.set(self.sheet_names[0] if self.sheet_names else "")
        except Exception as e:
            messagebox.showerror("读取失败", str(e))

    def run(self):
        if not self.input_path:
            messagebox.showwarning("提示", "请先选择客户消费明细 Excel。"); return
        try:
            months = int(self.month_var.get())
            if not 1 <= months <= 120: raise ValueError
        except ValueError:
            messagebox.showwarning("月份错误", "月份数必须为1至120的整数。"); return
        level = self.level_var.get()
        if level not in LEVELS:
            messagebox.showwarning("等级错误", "请选择有效产品等级。"); return
        if not self.keywords:
            if not messagebox.askyesno("确认", "当前黑名单为空，将不会过滤任何产品，是否继续？"): return

        sheet = self.sheet_var.get().strip()
        if not sheet:
            messagebox.showwarning("提示", "请先选择账单工作表。")
            return

        stem = os.path.splitext(os.path.basename(self.input_path))[0]
        output = filedialog.asksaveasfilename(
            title="保存计费报告",
            initialdir=os.path.dirname(self.input_path),
            initialfile=stem + "_阶梯计费结果.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")]
        )
        if not output: return
        self.run_btn.config(state="disabled", text="正在计算，请稍候…")
        self.root.update_idletasks()
        try:
            result = create_report(self.input_path, sheet, months, level, self.keywords, output)
            messagebox.showinfo(
                "计算完成",
                f"报告已生成：\n{output}\n\n"
                f"产品等级：{level}\n账单月份：{months}\n"
                f"过滤后累计消费：¥{result['eligible_total']:,.2f}\n"
                f"月均云服务消费：¥{result['monthly']:,.2f}\n"
                f"月均阶梯计费：¥{result['tier']:,.2f}\n"
                f"最终月度应收：¥{result['monthly_final']:,.2f}\n"
                f"本账期应收总额：¥{result['period_final']:,.2f}\n"
                f"备注：{result['note']}"
            )
            try: os.startfile(output)
            except Exception: pass
        except Exception as e:
            messagebox.showerror("计算失败", f"{e}\n\n{traceback.format_exc()}")
        finally:
            self.run_btn.config(state="normal", text="计算并导出 Excel 报告")

def main():
    root = tk.Tk()
    app = BillingGUI(root)
    app.load_keywords()
    app.populate_list()
    root.mainloop()

if __name__ == "__main__":
    main()
