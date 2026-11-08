# -*- coding: utf-8 -*-
"""企业级/商业级/尊享级阶梯计费工具 - Windows GUI"""
import os
import re
import sys
import json
import traceback
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.cell import WriteOnlyCell
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
    """统一产品类型和黑名单关键词格式，便于做包含匹配。"""
    s = str(value or "").strip().lower()
    return re.sub(r"[\s\u3000\u00a0\-_—（）()\[\]【】,，.。/\\:：]+", "", s)


def build_blacklist_matcher(keywords):
    """
    将黑名单编译为单个正则表达式。
    相比每行逐个遍历全部关键词，单次正则扫描可显著减少大账单的匹配开销。
    """
    mapping = {}
    for keyword in keywords:
        original = str(keyword or "").strip()
        normalized = normalize(original)
        if normalized and normalized not in mapping:
            mapping[normalized] = original

    if not mapping:
        return None, {}

    # 长关键词优先，避免短关键词抢先命中造成命中原因不够具体。
    ordered = sorted(mapping, key=len, reverse=True)
    pattern = re.compile("|".join(re.escape(item) for item in ordered))
    return pattern, mapping


def match_blacklist(product_type, matcher):
    if matcher is None or product_type is None:
        return None
    normalized_product = normalize(product_type)
    if not normalized_product:
        return None
    pattern, mapping = matcher
    found = pattern.search(normalized_product)
    return mapping.get(found.group(0)) if found else None


def amount(value):
    """用 Decimal 解析金额，避免浮点数累计误差。"""
    if value is None or value == "":
        return Decimal("0")
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    s = str(value).strip()
    for token in ("¥", "￥", "元", ",", "，", " "):
        s = s.replace(token, "")
    if not s:
        return Decimal("0")
    try:
        return Decimal(s)
    except InvalidOperation:
        raise ValueError(f"官网价存在无法识别的金额：{value!r}")


def money_float(value):
    return float(value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def progressive_fee(consumption, level):
    """按单月消费额分段累进计算阶梯费用。"""
    remaining = max(Decimal(str(consumption)), Decimal("0"))
    previous = Decimal("0")
    fee = Decimal("0")
    for upper, rate in LEVELS[level]["tiers"]:
        rate_decimal = Decimal(str(rate))
        if upper is None:
            portion = max(remaining - previous, Decimal("0"))
        else:
            upper_decimal = Decimal(str(upper))
            portion = max(min(remaining, upper_decimal) - previous, Decimal("0"))
        fee += portion * rate_decimal
        if upper is None or remaining <= Decimal(str(upper)):
            break
        previous = Decimal(str(upper))
    return fee.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def find_col(headers, names):
    cleaned = [re.sub(r"\s+", "", str(x or "")).lower() for x in headers]
    for name in names:
        target = re.sub(r"\s+", "", name).lower()
        if target in cleaned:
            return cleaned.index(target)
    return None


def create_report(input_path, sheet_name, level, keywords, output_path):
    """
    单月账单高速处理：
    - 只读模式逐行读取，不把5万+行全部装入内存；
    - 黑名单一次编译，逐行只做一次正则搜索；
    - write_only 模式流式写出结果，避免逐单元格创建普通工作簿。
    """
    if level not in LEVELS:
        raise ValueError("产品等级无效，请选择企业级、商业级或尊享级。")

    keep_vba = input_path.lower().endswith(".xlsm")
    input_wb = load_workbook(
        input_path, data_only=True, read_only=True, keep_vba=keep_vba
    )
    try:
        if sheet_name not in input_wb.sheetnames:
            raise ValueError(f"找不到工作表：{sheet_name}")
        input_ws = input_wb[sheet_name]
        rows = input_ws.iter_rows(values_only=True)

        try:
            headers = list(next(rows))
        except StopIteration:
            raise ValueError("选择的工作表为空。")

        type_idx = find_col(headers, ["产品类型"])
        price_idx = find_col(
            headers,
            ["官网价(¥)", "官网价（¥）", "官网价", "官网价(元)", "官网价（元）"]
        )
        if type_idx is None:
            raise ValueError("没有找到“产品类型”列，请检查第一行表头。")
        if price_idx is None:
            raise ValueError("没有找到“官网价(¥)”列，请检查第一行表头。")

        matcher = build_blacklist_matcher(keywords)
        all_total = Decimal("0")
        filtered_total = Decimal("0")
        eligible_total = Decimal("0")
        filtered_count = 0
        eligible_count = 0
        imported_count = 0

        # write_only 工作簿：结果逐行写入，显著降低大文件写出时的内存和耗时。
        out_wb = Workbook(write_only=True)
        summary_ws = out_wb.create_sheet("计费汇总")
        detail_ws = out_wb.create_sheet("消费明细")
        rules_ws = out_wb.create_sheet("计费规则")

        detail_headers = headers + ["是否过滤", "过滤命中项", "计入消费官网价(¥)"]
        header_cells = []
        for value in detail_headers:
            cell = WriteOnlyCell(detail_ws, value=value)
            cell.font = Font(bold=True, color=WHITE, name="微软雅黑", size=9)
            cell.fill = PatternFill("solid", fgColor=NAVY)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            header_cells.append(cell)
        detail_ws.append(header_cells)
        detail_ws.freeze_panes = "A2"

        # 核心：单次遍历源数据，不缓存全部明细行。
        for row in rows:
            if not row or not any(v is not None and str(v).strip() for v in row):
                continue
            imported_count += 1
            row_values = list(row[:len(headers)])
            if len(row_values) < len(headers):
                row_values.extend([None] * (len(headers) - len(row_values)))

            price = amount(row_values[price_idx])
            all_total += price

            product_type = row_values[type_idx]
            hit = match_blacklist(product_type, matcher)
            if hit:
                filtered_total += price
                filtered_count += 1
                included = Decimal("0")
                status = "是"
            else:
                eligible_total += price
                eligible_count += 1
                included = price
                status = "否"

            detail_ws.append(row_values + [status, hit or "", money_float(included)])

        # 当前账单为单月账单：不再除以月份数，也不再乘回月份数。
        tier = progressive_fee(eligible_total, level)
        floor = Decimal(str(LEVELS[level]["minimum"]))
        monthly_final = max(tier, floor)
        over_floor = max(tier - floor, Decimal("0"))
        note = "按对应等级月保底价计费" if tier <= floor else "阶梯计费金额高于保底价，按阶梯计费金额计费"

        # 汇总页：简洁版式，突出最终月度应收。
        summary_ws.sheet_properties.tabColor = "4472C4"
        summary_ws.sheet_view.showGridLines = False
        summary_ws.column_dimensions["A"].width = 27
        summary_ws.column_dimensions["B"].width = 21
        summary_ws.column_dimensions["C"].width = 48

        title_fill = PatternFill("solid", fgColor="17365D")
        header_fill = PatternFill("solid", fgColor="4472C4")
        label_fill = PatternFill("solid", fgColor="EAF1F8")
        result_fill = PatternFill("solid", fgColor="E2F0D9")
        # 最终应收仅突出“计算结果”单元格，避免整行视觉权重过高。
        final_fill = PatternFill("solid", fgColor="FFF2CC")
        final_font = Font(name="微软雅黑", size=12, bold=True, color="7F6000")
        white_bold = Font(name="微软雅黑", bold=True, color="FFFFFF", size=11)
        normal_font = Font(name="微软雅黑", size=10, color="243247")
        label_font = Font(name="微软雅黑", size=10, bold=True, color="17365D")
        thin_gray = Side(style="thin", color="D9E2F3")
        cell_border = Border(bottom=thin_gray)

        # 标题行
        title_cells = []
        for value in ("阶梯计费汇总", "", ""):
            cell = WriteOnlyCell(summary_ws, value=value)
            cell.font = Font(name="微软雅黑", size=15, bold=True, color="FFFFFF")
            cell.fill = title_fill
            cell.alignment = Alignment(vertical="center")
            title_cells.append(cell)
        summary_ws.append(title_cells)
        summary_ws.row_dimensions[1].height = 32
        summary_ws.append([])

        # 删除不需要展示的过程字段，只保留关键计费结果。
        summary_rows = [
            ("项目", "计算结果", "说明"),
            ("产品等级", level, "本次选择的计费等级"),
            ("当月云服务消费总额", money_float(eligible_total), "过滤黑名单后的官网价总额，按折前价计算"),
            ("阶梯计费金额", money_float(tier), "按所选等级对当月云服务消费总额分段累进计算"),
            ("最终月度应收", money_float(monthly_final), "阶梯计费金额与月保底价取高者"),
            ("超出保底价金额", money_float(over_floor), "阶梯计费金额超过月保底价的差额"),
            ("计费备注", note, "自动判断"),
            ("过滤明细行数", filtered_count, "命中黑名单的行数"),
            ("有效明细行数", eligible_count, "未命中黑名单的行数"),
        ]
        for idx, row in enumerate(summary_rows):
            cells = []
            is_header = idx == 0
            is_final = row[0] == "最终月度应收"
            for col_idx, value in enumerate(row, start=1):
                cell = WriteOnlyCell(summary_ws, value=value)
                cell.alignment = Alignment(
                    vertical="center",
                    horizontal="right" if col_idx == 2 and isinstance(value, (int, float)) else "left",
                    wrap_text=True
                )
                cell.border = cell_border
                if is_header:
                    cell.font = white_bold
                    cell.fill = header_fill
                else:
                    # 最终月度应收所在行，其“项目”和“说明”列仍沿用普通行样式；
                    # 只有“计算结果”列使用黄色高亮。
                    if col_idx == 1:
                        cell.font = label_font
                        cell.fill = label_fill
                    elif col_idx == 2:
                        cell.font = final_font if is_final else normal_font
                        cell.fill = final_fill if is_final else result_fill
                    else:
                        cell.font = normal_font
                        cell.fill = PatternFill(fill_type=None)

                if col_idx == 2 and isinstance(value, (int, float)):
                    if row[0] in ("过滤明细行数", "有效明细行数"):
                        cell.number_format = '#,##0" 行"'
                    else:
                        cell.number_format = '¥#,##0.00'
                cells.append(cell)
            summary_ws.append(cells)
            actual_row = idx + 3
            summary_ws.row_dimensions[actual_row].height = 27 if is_final else 25
        summary_ws.freeze_panes = "A4"

        # 规则说明页
        tier_descriptions = []
        previous = 0
        for upper, rate in LEVELS[level]["tiers"]:
            if upper is None:
                tier_descriptions.append(f"超过{previous:,.0f}元部分 × {rate:.0%}")
            else:
                tier_descriptions.append(f"{previous:,.0f}至{upper:,.0f}元部分 × {rate:.0%}")
                previous = upper
        rule_rows = [
            ("项目", "本次计算口径"),
            ("产品等级", level),
            ("计费周期", "单月账单"),
            ("消费基数", "过滤黑名单后的官网价(¥)总额，折前价"),
            ("阶梯规则", "；".join(tier_descriptions)),
            ("月保底价", money_float(floor)),
            ("最终月度应收", "MAX(阶梯计费金额, 月保底价)"),
            ("黑名单关键词", "；".join(keywords)),
        ]
        for idx, row in enumerate(rule_rows):
            cells = []
            for value in row:
                cell = WriteOnlyCell(rules_ws, value=value)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                if idx == 0:
                    cell.font = Font(bold=True, color=WHITE)
                    cell.fill = PatternFill("solid", fgColor=NAVY)
                cells.append(cell)
            rules_ws.append(cells)

        # 写出前设置列宽和筛选区域。
        for col_idx, title in enumerate(detail_headers, 1):
            detail_ws.column_dimensions[get_column_letter(col_idx)].width = min(
                max(len(str(title)) + 3, 12), 32
            )
        detail_ws.auto_filter.ref = f"A1:{get_column_letter(len(detail_headers))}{imported_count + 1}"

        out_wb.save(output_path)
    finally:
        input_wb.close()

    return {
        "all_total": money_float(all_total),
        "filtered_total": money_float(filtered_total),
        "eligible_total": money_float(eligible_total),
        "tier": money_float(tier),
        "floor": money_float(floor),
        "monthly_final": money_float(monthly_final),
        "over_floor": money_float(over_floor),
        "filtered_count": filtered_count,
        "eligible_count": eligible_count,
        "imported_count": imported_count,
        "note": note,
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
        ttk.Label(row, text="计费周期：单月账单（直接按当月官网价总额计算）",
                  foreground="#555").pack(side="left", padx=(18, 0))

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
            result = create_report(self.input_path, sheet, level, self.keywords, output)
            messagebox.showinfo(
                "计算完成",
                f"产品等级：{level}\n"
                f"最终月度应收：¥{result['monthly_final']:,.2f}\n"
                f"计费备注：{result['note']}"
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
