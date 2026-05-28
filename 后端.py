import os
import sys
import socket
from datetime import datetime

import pandas as pd
from flask import Flask, render_template, jsonify, request, send_from_directory

# =========================
# 路径设置
# =========================
# 源码运行时：
#   RESOURCE_DIR = 后端.py 所在目录
#   BASE_DIR = 后端.py 所在目录
#
# EXE运行时：
#   RESOURCE_DIR = PyInstaller临时目录，用来读取 templates/static
#   BASE_DIR = EXE所在目录，用来读取 data/库存数据.xlsx
# =========================

if getattr(sys, "frozen", False):
    RESOURCE_DIR = sys._MEIPASS
    BASE_DIR = os.path.dirname(sys.executable)
else:
    RESOURCE_DIR = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR = RESOURCE_DIR

app = Flask(
    __name__,
    template_folder=os.path.join(RESOURCE_DIR, "templates"),
    static_folder=os.path.join(RESOURCE_DIR, "static")
)

# 同时兼容 Excel 2010 的 .xlsx 和老格式 .xls
DATA_FILE_XLSX = os.path.join(BASE_DIR, "data", "库存数据.xlsx")
DATA_FILE_XLS = os.path.join(BASE_DIR, "data", "库存数据.xls")

if os.path.exists(DATA_FILE_XLSX):
    DATA_FILE = DATA_FILE_XLSX
elif os.path.exists(DATA_FILE_XLS):
    DATA_FILE = DATA_FILE_XLS
else:
    DATA_FILE = DATA_FILE_XLSX

# Excel真实表头，不包含“序号”，序号由程序自动生成
REQUIRED_COLUMNS = [
    "产品名称",
    "产品图号",
    "产品数量",
    "产品单价",
    "产品金额",
    "产品单位当量（g）",
    "产品当量（t）",
    "产品来源",
    "产品入库日期",
    "库房名称1",
    "库房名称2",
    "库房名称3"
]

OPTIONAL_COLUMNS = [
    "库房名称4"
]

NEW_INVENTORY_SHEETS = ["火工品"]

NEW_COLUMN_MAP = {
    "名称": "产品名称",
    "图号": "产品图号",
    "数量": "产品数量",
    "单价": "产品单价",
    "金额": "产品金额",
    "单位当量（g）": "产品单位当量（g）",
    "产品当量（kg）": "产品当量（kg）",
    "批次": "产品来源",
    "入库日期": "产品入库日期",
    "产品类型": "库房名称1",
    "库房类型": "库房名称2",
    "库房编号": "库房名称3"
}


def get_local_ip():
    """
    获取本机局域网IP，用于启动时提示访问地址
    """
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        local_ip = s.getsockname()[0]
        s.close()
        return local_ip
    except Exception:
        return "127.0.0.1"


def clean_number_series(series):
    """
    清洗金额、数量、当量等数值字段：
    处理逗号、人民币符号、空格、中文空格、单位等问题
    """
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("，", "", regex=False)
        .str.replace("￥", "", regex=False)
        .str.replace("元", "", regex=False)
        .str.replace("吨", "", regex=False)
        .str.replace("t", "", regex=False)
        .str.replace("T", "", regex=False)
        .str.replace("g", "", regex=False)
        .str.replace("G", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace("\u3000", "", regex=False),
        errors="coerce"
    ).fillna(0)


def normalize_warehouse_code(value):
    code = str(value).strip()
    if code.lower() == "nan":
        return ""
    if code.endswith(".0"):
        return code[:-2]
    return code


def normalize_columns(raw_df):
    raw_df.columns = [
        str(col).strip()
        .replace(" ", "")
        .replace("\u3000", "")
        .replace("\n", "")
        .replace("\r", "")
        for col in raw_df.columns
    ]
    return raw_df


def read_inventory_excel():
    """
    读取并兼容两种库存结构：
    1. 旧版：库存数据 sheet，字段为产品名称/产品图号...
    2. 新版：火工品 sheet，字段为名称/图号/数量...
    """
    try:
        df = pd.read_excel(DATA_FILE, sheet_name="库存数据")
        df = normalize_columns(df)
        source_format = "old"

        # 兼容旧Excel：库存数据sheet缺少产品名称/产品图号时，
        # 从完整的库存数据1 sheet按行补齐这两列。
        name_cols = ["产品名称", "产品图号"]
        if any(col not in df.columns for col in name_cols):
            try:
                full_df = pd.read_excel(DATA_FILE, sheet_name="库存数据1")
                full_df = normalize_columns(full_df)

                if all(col in full_df.columns for col in name_cols) and len(full_df) >= len(df):
                    for col in name_cols:
                        if col not in df.columns:
                            df.insert(0 if col == "产品名称" else 1, col, full_df[col].iloc[:len(df)].values)
            except Exception:
                pass

        return df, source_format

    except ValueError:
        pass

    excel_file = pd.ExcelFile(DATA_FILE)
    frames = []

    for sheet_name in NEW_INVENTORY_SHEETS:
        if sheet_name not in excel_file.sheet_names:
            continue

        sheet_df = pd.read_excel(DATA_FILE, sheet_name=sheet_name)
        if sheet_df.empty:
            continue

        sheet_df = normalize_columns(sheet_df)
        sheet_df = sheet_df.dropna(how="all")
        if sheet_df.empty:
            continue

        sheet_df = sheet_df.rename(columns=NEW_COLUMN_MAP)
        sheet_df["库存Sheet"] = sheet_name

        if "库房名称2" not in sheet_df.columns:
            sheet_df["库房名称2"] = sheet_name
        else:
            sheet_df["库房名称2"] = sheet_df["库房名称2"].fillna("")
            sheet_df.loc[sheet_df["库房名称2"].astype(str).str.strip() == "", "库房名称2"] = sheet_name

        frames.append(sheet_df)

    if not frames:
        # 最后兜底读取第一个sheet，让缺字段提示能展示真实表头。
        fallback_df = pd.read_excel(DATA_FILE)
        return normalize_columns(fallback_df), "old"

    return pd.concat(frames, ignore_index=True), "new"


def read_inventory_data():
    """
    读取库存Excel数据
    """
    if not os.path.exists(DATA_FILE):
        return pd.DataFrame(columns=REQUIRED_COLUMNS), f"未找到数据文件：{DATA_FILE}"

    try:
        df, source_format = read_inventory_excel()

        if source_format == "new" and "产品当量（kg）" in df.columns:
            df["产品当量（t）"] = clean_number_series(df["产品当量（kg）"]) / 1000

        missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
        if missing_cols:
            return pd.DataFrame(columns=REQUIRED_COLUMNS + OPTIONAL_COLUMNS), (
                f"Excel缺少字段：{missing_cols}；当前识别到的字段：{list(df.columns)}"
            )

        for col in OPTIONAL_COLUMNS:
            if col not in df.columns:
                df[col] = ""

        if "产品当量（kg）" not in df.columns:
            df["产品当量（kg）"] = ""

        df = df[REQUIRED_COLUMNS + OPTIONAL_COLUMNS + ["产品当量（kg）"]].copy()

        # 删除完全空行
        df = df.dropna(how="all").reset_index(drop=True)

        # 自动生成序号
        df.insert(0, "序号", range(1, len(df) + 1))

        # 文本字段处理
        text_cols = [
            "序号",
            "产品名称",
            "产品图号",
            "产品来源",
            "库房名称1",
            "库房名称2",
            "库房名称3",
            "库房名称4"
        ]

        for col in text_cols:
            df[col] = (
                df[col]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.replace(" ", "", regex=False)
                .str.replace("\u3000", "", regex=False)
            )

        df["库房名称3"] = df["库房名称3"].apply(normalize_warehouse_code)

        # 数值字段处理
        number_cols = [
            "产品数量",
            "产品单价",
            "产品金额",
            "产品单位当量（g）",
            "产品当量（t）",
            "产品当量（kg）"
        ]

        for col in number_cols:
            df[col] = clean_number_series(df[col])

        # 如果产品金额为空或为0，则自动用 产品数量 × 产品单价 补充计算
        df["产品金额"] = df.apply(
            lambda row: round(row["产品数量"] * row["产品单价"], 2)
            if row["产品金额"] == 0 and row["产品数量"] > 0 and row["产品单价"] > 0
            else round(row["产品金额"], 2),
            axis=1
        )

        # 如果产品当量为空或为0，则自动用 产品单位当量 × 产品数量 / 1000000 补充计算
        df["产品当量（t）"] = df.apply(
            lambda row: round(row["产品单位当量（g）"] * row["产品数量"] / 1000000, 6)
            if row["产品当量（t）"] == 0 and row["产品单位当量（g）"] > 0 and row["产品数量"] > 0
            else round(row["产品当量（t）"], 6),
            axis=1
        )

        # 顶部“库存总当量（kg）”直接使用新版Excel的产品当量（kg）字段；
        # 老版Excel没有该列时，再由吨换算为kg兜底。
        df["产品当量（kg）"] = df.apply(
            lambda row: round(row["产品当量（t）"] * 1000, 2)
            if row["产品当量（kg）"] == 0 and row["产品当量（t）"] > 0
            else round(row["产品当量（kg）"], 2),
            axis=1
        )

        # 日期字段处理
        df["产品入库日期"] = pd.to_datetime(df["产品入库日期"], errors="coerce")
        df["入库年度"] = df["产品入库日期"].dt.year
        df["产品入库日期"] = df["产品入库日期"].dt.strftime("%Y-%m-%d")
        df["产品入库日期"] = df["产品入库日期"].fillna("")
        df["入库年度"] = df["入库年度"].fillna("未知").astype(str)

        return df, None

    except Exception as e:
        return pd.DataFrame(columns=REQUIRED_COLUMNS + OPTIONAL_COLUMNS), f"读取Excel失败：{str(e)}"

def read_trend_data():
    """
    读取第二个Sheet：交付及库存走势图
    表头要求：
    月份、交付金额、库存金额
    金额单位：亿元
    """
    try:
        trend_df = pd.read_excel(DATA_FILE, sheet_name="交付及库存走势图")

        trend_df = normalize_columns(trend_df)

        required_cols = ["月份", "交付金额", "库存金额"]
        missing_cols = [col for col in required_cols if col not in trend_df.columns]

        if missing_cols:
            return pd.DataFrame(columns=required_cols), f"第二个Sheet缺少字段：{missing_cols}"

        trend_df = trend_df[required_cols].copy()

        trend_df["月份"] = trend_df["月份"].fillna("").astype(str).str.strip()
        trend_df["交付金额"] = clean_number_series(trend_df["交付金额"])
        trend_df["库存金额"] = clean_number_series(trend_df["库存金额"])

        trend_df = trend_df[trend_df["月份"] != ""].copy()

        return trend_df, None

    except Exception as e:
        return pd.DataFrame(columns=["月份", "交付金额", "库存金额"]), f"读取交付及库存走势图失败：{str(e)}"


def read_prediction_data():
    """
    读取入库预测sheet：
    第一行：月份、6月、7月...
    后续行：205库房、207库房、208库房及对应预测当量
    """
    try:
        prediction_df = pd.read_excel(DATA_FILE, sheet_name="入库预测")
        prediction_df = normalize_columns(prediction_df)
        prediction_df = prediction_df.dropna(how="all")

        if prediction_df.empty or prediction_df.shape[1] < 2:
            return [], None

        months = [
            str(value).strip()
            for value in prediction_df.columns[1:].tolist()
            if str(value).strip() and str(value).strip().lower() != "nan"
        ]

        result = []
        target_warehouses = ["205库房", "207库房", "208库房"]
        warehouse_col = prediction_df.columns[0]

        for warehouse_name in target_warehouses:
            matched = prediction_df[
                prediction_df[warehouse_col].fillna("").astype(str).str.strip() == warehouse_name
            ]

            values = []
            if not matched.empty:
                raw_values = matched.iloc[0, 1:1 + len(months)]
                values = clean_number_series(raw_values).round(2).tolist()
            else:
                values = [0 for _ in months]

            result.append({
                "库房名称": warehouse_name.replace("库房", ""),
                "月份": months,
                "预测当量kg": values
            })

        return result, None

    except Exception as e:
        return [], f"读取入库预测失败：{str(e)}"


def classify_warehouse2(warehouse2):
    """
    库房名称2合并规则：
    伞类1库、伞类2库 -> 伞类
    个防1库、个防2库 -> 个防
    火工1库、火工2库 -> 火工
    座椅1库、座椅2库 -> 座椅
    其他 xxx1库、xxx2库 -> xxx
    """
    wh2 = str(warehouse2).strip()

    if not wh2:
        return "未填写"

    if wh2.endswith("1库") or wh2.endswith("2库"):
        return wh2[:-2]

    return wh2


@app.route("/")
def dashboard():
    return render_template("dashboard.html")


def get_layout_file_info(warehouse_name):
    """
    查找 data 目录下指定库房布局图。
    优先返回浏览器可直接显示的图片；如果只有 vsdx，则返回文件链接供打开/下载。
    """
    data_dir = os.path.join(BASE_DIR, "data")
    if not os.path.isdir(data_dir):
        return None

    base_name = f"{warehouse_name}库房布局图"
    image_exts = [".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp", ".svg"]
    file_exts = image_exts + [".vsdx", ".vsd", ".pdf"]

    for ext in image_exts:
        filename = base_name + ext
        if os.path.exists(os.path.join(data_dir, filename)):
            return {
                "warehouse": warehouse_name,
                "filename": filename,
                "url": f"/api/layout/{warehouse_name}",
                "kind": "image",
                "can_preview": True
            }

    for ext in file_exts:
        filename = base_name + ext
        if os.path.exists(os.path.join(data_dir, filename)):
            return {
                "warehouse": warehouse_name,
                "filename": filename,
                "url": f"/api/layout/{warehouse_name}",
                "kind": ext.replace(".", ""),
                "can_preview": ext.lower() in [".pdf"]
            }

    return None


@app.route("/api/layout/<warehouse_name>")
def api_layout_file(warehouse_name):
    info = get_layout_file_info(warehouse_name)
    if not info:
        return "未找到布局图文件", 404

    return send_from_directory(
        os.path.join(BASE_DIR, "data"),
        info["filename"],
        as_attachment=False
    )


@app.route("/api/inventory")
def api_inventory():
    df, error = read_inventory_data()
    trend_df, trend_error = read_trend_data()
    prediction_chart, prediction_error = read_prediction_data()


    if error:        return jsonify({
            "success": False,
            "error": error,
            "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "summary": {
                "total_amount": 0,
                "total_equivalent": 0,
                "total_quantity": 0
            },
            "charts": {
                "warehouse_stack_chart": [],
                "explosive_equivalent_chart": [],
                "trend_chart": [],
                "age_year_chart": [],
                "prediction_chart": []
            },
            "filters": {
                "warehouse1_options": [],
                "warehouse2_options": [],
                "warehouse3_options": []
            },
        })

    keyword = request.args.get("keyword", "").strip()
    warehouse1 = request.args.get("warehouse1", "成品").strip()
    warehouse2 = request.args.get("warehouse2", "").strip()
    warehouse3 = request.args.get("warehouse3", "").strip()

    filtered_df = df.copy()

    # 默认展示库房名称1包含“成品”的数据
    if warehouse1:
        filtered_df = filtered_df[
            filtered_df["库房名称1"].astype(str).str.contains(warehouse1, na=False)
        ]

    # 搜索产品名称 / 产品图号
    if keyword:
        filtered_df = filtered_df[
            filtered_df["产品名称"].astype(str).str.contains(keyword, case=False, na=False) |
            filtered_df["产品图号"].astype(str).str.contains(keyword, case=False, na=False)
        ]

    # 库房名称2筛选
    if warehouse2:
        filtered_df = filtered_df[filtered_df["库房名称2"] == warehouse2]

    # 库房名称3筛选
    if warehouse3:
        filtered_df = filtered_df[filtered_df["库房名称3"] == warehouse3]

    # =========================
    # 顶部指标
    # =========================

    # 库存总金额：Excel“金额”列全量合计
    total_amount = float(df["产品金额"].sum())

    # 产品总数量：Excel“数量”列全量合计
    total_quantity = float(df["产品数量"].sum())

    # 库存总当量：Excel“产品当量（kg）”列全量合计
    total_equivalent = float(df["产品当量（kg）"].sum())

    # =========================
    # 第一模块：库房大类数量和金额
    # =========================

    module_df = filtered_df.copy()

    def get_warehouse_big_type(warehouse2):
        wh2 = str(warehouse2).strip()

        if wh2.endswith("1库") or wh2.endswith("2库"):
            return wh2[:-2]

        return wh2 if wh2 else "未填写"

    def get_warehouse_sub_type(warehouse2):
        wh2 = str(warehouse2).strip()

        if wh2.endswith("1库"):
            return "1库"

        if wh2.endswith("2库"):
            return "2库"

        return "其他"

    module_df["库房大类"] = module_df["库房名称2"].apply(get_warehouse_big_type)
    module_df["库房分库"] = module_df["库房名称2"].apply(get_warehouse_sub_type)

    warehouse_stack_df = (
        module_df.groupby(["库房大类", "库房分库"])
        .agg(
            产品数量=("产品数量", "sum"),
            产品金额=("产品金额", "sum")
        )
        .reset_index()
    )

    # 固定显示顺序
    warehouse_order = ["火工", "座椅", "个防", "伞类"]

    # 如果以后还有其他库房大类，也自动追加到后面
    existing_types = warehouse_stack_df["库房大类"].dropna().unique().tolist()
    extra_types = [x for x in existing_types if x not in warehouse_order]
    final_order = warehouse_order + extra_types

    warehouse_stack_chart = []

    for big_type in final_order:
        temp = warehouse_stack_df[warehouse_stack_df["库房大类"] == big_type]

        if temp.empty:
            continue

        qty_1 = float(temp[temp["库房分库"] == "1库"]["产品数量"].sum())
        qty_2 = float(temp[temp["库房分库"] == "2库"]["产品数量"].sum())
        qty_other = float(temp[temp["库房分库"] == "其他"]["产品数量"].sum())

        amount_1 = float(temp[temp["库房分库"] == "1库"]["产品金额"].sum())
        amount_2 = float(temp[temp["库房分库"] == "2库"]["产品金额"].sum())
        amount_other = float(temp[temp["库房分库"] == "其他"]["产品金额"].sum())

        warehouse_stack_chart.append({
            "库房大类": big_type,
            "数量_1库": qty_1,
            "数量_2库": qty_2,
            "数量_其他": qty_other,
            "金额_1库": amount_1,
            "金额_2库": amount_2,
            "金额_其他": amount_other
        })


    # =========================
    # 火工库按库房名称3统计产品当量
    # 单位：kg
    # 警戒值：按库房名称3设定
    # =========================

    explosive_warning_limit = {
        "205": 500,
        "207": 4000,
        "207-1": 50,
        "208": 4000,
        "270": 6890
    }

    # 左侧火工库房当量统计：按火工品sheet里的全部库房编号统计。
    # 宏伟航空器只作为库房名称4的分项颜色显示，不作为横坐标分类显示。
    default_explosive_order = ["208", "207", "207-1", "205", "270"]
    existing_explosive_codes = [
        code
        for code in df["库房名称3"].dropna().astype(str).map(normalize_warehouse_code).unique().tolist()
        if code
    ]
    explosive_order = default_explosive_order + [
        code for code in existing_explosive_codes if code not in default_explosive_order
    ]

    explosive_df = df[
        df["库房名称3"].astype(str) != ""
    ].copy()

    if len(explosive_df) > 0:
        explosive_df["是否宏伟航空器"] = explosive_df["库房名称4"].astype(str).str.strip() == "宏伟航空器"

        explosive_group = (
            explosive_df.groupby("库房名称3")
            .agg(
                产品当量吨=("产品当量（t）", "sum"),
                宏伟航空器当量吨=("产品当量（t）", lambda s: s[explosive_df.loc[s.index, "是否宏伟航空器"]].sum()),
                产品数量=("产品数量", "sum"),
                产品金额=("产品金额", "sum")
            )
            .reset_index()
        )

        explosive_group["库房名称3"] = explosive_group["库房名称3"].replace("", "未填写")

        explosive_equivalent_list = []

        for wh3 in explosive_order:
            temp = explosive_group[explosive_group["库房名称3"] == wh3]

            if len(temp) > 0:
                actual_kg = float(temp["产品当量吨"].sum()) * 1000
                hongwei_kg = float(temp["宏伟航空器当量吨"].sum()) * 1000
                quantity = float(temp["产品数量"].sum())
                amount = float(temp["产品金额"].sum())
            else:
                actual_kg = 0
                hongwei_kg = 0
                quantity = 0
                amount = 0

            normal_kg = max(actual_kg - hongwei_kg, 0)

            warehouse_detail_df = explosive_df[
                explosive_df["库房名称3"].replace("", "未填写") == wh3
            ].copy()

            # 宏伟没有警戒值，默认为0
            limit_kg = explosive_warning_limit.get(wh3, 0)
            warning_kg = limit_kg * 0.9 if limit_kg > 0 else 0

            explosive_equivalent_list.append({
                "分类名称": wh3,
                "产品当量kg": round(actual_kg, 2),
                "普通产品当量kg": round(normal_kg, 2),
                "宏伟航空器当量kg": round(hongwei_kg, 2),
                "警戒值kg": round(limit_kg, 2),
                "百分之九十警戒值kg": round(warning_kg, 2),
                "是否超90警戒线": actual_kg >= warning_kg if limit_kg > 0 else False,
                "产品数量": quantity,
                "产品金额": amount,
                "布局图": get_layout_file_info(wh3) if wh3 == "208" else None,
                "当量明细": warehouse_detail_df.to_dict(orient="records")
            })

        explosive_equivalent_chart = pd.DataFrame(explosive_equivalent_list)

    else:
        explosive_equivalent_chart = pd.DataFrame([
            {
                "分类名称": wh3,
                "产品当量kg": 0,
                "普通产品当量kg": 0,
                "宏伟航空器当量kg": 0,
                "警戒值kg": explosive_warning_limit.get(wh3, 0),
                "百分之九十警戒值kg": round(explosive_warning_limit.get(wh3, 0) * 0.9, 2),
                "是否超90警戒线": False,
                "产品数量": 0,
                "产品金额": 0,
                "布局图": get_layout_file_info(wh3) if wh3 == "208" else None,
                "当量明细": []
            }
            for wh3 in explosive_order
        ])

    # =========================
    # 第二模块：成品库库龄年度分析
    # =========================

    age_df = df[
        df["库房名称1"].astype(str).str.contains("成品", na=False)
    ].copy()

    age_year_chart = (
        age_df.groupby("入库年度")
        .agg(
            产品数量=("产品数量", "sum"),
            产品金额=("产品金额", "sum")
        )
        .reset_index()
        .sort_values("入库年度")
    )

    # =========================
    # 第二模块：交付走势与库存金额走势图
    # =========================

    if trend_error:
        trend_chart = []
    else:
        trend_chart = trend_df.to_dict(orient="records")

    # =========================
    # 筛选项
    # =========================

    warehouse1_options = sorted(df["库房名称1"].dropna().astype(str).unique().tolist())
    warehouse2_options = sorted(df["库房名称2"].dropna().astype(str).unique().tolist())
    warehouse3_options = sorted(df["库房名称3"].dropna().astype(str).unique().tolist())

    return jsonify({
        "success": True,
        "error": "",
        "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "summary": {
            "total_amount": total_amount,
            "total_equivalent": total_equivalent,
            "total_quantity": total_quantity
        },
        "charts": {
            "warehouse_stack_chart": warehouse_stack_chart,
            "explosive_equivalent_chart": explosive_equivalent_chart.to_dict(orient="records"),
            "trend_chart": trend_chart,
            "age_year_chart": age_year_chart.to_dict(orient="records"),
            "prediction_chart": [] if prediction_error else prediction_chart
        },
        "filters": {
            "warehouse1_options": warehouse1_options,
            "warehouse2_options": warehouse2_options,
            "warehouse3_options": warehouse3_options
        }
    })


if __name__ == "__main__":
    ip = get_local_ip()
    print("=" * 60)
    print("火工品库房看板已启动")
    print(f"本机访问：http://127.0.0.1:8000")
    print(f"局域网访问：http://{ip}:8000")
    print("请确保防火墙允许8000端口访问")
    print("=" * 60)

    app.run(host="0.0.0.0", port=8000, debug=False)
