from pathlib import Path
import json
import base64

import pandas as pd


ROOT = Path(__file__).resolve().parent
EXCEL_PATH = ROOT / "作图数据整理.xlsx"
PBROE_PATH = ROOT / "PB-ROE和资产组合净值.xlsx"
LOGO_PATH = ROOT / "logo.png"
# 输出的主页面文件名（GitHub Pages 默认会使用 index.html 作为首页）
DASHBOARD_HTML = ROOT / "dashboard.html"
INDEX_HTML = ROOT / "index.html"


def sheet_to_records(df):
    """Convert DataFrame to plain Python records with JSON‑friendly values."""

    def convert_value(v):
        # Convert timestamps / dates to string, only keep date part (YYYY-MM-DD)
        if hasattr(v, "isoformat"):
            # 如果是日期/时间类型，只返回日期部分，不包含时间
            if hasattr(v, "date"):
                # datetime 类型，只取日期部分
                return v.date().isoformat()
            else:
                # date 类型，直接格式化
                return v.isoformat()
        
        # 处理字符串格式的日期（包含时间部分的情况）
        if isinstance(v, str):
            # 如果字符串是 ISO 格式日期时间（如 "2016-01-04T00:00:00"），只保留日期部分
            if "T" in v and len(v) > 10:
                # 提取日期部分（YYYY-MM-DD）
                date_part = v.split("T")[0]
                # 验证是否是有效的日期格式
                try:
                    from datetime import datetime
                    datetime.strptime(date_part, "%Y-%m-%d")
                    return date_part
                except ValueError:
                    pass
        
        return v

    records = []
    for _, row in df.iterrows():
        rec = {}
        for k, v in row.items():
            rec[str(k)] = convert_value(v)
        records.append(rec)
    return records


def build_data_and_config():
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"未找到文件: {EXCEL_PATH}")

    xls = pd.ExcelFile(EXCEL_PATH)

    # 读取各个 sheet
    sheets_needed = [
        "风险图1-新高个股占比",
        "资金图1",
        "资金图2",
        "因子图1",
        "因子图2",
    ]

    data_by_sheet = {}
    for name in sheets_needed:
        if name not in xls.sheet_names:
            continue
        df = pd.read_excel(EXCEL_PATH, sheet_name=name)
        data_by_sheet[name] = sheet_to_records(df)
    
    # 读取PB-ROE数据
    if PBROE_PATH.exists():
        pb_roe_xls = pd.ExcelFile(PBROE_PATH)
        if "PB-ROE" in pb_roe_xls.sheet_names:
            pb_roe_df = pd.read_excel(PBROE_PATH, sheet_name="PB-ROE")
            # 重命名列，使用Unnamed: 1作为名称
            pb_roe_df = pb_roe_df.rename(columns={"Unnamed: 1": "名称"})
            data_by_sheet["PB-ROE"] = sheet_to_records(pb_roe_df)
        
        # 读取资产配置净值数据
        if "资产配置净值" in pb_roe_xls.sheet_names:
            asset_df = pd.read_excel(PBROE_PATH, sheet_name="资产配置净值")
            data_by_sheet["资产配置净值"] = sheet_to_records(asset_df)

    # 图表配置：说明每张图的 x 轴、y 轴字段及含义
    chart_configs = {
        "因子": [
            {
                "id": "factor1",
                "title": "外币资金占货币资金比",
                "displayTitle": "外币资金占货币资金比因子 vs 基准",
                "sheet": "因子图1",
                "x": "trade_dt",
                "description": "• 外币资金占货币资金比因子：反映公司外币资金在货币资金中的占比情况，用于描述公司海外业务形成的资金结构特征\n• 基准：用于对比的基准线",
                "lines": [
                    {
                        "name": "外币资金占货币资金比因子",
                        "field": "外币资金占货币资金比",
                        "axis": "y1",
                    },
                    {"name": "基准", "field": "基准", "axis": "y1"},
                ],
            },
            {
                "id": "factor2",
                "title": "主要客户占比稳定性",
                "displayTitle": "主要客户占比稳定性因子 vs 基准",
                "sheet": "因子图2",
                "x": "trade_dt",
                "description": "• 主要客户占比稳定性因子：衡量公司主要客户集中度的稳定性，反映公司客户结构的健康程度和业务风险\n• 基准：用于对比的基准线",
                "lines": [
                    {
                        "name": "主要客户占比稳定性因子",
                        "field": "主要客户占比稳定性",
                        "axis": "y1",
                    },
                    {"name": "基准", "field": "基准", "axis": "y1"},
                ],
            },
        ],
        "资金": [
            {
                "id": "fund1",
                "title": "主力累计净买入",
                "displayTitle": "主力累计净买入 vs Wind全A收盘价",
                "sheet": "资金图1",
                "x": "Unnamed: 0",
                "description": "• 主力累计净买入(亿元)：反映大资金的累计净买入情况，单位为亿元，可用于判断市场资金面和情绪变化\n• Wind全A收盘价（右轴）：Wind全A指数的收盘价，显示在右轴",
                "lines": [
                    {
                        "name": "主力累计净买入(亿元)",
                        "field": "主力累计净买入(亿元)",
                        "axis": "y1",
                    },
                    {
                        "name": "Wind全A收盘价（右轴）",
                        "field": "万得全A(右轴)",
                        "axis": "y2",
                    },
                ],
            },
            {
                "id": "fund2",
                "title": "博弈 / 存量",
                "displayTitle": "博弈/存量 vs Wind全A收盘价",
                "sheet": "资金图2",
                "x": "Unnamed: 0",
                "description": "• 博弈/存量：反映市场交易活跃度和资金利用效率。计算方式：博弈/存量 = 成交额 / 保证金余额\n• 上轨80分位数：80分位数参考线，用于判断博弈/存量的高位水平\n• 下轨20分位数：20分位数参考线，用于判断博弈/存量的低位水平\n• Wind全A收盘价（右轴）：Wind全A指数的收盘价，显示在右轴",
                "lines": [
                    {
                        "name": "博弈/存量",
                        "field": "博弈/存量",
                        "axis": "y1",
                    },
                    {
                        "name": "上轨80分位数",
                        "field": "上轨80分位数",
                        "axis": "y1",
                    },
                    {
                        "name": "下轨20分位数",
                        "field": "下轨20分位数",
                        "axis": "y1",
                    },
                    {
                        "name": "Wind全A收盘价（右轴）",
                        "field": "wind全A收盘价(右)",
                        "axis": "y2",
                    },
                ],
            },
        ],
        "风险": [
            {
                "id": "risk1",
                "title": "新高个股占比",
                "displayTitle": "新高个股占比 vs 上证综指",
                "sheet": "风险图1-新高个股占比",
                "x": "date",
                "description": "• 新高个股占比：反映市场创新高股票的比例，可用于判断市场情绪和趋势强度\n• 上证综指：上海证券交易所综合股价指数，显示在右轴",
                "lines": [
                    {"name": "新高个股占比", "field": "占比", "axis": "y1"},
                    {"name": "上证综指", "field": "上证综指(右)", "axis": "y2"},
                ],
            },
            {
                "id": "risk2",
                "title": "预测 PB - ROE",
                "sheet": "PB-ROE",
                "type": "scatter",
                "x": "预测ROE",
                "y": "预测PB",
                "text": "名称",
                "description": "• X轴（预测ROE）：预测的净资产收益率，反映公司的盈利能力\n• Y轴（预测PB）：预测的市净率，反映公司的估值水平\n• 散点：每个散点代表一只股票，散点上的文字为股票名称",
            },
        ],
        "量化资产配置": [
            {
                "id": "asset1",
                "title": "中低波组合净值",
                "displayTitle": "中低波组合净值 vs 基准",
                "sheet": "资产配置净值",
                "x": "date",
                "description": "• 中波组合：中波动率组合的净值表现\n• 低波组合：低波动率组合的净值表现\n• 基准：资产风险平价基准，用于对比的基准线",
                "lines": [
                    {
                        "name": "中波组合",
                        "field": "中波组合",
                        "axis": "y1",
                    },
                    {
                        "name": "低波组合",
                        "field": "低波组合",
                        "axis": "y1",
                    },
                    {
                        "name": "基准",
                        "field": "资产风险平价",
                        "axis": "y1",
                    },
                ],
            },
        ],
    }

    return data_by_sheet, chart_configs


def build_html(data_by_sheet, chart_configs):
    # 读取logo并转换为base64
    logo_base64 = ""
    if LOGO_PATH.exists():
        with open(LOGO_PATH, "rb") as f:
            logo_data = f.read()
            logo_base64 = base64.b64encode(logo_data).decode("utf-8")
    
    # 嵌入到前端的 JSON 字符串
    data_json = json.dumps(data_by_sheet, ensure_ascii=False)
    config_json = json.dumps(chart_configs, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes" />
  <title>量化分析监控看板</title>
  <script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <style>
    * {{
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }}
    html {{
      min-width: 400px;
      overflow-x: hidden;
    }}
    body {{
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, 'Noto Sans', sans-serif;
      background-color: #eaeaea;
      color: #2f3640;
      height: 100vh;
      display: flex;
      flex-direction: column;
      margin: 0;
      padding: 0;
      min-width: 400px;
      overflow-x: hidden;
    }}
    .header {{
      height: 70px;
      background: #eaeaea;
      display: flex;
      align-items: center;
      padding: 0 24px;
      flex-shrink: 0;
    }}
    .header-left {{
      display: flex;
      align-items: center;
    }}
    .header-logo {{
      height: 46px;
      margin-right: 16px;
    }}
    .header-title-main {{
      font-size: 18px;
      font-weight: 600;
      margin-bottom: 2px;
      letter-spacing: 2px;
    }}
    .header-title-sub {{
      font-size: 12px;
      color: #7f8c8d;
      letter-spacing: 0.1px;
    }}
    .layout {{
      flex: 1;
      display: flex;
      min-height: 0;
      gap: 10px;
      padding: 10px;
      background: #eaeaea;
    }}
    .sidebar {{
      width: 260px;
      background: #ffffff;
      color: #2f3640;
      padding: 0;
      display: flex;
      flex-direction: column;
      border-radius: 4px;
      flex-shrink: 0;
      transition: transform 0.3s ease;
      position: relative;
      overflow: hidden;
    }}
    .sidebar-content {{
      flex: 1;
      padding: 20px 16px;
      overflow-y: auto;
      min-height: 0;
    }}
    .sidebar-content::-webkit-scrollbar {{
      width: 8px;
    }}
    .sidebar-content::-webkit-scrollbar-track {{
      background: transparent;
      border: none;
    }}
    .sidebar-content::-webkit-scrollbar-thumb {{
      background: #d0d0d0;
      border-radius: 4px;
      border: none;
    }}
    .sidebar-content::-webkit-scrollbar-thumb:hover {{
      background: #b0b0b0;
    }}
    .sidebar h1 {{
      font-size: 18px;
      margin-bottom: 16px;
      font-weight: 600;
      color: #2f3640;
    }}
    .nav-browse-title {{
      font-size: 18px;
      font-weight: bold;
      margin: 0 0 20px 0;
      color: #333333;
      padding-bottom: 8px;
      border-bottom: 1px solid #eaeaea;
    }}
    .nav-section-title {{
      font-size: 16px;
      font-weight: 600;
      margin: 16px 0 6px;
      text-transform: uppercase;
      letter-spacing: 1px;
      color: #333333;
    }}
    .nav-list {{
      list-style: none;
    }}
    .nav-item {{
      padding: 8px 10px 8px 0;
      margin-bottom: 4px;
      border-radius: 0;
      font-size: 15px;
      cursor: pointer;
      transition: color 0.15s ease, font-weight 0.15s ease;
      color: #666666;
      background: transparent;
      display: flex;
      align-items: center;
      gap: 10px;
    }}
    .nav-item:hover {{
      color: #005bac;
    }}
    .nav-item:hover .nav-icon {{
      color: #005bac;
    }}
    .nav-item.active {{
      color: #005bac;
      font-weight: 600;
      background: transparent;
    }}
    .nav-item.active .nav-icon {{
      color: #005bac;
    }}
    .nav-icon {{
      width: 16px;
      height: 16px;
      flex-shrink: 0;
      color: #666666;
      transition: color 0.15s ease;
    }}
    .nav-icon svg {{
      width: 100%;
      height: 100%;
    }}
    .sidebar-footer {{
      flex-shrink: 0;
      padding: 8px 16px;
      background: #ffffff;
      display: flex;
      flex-direction: row;
      gap: 70px;
      justify-content: center;
      z-index: 10;
      position: relative;
    }}
    .sidebar-footer::before {{
      content: '';
      position: absolute;
      top: 0;
      left: 16px;
      right: 16px;
      height: 1px;
      background-color: #eaeaea;
    }}
    .sidebar-action {{
      display: flex;
      align-items: center;
      gap: 8px;
      cursor: pointer;
      color: #666666;
      transition: color 0.15s ease;
      font-size: 15px;
      padding: 8px 0;
    }}
    .sidebar-action:hover {{
      color: #005bac;
    }}
    .sidebar-action-icon {{
      width: 18px;
      height: 18px;
      flex-shrink: 0;
    }}
    .sidebar-action-icon svg {{
      width: 100%;
      height: 100%;
    }}
    .share-modal {{
      display: none;
      position: fixed;
      z-index: 1000;
    }}
    .share-modal.active {{
      display: block;
    }}
    .share-modal-content {{
      background-color: #f8f9fa;
      padding: 12px 16px;
      border-radius: 4px;
      max-width: 400px;
      min-width: 300px;
      border: 1px solid #e9ecef;
      position: relative;
      box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
    }}
    .share-modal-header {{
      font-size: 14px;
      font-weight: 400;
      color: #666666;
      margin-bottom: 12px;
    }}
    .share-modal-url {{
      background-color: #ffffff;
      border: 1px solid #e9ecef;
      border-radius: 4px;
      padding: 12px;
      font-size: 14px;
      color: #666666;
      user-select: all;
      cursor: text;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }}
    .main {{
      flex: 1;
      padding: 16px 20px;
      display: flex;
      flex-direction: column;
      overflow: hidden;
      background: #ffffff;
      border-radius: 4px;
      min-width: 0;
    }}
    .sections-container {{
      flex: 1;
      display: flex;
      flex-direction: column;
      gap: 0;
      overflow: hidden;
      padding-right: 8px;
    }}
    .section {{
      display: none;
      flex-direction: column;
      background: #ffffff;
      border-radius: 4px;
      padding: 4px 10px 4px 10px;
      flex: 1;
      min-height: 0;
    }}
    .section.active {{
      display: flex;
    }}
    .section-title {{
      font-size: 18px;
      font-weight: bold;
      margin-bottom: 20px;
      color: #333333;
      padding-bottom: 8px;
      border-bottom: 1px solid #eaeaea;
    }}
    .section-charts {{
      flex: 1;
      display: flex;
      flex-direction: column;
      gap: 24px;
      overflow-y: auto;
      min-height: 0;
    }}
    .section-charts::-webkit-scrollbar {{
      width: 8px;
    }}
    .section-charts::-webkit-scrollbar-track {{
      background: transparent;
      border: none;
    }}
    .section-charts::-webkit-scrollbar-thumb {{
      background: #d0d0d0;
      border-radius: 4px;
      border: none;
    }}
    .section-charts::-webkit-scrollbar-thumb:hover {{
      background: #b0b0b0;
    }}
    /* 隐藏rangeslider中的折线图，只显示灰色控制条 */
    .js-plotly-plot .rangeslider-container {{
      height: 20px !important;
      min-height: 20px !important;
    }}
    .js-plotly-plot .rangeslider-container svg {{
      display: none !important;
    }}
    .js-plotly-plot .rangeslider-container .js-plotly-plot {{
      display: none !important;
    }}
    .chart-wrapper {{
      display: flex;
      flex-direction: column;
      background: #ffffff;
      border-radius: 4px;
      /* 让标题更贴近左边界：缩小内边距 */
      padding: 0px 16px 16px 0px;
    }}
    .chart-header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 12px;
    }}
    .chart-title {{
      font-size: 16px;
      font-weight: 600;
      color: #333333;
    }}
    .chart-download {{
      display: flex;
      align-items: center;
      gap: 6px;
      cursor: pointer;
      color: #666666;
      transition: color 0.15s ease;
      flex-shrink: 0;
      font-size: 15px;
    }}
    .chart-download:hover {{
      color: #005bac;
    }}
    .chart-download-icon {{
      width: 20px;
      height: 20px;
      flex-shrink: 0;
    }}
    .chart-download-icon svg {{
      width: 100%;
      height: 100%;
    }}
    .chart-info {{
      display: flex;
      align-items: center;
      gap: 6px;
      cursor: pointer;
      color: #666666;
      transition: color 0.15s ease;
      flex-shrink: 0;
      font-size: 15px;
      margin-right: 16px;
    }}
    .chart-info:hover {{
      color: #005bac;
    }}
    .chart-info-icon {{
      width: 20px;
      height: 20px;
      flex-shrink: 0;
    }}
    .chart-info-icon svg {{
      width: 100%;
      height: 100%;
    }}
    .chart-info-panel {{
      display: none;
      background-color: #f8f9fa;
      border: 1px solid #e9ecef;
      border-radius: 4px;
      padding: 12px 16px;
      margin-top: 8px;
      margin-bottom: 12px;
      font-size: 14px;
      color: #666666;
      line-height: 1.6;
    }}
    .chart-info-panel .info-line {{
      display: flex;
      align-items: flex-start;
      margin-bottom: 8px;
    }}
    .chart-info-panel .info-line:last-child {{
      margin-bottom: 0;
    }}
    .chart-info-panel .color-indicator {{
      width: 12px;
      height: 12px;
      border-radius: 2px;
      flex-shrink: 0;
      margin-right: 8px;
      margin-top: 3px;
    }}
    .chart-info-panel .info-text {{
      flex: 1;
    }}
    .chart-info-panel.active {{
      display: block;
    }}
    .chart-subtitle {{
      font-size: 12px;
      color: #7f8c8d;
    }}
    .chart-container {{
      width: 100%;
      height: 500px;
      min-width: 300px;
      max-width: 100%;
      background: #ffffff;
      border-radius: 4px;
      padding: 10px;
      box-sizing: border-box;
    }}
    .chart {{
      width: 100%;
      height: 100%;
      min-width: 300px;
    }}
    
    /* 响应式布局：窗口缩小时隐藏侧边栏 */
    @media (max-width: 900px) {{
      .sidebar {{
        display: none;
      }}
      .layout {{
        gap: 0;
      }}
    }}
    /* 移动端优化 */
    @media (max-width: 768px) {{
      .header {{
        height: 60px;
        padding: 0 12px;
      }}
      .header-logo {{
        height: 36px;
        margin-right: 12px;
      }}
      .header-title-main {{
        font-size: 14px;
      }}
      .header-title-sub {{
        font-size: 11px;
      }}
      .main {{
        padding: 12px;
      }}
      .section {{
        padding: 4px 12px 16px 12px;
      }}
      .section-title {{
        font-size: 16px;
        margin-bottom: 16px;
      }}
      .chart-header {{
        margin-bottom: 10px;
      }}
      .chart-title {{
        font-size: 15px;
      }}
      .chart-download, .chart-info {{
        font-size: 13px;
        gap: 4px;
      }}
      .chart-download-icon, .chart-info-icon {{
        width: 18px;
        height: 18px;
      }}
      .chart-info-panel {{
        font-size: 13px;
        padding: 10px 12px;
      }}
      .chart-wrapper {{
        margin-bottom: 24px;
        padding: 12px;
      }}
      .chart-container {{
        height: 400px;
        padding: 8px;
      }}
    }}
    @media (max-width: 480px) {{
      .header {{
        height: 56px;
        padding: 0 10px;
      }}
      .header-logo {{
        height: 32px;
        margin-right: 10px;
      }}
      .header-title-main {{
        font-size: 13px;
      }}
      .header-title-sub {{
        font-size: 10px;
      }}
      .main {{
        padding: 10px;
      }}
      .section {{
        padding: 4px 10px 12px 10px;
      }}
      .section-title {{
        font-size: 15px;
        margin-bottom: 12px;
      }}
      .chart-title {{
        font-size: 14px;
      }}
      .chart-download, .chart-info {{
        font-size: 12px;
      }}
      .chart-wrapper {{
        margin-bottom: 20px;
        padding: 10px;
      }}
      .chart-container {{
        height: 350px;
        padding: 6px;
      }}
    }}
  </style>
</head>
<body>
  <div class="header">
    <div class="header-left">
      <img src="data:image/png;base64,{logo_base64}" alt="NORTHEAST SECURITIES" class="header-logo">
      <div>
        <div class="header-title-main">量化分析监控看板</div>
        <div class="header-title-sub">Quant Analysis Dashboard</div>
      </div>
    </div>
  </div>
  <div class="layout">
  <div class="sidebar">
    <div id="nav" class="sidebar-content"></div>
  </div>
  <div class="main">
    <div class="sections-container" id="sections-container"></div>
  </div>
  </div>

  <script>
    const dataBySheet = {data_json};
    const chartConfigs = {config_json};
    // 记录每个图是否已经创建，避免重复创建以及在隐藏状态下创建导致尺寸异常
    const chartInstances = {{}};

    // 下载图表数据为CSV（原始Excel表格数据）
    function downloadChartData(cfg) {{
      const records = dataBySheet[cfg.sheet] || [];
      if (records.length === 0) {{
        alert('没有可下载的数据');
        return;
      }}

      // 获取所有列名（从第一条记录中获取所有键）
      const headers = Object.keys(records[0] || {{}});
      if (headers.length === 0) {{
        alert('没有可下载的数据');
        return;
      }}

      // 构建CSV内容，包含所有原始列
      let csvContent = '';
      csvContent += headers.join(',') + '\\n';
      
      records.forEach(record => {{
        const row = headers.map(header => {{
          const value = record[header];
          // 处理包含逗号、引号或换行符的值
          if (value === null || value === undefined) {{
            return '';
          }}
          const strValue = String(value);
          // 如果包含逗号、引号或换行符，需要用引号包裹并转义引号
          if (strValue.includes(',') || strValue.includes('"') || strValue.includes('\\n')) {{
            return '"' + strValue.replace(/"/g, '""') + '"';
          }}
          return strValue;
        }});
        csvContent += row.join(',') + '\\n';
      }});

      // 创建下载链接
      const blob = new Blob(['\\ufeff' + csvContent], {{ type: 'text/csv;charset=utf-8;' }});
      const link = document.createElement('a');
      const url = URL.createObjectURL(blob);
      link.setAttribute('href', url);
      link.setAttribute('download', `${{cfg.title || cfg.id}}.csv`);
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }}

    // 根据线条名称获取颜色
    function getLineColor(cfg, lineName) {{
      if (cfg.id === 'fund2' && lineName === '上轨80分位数') {{
        return '#FFABAB';
      }} else if (cfg.id === 'fund2' && lineName === '下轨20分位数') {{
        return '#B6E880';
      }} else if (cfg.id === 'asset1' && lineName === '低波组合') {{
        return '#B6E880';
      }} else if (cfg.id === 'asset1' && lineName === '中波组合') {{
        return '#005bac';
      }} else if (cfg.id === 'asset1' && lineName === '基准') {{
        return '#FFCB05';
      }} else if (lineName === '基准' || lineName === 'Wind全A收盘价（右轴）' || lineName === 'Wind全A收盘价' || lineName === '上证综指') {{
        return '#FFCB05';
      }} else {{
        return '#005bac';
      }}
    }}

    // 解析说明文本并生成带颜色标识的HTML
    function parseDescription(cfg, description) {{
      if (!description) return '';
      
      const lines = description.split('\\n');
      let html = '';
      
      for (let line of lines) {{
        if (line.trim() === '') {{
          html += '<br>';
          continue;
        }}
        
        // 检查是否是线条说明行（以•开头）
        if (line.trim().startsWith('•')) {{
          const text = line.trim().substring(1).trim();
          
          // 尝试匹配线条名称（去掉颜色文字说明）
          let lineName = text;
          let color = null;
          
          // 查找对应的线条名称
          if (cfg.lines) {{
            for (let lineCfg of cfg.lines) {{
              if (text.includes(lineCfg.name)) {{
                lineName = lineCfg.name;
                color = getLineColor(cfg, lineName);
                break;
              }}
            }}
          }}
          
          // 如果没有找到匹配的线条，尝试从文本中提取（去掉括号中的颜色说明）
          if (!color) {{
            const match = text.match(/^([^（(]+)/);
            if (match) {{
              const name = match[1].trim();
              if (cfg.lines) {{
                for (let lineCfg of cfg.lines) {{
                  if (name.includes(lineCfg.name) || lineCfg.name.includes(name)) {{
                    lineName = lineCfg.name;
                    color = getLineColor(cfg, lineName);
                    break;
                  }}
                }}
              }}
            }}
          }}
          
          // 如果找到了颜色，生成带颜色标识的HTML
          if (color) {{
            // 去掉文本中的颜色说明（蓝色）、（黄色）等
            const cleanText = text.replace(/（[^）]+）|\\([^)]+\\)/g, '').trim();
            html += `<div class="info-line"><span class="color-indicator" style="background-color: ${{color}}"></span><span class="info-text">${{cleanText}}</span></div>`;
          }} else {{
            html += `<div class="info-line"><span class="info-text">${{text}}</span></div>`;
          }}
        }} else {{
          // 非线条说明行，直接显示
          html += `<div class="info-line"><span class="info-text">${{line}}</span></div>`;
        }}
      }}
      
      return html;
    }}

    // 显示/隐藏图表数据说明
    function showChartInfo(cfg) {{
      if (!cfg.description) return;
      
      const panelId = `chart-info-panel-${{cfg.id}}`;
      const panel = document.getElementById(panelId);
      
      if (panel) {{
        // 切换显示/隐藏
        panel.classList.toggle('active');
      }}
    }}

    // 下载所有数据
    // 清理 Excel sheet 名称中的非法字符
    function sanitizeSheetName(name) {{
      if (!name) return 'Sheet';
      // Excel sheet 名称不能包含: \ / ? * [ ]
      let sanitized = name.replace(/[\\/?:*\\[\\]]/g, '');
      // 移除首尾的单引号
      sanitized = sanitized.replace(/^['"]+|['"]+$/g, '');
      // 如果清理后为空，使用默认名称
      if (!sanitized || sanitized.trim() === '') {{
        sanitized = 'Sheet';
      }}
      // Excel限制sheet名称最多31个字符
      if (sanitized.length > 31) {{
        sanitized = sanitized.substring(0, 31);
      }}
      return sanitized;
    }}

    function downloadAllData() {{
      try {{
        // 检查xlsx库是否加载
        if (typeof XLSX === 'undefined') {{
          alert('Excel库加载失败，请刷新页面重试');
          console.error('XLSX library is not loaded');
          return;
        }}
        
        // 创建工作簿
        const workbook = XLSX.utils.book_new();
        
        // 用于跟踪已使用的 sheet 名称，避免重复
        const usedSheetNames = new Set();
        
        // 收集所有图表的数据，每个图表作为一个sheet
        let hasData = false;
        
        Object.keys(chartConfigs).forEach(category => {{
          chartConfigs[category].forEach((cfg, idx) => {{
            const records = dataBySheet[cfg.sheet] || [];
            if (records.length > 0) {{
              hasData = true;
              
              // 获取所有列名
              const headers = Object.keys(records[0] || {{}});
              if (headers.length > 0) {{
                // 准备数据：第一行是表头，后面是数据行
                const sheetData = [headers];
                
                // 识别日期列（通过列名判断）
                const dateColumns = new Set();
                headers.forEach(header => {{
                  const headerLower = String(header).toLowerCase();
                  if (headerLower.includes('date') || headerLower.includes('日期') || 
                      headerLower.includes('trade_dt') || headerLower === 'date' ||
                      headerLower === 'trade_dt' || headerLower.includes('dt')) {{
                    dateColumns.add(header);
                  }}
                }});
                
                records.forEach(record => {{
                  const row = headers.map(header => {{
                    let value = record[header];
                    
                    // 处理 null 和 undefined
                    if (value === null || value === undefined) {{
                      return '';
                    }}
                    
                    // 日期列：保持原始格式，不进行数字转换
                    if (dateColumns.has(header)) {{
                      // 如果是日期列，保持原始字符串格式
                      return String(value);
                    }}
                    
                    // 处理日期格式字符串（非日期列但可能是日期值）
                    if (typeof value === 'string') {{
                      // 检查是否是日期格式（ISO 格式或其他常见日期格式）
                      const datePatterns = [
                        /^\\d{4}-\\d{2}-\\d{2}/,  // ISO: 2024-01-01
                        /^\\d{4}\\/\\d{2}\\/\\d{2}/,  // 2024/01/01
                        /^\\d{4}\\.\\d{2}\\.\\d{2}/,  // 2024.01.01
                        /^\\d{4}-\\d{2}-\\d{2}T/,  // ISO with time
                      ];
                      
                      const isDateString = datePatterns.some(pattern => pattern.test(value));
                      if (isDateString) {{
                        // 保持日期字符串格式
                        return value;
                      }}
                      
                      // 检查是否是纯数字字符串（不是日期）
                      const numValue = parseFloat(value);
                      if (!isNaN(numValue) && isFinite(numValue) && value.trim() !== '' && 
                          !isDateString && !value.includes('-') && !value.includes('/')) {{
                        // 如果字符串是纯数字且不是日期，转换为数字类型
                        return numValue;
                      }}
                    }}
                    
                    // 处理数字：确保是有效的数字
                    if (typeof value === 'number') {{
                      if (isNaN(value) || !isFinite(value)) {{
                        return '';
                      }}
                      return value;
                    }}
                    
                    // 其他情况保持原值
                    return value;
                  }});
                  sheetData.push(row);
                }});
                
                // 创建工作表
                const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
                
                // 设置列宽（可选，但有助于 Excel 正确识别格式）
                const colWidths = headers.map(() => ({{ wch: 15 }}));
                worksheet['!cols'] = colWidths;
                
                // 生成sheet名称：根据category和顺序索引生成标识，然后加上title
                // 使用索引+1作为图表编号（从1开始）
                const chartNum = (idx + 1).toString();
                
                // 根据category生成标识前缀
                let prefix = '';
                if (category === '因子') {{
                  prefix = '因子图' + chartNum;
                }} else if (category === '资金') {{
                  prefix = '资金图' + chartNum;
                }} else if (category === '风险') {{
                  prefix = '风险图' + chartNum;
                }} else if (category === '量化资产配置') {{
                  prefix = '量化资产配置图' + chartNum;
                }}
                
                // 组合标识和标题
                let sheetName = prefix;
                if (cfg.title) {{
                  // 如果sheet名称已经包含标题，直接使用sheet名称
                  if (cfg.sheet && cfg.sheet.includes(cfg.title)) {{
                    sheetName = cfg.sheet;
                  }} else {{
                    sheetName = prefix + '-' + cfg.title;
                  }}
                }} else {{
                  // 如果没有title，使用sheet名称
                  sheetName = cfg.sheet || prefix;
                }}
                
                // 清理 sheet 名称中的非法字符
                sheetName = sanitizeSheetName(sheetName);
                
                // 处理重复的 sheet 名称
                let finalSheetName = sheetName;
                let counter = 1;
                while (usedSheetNames.has(finalSheetName)) {{
                  const suffix = counter.toString();
                  const maxLength = 31 - suffix.length - 1;
                  finalSheetName = sheetName.substring(0, maxLength) + '_' + suffix;
                  counter++;
                }}
                usedSheetNames.add(finalSheetName);
                
                // 将工作表添加到工作簿
                XLSX.utils.book_append_sheet(workbook, worksheet, finalSheetName);
              }}
            }}
          }});
        }});
        
        if (!hasData) {{
          alert('没有可下载的数据');
          console.warn('No data available for download');
          return;
        }}
        
        // 生成Excel文件并下载
        XLSX.writeFile(workbook, '所有图表数据.xlsx', {{
          bookType: 'xlsx'
        }});
        console.log('Excel file downloaded successfully');
      }} catch (error) {{
        console.error('Error in downloadAllData:', error);
        alert('下载失败：' + error.message);
      }}
    }}

    // 显示/隐藏分享对话框
    function showShareModal() {{
      // 找到分享按钮
      const shareBtn = document.getElementById('share-button');
      
      if (!shareBtn) {{
        console.error('分享按钮未找到');
        return;
      }}
      
      // 检查是否已存在分享对话框
      let modal = document.getElementById('share-modal');
      if (!modal) {{
        // 创建分享对话框
        modal = document.createElement('div');
        modal.id = 'share-modal';
        modal.className = 'share-modal';
        
        const content = document.createElement('div');
        content.className = 'share-modal-content';
        
        const header = document.createElement('div');
        header.className = 'share-modal-header';
        header.textContent = '复制并粘贴此URL以分享：';
        
        const urlDiv = document.createElement('div');
        urlDiv.className = 'share-modal-url';
        urlDiv.id = 'share-url';
        urlDiv.textContent = window.location.href;
        
        // 点击URL自动选中
        urlDiv.onclick = function(e) {{
          e.stopPropagation();
          const range = document.createRange();
          range.selectNodeContents(this);
          const selection = window.getSelection();
          selection.removeAllRanges();
          selection.addRange(range);
        }};
        
        content.appendChild(header);
        content.appendChild(urlDiv);
        modal.appendChild(content);
        
        document.body.appendChild(modal);
      }}
      
      // 如果弹窗已经显示，则关闭它
      if (modal.classList.contains('active')) {{
        modal.classList.remove('active');
        return;
      }}
      
      // 更新URL
      const urlDiv = document.getElementById('share-url');
      if (urlDiv) {{
        urlDiv.textContent = window.location.href;
      }}
      
      // 先显示对话框（隐藏状态），以便计算尺寸
      modal.style.visibility = 'hidden';
      modal.classList.add('active');
      
      // 使用setTimeout确保DOM已渲染
      setTimeout(() => {{
        // 计算按钮位置
        const btnRect = shareBtn.getBoundingClientRect();
        const modalContent = modal.querySelector('.share-modal-content');
        const contentHeight = modalContent ? modalContent.offsetHeight : 150;
        const contentWidth = modalContent ? modalContent.offsetWidth : 400;
        
        // 定位在按钮上方，居中对齐
        const left = btnRect.left + (btnRect.width / 2) - (contentWidth / 2);
        const top = btnRect.top - contentHeight - 15; // 按钮上方5px
        
        // 确保对话框不超出屏幕左边界，留出空隙
        const finalLeft = Math.max(15, left);
        
        // 如果上方空间不够，显示在按钮下方
        const finalTop = top < 10 ? btnRect.bottom + 10 : top;
        
        modal.style.left = finalLeft + 'px';
        modal.style.top = finalTop + 'px';
        modal.style.visibility = 'visible';
      }}, 0);
      
      // 点击其他地方关闭弹窗
      const closeModal = (e) => {{
        if (!modal.contains(e.target) && !shareBtn.contains(e.target)) {{
          modal.classList.remove('active');
          document.removeEventListener('click', closeModal);
        }}
      }};
      
      // 延迟添加事件监听，避免立即触发
      setTimeout(() => {{
        document.addEventListener('click', closeModal);
      }}, 100);
    }}

    function createChart(cfg, containerId) {{
      const records = dataBySheet[cfg.sheet] || [];
      
      let traces, layout;
      let hasY2 = false; // 默认值，避免作用域问题
      
      if (cfg.type === 'scatter') {{
        // 散点图模式
        const x = records.map(r => parseFloat(r[cfg.x]) || 0);
        const y = records.map(r => parseFloat(r[cfg.y]) || 0);
        const text = cfg.text ? records.map(r => r[cfg.text] || '') : [];
        
        // 创建外环和内圈，颜色相同，中间有间隙
        traces = [
          // 外环（空心圆）
          {{
            x: x,
            y: y,
            mode: 'markers',
            type: 'scatter',
            marker: {{
              size: 12,
              color: '#005bac',
              line: {{
                color: '#005bac',
                width: 1
              }},
              symbol: 'circle-open',
              opacity: 1
            }},
            name: 'PB-ROE',
            showlegend: false,
            hoverinfo: 'skip'
          }},
          // 内圈（实心圆）+ 文字标签
          {{
            x: x,
            y: y,
            mode: 'markers+text',
            type: 'scatter',
            text: text,
            textposition: 'top center',
            textfont: {{
              size: 10,
              color: '#2f3640'
            }},
            marker: {{
              size: 8,
              color: '#005bac',
              line: {{
                color: 'transparent',
                width: 0
              }},
              opacity: 1
            }},
            name: 'PB-ROE',
            showlegend: false,
            hovertemplate: '<b>%{{text}}</b><br>' +
                          '%{{xaxis.title.text}}: %{{x:,.4f}}, %{{yaxis.title.text}}: %{{y:,.4f}}<extra></extra>'
          }}
        ];
        
        layout = {{
          margin: {{ t: 20, r: 15, b: 60, l: 60 }},
          showlegend: false,
          dragmode: 'zoom',
          plot_bgcolor: 'white',
          paper_bgcolor: 'white',
          hoverlabel: {{
            bgcolor: 'white',
            bordercolor: '#d0d0d0',
            font: {{ color: '#666666' }}
          }},
          xaxis: {{
            title: cfg.x || '预测ROE',
            showgrid: true,
            gridcolor: '#ecf0f1',
            type: 'linear',
            showline: true,
            linecolor: '#ecf0f1',
            linewidth: 1,
            mirror: true,
            titlefont: {{ color: '#666666', size: 15 }},
            tickfont: {{ color: '#666666' }},
            zeroline: false,
            hoverformat: '.4f'
          }},
          yaxis: {{
            title: {{
              text: cfg.y || '预测PB',
              standoff: 30
            }},
            showgrid: true,
            gridcolor: '#ecf0f1',
            showline: true,
            linecolor: '#ecf0f1',
            linewidth: 1,
            mirror: true,
            titlefont: {{ color: '#666666', size: 15 }},
            tickfont: {{ color: '#666666' }},
            zeroline: false,
            hoverformat: '.4f'
          }},
          hovermode: 'closest',
        }};
      }} else {{
        // 折线图模式（原有逻辑）
        const x = records.map(r => r[cfg.x]);
        traces = cfg.lines.map(lineCfg => {{
          const y = records.map(r => r[lineCfg.field]);
          const axisName = lineCfg.axis === 'y2' ? 'y2' : 'y';
          const trace = {{
            x,
            y,
            mode: 'lines',
            name: lineCfg.name,
            yaxis: axisName,
            hovertemplate: '%{{fullData.name}}: %{{y:,.4f}}<extra></extra>'
          }};
          
          // 设置颜色：基准和右轴为#FFCB05，其他为#005bac
          // 特殊处理fund2图表的上下轨颜色
          if (cfg.id === 'fund2' && lineCfg.name === '上轨80分位数') {{
            trace.line = {{ color: '#FFABAB' }};
          }} else if (cfg.id === 'fund2' && lineCfg.name === '下轨20分位数') {{
            trace.line = {{ color: '#B6E880' }};
          }} else if (cfg.id === 'asset1' && lineCfg.name === '低波组合') {{
            trace.line = {{ color: '#B6E880' }};
          }} else if (cfg.id === 'asset1' && lineCfg.name === '中波组合') {{
            trace.line = {{ color: '#005bac' }};
          }} else if (cfg.id === 'asset1' && lineCfg.name === '基准') {{
            trace.line = {{ color: '#FFCB05' }};
          }} else if (lineCfg.name === '基准' || lineCfg.name === 'Wind全A收盘价（右轴）' || lineCfg.name === 'Wind全A收盘价' || lineCfg.axis === 'y2') {{
            trace.line = {{ color: '#FFCB05' }};
          }} else {{
            trace.line = {{ color: '#005bac' }};
          }}
          
          return trace;
        }});

        const hasY2 = cfg.lines.some(l => l.axis === 'y2');
        
        // 针对"外币资金占货币资金比"和"主要客户占比稳定性"图表的特殊配置（坐标轴标题）
        const isFactor1Chart = cfg.id === 'factor1';
        const isFactor2Chart = cfg.id === 'factor2';
        const isFactorChart = isFactor1Chart || isFactor2Chart;
        const isFund1Chart = cfg.id === 'fund1';
        const isFund2Chart = cfg.id === 'fund2';
        const isRisk1Chart = cfg.id === 'risk1';
        const isAsset1Chart = cfg.id === 'asset1';
        // 如果没有右轴，右侧 margin 可以更紧凑
        const rightMargin = hasY2 ? 70 : 15;

        layout = {{
          margin: {{ t: 30, r: rightMargin, b: 80, l: 70 }},
          showlegend: false,
          dragmode: 'zoom',
          plot_bgcolor: 'white',
          paper_bgcolor: 'white',
          hoverlabel: {{
            bgcolor: 'white',
            bordercolor: '#d0d0d0',
            font: {{ color: '#666666' }}
          }},
          xaxis: {{
            title: (isFactorChart || isFund1Chart || isFund2Chart || isRisk1Chart || isAsset1Chart) ? '日期' : cfg.x,
            showgrid: true,
            gridcolor: '#ecf0f1',
            type: 'date',
            showline: true,
            linecolor: '#ecf0f1',
            linewidth: 1,
            mirror: true,
            titlefont: {{ color: '#666666', size: 15 }},
            tickfont: {{ color: '#666666' }},
            zeroline: false,
            hoverformat: '%Y-%m-%d',
            rangeslider: {{
              visible: false
            }},
            rangeselector: {{
              buttons: [
                {{ count: 1, label: '1M', step: 'month', stepmode: 'backward' }},
                {{ count: 3, label: '3M', step: 'month', stepmode: 'backward' }},
                {{ count: 6, label: '6M', step: 'month', stepmode: 'backward' }},
                {{ count: 1, label: '1Y', step: 'year', stepmode: 'backward' }},
                {{ step: 'all', label: '全部' }}
              ],
              font: {{ color: '#666666' }}
            }},
          }},
          yaxis: {{
            title: {{
              text: isFactorChart ? '净值' : (isFund1Chart ? '主力累计净买入(亿元)' : (isFund2Chart ? '博弈/存量' : (isRisk1Chart ? '新高个股占比' : (isAsset1Chart ? '净值' : '')))),
              standoff: 30
            }},
            showgrid: true,
            gridcolor: '#ecf0f1',
            showline: true,
            linecolor: '#ecf0f1',
            linewidth: 1,
            mirror: true,
            titlefont: {{ color: '#666666', size: 15 }},
            tickfont: {{ color: '#666666' }},
            zeroline: false,
            hoverformat: '.4f'
          }},
          hovermode: 'x unified',
        }};

        if (hasY2) {{
          const y2TitleText = isFund1Chart ? 'Wind全A收盘价' : (isFund2Chart ? 'Wind全A收盘价' : (isRisk1Chart ? '上证综指' : ''));
          layout.yaxis2 = {{
            overlaying: 'y',
            side: 'right',
            showgrid: false,
            showline: true,
            linecolor: '#ecf0f1',
            linewidth: 1,
            mirror: true,
            title: y2TitleText ? {{ text: y2TitleText, standoff: 15 }} : undefined,
            titlefont: {{ color: '#666666', size: 15 }},
            tickfont: {{ color: '#666666' }},
            zeroline: false,
            hoverformat: '.4f'
          }};
        }}
      }}

      Plotly.newPlot(containerId, traces, layout, {{
        responsive: true,
        displaylogo: false,
        modeBarButtonsToRemove: ['toImage']
      }}).then(() => {{
        // 标记该图表已创建
        chartInstances[cfg.id] = true;

        // 对于有双坐标轴的图表，手动触发resize以确保正确布局
        if (hasY2) {{
          setTimeout(() => {{
            Plotly.Plots.resize(containerId);
          }}, 100);
        }}
        
        // 隐藏rangeslider中的折线图，只显示灰色控制条
        if (cfg.type !== 'scatter') {{
          setTimeout(() => {{
            const plotDiv = document.getElementById(containerId);
            if (plotDiv) {{
              // 查找rangeslider容器
              const rangesliderContainer = plotDiv.querySelector('.rangeslider-container');
              if (rangesliderContainer) {{
                // 隐藏所有svg图表
                const svgs = rangesliderContainer.querySelectorAll('svg');
                svgs.forEach(svg => svg.style.display = 'none');
                // 设置容器高度为20px（所有图表统一）
                rangesliderContainer.style.height = '20px';
                rangesliderContainer.style.minHeight = '20px';
              }}
              // 再次触发resize以确保布局正确
              if (hasY2) {{
                Plotly.Plots.resize(containerId);
              }}
            }}
          }}, 300);
        }}

        // 对于有基准线的折线图（因子图和量化资产配置图），添加时间范围选择时的归一化处理
        const hasBenchmark = cfg.lines && cfg.lines.some(l => l.name === '基准');
        const isBenchmarkChart = (cfg.id === 'factor1' || cfg.id === 'factor2' || cfg.id === 'asset1');
        if (hasBenchmark && isBenchmarkChart && cfg.type !== 'scatter') {{
          // 延迟绑定事件，确保图表完全加载
          setTimeout(() => {{
            const plotDiv = document.getElementById(containerId);
            if (!plotDiv) return;

            // 保存原始数据
            const originalData = {{
              x: records.map(r => r[cfg.x]),
              traces: cfg.lines.map(lineCfg => ({{
                name: lineCfg.name,
                field: lineCfg.field,
                y: records.map(r => {{
                  const val = r[lineCfg.field];
                  return val !== null && val !== undefined ? parseFloat(val) : null;
                }})
              }}))
            }};

            // 标记是否已归一化
            let isNormalized = false;

            // 归一化函数
            function normalizeData(rangeStart, rangeEnd) {{
              // 找到起始点索引：优先找范围内的第一个点，否则找最接近的点
              let startIndex = -1;
              let minDiff = Infinity;
              let closestIndex = -1;
              
              for (let i = 0; i < originalData.x.length; i++) {{
                const dateStr = originalData.x[i];
                let date = null;
                
                // 尝试解析日期
                if (typeof dateStr === 'string') {{
                  date = new Date(dateStr);
                }} else if (dateStr instanceof Date) {{
                  date = dateStr;
                }}
                
                if (date && !isNaN(date.getTime())) {{
                  // 优先找范围内的第一个点
                  if (date >= rangeStart && date <= rangeEnd) {{
                    startIndex = i;
                    break;
                  }}
                  
                  // 同时记录最接近的点
                  const diff = Math.abs(date.getTime() - rangeStart.getTime());
                  if (diff < minDiff) {{
                    minDiff = diff;
                    closestIndex = i;
                  }}
                }}
              }}

              // 如果没找到范围内的点，使用最接近的点
              if (startIndex < 0 && closestIndex >= 0) {{
                startIndex = closestIndex;
              }}

              // 如果找到了起始点，进行归一化
              if (startIndex >= 0) {{
                // 准备所有更新数据，使用新的归一化逻辑
                const yArrays = [];
                const indices = [];
                cfg.lines.forEach((lineCfg, idx) => {{
                  const originalTrace = originalData.traces.find(t => t.name === lineCfg.name);
                  if (originalTrace) {{
                    // 初始化数组，复制原始值
                    const normalizedY = originalTrace.y.slice();
                    
                    // 从起始点开始，逐日计算归一化值
                    // 起始点：所有净值都取1
                    // 后续每天：今日归一化值 = 昨日归一化值 * (今日原始值 / 昨日原始值)
                    // 起始点之前：反向计算，今日归一化值 = 明日归一化值 * (今日原始值 / 明日原始值)
                    
                    // 先处理起始点及之后的数据（正向）
                    for (let i = startIndex; i < originalTrace.y.length; i++) {{
                      const currentValue = originalTrace.y[i];
                      
                      if (currentValue === null || currentValue === undefined || isNaN(currentValue)) {{
                        continue;
                      }}
                      
                      if (i === startIndex) {{
                        // 起始点：设为1
                        normalizedY[i] = 1;
                      }} else {{
                        // 起始点之后的数据：按涨跌幅累乘
                        const prevValue = originalTrace.y[i - 1];
                        if (prevValue !== null && prevValue !== undefined && !isNaN(prevValue) && prevValue !== 0) {{
                          // 今日归一化值 = 昨日归一化值 * (今日原始值 / 昨日原始值)
                          const prevNormalized = normalizedY[i - 1];
                          if (prevNormalized !== null && prevNormalized !== undefined && !isNaN(prevNormalized)) {{
                            normalizedY[i] = prevNormalized * (currentValue / prevValue);
                          }}
                        }}
                      }}
                    }}
                    
                    // 再处理起始点之前的数据（反向）
                    for (let i = startIndex - 1; i >= 0; i--) {{
                      const currentValue = originalTrace.y[i];
                      
                      if (currentValue === null || currentValue === undefined || isNaN(currentValue)) {{
                        continue;
                      }}
                      
                      // 起始点之前的数据：反向计算
                      const nextValue = originalTrace.y[i + 1];
                      if (nextValue !== null && nextValue !== undefined && !isNaN(nextValue) && nextValue !== 0) {{
                        // 今日归一化值 = 明日归一化值 * (今日原始值 / 明日原始值)
                        const nextNormalized = normalizedY[i + 1];
                        if (nextNormalized !== null && nextNormalized !== undefined && !isNaN(nextNormalized)) {{
                          normalizedY[i] = nextNormalized * (currentValue / nextValue);
                        }}
                      }}
                    }}
                    
                    yArrays.push(normalizedY);
                    indices.push(idx);
                  }}
                }});
                
                // 批量更新所有 traces，使用正确的格式
                if (yArrays.length > 0) {{
                  Plotly.restyle(containerId, {{ y: yArrays }}, indices);
                  isNormalized = true;
                }}
              }}
            }}

            // 恢复原始数据函数
            function restoreData() {{
              if (isNormalized) {{
                // 准备所有恢复数据
                const yArrays = [];
                const indices = [];
                cfg.lines.forEach((lineCfg, idx) => {{
                  const originalTrace = originalData.traces.find(t => t.name === lineCfg.name);
                  if (originalTrace) {{
                    yArrays.push(originalTrace.y);
                    indices.push(idx);
                  }}
                }});
                
                // 批量恢复所有 traces
                if (yArrays.length > 0) {{
                  Plotly.restyle(containerId, {{ y: yArrays }}, indices);
                  isNormalized = false;
                }}
              }}
            }}

            // 保存原始数据的完整范围
            const originalXMin = originalData.x.length > 0 ? new Date(originalData.x[0]) : null;
            const originalXMax = originalData.x.length > 0 ? new Date(originalData.x[originalData.x.length - 1]) : null;

            // 监听布局变化事件
            plotDiv.on('plotly_relayout', function(eventData) {{
              // 延迟处理，确保范围已更新
              setTimeout(() => {{
                try {{
                  // 检查是否是恢复到全部范围（点击"全部"按钮）
                  if (eventData['xaxis.autorange'] === true) {{
                    restoreData();
                    return;
                  }}

                  // 检查当前布局中的范围
                  const layout = plotDiv.layout;
                  if (!layout || !layout.xaxis) return;

                  let rangeStart = null;
                  let rangeEnd = null;
                  
                  // 优先从事件数据中获取范围
                  if (eventData['xaxis.range[0]'] !== undefined && eventData['xaxis.range[1]'] !== undefined) {{
                    rangeStart = new Date(eventData['xaxis.range[0]']);
                    rangeEnd = new Date(eventData['xaxis.range[1]']);
                  }} else if (eventData['xaxis.range'] && Array.isArray(eventData['xaxis.range'])) {{
                    rangeStart = new Date(eventData['xaxis.range'][0]);
                    rangeEnd = new Date(eventData['xaxis.range'][1]);
                  }} else if (layout.xaxis.range && Array.isArray(layout.xaxis.range)) {{
                    // 从布局中获取当前范围
                    rangeStart = new Date(layout.xaxis.range[0]);
                    rangeEnd = new Date(layout.xaxis.range[1]);
                  }} else {{
                    return; // 无法获取范围，退出
                  }}

                  // 检查是否是全部范围（通过比较与原始数据的范围）
                  if (originalXMin && originalXMax) {{
                    const timeDiff = 86400000; // 1天的毫秒数，允许误差
                    const isFullRange = Math.abs(rangeStart.getTime() - originalXMin.getTime()) < timeDiff && 
                                       Math.abs(rangeEnd.getTime() - originalXMax.getTime()) < timeDiff;
                    
                    if (isFullRange) {{
                      restoreData();
                      return;
                    }}
                  }}

                  // 如果有有效的时间范围，进行归一化
                  if (rangeStart && rangeEnd && !isNaN(rangeStart.getTime()) && !isNaN(rangeEnd.getTime())) {{
                    normalizeData(rangeStart, rangeEnd);
                  }}
                }} catch (e) {{
                  console.error('Error in normalize handler:', e);
                }}
              }}, 300);
            }});

            // 额外监听：直接监听 rangeselector 按钮点击（备用方案）
            setTimeout(() => {{
              const plotDiv = document.getElementById(containerId);
              if (!plotDiv) return;
              
              // 查找 rangeselector 按钮
              const rangeSelector = plotDiv.querySelector('.rangeselector');
              if (rangeSelector) {{
                const buttons = rangeSelector.querySelectorAll('button');
                buttons.forEach(button => {{
                  button.addEventListener('click', function() {{
                    setTimeout(() => {{
                      const layout = plotDiv.layout;
                      if (layout && layout.xaxis && layout.xaxis.range && Array.isArray(layout.xaxis.range)) {{
                        const rangeStart = new Date(layout.xaxis.range[0]);
                        const rangeEnd = new Date(layout.xaxis.range[1]);
                        
                        // 检查是否是全部范围
                        if (originalXMin && originalXMax) {{
                          const timeDiff = 86400000;
                          const isFullRange = Math.abs(rangeStart.getTime() - originalXMin.getTime()) < timeDiff && 
                                             Math.abs(rangeEnd.getTime() - originalXMax.getTime()) < timeDiff;
                          
                          if (isFullRange) {{
                            restoreData();
                          }} else if (!isNaN(rangeStart.getTime()) && !isNaN(rangeEnd.getTime())) {{
                            normalizeData(rangeStart, rangeEnd);
                          }}
                        }}
                      }}
                    }}, 400);
                  }});
                }});
              }}
            }}, 1000);
          }}, 500);
        }}
      }});
    }}

    function buildSections() {{
      const container = document.getElementById('sections-container');
      container.innerHTML = '';

      let firstCategory = null;

      Object.keys(chartConfigs).forEach((category, catIdx) => {{
        const section = document.createElement('div');
        section.className = 'section';
        section.id = `section-${{category}}`;
        
        // 第一个板块默认显示
        if (catIdx === 0) {{
          section.classList.add('active');
          firstCategory = category;
        }}

        const sectionTitle = document.createElement('div');
        sectionTitle.className = 'section-title';
        sectionTitle.textContent = category;
        section.appendChild(sectionTitle);

        const chartsContainer = document.createElement('div');
        chartsContainer.className = 'section-charts';

        chartConfigs[category].forEach((cfg, idx) => {{
          const chartWrapper = document.createElement('div');
          chartWrapper.className = 'chart-wrapper';
          chartWrapper.id = `chart-${{cfg.id}}`;

          const chartHeader = document.createElement('div');
          chartHeader.className = 'chart-header';

          const chartTitle = document.createElement('div');
          chartTitle.className = 'chart-title';
          chartTitle.textContent = cfg.displayTitle || cfg.title;
          chartHeader.appendChild(chartTitle);

          // 创建右侧操作区域容器
          const chartActions = document.createElement('div');
          chartActions.style.display = 'flex';
          chartActions.style.alignItems = 'center';
          chartActions.style.gap = '16px';

          // 添加数据说明图标和文字
          if (cfg.description) {{
            const infoIcon = document.createElement('div');
            infoIcon.className = 'chart-info';
            infoIcon.onclick = () => showChartInfo(cfg);
            infoIcon.title = '说明';
            
            const infoIconSvg = document.createElement('div');
            infoIconSvg.className = 'chart-info-icon';
            infoIconSvg.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>';
            infoIcon.appendChild(infoIconSvg);
            
            const infoText = document.createElement('span');
            infoText.textContent = '说明';
            infoIcon.appendChild(infoText);
            
            chartActions.appendChild(infoIcon);
          }}

          // 添加下载图标和文字
          const downloadIcon = document.createElement('div');
          downloadIcon.className = 'chart-download';
          downloadIcon.onclick = () => downloadChartData(cfg);
          downloadIcon.title = '下载数据';
          
          const iconSvg = document.createElement('div');
          iconSvg.className = 'chart-download-icon';
          iconSvg.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>';
          downloadIcon.appendChild(iconSvg);
          
          const downloadText = document.createElement('span');
          downloadText.textContent = '下载';
          downloadIcon.appendChild(downloadText);
          
          chartActions.appendChild(downloadIcon);
          chartHeader.appendChild(chartActions);

          chartWrapper.appendChild(chartHeader);

          // 添加数据说明面板
          if (cfg.description) {{
            const infoPanel = document.createElement('div');
            infoPanel.id = `chart-info-panel-${{cfg.id}}`;
            infoPanel.className = 'chart-info-panel';
            infoPanel.innerHTML = parseDescription(cfg, cfg.description);
            chartWrapper.appendChild(infoPanel);
          }}

          const chartContainer = document.createElement('div');
          chartContainer.className = 'chart-container';
          const chartDiv = document.createElement('div');
          chartDiv.className = 'chart';
          chartDiv.id = `plot-${{cfg.id}}`;
          chartContainer.appendChild(chartDiv);
          chartWrapper.appendChild(chartContainer);

          chartsContainer.appendChild(chartWrapper);
        }});

        section.appendChild(chartsContainer);
        container.appendChild(section);
      }});

      // 仅为第一个板块创建图表，其余板块在首次点击导航时再创建，避免在隐藏状态下绘制导致初始尺寸过小
      if (firstCategory && chartConfigs[firstCategory].length > 0) {{
        chartConfigs[firstCategory].forEach(cfg => {{
          createChart(cfg, `plot-${{cfg.id}}`);
        }});
      }}
      
      // 默认激活第一个图表的导航项
      if (firstCategory && chartConfigs[firstCategory].length > 0) {{
        const firstChartId = chartConfigs[firstCategory][0].id;
        const firstNavItem = document.querySelector(`[data-chart-id="${{firstChartId}}"]`);
        if (firstNavItem) {{
          firstNavItem.classList.add('active');
        }}
      }}
    }}

    function buildNav() {{
      const navRoot = document.getElementById('nav');
      navRoot.innerHTML = '';

      // 添加"浏览"标题
      const browseTitle = document.createElement('div');
      browseTitle.className = 'nav-browse-title';
      browseTitle.textContent = '浏览';
      navRoot.appendChild(browseTitle);

      Object.keys(chartConfigs).forEach(category => {{
        const sectionTitle = document.createElement('div');
        sectionTitle.className = 'nav-section-title';
        sectionTitle.textContent = category;
        navRoot.appendChild(sectionTitle);

        const ul = document.createElement('ul');
        ul.className = 'nav-list';

        chartConfigs[category].forEach((cfg, idx) => {{
          const li = document.createElement('li');
          li.className = 'nav-item';
          li.dataset.category = category;
          li.dataset.index = idx;
          li.dataset.chartId = cfg.id;
          li.onclick = () => scrollToChart(category, cfg.id);
          
          // 添加图标
          const icon = document.createElement('div');
          icon.className = 'nav-icon';
          if (cfg.type === 'scatter') {{
            // 散点图图标 - 分散排列，不是斜线
            icon.innerHTML = '<svg viewBox="0 0 16 16" fill="currentColor"><circle cx="3" cy="4" r="1.5"/><circle cx="10" cy="8" r="1.5"/><circle cx="6" cy="13" r="1.5"/></svg>';
          }} else {{
            // 折线图图标
            icon.innerHTML = '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="2,12 5,8 8,10 12,4 14,6"/></svg>';
          }}
          li.appendChild(icon);
          
          // 添加文字
          const text = document.createElement('span');
          text.textContent = cfg.title;
          li.appendChild(text);
          
          ul.appendChild(li);
        }});

        navRoot.appendChild(ul);
      }});

      // 添加底部操作区域
      const footer = document.createElement('div');
      footer.className = 'sidebar-footer';

      // 下载按钮
      const downloadBtn = document.createElement('div');
      downloadBtn.className = 'sidebar-action';
      downloadBtn.onclick = function(e) {{
        e.preventDefault();
        e.stopPropagation();
        downloadAllData();
      }};
      const downloadIcon = document.createElement('div');
      downloadIcon.className = 'sidebar-action-icon';
      downloadIcon.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>';
      downloadIcon.style.pointerEvents = 'none';
      downloadBtn.appendChild(downloadIcon);
      const downloadText = document.createElement('span');
      downloadText.textContent = '下载';
      downloadText.style.pointerEvents = 'none';
      downloadBtn.appendChild(downloadText);
      footer.appendChild(downloadBtn);

      // 分享按钮
      const shareBtn = document.createElement('div');
      shareBtn.className = 'sidebar-action';
      shareBtn.id = 'share-button';
      shareBtn.onclick = showShareModal;
      const shareIcon = document.createElement('div');
      shareIcon.className = 'sidebar-action-icon';
      shareIcon.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="18" cy="5" r="3"></circle><circle cx="6" cy="12" r="3"></circle><circle cx="18" cy="19" r="3"></circle><line x1="8.59" y1="13.51" x2="15.42" y2="17.49"></line><line x1="15.41" y1="6.51" x2="8.59" y2="10.49"></line></svg>';
      shareBtn.appendChild(shareIcon);
      const shareText = document.createElement('span');
      shareText.textContent = '分享';
      shareBtn.appendChild(shareText);
      footer.appendChild(shareBtn);

      // 将footer添加到sidebar，而不是navRoot
      const sidebar = document.querySelector('.sidebar');
      // 如果footer已存在，先移除
      const existingFooter = sidebar.querySelector('.sidebar-footer');
      if (existingFooter) {{
        existingFooter.remove();
      }}
      sidebar.appendChild(footer);
    }}

    function scrollToChart(category, chartId) {{
      // 激活状态样式
      document.querySelectorAll('.nav-item').forEach(el => {{
        el.classList.remove('active');
      }});
      document.querySelectorAll('.nav-item').forEach(el => {{
        if (el.dataset.chartId === chartId) {{
          el.classList.add('active');
        }}
      }});

      // 隐藏所有板块，只显示当前板块
      document.querySelectorAll('.section').forEach(section => {{
        section.classList.remove('active');
      }});
      
      const targetSection = document.getElementById(`section-${{category}}`);
      if (targetSection) {{
        targetSection.classList.add('active');

        // 当前板块的所有图表在可见状态下一次性创建，避免逐个点击才渲染
        if (chartConfigs[category]) {{
          chartConfigs[category].forEach(cfg => {{
            if (!chartInstances[cfg.id]) {{
              createChart(cfg, `plot-${{cfg.id}}`);
            }}
          }});
        }}
        
        // 找到目标图表元素并滚动到它
        setTimeout(() => {{
          const targetChart = document.getElementById(`chart-${{chartId}}`);
          const sectionCharts = targetSection.querySelector('.section-charts');
          
          if (targetChart && sectionCharts) {{
            const containerRect = sectionCharts.getBoundingClientRect();
            const chartRect = targetChart.getBoundingClientRect();
            const scrollTop = sectionCharts.scrollTop;
            const targetScrollTop = scrollTop + chartRect.top - containerRect.top - 20;
            
            sectionCharts.scrollTo({{
              top: targetScrollTop,
              behavior: 'smooth'
            }});
            
            // 触发该板块内所有图表的resize，确保所有图表都能正确调整大小
            if (chartConfigs[category]) {{
              chartConfigs[category].forEach(cfg => {{
                const chartDiv = document.getElementById(`plot-${{cfg.id}}`);
                if (chartDiv && chartInstances[cfg.id]) {{
                  setTimeout(() => {{
                    Plotly.Plots.resize(`plot-${{cfg.id}}`);
                  }}, 100);
                }}
              }});
            }}
          }}
        }}, 50);
      }}
    }}

    window.addEventListener('resize', () => {{
      Object.keys(chartConfigs).forEach(category => {{
        chartConfigs[category].forEach(cfg => {{
          const chartDiv = document.getElementById(`plot-${{cfg.id}}`);
          if (chartDiv && chartInstances[cfg.id]) {{
            Plotly.Plots.resize(`plot-${{cfg.id}}`);
          }}
        }});
      }});
    }});

    document.addEventListener('DOMContentLoaded', () => {{
      buildNav();
      buildSections();
      
      // 页面加载完成后，对所有图表进行一次resize，确保双坐标轴图表正确布局
      setTimeout(() => {{
        Object.keys(chartConfigs).forEach(category => {{
          chartConfigs[category].forEach(cfg => {{
            const chartDiv = document.getElementById(`plot-${{cfg.id}}`);
            if (chartDiv && chartInstances[cfg.id]) {{
              Plotly.Plots.resize(`plot-${{cfg.id}}`);
            }}
          }});
        }});
      }}, 500);
    }});
  </script>
</body>
</html>
"""
    return html


def main():
    data_by_sheet, chart_configs = build_data_and_config()
    html = build_html(data_by_sheet, chart_configs)

    # 同时生成 dashboard.html 和 index.html，内容完全一致
    DASHBOARD_HTML.write_text(html, encoding="utf-8")
    INDEX_HTML.write_text(html, encoding="utf-8")
    print(f"已生成网页文件: {DASHBOARD_HTML}")
    print(f"已生成首页文件: {INDEX_HTML}")


if __name__ == "__main__":
    main()


