import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { pathToFileURL } from "node:url";

const ARTIFACT_TOOL_MODULE =
  process.env.OAI_ARTIFACT_TOOL_MODULE ||
  "/Users/cp/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const { FileBlob, SpreadsheetFile, Workbook } = await import(pathToFileURL(ARTIFACT_TOOL_MODULE).href);

const REQUIRED_FIELDS = {
  materialName: ["素材名称", "素材视频名称"],
  materialId: ["素材ID"],
  createdAt: ["素材创建时间"],
  spend: ["整体消耗"],
  roi: ["整体支付ROI"],
};

const OPTIONAL_FIELDS = {
  platformEval: ["素材评估"],
  duration: ["素材时长"],
  source: ["素材来源"],
  rawTags: ["标签"],
  gmv: ["整体成交金额"],
  ctr: ["整体点击率"],
  play3s: ["3秒播放率"],
  completion: ["视频完播率"],
  cvr: ["整体转化率"],
  impressions: ["整体展示次数"],
  clicks: ["整体点击次数"],
  planCount: ["素材关联的计划数量"],
  productCount: ["素材关联的商品数量"],
};

const SCRIPT_STYLES = ["价格福利型", "痛点科普型", "种草体验型", "人群点名型", "场景代入型", "测评对比型", "待补充"];
const HOOK_TYPES = ["惊讶反差", "价格锚点", "痛点直击", "人群点名", "信任背书", "场景问题", "待补充"];

function argValue(flag) {
  const i = process.argv.indexOf(flag);
  return i >= 0 ? process.argv[i + 1] : null;
}

async function latestDownload() {
  const downloads = path.join(os.homedir(), "Downloads");
  const files = await fs.readdir(downloads);
  const candidates = [];
  for (const file of files) {
    if (!/^全域数据_素材分析_视频.*\.xlsx$/.test(file)) continue;
    const full = path.join(downloads, file);
    const stat = await fs.stat(full);
    candidates.push({ full, mtimeMs: stat.mtimeMs });
  }
  candidates.sort((a, b) => b.mtimeMs - a.mtimeMs);
  if (!candidates.length) {
    throw new Error("未在 ~/Downloads 找到「全域数据_素材分析_视频*.xlsx」。请传入文件路径。");
  }
  return candidates[0].full;
}

function parseNumber(value) {
  if (value === null || value === undefined || value === "-" || value === "") return 0;
  if (typeof value === "number") return value;
  const cleaned = String(value).replace(/,/g, "").replace(/%/g, "").trim();
  const num = Number(cleaned);
  return Number.isFinite(num) ? num : 0;
}

function parsePercent(value) {
  if (value === null || value === undefined || value === "-" || value === "") return 0;
  if (typeof value === "number") return value > 1 ? value / 100 : value;
  const text = String(value).trim();
  const num = parseNumber(text);
  return text.includes("%") || num > 1 ? num / 100 : num;
}

function csvEscape(value) {
  if (value === null || value === undefined) return "";
  const text = String(value);
  if (/[",\n\r]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
  return text;
}

function toCsv(rows) {
  return "\uFEFF" + rows.map((row) => row.map(csvEscape).join(",")).join("\n");
}

function styleRange(range, opts = {}) {
  if (opts.fill) range.format.fill = opts.fill;
  if (opts.fontColor || opts.bold || opts.size) {
    range.format.font = {
      color: opts.fontColor,
      bold: opts.bold,
      size: opts.size,
    };
  }
  if (opts.border) range.format.borders = opts.border;
  if (opts.wrap !== undefined) range.format.wrapText = opts.wrap;
  if (opts.hAlign) range.format.horizontalAlignment = opts.hAlign;
  if (opts.vAlign) range.format.verticalAlignment = opts.vAlign;
}

function inferScriptStyle(name, rawTags) {
  const text = `${name || ""} ${rawTags || ""}`;
  if (/618|大促|活动|机制|福利|薅|便宜|价|送|囤|到手|平价/.test(text)) return "价格福利型";
  if (/牙龈|口腔|专业|权威|护理|修复|萎缩|抗敏|防蛀/.test(text)) return "痛点科普型";
  if (/闺蜜|朋友|亲测|自用|分享|没骗|推荐|好物/.test(text)) return "种草体验型";
  if (/请大数据|每一个|女朋友|人群|所有/.test(text)) return "人群点名型";
  if (/早上|晚上|上班|出门|吃饭|刷牙|场景/.test(text)) return "场景代入型";
  if (/测评|对比|升级|旧款|新版/.test(text)) return "测评对比型";
  return "待补充";
}

function inferHookType(name) {
  const text = name || "";
  if (/天呐|OMG|竟然|简直|没见过|背刺|离谱|宝藏/.test(text)) return "惊讶反差";
  if (/便宜|价|福利|送|薅|才|活动|平价|机制/.test(text)) return "价格锚点";
  if (/牙龈|口腔|护理|修复|萎缩|抗敏|防蛀/.test(text)) return "痛点直击";
  if (/请大数据|每一个|女朋友|买|所有/.test(text)) return "人群点名";
  if (/闺蜜|朋友|没骗|推荐|权威/.test(text)) return "信任背书";
  if (/为什么|怎么|原来|别再|注意/.test(text)) return "场景问题";
  return "待补充";
}

function recommendation(record) {
  const { spend, roi, ctr, play3s, cvr } = record;
  if (spend >= 300 && roi >= 1.8 && ctr >= 0.06 && cvr >= 0.055) return "加预算放量，拆解开头与利益点复刻";
  if (roi >= 1.8 && spend < 200) return "小预算加测，优先复制同款脚本结构";
  if (spend >= 300 && roi < 1.3) return "控预算，重剪前3秒或替换商品利益点";
  if (ctr < 0.04) return "优化封面/标题/前3秒钩子";
  if (play3s < 0.25) return "开头信息密度不足，前3秒重剪";
  if (cvr < 0.04) return "保留流量钩子，强化成交理由和价格机制";
  return "继续观察，补齐脚本标签后复盘";
}

function timestamp() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

function colName(n) {
  let name = "";
  let value = n;
  while (value > 0) {
    const rem = (value - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

function validateHeaders(headers) {
  const missing = Object.values(REQUIRED_FIELDS).filter((aliases) => !aliases.some((h) => headers.includes(h)));
  if (missing.length) throw new Error(`源表缺少必要表头：${missing.map((aliases) => aliases.join("/")).join("、")}`);
}

function findHeaderIndex(headers, aliases) {
  for (const alias of aliases) {
    const i = headers.indexOf(alias);
    if (i >= 0) return i;
  }
  return -1;
}

function buildRecords(sourceValues) {
  const headers = sourceValues[0];
  validateHeaders(headers);
  const required = Object.fromEntries(Object.entries(REQUIRED_FIELDS).map(([key, aliases]) => [key, findHeaderIndex(headers, aliases)]));
  const optional = Object.fromEntries(Object.entries(OPTIONAL_FIELDS).map(([key, aliases]) => [key, findHeaderIndex(headers, aliases)]));
  const optionalValue = (row, key, fallback = "") => {
    const i = optional[key];
    return i >= 0 ? row[i] : fallback;
  };
  const optionalNumber = (row, key) => {
    const i = optional[key];
    return i >= 0 ? parseNumber(row[i]) : 0;
  };
  const optionalPercent = (row, key) => {
    const i = optional[key];
    return i >= 0 ? parsePercent(row[i]) : 0;
  };
  return sourceValues.slice(1).filter((r) => r[required.materialId]).map((r) => {
    const spend = parseNumber(r[required.spend]);
    const roi = parseNumber(r[required.roi]);
    const record = {
      materialName: r[required.materialName],
      materialId: String(r[required.materialId]),
      platformEval: optionalValue(r, "platformEval"),
      duration: optionalValue(r, "duration"),
      createdAt: r[required.createdAt],
      source: optionalValue(r, "source"),
      rawTags: optionalValue(r, "rawTags") === "-" ? "" : optionalValue(r, "rawTags"),
      spend,
      gmv: optionalNumber(r, "gmv") || spend * roi,
      roi,
      ctr: optionalPercent(r, "ctr"),
      play3s: optionalPercent(r, "play3s"),
      completion: optionalPercent(r, "completion"),
      cvr: optionalPercent(r, "cvr"),
      impressions: optionalNumber(r, "impressions"),
      clicks: optionalNumber(r, "clicks"),
      planCount: optionalNumber(r, "planCount"),
      productCount: optionalNumber(r, "productCount"),
    };
    record.scriptStyle = inferScriptStyle(record.materialName, record.rawTags);
    record.hookType = inferHookType(record.materialName);
    record.recommendation = recommendation(record);
    return record;
  });
}

function buildWorkbook({ sourcePath, sourceValues, records }) {
  const workbook = Workbook.create();
  const main = workbook.worksheets.add("飞书导入主表");
  const dashboard = workbook.worksheets.add("素材看板");
  const chartData = workbook.worksheets.add("图表数据");
  const dict = workbook.worksheets.add("字段字典");
  const views = workbook.worksheets.add("飞书视图建议");
  const raw = workbook.worksheets.add("原始数据");

  for (const ws of [main, dashboard, chartData, dict, views, raw]) ws.showGridLines = false;

  const mainHeaders = [
    "素材ID",
    "素材名称",
    "平台素材评估",
    "素材时长",
    "素材创建时间",
    "素材来源",
    "原始标签",
    "消耗金额",
    "GMV",
    "ROI",
    "CTR",
    "3s播放率",
    "完播率",
    "CVR",
    "曝光量",
    "点击量",
    "关联计划数",
    "关联商品数",
    "素材值",
    "素材评级",
    "跑量类型",
    "脚本风格",
    "开头钩子类型",
    "优化方向",
    "负责人",
    "复盘状态",
    "备注",
  ];

  const lastRow = records.length + 1;
  main.getRange("A1:AA1").values = [mainHeaders];
  main.getRange(`A2:A${lastRow}`).setNumberFormat("@");
  main.getRange(`A2:R${lastRow}`).values = records.map((r) => [
    r.materialId,
    r.materialName,
    r.platformEval,
    r.duration,
    r.createdAt,
    r.source,
    r.rawTags,
    r.spend,
    r.gmv,
    r.roi,
    r.ctr,
    r.play3s,
    r.completion,
    r.cvr,
    r.impressions,
    r.clicks,
    r.planCount,
    r.productCount,
  ]);
  main.getRange(`V2:X${lastRow}`).values = records.map((r) => [r.scriptStyle, r.hookType, r.recommendation]);
  main.getRange(`Y2:AA${lastRow}`).values = records.map(() => ["", "待复盘", ""]);
  main.getRange("S2").formulas = [["=IFERROR(H2*J2*L2*K2*N2,0)"]];
  main.getRange(`S2:S${lastRow}`).fillDown();
  main.getRange("T2").formulas = [[`=IF(S2<=0,"D",IF(RANK.EQ(S2,$S$2:$S$${lastRow},0)<=ROUNDUP(COUNTIF($S$2:$S$${lastRow},">0")*0.15,0),"A",IF(RANK.EQ(S2,$S$2:$S$${lastRow},0)<=ROUNDUP(COUNTIF($S$2:$S$${lastRow},">0")*0.4,0),"B",IF(RANK.EQ(S2,$S$2:$S$${lastRow},0)<=ROUNDUP(COUNTIF($S$2:$S$${lastRow},">0")*0.7,0),"C","D"))))`]];
  main.getRange(`T2:T${lastRow}`).fillDown();
  main.getRange("U2").formulas = [[`=IF(T2="A","头部跑量",IF(AND(T2="B",H2>=PERCENTILE.INC($H$2:$H$${lastRow},0.6)),"稳定放量",IF(AND(T2="B",H2<PERCENTILE.INC($H$2:$H$${lastRow},0.6)),"高潜待放量",IF(AND(T2<>"A",H2>=PERCENTILE.INC($H$2:$H$${lastRow},0.75)),"低效烧钱","待观察"))))`]];
  main.getRange(`U2:U${lastRow}`).fillDown();
  main.tables.add(`A1:AA${lastRow}`, true).name = "tblFeishuMaterials";
  main.freezePanes.freezeRows(1);
  main.freezePanes.freezeColumns(2);
  styleRange(main.getRange("A1:AA1"), { fill: "#23395B", fontColor: "#FFFFFF", bold: true, hAlign: "center", vAlign: "center", wrap: true });
  styleRange(main.getRange(`A1:AA${lastRow}`), { border: { preset: "inside", style: "thin", color: "#D9E2EC" }, vAlign: "center" });
  main.getRange("A:A").format.columnWidth = 28;
  main.getRange("B:B").format.columnWidth = 42;
  main.getRange("C:G").format.columnWidth = 16;
  main.getRange("H:R").format.columnWidth = 13;
  main.getRange("S:U").format.columnWidth = 14;
  main.getRange("V:X").format.columnWidth = 18;
  main.getRange("Y:AA").format.columnWidth = 14;
  main.getRange(`H2:I${lastRow}`).setNumberFormat("#,##0.00");
  main.getRange(`J2:J${lastRow}`).setNumberFormat("0.00");
  main.getRange(`K2:N${lastRow}`).setNumberFormat("0.00%");
  main.getRange(`O2:R${lastRow}`).setNumberFormat("#,##0");
  main.getRange(`S2:S${lastRow}`).setNumberFormat("#,##0.0000");
  main.getRange(`B2:B${lastRow}`).format.wrapText = true;
  main.getRange(`X2:X${lastRow}`).format.wrapText = true;
  main.getRange(`V2:V${lastRow}`).dataValidation = { rule: { type: "list", values: SCRIPT_STYLES } };
  main.getRange(`W2:W${lastRow}`).dataValidation = { rule: { type: "list", values: HOOK_TYPES } };
  main.getRange(`Z2:Z${lastRow}`).dataValidation = { rule: { type: "list", values: ["待复盘", "已复盘", "已复刻", "已停投", "待二剪"] } };

  dashboard.getRange("A1:H1").values = [["千川素材跑量看板", "", "", "", "", "", "", ""]];
  styleRange(dashboard.getRange("A1:H1"), { fill: "#23395B", fontColor: "#FFFFFF", bold: true, size: 14, hAlign: "center" });
  dashboard.getRange("A3:H4").values = [
    ["素材数", "总消耗", "总GMV", "整体ROI", "平均CTR", "平均3s播放率", "平均CVR", "A级素材数"],
    [
      `=COUNTA('飞书导入主表'!A2:A${lastRow})`,
      `=SUM('飞书导入主表'!H2:H${lastRow})`,
      `=SUM('飞书导入主表'!I2:I${lastRow})`,
      `=IFERROR(C4/B4,0)`,
      `=AVERAGE('飞书导入主表'!K2:K${lastRow})`,
      `=AVERAGE('飞书导入主表'!L2:L${lastRow})`,
      `=AVERAGE('飞书导入主表'!N2:N${lastRow})`,
      `=COUNTIF('飞书导入主表'!T2:T${lastRow},"A")`,
    ],
  ];
  styleRange(dashboard.getRange("A3:H3"), { fill: "#D9EAF7", bold: true, hAlign: "center" });
  styleRange(dashboard.getRange("A4:H4"), { fill: "#F8FBFD", bold: true, hAlign: "center" });
  dashboard.getRange("B4:C4").setNumberFormat("#,##0.00");
  dashboard.getRange("D4:D4").setNumberFormat("0.00");
  dashboard.getRange("E4:G4").setNumberFormat("0.00%");
  dashboard.getRange("A6:F6").values = [["素材评级", "素材数", "消耗", "GMV", "平均ROI", "运营动作"]];
  dashboard.getRange("A7:F10").values = [
    ["A", `=COUNTIF('飞书导入主表'!T2:T${lastRow},A7)`, `=SUMIF('飞书导入主表'!T2:T${lastRow},A7,'飞书导入主表'!H2:H${lastRow})`, `=SUMIF('飞书导入主表'!T2:T${lastRow},A7,'飞书导入主表'!I2:I${lastRow})`, `=IFERROR(AVERAGEIF('飞书导入主表'!T2:T${lastRow},A7,'飞书导入主表'!J2:J${lastRow}),0)`, "加预算、拆脚本、快速复刻"],
    ["B", `=COUNTIF('飞书导入主表'!T2:T${lastRow},A8)`, `=SUMIF('飞书导入主表'!T2:T${lastRow},A8,'飞书导入主表'!H2:H${lastRow})`, `=SUMIF('飞书导入主表'!T2:T${lastRow},A8,'飞书导入主表'!I2:I${lastRow})`, `=IFERROR(AVERAGEIF('飞书导入主表'!T2:T${lastRow},A8,'飞书导入主表'!J2:J${lastRow}),0)`, "小幅放量，验证稳定性"],
    ["C", `=COUNTIF('飞书导入主表'!T2:T${lastRow},A9)`, `=SUMIF('飞书导入主表'!T2:T${lastRow},A9,'飞书导入主表'!H2:H${lastRow})`, `=SUMIF('飞书导入主表'!T2:T${lastRow},A9,'飞书导入主表'!I2:I${lastRow})`, `=IFERROR(AVERAGEIF('飞书导入主表'!T2:T${lastRow},A9,'飞书导入主表'!J2:J${lastRow}),0)`, "保留标签，优先二剪"],
    ["D", `=COUNTIF('飞书导入主表'!T2:T${lastRow},A10)`, `=SUMIF('飞书导入主表'!T2:T${lastRow},A10,'飞书导入主表'!H2:H${lastRow})`, `=SUMIF('飞书导入主表'!T2:T${lastRow},A10,'飞书导入主表'!I2:I${lastRow})`, `=IFERROR(AVERAGEIF('飞书导入主表'!T2:T${lastRow},A10,'飞书导入主表'!J2:J${lastRow}),0)`, "停投或重做前3秒"],
  ];
  styleRange(dashboard.getRange("A6:F6"), { fill: "#23395B", fontColor: "#FFFFFF", bold: true, hAlign: "center" });
  styleRange(dashboard.getRange("A7:F10"), { border: { preset: "inside", style: "thin", color: "#D9E2EC" } });
  dashboard.getRange("C7:D10").setNumberFormat("#,##0.00");
  dashboard.getRange("E7:E10").setNumberFormat("0.00");

  chartData.getRange("A1:B1").values = [["脚本风格", "消耗金额"]];
  chartData.getRange(`A2:A${SCRIPT_STYLES.length + 1}`).values = SCRIPT_STYLES.map((v) => [v]);
  chartData.getRange("B2").formulas = [[`=SUMIF('飞书导入主表'!$V$2:$V$${lastRow},A2,'飞书导入主表'!$H$2:$H$${lastRow})`]];
  chartData.getRange(`B2:B${SCRIPT_STYLES.length + 1}`).fillDown();
  const hookStart = SCRIPT_STYLES.length + 4;
  chartData.getRange(`A${hookStart}:B${hookStart}`).values = [["开头钩子类型", "消耗金额"]];
  chartData.getRange(`A${hookStart + 1}:A${hookStart + HOOK_TYPES.length}`).values = HOOK_TYPES.map((v) => [v]);
  chartData.getRange(`B${hookStart + 1}`).formulas = [[`=SUMIF('飞书导入主表'!$W$2:$W$${lastRow},A${hookStart + 1},'飞书导入主表'!$H$2:$H$${lastRow})`]];
  chartData.getRange(`B${hookStart + 1}:B${hookStart + HOOK_TYPES.length}`).fillDown();
  styleRange(chartData.getRange("A1:B1"), { fill: "#D9EAF7", bold: true, hAlign: "center" });
  styleRange(chartData.getRange(`A${hookStart}:B${hookStart}`), { fill: "#D9EAF7", bold: true, hAlign: "center" });
  chartData.getRange("A:B").format.columnWidth = 18;
  chartData.getRange(`B2:B${hookStart + HOOK_TYPES.length}`).setNumberFormat("#,##0.00");

  dashboard.getRange("A12:B12").values = [["脚本风格", "消耗金额"]];
  dashboard.getRange(`A13:B${12 + SCRIPT_STYLES.length}`).formulas = SCRIPT_STYLES.map((_, i) => [
    `='图表数据'!A${i + 2}`,
    `='图表数据'!B${i + 2}`,
  ]);
  dashboard.getRange("D12:E12").values = [["开头钩子类型", "消耗金额"]];
  dashboard.getRange(`D13:E${12 + HOOK_TYPES.length}`).formulas = HOOK_TYPES.map((_, i) => [
    `='图表数据'!A${hookStart + i + 1}`,
    `='图表数据'!B${hookStart + i + 1}`,
  ]);
  styleRange(dashboard.getRange("A12:B12"), { fill: "#D9EAF7", bold: true, hAlign: "center" });
  styleRange(dashboard.getRange("D12:E12"), { fill: "#D9EAF7", bold: true, hAlign: "center" });
  dashboard.getRange(`B13:B${12 + SCRIPT_STYLES.length}`).setNumberFormat("#,##0.00");
  dashboard.getRange(`E13:E${12 + HOOK_TYPES.length}`).setNumberFormat("#,##0.00");
  try {
    const chart1 = dashboard.charts.add("pie", dashboard.getRange(`A12:B${12 + SCRIPT_STYLES.length}`), "Auto");
    chart1.title.text = "脚本风格消耗占比";
    chart1.setPosition(dashboard.getRange("G12:K25"));
    chart1.width = 430;
    chart1.height = 260;
    const chart2 = dashboard.charts.add("pie", dashboard.getRange(`D12:E${12 + HOOK_TYPES.length}`), "Auto");
    chart2.title.text = "开头钩子类型消耗占比";
    chart2.setPosition(dashboard.getRange("G27:K40"));
    chart2.width = 430;
    chart2.height = 260;
  } catch {
    dashboard.getRange("G12").values = [["图表创建失败：请在飞书中基于左侧汇总表插入饼图。"]];
  }

  const dictRows = [
    ["字段", "飞书字段类型建议", "来源/公式", "填写人", "说明"],
    ["素材ID", "文本/主字段", "源表：素材ID", "系统", "唯一主键，不建议编辑"],
    ["消耗金额", "数字/货币", "源表：整体消耗", "系统", "投放成本"],
    ["ROI", "数字", "源表：整体支付ROI", "系统", "成交金额/消耗"],
    ["CTR", "百分比", "源表：整体点击率", "系统", "点击效率"],
    ["3s播放率", "百分比", "源表：3秒播放率", "系统", "开头留人能力"],
    ["完播率", "百分比", "源表：视频完播率", "系统", "完整观看能力"],
    ["GMV", "数字/货币", "源表：整体成交金额", "系统", "成交结果"],
    ["CVR", "百分比", "源表：整体转化率", "系统", "点击后成交能力"],
    ["素材值", "公式", "消耗金额*ROI*3s播放率*CTR*CVR", "系统", "用于综合衡量跑量质量"],
    ["素材评级", "公式/单选", "仅对素材值>0按排名分层：前15%=A，15%-40%=B，40%-70%=C，其余=D", "系统", "快速圈定头部素材"],
    ["脚本风格", "单选/多选", "标题与标签初筛，可人工修正", "编导", SCRIPT_STYLES.join("、")],
    ["开头钩子类型", "单选/多选", "标题初筛，可人工修正", "编导", HOOK_TYPES.join("、")],
    ["复盘状态", "单选", "待复盘/已复盘/已复刻/已停投/待二剪", "运营", "管理素材优化流转"],
  ];
  dict.getRange(`A1:E${dictRows.length}`).values = dictRows;
  dict.tables.add(`A1:E${dictRows.length}`, true).name = "tblFieldDictionary";
  styleRange(dict.getRange("A1:E1"), { fill: "#23395B", fontColor: "#FFFFFF", bold: true, hAlign: "center" });
  styleRange(dict.getRange(`A1:E${dictRows.length}`), { border: { preset: "inside", style: "thin", color: "#D9E2EC" }, wrap: true, vAlign: "center" });
  dict.getRange("A:E").format.columnWidth = 24;
  dict.getRange("C:E").format.columnWidth = 34;

  const viewRows = [
    ["视图名称", "面向角色", "筛选条件", "分组", "排序", "建议动作"],
    ["01_头部跑量素材", "投放", "素材评级=A", "脚本风格", "素材值降序", "提高预算，记录放量时间点，拆分计划测试"],
    ["02_高潜待放量素材", "投放", "跑量类型=高潜待放量", "开头钩子类型", "ROI降序", "小预算加测，确认可持续性"],
    ["03_低效烧钱素材", "投放/运营", "跑量类型=低效烧钱", "素材来源", "消耗降序", "停投、换钩子、换价格机制"],
    ["04_编导复刻脚本", "编导", "素材评级=A/B", "脚本风格", "完播率降序", "沉淀前3秒、主体卖点、结尾转化结构"],
    ["05_开头钩子复盘", "编导", "3s播放率或CTR异常", "开头钩子类型", "CTR降序", "判断钩子吸睛但不成交、或不吸睛"],
    ["06_运营转化问题", "运营", "CTR高且CVR低", "脚本风格", "CVR升序", "检查商品卡、价格、承接页和利益点一致性"],
  ];
  views.getRange(`A1:F${viewRows.length}`).values = viewRows;
  views.tables.add(`A1:F${viewRows.length}`, true).name = "tblFeishuViews";
  styleRange(views.getRange("A1:F1"), { fill: "#23395B", fontColor: "#FFFFFF", bold: true, hAlign: "center" });
  styleRange(views.getRange(`A1:F${viewRows.length}`), { border: { preset: "inside", style: "thin", color: "#D9E2EC" }, wrap: true, vAlign: "center" });
  views.getRange("A:F").format.columnWidth = 24;
  views.getRange("C:F").format.columnWidth = 34;

  raw.getRangeByIndexes(0, 0, sourceValues.length, sourceValues[0].length).values = sourceValues;
  raw.tables.add(`A1:${colName(sourceValues[0].length)}${sourceValues.length}`, true).name = "tblRawQianchuan";
  raw.freezePanes.freezeRows(1);
  styleRange(raw.getRange("A1:R1"), { fill: "#4A5568", fontColor: "#FFFFFF", bold: true, hAlign: "center" });
  raw.getRange("A:R").format.columnWidth = 16;
  raw.getRange("A:A").format.columnWidth = 42;
  raw.getRange("B:B").format.columnWidth = 24;
  raw.getRange(`B2:B${sourceValues.length}`).setNumberFormat("@");

  return workbook;
}

async function exportImportPack(workbook, outputDir) {
  const packDir = path.join(outputDir, "飞书多维表格导入包");
  await fs.mkdir(packDir, { recursive: true });
  const exports = [
    ["飞书导入主表", "01_素材跑量主表_导入飞书.csv"],
    ["字段字典", "02_字段配置说明.csv"],
    ["飞书视图建议", "03_推荐视图配置.csv"],
    ["图表数据", "04_饼图数据_脚本风格与开头钩子.csv"],
  ];
  for (const [sheetName, fileName] of exports) {
    const values = workbook.worksheets.getItem(sheetName).getUsedRange(true).values;
    await fs.writeFile(path.join(packDir, fileName), toCsv(values), "utf8");
  }
  const readme = [
    "# 飞书多维表格导入说明",
    "",
    "1. 在飞书新建多维表格，选择从 CSV / Excel 导入。",
    "2. 上传 `01_素材跑量主表_导入飞书.csv` 作为主表。",
    "3. 将 `素材ID` 设置为文本字段，并设为主字段或唯一识别字段。",
    "4. 按 `02_字段配置说明.csv` 调整字段类型。",
    "5. 按 `03_推荐视图配置.csv` 新建投放、编导、运营视图。",
    "6. 用 `04_饼图数据_脚本风格与开头钩子.csv` 在飞书仪表盘里创建两个饼图：脚本风格消耗占比、开头钩子类型消耗占比。",
    "",
    "素材值公式：消耗金额 * ROI * 3s播放率 * CTR * CVR。",
  ].join("\n");
  await fs.writeFile(path.join(packDir, "README_导入说明.md"), readme, "utf8");
  return packDir;
}

async function main() {
  const inputPath = process.argv[2] && !process.argv[2].startsWith("--") ? process.argv[2] : await latestDownload();
  const outArg = argValue("--out");
  const outputDir = outArg ? path.resolve(outArg) : path.resolve("outputs", "qianchuan_material_pipeline", timestamp());
  await fs.mkdir(outputDir, { recursive: true });

  const sourceWb = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
  const sourceSheet = sourceWb.worksheets.getItemAt(0);
  const sourceValues = sourceSheet.getUsedRange().values;
  const records = buildRecords(sourceValues);
  const workbook = buildWorkbook({ sourcePath: inputPath, sourceValues, records });

  for (const [sheetName, range, fileName] of [
    ["飞书导入主表", "A1:AA20", "preview_main.png"],
    ["素材看板", "A1:K40", "preview_dashboard.png"],
    ["图表数据", "A1:B20", "preview_chart_data.png"],
  ]) {
    const blob = await workbook.render({ sheetName, range, scale: 1, format: "png" });
    await fs.writeFile(path.join(outputDir, fileName), new Uint8Array(await blob.arrayBuffer()));
  }

  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: "final formula error scan",
    maxChars: 4000,
  });
  const errorText = errors.ndjson;
  if (!/matched 0 entries/.test(errorText)) {
    await fs.writeFile(path.join(outputDir, "formula_errors.ndjson"), errorText, "utf8");
    throw new Error(`公式错误扫描发现问题，详见 ${path.join(outputDir, "formula_errors.ndjson")}`);
  }

  const xlsxPath = path.join(outputDir, "千川素材跑量分析_飞书多维表格设计.xlsx");
  const xlsx = await SpreadsheetFile.exportXlsx(workbook);
  await xlsx.save(xlsxPath);
  const packDir = await exportImportPack(workbook, outputDir);

  const summary = {
    inputPath,
    outputDir,
    xlsxPath,
    packDir,
    rows: records.length,
    totalSpend: records.reduce((sum, r) => sum + r.spend, 0),
    totalGmv: records.reduce((sum, r) => sum + r.gmv, 0),
  };
  await fs.writeFile(path.join(outputDir, "summary.json"), JSON.stringify(summary, null, 2), "utf8");
  console.log(JSON.stringify(summary, null, 2));
}

await main();
