---
name: 仓位管理计算器
description: 生成/换品种「交易仓位管理与盈利计算器」：一次产出 Excel + 网页两个版本，两表联动、6个浅蓝格可改、其余锁定、亏损红字警示。当用户说"做仓位计算器 / 仓位管理计算器 / 盈利计算器 / 换成原油/美日/外汇/其他品种 / 用计算器套用某品种 / 重新生成计算器"时使用。已封装好今天打磨成型的最终样式，换品种只需给品种名和点值，不要重新设计表格。
---

# 仓位管理计算器 Skill

把今天打磨成型的「仓位管理 + 盈利计算器」封装成可复用生成器。**一次生成 Excel + 网页两个版本**，换品种只改点值参数，公式/配色/字体/保护/警示全部自动套用——不要重新设计表格。

## 什么时候用

- 用户要一个仓位管理/盈利计算器（Excel 版、网页版或两版都要）
- 用户要**换品种**："换成原油""做外汇美日的""换成欧美/白银""用这个计算器套XX品种"
- 用户要重新生成/恢复默认（黄金）版

## 核心：品种点值（point_value）

点值 = **1 个点 × 1 手 = 多少美元**，是整个计算器的唯一品种相关参数。

| 品种 | 参数 `--point-value` | 说明文字 `--note` |
|---|---|---|
| 黄金 XAUUSD | 100 | 1手=100盎司，1个点(1美元)=$100/手 |
| 白银 XAGUSD | 50 | 1手=5000盎司，1个点=$50/手 |
| 原油 USOIL | 10 | 1手=1000桶，1个点(0.01美元)=$10/手 |
| 欧美 EURUSD | 10 | 1标准手=10万，1个点(pip)=$10/手 |
| 英镑 GBPUSD | 10 | 1标准手=10万，1个点(pip)=$10/手 |
| 美日 USDJPY | 6.7 | 1标准手=10万美日，1个点(pip)≈$6.7/手（随汇率波动） |
| 澳美 AUDUSD | 6.5 | 1标准手=10万，1个点(pip)≈$6.5/手 |

> 点值以用户平台为准；不确定就问用户，或用表格里的常见值并说明"以你的平台为准"。

## 工作流程

### 第 1 步：确定参数

从用户话里提取：品种名、点值、合约说明。默认（用户没说品种）= 黄金、点值100、说明"1手=100盎司，1个点(1美元)=$100/手"。
默认值：总仓位1000 / 单笔止损5% / 盈亏比3 / 止损范围5点 / 交易次数10 / 胜率50%。用户要改默认值就加对应参数。

### 第 2 步：运行生成脚本

```bash
python "C:\Users\Administrator\.workbuddy\skills\仓位管理计算器\scripts\build_calculator.py" --outdir "<输出目录>"
```

换品种示例（**注意：说明文字里有 `$`，Bash 里必须用单引号包住，否则 `$10` 会被当变量吃掉**）：

```bash
# 原油
python "C:\Users\Administrator\.workbuddy\skills\仓位管理计算器\scripts\build_calculator.py" --outdir "<目录>" \
  --name 原油 --point-value 10 --note '1手=1000桶，1个点(0.01美元)=$10/手'

# 美日
python "C:\Users\Administrator\.workbuddy\skills\仓位管理计算器\scripts\build_calculator.py" --outdir "<目录>" \
  --name 美日 --point-value 6.7 --note '1标准手=10万美日，1个点(pip)≈$6.7/手'
```

其他参数：`--only excel` / `--only html` 只生成一种；`--capital --stop-pct --rr --stop-range --trades --win-rate` 改默认值；`--prefix` 改文件名前缀。

### 第 3 步：交付

用 present_files 同时打开生成的 `.xlsx` 和 `.html`。Excel 版需打 AIGC 合规标识：

```bash
python "C:/Users/Administrator/.workbuddy/plugins/cache/workbuddy-builtin/workbuddy-sheetagent/5.9.2-wb.42136757.g0b5051be.hd4338c001c60/skills/workbuddy-sheet-generation/scripts/aigc_label.py" "<xlsx路径>" --platform WORKBUDDY
```

## 写文件被占用时的处理（重要）

Excel 若正在预览/WPS 打开，会锁文件报 `PermissionError` 或 `file is occupied`。解决顺序：

1. 查后台实例：`python3 "D:/WorkBuddy/resources/app.asar.unpacked/resources/plugins/workbuddy-builtin/skills/workbuddy-local-office-edit/edsdk.py" call get_pool_status`
2. 有实例就 `close_file file_id=<id>`（用 force=true）；pool 为空还锁着 = 前端预览面板持有，请用户点预览的 × 关闭
3. 关闭后再跑脚本覆盖

## 成品已固化的特性（勿改，都是用户逐条确认过的）

**Excel 版**（单 Sheet「计算器」，两个表纵向排列）：
- 配色：标题黄底 #FFC000、表头蓝底 #4472C4、输入格浅蓝 #D9E2F3、输出格深绿 #548235、medium 边框
- 字体：微软雅黑，标题16pt、表头12pt、数值13pt（WPS 里够黑够大）
- 合并单元格必须先设边框再 merge，否则边框缺失
- **工作表保护已开启**：只有 6 个浅蓝格可改（A3总仓位/B3止损/D3盈亏比/G3止损范围/B8交易次数/A10胜率），其余锁定防误碰；无密码，可在「审阅→撤销工作表保护」解开
- 公式：H3 = 止损金额÷(止损范围×点值)；表2 全部引用表1 联动
- 参考指标：实际收益率、盈亏平衡胜率 = 1/(1+盈亏比)，**说明文字不含写死的"25%"**（盈亏比一改就变）
- 条件格式：实际利润<0 红字加粗；胜率低于盈亏平衡胜率时胜率格红字

**网页版**：浏览器即开即用，6 个蓝框可改实时联动，副标题"XX为例 · 改蓝框即时联动"，利润为负红底、胜率不足红框，`POINT_VALUE` 常量控制点值。

## 环境

- Python：Bash 里 `python` 可用；PowerShell 里 PATH 的 python 是微软商店占位符（无效），真实路径 `C:\Users\Administrator\.workbuddy\binaries\python\versions\3.13.12\python.exe`
- 依赖：openpyxl（脚本缺时自动装）
