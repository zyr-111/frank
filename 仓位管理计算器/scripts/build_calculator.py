# -*- coding: utf-8 -*-
"""
仓位管理计算器 + 盈利计算器 生成器（参数化品种）

一次生成两个版本：
  - Excel：<名称>仓位管理与盈利计算器.xlsx（单Sheet两表联动，微软雅黑，边框完整，输入格保护）
  - 网页：<名称>仓位管理与盈利计算器.html（浏览器即开即用，蓝框可改，实时联动）

换品种只需改 --name / --point-value / --note 三个参数，公式、配色、字体、保护、警示全部自动套用。

用法：
  python build_calculator.py --outdir <目录>
  python build_calculator.py --outdir <目录> --name 原油 --point-value 10 --note "1手=1000桶，1个点=$10/手"
  python build_calculator.py --outdir <目录> --name 美日 --point-value 6.7 --note "1标准手=10万美日，1点(pip)≈$6.7/手"
  python build_calculator.py --outdir <目录> --only excel   # 只生成Excel
  python build_calculator.py --outdir <目录> --only html    # 只生成网页

品种点值参考（每点每手的美元价值）：
  黄金XAUUSD=100  白银XAGUSD=50  原油USOIL=10  欧美EURUSD=10  英镑GBPUSD=10
  美日USDJPY≈6.7  澳美AUDUSD=6.5  纳指/标普按合约规格  具体以你的平台为准
"""
import argparse
import os
import sys

try:
    import openpyxl
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "openpyxl>=3.1.0"])
    import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils.cell import range_boundaries


def xl_color(css_hex):
    return "FF" + css_hex.removeprefix("#").upper()


# ============ 固定样式（打磨成型的最终版，勿改）============
XL_TITLE_BG = xl_color("#FFC000")
XL_HEADER_BG = xl_color("#4472C4")
XL_INPUT_BG = xl_color("#D9E2F3")
XL_OUT_BG = xl_color("#548235")
XL_BORDER = xl_color("#7F7F7F")
XL_NOTE = xl_color("#595959")
XL_BLACK = xl_color("#000000")
FONT_NAME = "微软雅黑"

thin = Side(style="medium", color=XL_BORDER)
box = Border(left=thin, right=thin, top=thin, bottom=thin)

f_title = Font(name=FONT_NAME, bold=True, size=16, color=XL_BLACK)
f_head = Font(name=FONT_NAME, bold=True, size=12, color=xl_color("#FFFFFF"))
f_input = Font(name=FONT_NAME, bold=True, size=13, color=XL_BLACK)
f_calc = Font(name=FONT_NAME, bold=True, size=13, color=XL_BLACK)
f_out = Font(name=FONT_NAME, bold=True, size=13, color=xl_color("#FFFFFF"))
f_note = Font(name=FONT_NAME, italic=False, size=10, color=XL_NOTE)
center = Alignment(horizontal="center", vertical="center")
left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)


def put(ws, cell, value, font=None, fill=None, fmt=None, align=None, border=True):
    c = ws[cell]
    c.value = value
    if font:
        c.font = font
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if fmt:
        c.number_format = fmt
    c.alignment = align or center
    if border:
        c.border = box
    return c


def merge_put(ws, rng, value, font=None, fill=None, fmt=None, align=None):
    """先给范围内每个单元格设边框/底色，再合并（否则 openpyxl 会丢失非左上角样式）"""
    min_col, min_row, max_col, max_row = range_boundaries(rng)
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = box
            if fill:
                cell.fill = PatternFill("solid", fgColor=fill)
    ws.merge_cells(rng)
    first = ws.cell(row=min_row, column=min_col)
    first.value = value
    if font:
        first.font = font
    if fmt:
        first.number_format = fmt
    first.alignment = align or center
    return first


def build_excel(path, cfg):
    """生成 Excel 版。cfg 含 name / point_value / note / capital / stop_pct / rr / stop_range / trades / win_rate"""
    pv = cfg["point_value"]
    name = cfg["name"]
    wb = openpyxl.Workbook()
    wb.properties.title = "交易仓位管理与盈利计算器"
    ws = wb.active
    ws.title = "计算器"

    # ---- 表1 ----
    merge_put(ws, "A1:H1", "仓位管理计算器（以%s为例）" % name, font=f_title, fill=XL_TITLE_BG, align=center)
    heads1 = ["总仓位", "单笔止损", "单笔止损金额", "盈亏比", "单笔盈利", "单笔盈利金额", "止损范围", "单笔手数"]
    for i, h in enumerate(heads1):
        put(ws, f"{chr(65+i)}2", h, font=f_head, fill=XL_HEADER_BG)

    put(ws, "A3", cfg["capital"], font=f_input, fill=XL_INPUT_BG, fmt="#,##0")
    put(ws, "B3", cfg["stop_pct"], font=f_input, fill=XL_INPUT_BG, fmt="0.00%")
    put(ws, "C3", '=IF(OR(A3="",B3=""),"",A3*B3)', font=f_calc, fmt="$#,##0.00")
    put(ws, "D3", cfg["rr"], font=f_input, fill=XL_INPUT_BG, fmt="0.00")
    put(ws, "E3", '=IF(OR(B3="",D3=""),"",B3*D3)', font=f_calc, fmt="0.00%")
    put(ws, "F3", '=IF(OR(C3="",D3=""),"",C3*D3)', font=f_calc, fmt="$#,##0.00")
    put(ws, "G3", cfg["stop_range"], font=f_input, fill=XL_INPUT_BG, fmt="0.00")
    # 核心：点值 pv 写进 H3 公式
    put(ws, "H3", '=IF(OR(C3="",G3="",G3=0),"",C3/(G3*%s))' % fmt_num(pv),
        font=f_out, fill=XL_OUT_BG, fmt="0.00")

    note1 = "说明：浅蓝色4格可改（总仓位 / 单笔止损 / 盈亏比 / 止损范围），其余自动计算。%s 固定在H3公式中。" % cfg["note"]
    merge_put(ws, "A5:H5", note1, font=f_note, align=left_wrap)

    # ---- 表2 ----
    merge_put(ws, "A7:H7", "盈利计算器", font=f_title, fill=XL_TITLE_BG, align=center)
    put(ws, "A8", "交易次数", font=f_head, fill=XL_TITLE_BG)
    merge_put(ws, "B8:H8", cfg["trades"], font=f_input, fill=XL_INPUT_BG, fmt="#,##0",
              align=Alignment(horizontal="left", vertical="center"))
    heads2 = ["胜率", "盈利次数", "亏损次数", "总盈利", "总盈利金额", "总亏损", "总亏损金额", "实际利润"]
    for i, h in enumerate(heads2):
        put(ws, f"{chr(65+i)}9", h, font=f_head, fill=XL_HEADER_BG)

    put(ws, "A10", cfg["win_rate"], font=f_input, fill=XL_INPUT_BG, fmt="0.00%")
    put(ws, "B10", '=IF(OR(B8="",A10=""),"",ROUND(B8*A10,0))', font=f_calc, fmt="#,##0")
    put(ws, "C10", '=IF(OR(B8="",B10=""),"",B8-B10)', font=f_calc, fmt="#,##0")
    put(ws, "D10", '=IF(OR(B10="",E3=""),"",B10*E3)', font=f_calc, fmt="0.00%")
    put(ws, "E10", '=IF(OR(B10="",F3=""),"",B10*F3)', font=f_calc, fmt="$#,##0.00")
    put(ws, "F10", '=IF(OR(C10="",B3=""),"",C10*B3)', font=f_calc, fmt="0.00%")
    put(ws, "G10", '=IF(OR(C10="",C3=""),"",C10*C3)', font=f_calc, fmt="$#,##0.00")
    put(ws, "H10", '=IF(OR(E10="",G10=""),"",E10-G10)', font=f_out, fill=XL_OUT_BG, fmt="$#,##0.00")

    merge_put(ws, "A12:H12", "说明：浅蓝色2格可改（交易次数 / 胜率）；单笔盈利与止损金额自动取上方「仓位管理计算器」，改表1即联动。",
              font=f_note, align=left_wrap)

    # 参考指标
    merge_put(ws, "A14:B14", "实际收益率", font=f_head, fill=XL_HEADER_BG)
    put(ws, "C14", '=IF(OR(H10="",A3="",A3=0),"",H10/A3)', font=f_calc, fmt="0.00%")
    merge_put(ws, "D14:H14", "＝ 实际利润 ÷ 总仓位", font=f_note, align=left_wrap)
    merge_put(ws, "A15:B15", "盈亏平衡胜率", font=f_head, fill=XL_HEADER_BG)
    put(ws, "C15", '=IF(OR(D3="",D3=0),"",1/(1+D3))', font=f_calc, fmt="0.00%")
    merge_put(ws, "D15:H15", "＝ 1÷(1+盈亏比)，胜率低于此值则亏损", font=f_note, align=left_wrap)

    for col, w in zip("ABCDEFGH", [14, 13, 16, 11, 13, 16, 13, 13]):
        ws.column_dimensions[col].width = w
    for row, h in [(1, 36), (2, 40), (3, 32), (5, 32), (7, 36), (8, 32), (9, 40), (10, 32), (12, 32), (14, 30), (15, 30)]:
        ws.row_dimensions[row].height = h

    ws.conditional_formatting.add("H10", CellIsRule(operator="lessThan", formula=["0"],
        font=Font(name=FONT_NAME, color=xl_color("#9C0006"), bold=True)))
    ws.conditional_formatting.add("A10", FormulaRule(
        formula=['AND(ISNUMBER($A$10),ISNUMBER($C$15),$A$10<$C$15)'],
        font=Font(name=FONT_NAME, color=xl_color("#9C0006"), bold=True)))
    ws.freeze_panes = "A2"

    # 工作表保护：只开放6个浅蓝输入格
    for addr in ["A3", "B3", "D3", "G3", "B8", "A10"]:
        ws[addr].protection = Protection(locked=False)
    ws.protection.sheet = True
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False
    ws.protection.formatCells = True
    ws.protection.insertRows = True
    ws.protection.deleteRows = True

    wb.save(path)
    return path


def fmt_num(v):
    """数字转公式里的字符串，整数不带小数点"""
    if isinstance(v, int) or (isinstance(v, float) and v.is_integer()):
        return str(int(v))
    return str(v)


def build_html(path, cfg):
    """生成网页版。点值 pv 注入到 JS 常量与说明文字。"""
    pv = cfg["point_value"]
    name = cfg["name"]
    pv_s = fmt_num(pv)
    pct = lambda x: fmt_num(x * 100)

    html = HTML_TEMPLATE
    html = html.replace("{{TITLE}}", "交易仓位管理与盈利计算器")
    html = html.replace("{{NAME}}", name)
    html = html.replace("{{PV_S}}", pv_s)
    html = html.replace("{{NOTE}}", cfg["note"])
    html = html.replace("{{CAP}}", fmt_num(cfg["capital"]))
    html = html.replace("{{STOP_PCT}}", pct(cfg["stop_pct"]))
    html = html.replace("{{RR}}", fmt_num(cfg["rr"]))
    html = html.replace("{{STOP_RANGE}}", fmt_num(cfg["stop_range"]))
    html = html.replace("{{TRADES}}", fmt_num(cfg["trades"]))
    html = html.replace("{{WIN_RATE}}", pct(cfg["win_rate"]))
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


HTML_TEMPLATE = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{{TITLE}}</title>
<style>
  :root{
    --primary:#4472C4; --primary-dark:#2F5597;
    --input:#D9E2F3; --calc:#F7F9FC; --out:#548235; --out-text:#fff;
    --ref:#FFF2CC; --border:#D9DEE7; --text:#1A1A1A; --muted:#808080;
    --red:#9C0006; --green:#006100;
    --bg:#EEF1F6; --card:#fff;
  }
  *{box-sizing:border-box}
  body{margin:0;font-family:-apple-system,"PingFang SC","Microsoft YaHei",Segoe UI,Roboto,sans-serif;
    background:var(--bg);color:var(--text);line-height:1.5}
  .wrap{max-width:980px;margin:0 auto;padding:24px 16px 60px}
  header{display:flex;align-items:center;gap:14px;margin-bottom:18px}
  header .logo{width:40px;height:40px;border-radius:10px;background:linear-gradient(135deg,var(--primary),var(--primary-dark));
    display:flex;align-items:center;justify-content:center;color:#fff;font-weight:700;font-size:20px}
  header h1{font-size:20px;margin:0;font-weight:700}
  header .sub{font-size:12.5px;color:var(--muted);font-weight:400}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}
  @media(max-width:820px){.grid{grid-template-columns:1fr}}
  .card{background:var(--card);border:1px solid var(--border);border-radius:14px;padding:18px 18px 14px;
    box-shadow:0 1px 3px rgba(0,0,0,.04)}
  .card h2{margin:0 0 4px;font-size:15.5px;display:flex;align-items:center;gap:8px}
  .card h2 .tag{font-size:11px;font-weight:600;color:#fff;background:var(--primary);padding:2px 8px;border-radius:20px}
  .card .hint{font-size:12px;color:var(--muted);margin:0 0 14px}
  table{width:100%;border-collapse:collapse;font-size:13.5px}
  th{background:var(--primary);color:#fff;font-weight:600;padding:9px 8px;text-align:center;border:1px solid var(--primary)}
  td{padding:8px;border:1px solid var(--border);text-align:center;vertical-align:middle}
  td.label{background:#F2F5FA;font-weight:600;text-align:left;padding-left:10px;white-space:nowrap}
  input[type=number]{width:100%;border:1.5px solid var(--primary);background:var(--input);
    border-radius:6px;padding:6px 8px;font-size:14px;font-weight:600;color:var(--primary-dark);text-align:center}
  input[type=number]:focus{outline:none;border-color:var(--primary-dark);box-shadow:0 0 0 3px rgba(68,114,196,.18)}
  .out{background:var(--out);color:var(--out-text);font-weight:700;font-size:14px}
  .calc{background:var(--calc)}
  .ref{background:var(--ref)}
  .unit{color:var(--muted);font-size:11.5px}
  .neg{color:var(--red)!important}
  .pos{color:var(--green)}
  .bonus{margin-top:14px;border-top:1px dashed var(--border);padding-top:12px;display:flex;gap:24px;flex-wrap:wrap}
  .bonus .item{font-size:12.5px}
  .bonus .item b{display:block;font-size:17px;margin-top:2px}
  .footer{margin-top:18px;font-size:12px;color:var(--muted);text-align:center}
  .gold-note{font-size:11.5px;color:var(--muted);background:#FFFBE6;border:1px solid #FFE9A8;
    border-radius:6px;padding:6px 10px;margin-top:10px}
  .trades-row{display:flex;align-items:center;gap:10px;margin-bottom:10px}
  .trades-row label{font-weight:600;font-size:13px}
  .trades-row input{width:90px}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div class="logo">仓</div>
    <div>
      <h1>交易仓位管理与盈利计算器</h1>
      <div class="sub">{{NAME}}为例 · 改蓝框即时联动</div>
    </div>
  </header>

  <div class="grid">
    <!-- 表1 仓位管理计算器 -->
    <div class="card">
      <h2>仓位管理计算器 <span class="tag">以{{NAME}}为例</span></h2>
      <p class="hint">改下面 4 个蓝框，其余自动计算</p>
      <table>
        <tr>
          <th>总仓位</th><th>单笔止损</th><th>单笔止损金额</th><th>盈亏比</th>
        </tr>
        <tr>
          <td><input type="number" id="cap" value="{{CAP}}" step="1000"></td>
          <td><input type="number" id="stopPct" value="{{STOP_PCT}}" step="0.5"></td>
          <td class="calc" id="stopAmt">—</td>
          <td><input type="number" id="rr" value="{{RR}}" step="0.5"></td>
        </tr>
        <tr>
          <th>单笔盈利</th><th>单笔盈利金额</th><th>止损范围</th><th>单笔手数</th>
        </tr>
        <tr>
          <td class="calc" id="winPct">—</td>
          <td class="calc" id="winAmt">—</td>
          <td><input type="number" id="stopRange" value="{{STOP_RANGE}}" step="0.5"></td>
          <td class="out" id="lots">—</td>
        </tr>
      </table>
      <div class="gold-note">{{NOTE}}，即 <b>${{PV_S}}/手</b>。<br>手数 = 单笔止损金额 ÷（止损范围 × {{PV_S}}）。</div>
    </div>

    <!-- 表2 盈利计算器 -->
    <div class="card">
      <h2>盈利计算器 <span class="tag">联动表</span></h2>
      <p class="hint">改下面 2 个蓝框；单笔盈利/止损自动取表1</p>
      <div class="trades-row">
        <label>交易次数</label>
        <input type="number" id="trades" value="{{TRADES}}" step="1">
        <label style="margin-left:14px">胜率(%)</label>
        <input type="number" id="winRate" value="{{WIN_RATE}}" step="5">
      </div>
      <table>
        <tr>
          <th>盈利次数</th><th>亏损次数</th><th>总盈利</th><th>总盈利金额</th>
        </tr>
        <tr>
          <td class="calc" id="winN">—</td>
          <td class="calc" id="lossN">—</td>
          <td class="calc" id="totWinPct">—</td>
          <td class="calc" id="totWinAmt">—</td>
        </tr>
        <tr>
          <th>总亏损</th><th>总亏损金额</th><th colspan="2">实际利润</th>
        </tr>
        <tr>
          <td class="calc" id="totLossPct">—</td>
          <td class="calc" id="totLossAmt">—</td>
          <td class="out" id="profit" colspan="2">—</td>
        </tr>
      </table>
      <div class="bonus">
        <div class="item">实际收益率<b id="roi">—</b></div>
        <div class="item">盈亏平衡胜率<b id="breakeven">—</b><span class="unit">（低于此值则亏损）</span></div>
        <div class="item">单笔盈利金额（取表1）<b id="refWin" style="font-size:14px">—</b></div>
        <div class="item">单笔止损金额（取表1）<b id="refLoss" style="font-size:14px">—</b></div>
      </div>
    </div>
  </div>

  <div class="footer">
    计算逻辑：单笔止损金额 = 总仓位×止损比例 · 单笔手数 = 止损金额÷(止损范围×{{PV_S}}) ·
    实际利润 = 盈利次数×单笔盈利 − 亏损次数×单笔止损 · 仅供学习，不构成投资建议
  </div>
</div>

<script>
const $ = id => document.getElementById(id);
const POINT_VALUE = {{PV_S}};   // 品种点值：每点每手美元
const fmtUSD = n => n==null||isNaN(n)?'—':'$'+n.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtInt = n => n==null||isNaN(n)?'—':n.toLocaleString('en-US');
const fmtPct = (n,d=2) => n==null||isNaN(n)?'—':(n*100).toFixed(d)+'%';
const num = id => { const v=parseFloat($(id).value); return isNaN(v)?0:v; };

function calc(){
  // 表1
  const cap = num('cap'), stopPct = num('stopPct')/100, rr = num('rr'), stopRange = num('stopRange');
  const stopAmt = cap*stopPct;
  const winPct = stopPct*rr;
  const winAmt = stopAmt*rr;
  const perLotRisk = stopRange*POINT_VALUE;
  const lots = perLotRisk>0 ? stopAmt/perLotRisk : 0;

  $('stopAmt').textContent = fmtUSD(stopAmt);
  $('winPct').textContent = fmtPct(winPct);
  $('winAmt').textContent = fmtUSD(winAmt);
  $('lots').textContent = lots.toFixed(2)+' 手';

  // 表2（联动表1）
  const trades = num('trades'), winRate = num('winRate')/100;
  const winN = Math.round(trades*winRate);
  const lossN = trades-winN;
  const totWinPct = winN*winPct;
  const totWinAmt = winN*winAmt;
  const totLossPct = lossN*stopPct;
  const totLossAmt = lossN*stopAmt;
  const profit = totWinAmt-totLossAmt;
  const roi = cap>0 ? profit/cap : 0;
  const breakeven = rr>0 ? 1/(1+rr) : 0;

  $('winN').textContent = fmtInt(winN);
  $('lossN').textContent = fmtInt(lossN);
  $('totWinPct').textContent = fmtPct(totWinPct);
  $('totWinAmt').textContent = fmtUSD(totWinAmt);
  $('totLossPct').textContent = fmtPct(totLossPct);
  $('totLossAmt').textContent = fmtUSD(totLossAmt);
  const pCell=$('profit');
  pCell.textContent = fmtUSD(profit);
  pCell.style.background = profit<0 ? '#9C0006' : 'var(--out)';

  $('roi').textContent = fmtPct(roi);
  $('breakeven').textContent = fmtPct(breakeven);
  $('refWin').textContent = fmtUSD(winAmt);
  $('refLoss').textContent = fmtUSD(stopAmt);

  // 胜率低于平衡胜率 → 胜率输入框红字警示
  const wrCell=$('winRate');
  if(winRate>0 && breakeven>0 && winRate<breakeven){
    wrCell.style.color='var(--red)'; wrCell.style.borderColor='var(--red)';
  }else{
    wrCell.style.color='var(--primary-dark)'; wrCell.style.borderColor='var(--primary)';
  }
}

['cap','stopPct','rr','stopRange','trades','winRate'].forEach(id=>{
  $(id).addEventListener('input', calc);
});
calc();
</script>
</body>
</html>
'''


def main():
    ap = argparse.ArgumentParser(description="生成仓位管理计算器（Excel + 网页），支持品种点值参数化")
    ap.add_argument("--outdir", required=True, help="输出目录")
    ap.add_argument("--name", default="黄金", help="品种名，如 黄金/原油/美日/欧美（默认黄金）")
    ap.add_argument("--point-value", type=float, default=100, help="每点每手美元价值，黄金=100 原油=10 美日=6.7（默认100）")
    ap.add_argument("--note", default="1手=100盎司，1个点(1美元)=$100/手", help="合约说明文字")
    ap.add_argument("--capital", type=float, default=1000, help="默认总仓位")
    ap.add_argument("--stop-pct", type=float, default=0.05, help="默认单笔止损比例")
    ap.add_argument("--rr", type=float, default=3, help="默认盈亏比")
    ap.add_argument("--stop-range", type=float, default=5, help="默认止损范围(点数)")
    ap.add_argument("--trades", type=int, default=10, help="默认交易次数")
    ap.add_argument("--win-rate", type=float, default=0.5, help="默认胜率")
    ap.add_argument("--only", choices=["excel", "html"], help="只生成其中一种")
    ap.add_argument("--prefix", default="", help="文件名前缀，如 '原油' 则生成 原油仓位管理与盈利计算器.xlsx")
    args = ap.parse_args()

    os.makedirs(args.outdir, exist_ok=True)
    prefix = args.prefix if args.prefix else ("" if args.name == "黄金" else args.name)
    base = prefix + "仓位管理与盈利计算器"
    cfg = dict(name=args.name, point_value=args.point_value, note=args.note,
               capital=args.capital, stop_pct=args.stop_pct, rr=args.rr,
               stop_range=args.stop_range, trades=args.trades, win_rate=args.win_rate)

    made = []
    if args.only in (None, "excel"):
        made.append(build_excel(os.path.join(args.outdir, base + ".xlsx"), cfg))
    if args.only in (None, "html"):
        made.append(build_html(os.path.join(args.outdir, base + ".html"), cfg))

    print("GENERATED:")
    for p in made:
        print("  " + p)
    print("品种=%s 点值=%s 说明=%s" % (args.name, fmt_num(args.point_value), args.note))


if __name__ == "__main__":
    main()
