"""
HTML Dashboard 导出模块

职责：生成单文件自包含 HTML Dashboard，用于内部监控。
"""

import pandas as pd
import json
from datetime import datetime
from typing import List, Dict


def _build_change_log(merged_df: pd.DataFrame) -> List[Dict]:
    """
    从 merged_df 中提取变更日志

    Args:
        merged_df: 合并后的 DataFrame

    Returns:
        变更日志列表，每项包含 date, type, name, detail
    """
    log_entries = []

    if merged_df.empty or "变更类型" not in merged_df.columns:
        return log_entries

    # 筛选变更类型 != "无变化" 的行
    changed_df = merged_df[merged_df["变更类型"] != "无变化"].copy()

    for _, row in changed_df.iterrows():
        entry = {
            "date": str(row.get("交易所系统更新日期", "")) if pd.notna(row.get("交易所系统更新日期")) else "",
            "type": row.get("变更类型", ""),
            "name": row.get("债券名称", ""),
            "detail": row.get("变更详情", "")
        }
        log_entries.append(entry)

    return log_entries


def export_html(merged_df: pd.DataFrame, output_path: str) -> None:
    """
    生成单文件自包含 HTML Dashboard

    将 merged_df 序列化为 JSON 内嵌到 <script> 标签中。
    日期列格式化为 'YYYY-MM-DD' 字符串，NaN/NaT 转为空字符串。
    变更日志从 merged_df 中的变更类型、变更详情字段提取。

    Args:
        merged_df: 合并后的 DataFrame（由 merger.merge 返回）
        output_path: 输出 HTML 文件路径
    """
    if merged_df.empty:
        print("警告: 数据为空，无法导出 HTML")
        return

    # 数据预处理：填充 NaN，格式化日期
    df_clean = merged_df.copy()

    # 将 NaN/NaT 替换为空字符串
    df_clean = df_clean.fillna('')

    # 日期列格式化为字符串
    date_cols = ['交易所系统更新日期', '受理日期', '预计发行时间']
    for col in date_cols:
        if col in df_clean.columns:
            df_clean[col] = df_clean[col].apply(
                lambda x: x.strftime('%Y-%m-%d') if isinstance(x, pd.Timestamp) else str(x) if x else ''
            )

    # 序列化为 JSON
    json_data = df_clean.to_json(orient='records', force_ascii=False, date_format='iso')

    # 构建变更日志
    change_log = _build_change_log(merged_df)
    change_log_json = json.dumps(change_log, ensure_ascii=False)

    # 获取最新更新日期
    update_dates = df_clean['交易所系统更新日期'].tolist()
    update_dates = [d for d in update_dates if d]
    latest_update = max(update_dates) if update_dates else datetime.now().strftime('%Y-%m-%d')

    # 统计数量
    sse_count = len(df_clean[df_clean['交易所'] == '上交所'])
    szse_count = len(df_clean[df_clean['交易所'] == '深交所'])
    total_count = len(df_clean)
    new_count = len(df_clean[df_clean['变更类型'] == '新增'])
    chg_count = len(df_clean[df_clean['变更类型'] == '状态变更'])
    total_amount = df_clean['拟发行金额(亿元)'].apply(lambda x: float(x) if x else 0).sum()

    # HTML 模板
    html_template = f'''<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>持有型ABS申报项目追踪</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
:root{{
  --p:#1a56a0;--pl:#e8f0fb;--pm:#4a7fc1;
  --g:#2e7d32;--gl:#e8f5e9;
  --o:#e65100;--ol:#fff3e0;
  --bg:#f0f4fa;--sur:#fff;--bdr:#dde3ec;
  --tx:#1a2333;--tm:#6b7a8d;
  --r6:6px;--r8:8px;
}}
html,body{{height:100%;overflow:hidden}}
body{{font-family:'Microsoft YaHei',PingFang SC,sans-serif;font-size:12px;background:var(--bg);color:var(--tx)}}
.hdr{{background:var(--p);padding:12px 20px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0}}
.hdr-l .t{{color:#fff;font-size:15px;font-weight:500}}
.hdr-l .s{{color:rgba(255,255,255,.7);font-size:11px;margin-top:2px}}
.hdr-r{{display:flex;gap:8px;align-items:center}}
.hbg{{background:rgba(255,255,255,.2);color:#fff;font-size:11px;padding:3px 10px;border-radius:20px}}
.btn{{font-size:11px;padding:5px 12px;border-radius:var(--r6);border:none;cursor:pointer;font-weight:500;transition:opacity .15s}}
.btn:hover{{opacity:.85}}
.btn-dl{{background:#fff;color:var(--p)}}
.btn-log{{background:rgba(255,255,255,.2);color:#fff;border:1px solid rgba(255,255,255,.4)}}
.page{{display:flex;flex-direction:column;height:calc(100vh - 44px)}}
.cards-row{{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;padding:10px 16px;background:var(--sur);border-bottom:0.5px solid var(--bdr);flex-shrink:0}}
.card{{background:var(--bg);border-radius:var(--r8);padding:10px 14px;border-left:3px solid var(--p)}}
.card.g{{border-left-color:var(--g)}}
.card.o{{border-left-color:var(--o)}}
.card .cl{{font-size:10px;color:var(--tm);margin-bottom:3px}}
.card .cv{{font-size:22px;font-weight:500;color:var(--p)}}
.card.g .cv{{color:var(--g)}}
.card.o .cv{{color:var(--o)}}
.layout{{display:flex;flex:1;overflow:hidden}}
.sidebar{{width:210px;flex-shrink:0;background:var(--sur);border-right:0.5px solid var(--bdr);display:flex;flex-direction:column;overflow-y:auto}}
.sb-sect{{padding:12px 14px;border-bottom:0.5px solid var(--bdr)}}
.sb-title{{font-size:11px;font-weight:500;color:var(--tm);text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px}}
.fl{{font-size:11px;color:var(--tm);margin:6px 0 3px}}
select.fsel{{width:100%;font-size:12px;padding:5px 8px;border:0.5px solid var(--bdr);border-radius:var(--r6);background:var(--sur);color:var(--tx);cursor:pointer}}
select.fsel:focus{{outline:none;border-color:var(--pm)}}
.reset-btn{{width:100%;margin-top:8px;padding:5px;border:0.5px solid var(--bdr);border-radius:var(--r6);background:var(--sur);color:var(--tm);cursor:pointer;font-size:11px}}
.reset-btn:hover{{background:var(--bg)}}
.sum-block{{padding:12px 14px;flex:1}}
.sum-title{{font-size:11px;font-weight:500;color:var(--tm);margin-bottom:8px}}
.sum-row{{display:flex;justify-content:space-between;align-items:center;padding:4px 0;border-bottom:0.5px solid #eef1f5;font-size:11px}}
.sum-row:last-child{{border-bottom:none;font-weight:500;padding-top:6px}}
.sum-label{{color:var(--tm);max-width:120px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.sum-val{{color:var(--tx);font-weight:500;font-size:12px}}
.sum-row.total .sum-label,.sum-row.total .sum-val{{color:var(--p)}}
.main{{flex:1;display:flex;flex-direction:column;overflow:hidden}}
.tbl-wrap{{flex:1;overflow:auto;padding:8px 14px 12px}}
.grp-hint{{font-size:11px;color:var(--p);padding:4px 0 2px;opacity:.8}}
table{{width:100%;border-collapse:collapse;background:var(--sur);margin-bottom:10px;border:0.5px solid var(--bdr);border-radius:var(--r8);overflow:hidden}}
th{{background:#edf2fa;color:var(--p);font-weight:500;font-size:11px;padding:7px 8px;text-align:left;border-bottom:0.5px solid var(--bdr);white-space:nowrap;cursor:pointer;user-select:none;position:sticky;top:0;z-index:2}}
th:hover{{background:#dce8f7}}
th.sort-a:after{{content:' ▲';font-size:9px}}
th.sort-d:after{{content:' ▼';font-size:9px}}
th.grp-active{{background:#d0e3f7;color:var(--p)}}
td{{padding:6px 8px;border-bottom:0.5px solid #eef1f5;font-size:12px;vertical-align:middle;white-space:nowrap;max-width:300px;overflow:hidden;text-overflow:ellipsis}}
tr.dr:hover{{background:#f5f8fd;cursor:pointer}}
tr.exp td{{background:#fafbfe;white-space:normal;max-width:none}}
.exp-inner{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;padding:8px 4px}}
.exp-item .el{{font-size:10px;color:var(--tm);margin-bottom:3px}}
.exp-item .ev{{font-size:12px;color:var(--tx)}}
.fsel-inline{{font-size:11px;padding:3px 5px;border:0.5px solid var(--pm);border-radius:4px;background:var(--pl);color:var(--p);cursor:pointer;max-width:150px;width:100%}}
.fsel-inline:focus{{outline:none}}
.badge{{display:inline-block;font-size:10px;padding:2px 7px;border-radius:10px;font-weight:500}}
.b-new{{background:#e8f5e9;color:#2e7d32}}
.b-chg{{background:#fff3e0;color:#e65100}}
.b-non{{background:#f0f0f0;color:#888}}
.grp-hdr{{background:var(--pl);padding:6px 10px;font-size:12px;font-weight:500;color:var(--p);border:0.5px solid var(--bdr);border-radius:var(--r6) var(--r6) 0 0;border-bottom:none;margin-top:10px;display:flex;align-items:center;gap:8px}}
.grp-hdr span.cnt{{background:var(--p);color:#fff;font-size:10px;padding:1px 7px;border-radius:10px}}
.grp-hdr span.amt{{font-size:11px;color:var(--pm);font-weight:400}}
.grp-table{{border-radius:0 0 var(--r8) var(--r8)}}
.footer{{font-size:11px;color:var(--tm);padding:2px 14px 6px;text-align:right;flex-shrink:0}}
.log-panel{{position:fixed;top:0;right:-460px;width:440px;height:100%;background:var(--sur);border-left:1px solid var(--bdr);z-index:200;transition:right .28s;display:flex;flex-direction:column;overflow:hidden;box-shadow:-4px 0 16px rgba(0,0,0,.08)}}
.log-panel.open{{right:0}}
.log-hdr{{background:var(--p);color:#fff;padding:12px 16px;display:flex;justify-content:space-between;align-items:center;font-size:13px;font-weight:500;flex-shrink:0}}
.log-close{{background:none;border:none;color:#fff;font-size:18px;cursor:pointer;line-height:1}}
.log-tabs{{display:flex;border-bottom:0.5px solid var(--bdr);flex-shrink:0}}
.log-tab{{flex:1;padding:8px;text-align:center;font-size:12px;cursor:pointer;border-bottom:2px solid transparent;color:var(--tm)}}
.log-tab.act{{border-bottom-color:var(--p);color:var(--p);font-weight:500}}
.log-body{{flex:1;overflow-y:auto;padding:10px 14px}}
.log-entry{{border-bottom:0.5px solid var(--bdr);padding:8px 0}}
.log-date{{color:var(--tm);font-size:10px;margin-bottom:3px}}
.log-name{{font-size:12px;color:var(--tx);font-weight:500;margin-bottom:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:360px}}
.log-detail{{font-size:11px;color:var(--tm)}}
.log-empty{{text-align:center;color:var(--tm);padding:30px 0;font-size:12px}}
.log-actions{{padding:10px 14px;border-top:0.5px solid var(--bdr);display:flex;gap:8px;flex-shrink:0}}
.log-dl-btn{{flex:1;font-size:12px;padding:7px;border:0.5px solid var(--p);border-radius:var(--r6);background:var(--pl);color:var(--p);cursor:pointer;text-align:center}}
.log-dl-btn:hover{{background:var(--p);color:#fff}}
</style>
</head>
<body>

<div class="hdr">
  <div class="hdr-l">
    <div class="t">持有型ABS申报项目追踪</div>
    <div class="s" id="upd-time">数据更新：{latest_update}</div>
  </div>
  <div class="hdr-r">
    <span class="hbg" id="sse-c">上交所 {sse_count} 条</span>
    <span class="hbg" id="szse-c">深交所 {szse_count} 条</span>
    <button class="btn btn-log" onclick="toggleLog()">变更日志</button>
    <button class="btn btn-dl" onclick="dlExcel()">下载 Excel</button>
  </div>
</div>

<div class="page">
  <div class="cards-row">
    <div class="card"><div class="cl">总项目数</div><div class="cv" id="c-tot">{total_count}</div></div>
    <div class="card g"><div class="cl">新增项目</div><div class="cv" id="c-new">{new_count}</div></div>
    <div class="card o"><div class="cl">状态变更</div><div class="cv" id="c-chg">{chg_count}</div></div>
    <div class="card"><div class="cl">合计拟发行（亿元）</div><div class="cv" id="c-amt">{total_amount:.2f}</div></div>
  </div>
  <div class="layout">
    <div class="sidebar">
      <div class="sb-sect">
        <div class="sb-title">筛选条件</div>
        <div class="fl">交易所</div>
        <select class="fsel" id="f-exch" onchange="render()">
          <option value="">全部</option>
          <option>上交所</option>
          <option>深交所</option>
        </select>
        <div class="fl">交易所审批状态</div>
        <select class="fsel" id="f-status" onchange="render()"><option value="">全部</option></select>
        <div class="fl">变更类型</div>
        <select class="fsel" id="f-change" onchange="render()">
          <option value="">全部</option>
          <option>新增</option>
          <option>状态变更</option>
          <option>无变化</option>
        </select>
        <div class="fl">跟进和投放状态</div>
        <select class="fsel" id="f-follow" onchange="render()"><option value="">全部</option></select>
        <button class="reset-btn" onclick="resetFilters()">重置筛选</button>
      </div>
      <div class="sum-block">
        <div class="sum-title">按跟进投放状态汇总（亿元）</div>
        <div id="sum-table"></div>
      </div>
    </div>
    <div class="main">
      <div class="tbl-wrap">
        <div id="tbl-area"></div>
      </div>
      <div class="footer" id="footer"></div>
    </div>
  </div>
</div>

<div class="log-panel" id="log-panel">
  <div class="log-hdr">
    <span>变更日志</span>
    <button class="log-close" onclick="toggleLog()">×</button>
  </div>
  <div class="log-tabs">
    <div class="log-tab act" onclick="switchLogTab('week',this)">本周变更</div>
    <div class="log-tab" onclick="switchLogTab('month',this)">本月变更</div>
    <div class="log-tab" onclick="switchLogTab('all',this)">全部记录</div>
  </div>
  <div class="log-body" id="log-body"></div>
  <div class="log-actions">
    <button class="log-dl-btn" onclick="dlLog('week')">下载本周日志</button>
    <button class="log-dl-btn" onclick="dlLog('all')">下载全部日志</button>
  </div>
</div>

<script>
// ===== DATA (Python will inject JSON here) =====
const DATA = {json_data};

const FOLLOW_OPTS = ["已跟进，跟进中","已跟进，不参与","已跟进，未参与","重点跟进,推动审批","已投放","已跟进，终止发行","已跟进，根据中",""];
let rows = DATA.map((r,i) => ({{...r,_id:i}}));
let sortCol=null, sortDir=1, groupCol=null, curLogTab='week';

function fmtD(d){{return d.toISOString().slice(0,10);}}
function addDays(d,n){{const x=new Date(d);x.setDate(x.getDate()+n);return x;}}
const today=new Date();

let changeLog = {change_log_json};

function initFilters(){{
  const sv=[...new Set(DATA.map(r=>r['交易所审批状态']))].filter(Boolean).sort();
  const fs=document.getElementById('f-status');
  fs.innerHTML='<option value="">全部</option>';
  sv.forEach(v=>{{const o=document.createElement('option');o.value=o.textContent=v;fs.appendChild(o);}});
  const fv=[...new Set(DATA.map(r=>r['跟进和投放状态']).filter(Boolean))].sort();
  const ff=document.getElementById('f-follow');
  ff.innerHTML='<option value="">全部</option>';
  fv.forEach(v=>{{const o=document.createElement('option');o.value=o.textContent=v;ff.appendChild(o);}});
}}

function getFiltered(){{
  const ex=document.getElementById('f-exch').value;
  const st=document.getElementById('f-status').value;
  const ch=document.getElementById('f-change').value;
  const fw=document.getElementById('f-follow').value;
  return rows.filter(r=>(!ex||r['交易所']===ex)&&(!st||r['交易所审批状态']===st)&&(!ch||r['变更类型']===ch)&&(!fw||r['跟进和投放状态']===fw));
}}

function sortData(arr){{
  let s=[...arr].sort((a,b)=>(b['交易所系统更新日期']||'').localeCompare(a['交易所系统更新日期']||''));
  if(sortCol) s.sort((a,b)=>{{
    const va=a[sortCol]??'',vb=b[sortCol]??'';
    return (typeof va==='number'?va-vb:String(va).localeCompare(String(vb),'zh'))*sortDir;
  }});
  return s;
}}

function badge(t){{
  if(t==='新增') return '<span class="badge b-new">新增</span>';
  if(t==='状态变更') return '<span class="badge b-chg">变更</span>';
  return '<span class="badge b-non">无变化</span>';
}}

function updateFollow(el, rowId, oldVal){{
  const newVal=el.value;
  if(newVal===oldVal) return;
  rows[rowId]['跟进和投放状态']=newVal;
  changeLog.unshift({{date:fmtD(new Date()),type:'跟进状态变更',name:rows[rowId]['债券名称'],detail:`跟进状态: ${{oldVal||'(空)'}} → ${{newVal}}`}});
  renderSummary(getFiltered().map(r=>rows[r._id]||r));
  renderLog();
}}

function renderSummary(filtered){{
  const map={{}};
  filtered.forEach(r=>{{
    const k=r['跟进和投放状态']||'（未填写）';
    if(!map[k]) map[k]={{cnt:0,amt:0}};
    map[k].cnt++;
    map[k].amt+=(parseFloat(r['拟发行金额(亿元)'])||0);
  }});
  const total=filtered.reduce((s,r)=>s+(parseFloat(r['拟发行金额(亿元)'])||0),0);
  const order=['已投放','重点跟进,推动审批','已跟进，跟进中','已跟进，未参与','已跟进，不参与','已跟进，终止发行','已跟进，根据中','（未填写）'];
  let html='';
  [...order,...Object.keys(map).filter(k=>!order.includes(k))].filter(k=>map[k]).forEach(k=>{{
    html+=`<div class="sum-row"><span class="sum-label" title="${{k}}">${{k}}</span><span class="sum-val">${{map[k].amt.toFixed(2)}}</span></div>`;
  }});
  html+=`<div class="sum-row total"><span class="sum-label">合计</span><span class="sum-val">${{total.toFixed(2)}}</span></div>`;
  document.getElementById('sum-table').innerHTML=html;
}}

const COLS=[
  {{k:'债券名称',l:'债券名称',w:'22%'}},
  {{k:'交易所',l:'交易所',w:'6%'}},
  {{k:'计划管理人',l:'管理人',w:'8%'}},
  {{k:'拟发行金额(亿元)',l:'金额(亿)',w:'6%'}},
  {{k:'交易所审批状态',l:'审批状态',w:'7%'}},
  {{k:'项目发行状态',l:'发行状态',w:'7%'}},
  {{k:'跟进和投放状态',l:'跟进和投放状态',w:'14%'}},
  {{k:'交易所系统更新日期',l:'更新日期',w:'8%'}},
  {{k:'受理日期',l:'受理日期',w:'8%'}},
  {{k:'变更类型',l:'变更',w:'6%'}},
];

function renderTable(data, container, isGrouped){{
  const t=document.createElement('table');
  if(isGrouped) t.classList.add('grp-table');
  const thead=t.createTHead();
  const hr=thead.insertRow();
  COLS.forEach(c=>{{
    const th=document.createElement('th');
    th.style.width=c.w;
    th.textContent=c.l;
    if(sortCol===c.k) th.classList.add(sortDir===1?'sort-a':'sort-d');
    if(groupCol===c.k) th.classList.add('grp-active');
    th.title='点击排序 · Shift+点击按此列分组';
    th.addEventListener('click',e=>{{
      if(e.shiftKey){{groupCol=(groupCol===c.k?null:c.k);render();return;}}
      if(sortCol===c.k) sortDir*=-1; else{{sortCol=c.k;sortDir=1;}}
      render();
    }});
    hr.appendChild(th);
  }});
  const tbody=t.createTBody();
  data.forEach(r=>{{
    const tr=tbody.insertRow();
    tr.className='dr';
    COLS.forEach(c=>{{
      const td=tr.insertCell();
      td.style.width=c.w;
      if(c.k==='变更类型'){{td.innerHTML=badge(r[c.k]);}}
      else if(c.k==='跟进和投放状态'){{
        const sel=document.createElement('select');
        sel.className='fsel-inline';
        FOLLOW_OPTS.filter(Boolean).forEach(opt=>{{
          const o=document.createElement('option');
          o.value=o.textContent=opt;
          if(opt===r[c.k]) o.selected=true;
          sel.appendChild(o);
        }});
        const oldVal=r[c.k];
        sel.onchange=()=>updateFollow(sel,r._id,oldVal);
        sel.onclick=e=>e.stopPropagation();
        td.appendChild(sel);
      }} else {{
        td.textContent=r[c.k]||'';
        td.title=String(r[c.k]||'');
      }}
    }});
    const expTr=tbody.insertRow();
    expTr.className='exp';
    expTr.style.display='none';
    const expTd=expTr.insertCell();
    expTd.colSpan=COLS.length;
    expTd.innerHTML=`<div class="exp-inner">
      <div class="exp-item"><div class="el">备注</div><div class="ev">${{r['备注']||'—'}}</div></div>
      <div class="exp-item"><div class="el">变更详情</div><div class="ev">${{r['变更详情']||'—'}}</div></div>
      <div class="exp-item"><div class="el">预计发行时间</div><div class="ev">${{r['预计发行时间']||'—'}}</div></div>
    </div>`;
    tr.onclick=()=>{{expTr.style.display=expTr.style.display==='none'?'table-row':'none';}};
  }});
  container.appendChild(t);
}}

function render(){{
  const filtered=getFiltered();
  const sorted=sortData(filtered);
  document.getElementById('c-tot').textContent=sorted.length;
  document.getElementById('c-new').textContent=sorted.filter(r=>r['变更类型']==='新增').length;
  document.getElementById('c-chg').textContent=sorted.filter(r=>r['变更类型']==='状态变更').length;
  document.getElementById('c-amt').textContent=sorted.reduce((s,r)=>s+(parseFloat(r['拟发行金额(亿元)'])||0),0).toFixed(2);
  document.getElementById('sse-c').textContent='上交所 '+sorted.filter(r=>r['交易所']==='上交所').length+' 条';
  document.getElementById('szse-c').textContent='深交所 '+sorted.filter(r=>r['交易所']==='深交所').length+' 条';
  const dates=sorted.map(r=>r['交易所系统更新日期']).filter(Boolean).sort().reverse();
  if(dates.length) document.getElementById('upd-time').textContent='数据更新：'+dates[0];
  renderSummary(sorted);
  const area=document.getElementById('tbl-area');
  area.innerHTML='';
  if(groupCol){{
    const hint=document.createElement('div');
    hint.className='grp-hint';
    hint.textContent='按「'+groupCol+'」分组 · Shift+点击表头可取消分组';
    area.appendChild(hint);
    const groups={{}};
    sorted.forEach(r=>{{const k=r[groupCol]||'(空)';if(!groups[k])groups[k]=[];groups[k].push(r);}});
    const keys=Object.keys(groups).sort((a,b)=>a.localeCompare(b,'zh'));
    keys.forEach(g=>{{
      const amt=groups[g].reduce((s,r)=>s+(parseFloat(r['拟发行金额(亿元)'])||0),0);
      const gh=document.createElement('div');
      gh.className='grp-hdr';
      gh.innerHTML=`${{g}}<span class="cnt">${{groups[g].length}} 条</span><span class="amt">合计 ${{amt.toFixed(2)}} 亿元</span>`;
      area.appendChild(gh);
      renderTable(groups[g],area,true);
    }});
  }} else {{
    renderTable(sorted,area,false);
  }}
  document.getElementById('footer').textContent='共 '+sorted.length+' 条记录'+(groupCol?' · 按「'+groupCol+'」分组':'');
  renderLog();
}}

function resetFilters(){{
  ['f-exch','f-status','f-change','f-follow'].forEach(id=>document.getElementById(id).value='');
  sortCol=null;sortDir=1;groupCol=null;render();
}}

function toggleLog(){{document.getElementById('log-panel').classList.toggle('open');}}

function switchLogTab(tab,el){{
  curLogTab=tab;
  document.querySelectorAll('.log-tab').forEach(t=>t.classList.remove('act'));
  el.classList.add('act');
  renderLog();
}}

function renderLog(){{
  const now=new Date();
  const weekAgo=addDays(now,-7);
  const monthAgo=addDays(now,-30);
  let entries=changeLog;
  if(curLogTab==='week') entries=changeLog.filter(e=>new Date(e.date)>=weekAgo);
  else if(curLogTab==='month') entries=changeLog.filter(e=>new Date(e.date)>=monthAgo);
  const body=document.getElementById('log-body');
  if(!entries.length){{body.innerHTML='<div class="log-empty">暂无变更记录</div>';return;}}
  body.innerHTML=entries.map(e=>`
    <div class="log-entry">
      <div class="log-date">${{e.date}}</div>
      <div style="margin-bottom:2px">${{badge(e.type==='新增'?'新增':e.type.includes('状态')?'状态变更':'无变化')}} <span style="font-size:10px;color:var(--tm)">${{e.type}}</span></div>
      <div class="log-name">${{e.name}}</div>
      <div class="log-detail">${{e.detail}}</div>
    </div>`).join('');
}}

function dlLog(scope){{
  let entries=changeLog;
  if(scope==='week') entries=changeLog.filter(e=>new Date(e.date)>=addDays(new Date(),-7));
  const lines=['\\uFEFF日期,变更类型,债券名称,变更详情'];
  entries.forEach(e=>lines.push([e.date,e.type,'"'+e.name+'"','"'+e.detail+'"'].join(',')));
  const blob=new Blob([lines.join('\\n')],{{type:'text/csv;charset=utf-8'}});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);
  a.download='ABS变更日志_'+fmtD(new Date()).replace(/-/g,'')+'.csv';a.click();
}}

function dlExcel(){{
  const sorted=sortData(getFiltered());
  const h=['债券名称','交易所','计划管理人','拟发行金额(亿元)','交易所审批状态','项目发行状态','跟进和投放状态','交易所系统更新日期','受理日期','预计发行时间','备注','变更类型','变更详情'];
  const lines=['\\uFEFF'+h.join('\\t')];
  sorted.forEach(r=>lines.push(h.map(k=>r[k]||'').join('\\t')));
  const blob=new Blob([lines.join('\\n')],{{type:'text/tab-separated-values;charset=utf-8'}});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);
  a.download='持有型ABS项目汇总_'+fmtD(new Date()).replace(/-/g,'')+'.xls';a.click();
}}

initFilters();
render();
</script>

</body></html>'''

    # 使用字符串替换（避免与 CSS 花括号冲突）
    html_content = html_template
    html_content = html_content.replace('{json_data}', json_data)
    html_content = html_content.replace('{change_log_json}', change_log_json)
    html_content = html_content.replace('{latest_update}', latest_update)
    html_content = html_content.replace('{sse_count}', str(sse_count))
    html_content = html_content.replace('{szse_count}', str(szse_count))
    html_content = html_content.replace('{total_count}', str(total_count))
    html_content = html_content.replace('{new_count}', str(new_count))
    html_content = html_content.replace('{chg_count}', str(chg_count))
    html_content = html_content.replace('{total_amount:.2f}', f'{total_amount:.2f}')
    html_content = html_content.replace('{total_amount}', f'{total_amount:.2f}')

    # 写入文件
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"HTML Dashboard 已导出: {output_path}")
