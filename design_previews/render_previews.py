"""生成可独立打开的界面方案，并用本机 Chrome 导出预览图。"""

from html import escape
from pathlib import Path
import subprocess

ROOT = Path(__file__).resolve().parent
CHROME = Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe")


def icon(name, size=18):
    paths = {
        "search": '<circle cx="10.5" cy="10.5" r="6.5"/><path d="m16 16 4.5 4.5"/>',
        "folder": '<path d="M3 6a2 2 0 0 1 2-2h5l2 3h7a2 2 0 0 1 2 2v10H3z"/>',
        "chevron": '<path d="m8 10 4 4 4-4"/>',
        "star": '<path d="m12 3 2.8 5.7 6.2.9-4.5 4.4 1.1 6.2-5.6-3-5.6 3 1.1-6.2L3 9.6l6.2-.9z"/>',
        "settings": '<path d="m10 3-.8 3-3 .8L3.8 9l1.4 2.8-1.4 2.7 2.4 2.3 3 .7.8 3h3l.8-3 3-.7 2.4-2.3-1.4-2.7L19.2 9l-2.4-2.2-3-.8L13 3z"/><circle cx="11.5" cy="11.8" r="3"/>',
        "help": '<circle cx="12" cy="12" r="9"/><path d="M9.5 9a2.5 2.5 0 1 1 4 2l-1.5 1.5V14m0 3h.01"/>',
        "close": '<path d="m7 7 10 10M7 17 17 7"/>',
        "copy": '<rect x="8" y="8" width="12" height="12" rx="2"/><path d="M16 8V4H4v12h4"/>',
        "export": '<path d="M12 15V3m-4 4 4-4 4 4M4 13v7h16v-7"/>',
        "open": '<path d="M13 3h8v8m0-8L10 14M9 5H4v15h15v-5"/>',
        "filter": '<path d="M3 5h18l-7 8v6l-4 2v-8z"/>',
        "table": '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 9h18M9 9v11"/>',
        "grid": '<rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/>',
        "check": '<path d="m5 12 4 4L19 6"/>',
        "clock": '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
        "more": '<circle cx="5" cy="12" r=".8"/><circle cx="12" cy="12" r=".8"/><circle cx="19" cy="12" r=".8"/>',
    }
    return f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{paths[name]}</svg>'


CSS = r'''
*{box-sizing:border-box}html,body{margin:0;width:1600px;height:1000px;overflow:hidden}body{font-family:"Microsoft YaHei UI","Segoe UI",sans-serif;color:#223048;background:#edf1f7;font-size:13px;-webkit-font-smoothing:antialiased}button,input{font:inherit}button{border:0;cursor:default}svg{flex-shrink:0;vertical-align:middle}.stage{padding:30px 32px 0}.window{height:914px;border:1px solid #cdd6e3;border-radius:12px;background:#f6f8fb;overflow:hidden;box-shadow:0 16px 48px #334b6c18,0 2px 5px #334b6c0b}.titlebar{height:40px;display:flex;align-items:center;justify-content:space-between;padding-left:16px;background:#f4f6fa;border-bottom:1px solid #e5eaf1}.brand{display:flex;align-items:center;gap:10px;font-size:12px;font-weight:600}.brand img{height:24px;width:24px;object-fit:contain}.title-actions{height:100%;display:flex}.title-actions span{width:46px;display:grid;place-items:center;color:#627186;font:18px "Segoe UI"}.toolbar{background:#fff;padding:20px 24px 17px;border-bottom:1px solid #e6ebf3}.toolbar-heading{display:flex;align-items:center;justify-content:space-between;margin-bottom:18px}.heading{font-size:22px;font-weight:700;letter-spacing:.2px}.heading span{font-size:12px;font-weight:400;color:#8995a7;margin-left:14px}.aux{display:flex;gap:18px;align-items:center;color:#6d7b8e}.aux span{display:flex;align-items:center;gap:6px}.folder-row{display:flex;align-items:center;gap:10px;margin-bottom:14px}.label{color:#7b899c;width:62px;font-size:12px;flex-shrink:0}.folder-input{height:38px;display:flex;align-items:center;gap:10px;flex:1;border:1px solid #dce3ed;border-radius:7px;background:#fff;padding:0 12px;color:#43516a}.folder-input .path{flex:1;font-family:"Segoe UI",sans-serif;font-size:13px}.subtle{color:#9ba7b7}.outline{height:36px;padding:0 14px;display:inline-flex;align-items:center;gap:7px;background:#fff;border:1px solid #dce3ed;border-radius:7px;color:#4a5b74}.pinned{display:flex;align-items:center;gap:5px;color:#8b6e19;background:#fff9e8;padding:8px 10px;border-radius:6px;font-size:12px}.search-row{display:flex;align-items:center;gap:12px}.modes{display:flex;align-items:center;padding:4px;background:#f1f4f8;border-radius:8px;gap:2px;flex-shrink:0}.mode{padding:8px 16px;color:#6e7b90;font-size:12px;border-radius:5px}.mode.active{background:#fff;color:#245ad6;font-weight:600;box-shadow:0 1px 4px #162c4815}.search-input{height:42px;flex:1;display:flex;align-items:center;gap:10px;border:1.5px solid #8eacf0;box-shadow:0 0 0 3px #edf3ff;border-radius:7px;padding:0 13px;background:#fff}.search-input .term{flex:1;font-size:14px;color:#24324a}.kbd{font-family:"Segoe UI",sans-serif;font-size:10px;color:#98a3b4;background:#f5f7fa;padding:2px 5px;border:1px solid #e8ecf2;border-radius:4px}.primary{height:42px;min-width:108px;border-radius:7px;display:inline-flex;align-items:center;justify-content:center;gap:8px;background:#3468e4;color:#fff;font-weight:600;box-shadow:0 2px 3px #2753b91c}.chips{margin-top:15px;display:flex;align-items:center;gap:8px;height:27px}.chips-label{color:#929daf;font-size:12px;margin-right:2px}.chip{height:27px;display:flex;align-items:center;gap:10px;border:1px solid #dce6f9;border-radius:5px;background:#f2f6ff;color:#5676ab;padding:0 9px;font-size:11px}.chip.basic{background:#f5f7fa;border-color:#e6ebf1;color:#748197}.chip .x{color:#98afd7;font-size:14px}.link{color:#507bd0;font-size:12px}.reset{color:#94a0b0;font-size:11px;margin-left:4px}.content{padding:18px 20px 0;height:603px;display:flex;gap:16px}.result{border:1px solid #e2e8f1;border-radius:9px;background:#fff;overflow:hidden;flex:1;min-width:0;display:flex;flex-direction:column}.result-bar{height:65px;display:flex;align-items:center;justify-content:space-between;padding:0 18px;border-bottom:1px solid #edf0f5;flex-shrink:0}.result-name{font-size:15px;font-weight:650;display:flex;align-items:center;gap:9px}.count{font:600 11px "Segoe UI",sans-serif;color:#4774d3;background:#edf3ff;padding:3px 7px;border-radius:5px}.result-sub{font-size:11px;color:#99a4b4;margin-top:5px}.result-tools{display:flex;align-items:center;gap:12px}.local-search{display:flex;gap:8px;align-items:center;border:1px solid #e5eaf2;border-radius:6px;height:30px;padding:0 9px;color:#9aa5b5;font-size:11px}.tiny-button{display:flex;gap:6px;align-items:center;color:#77869b;font-size:11px}.density{border-left:1px solid #e6ebf2;padding-left:12px}.table-scroll{flex:1;overflow:hidden;position:relative}.columns{height:37px;display:grid;grid-template-columns:285px 90px 145px 1fr 64px;align-items:center;padding:0 18px;background:#fafbfd;border-bottom:1px solid #edf0f5;color:#8b98ab;font-size:11px}.group{height:39px;padding:0 18px;display:flex;align-items:center;gap:8px;border-bottom:1px solid #edf0f5;background:#fbfcff}.arrow{font:12px "Segoe UI",sans-serif;color:#8191a8}.excel{height:19px;width:17px;background:#e5f3ec;border-radius:3px;color:#318263;display:inline-grid;place-items:center;font:bold 11px "Segoe UI"}.group-name{font-size:12px;font-weight:650;color:#425573}.group .group-count{color:#9ba7b8;font-size:10px;margin-left:4px}.group .group-path{margin-left:auto;color:#a6b0be;font:11px "Segoe UI",sans-serif}.sheet{height:29px;padding-left:42px;display:flex;gap:8px;align-items:center;color:#74839a;background:#fff;font-size:11px;border-bottom:1px solid #f1f4f8}.dot-grid{width:12px;height:12px;border:1px solid #a4b5c9;border-radius:2px;background:linear-gradient(90deg,transparent 45%,#a4b5c9 46%,#a4b5c9 53%,transparent 54%),linear-gradient(transparent 45%,#a4b5c9 46%,#a4b5c9 53%,transparent 54%)}.data-row{height:38px;padding:0 18px;display:grid;grid-template-columns:285px 90px 145px 1fr 64px;align-items:center;border-bottom:1px solid #f0f3f8;font-size:12px;color:#637189}.data-row:nth-child(even){background:#fdfefe}.data-row .tree{padding-left:44px;display:flex;gap:9px;align-items:center;color:#8d9bae;font-size:11px}.cell{font:12px "Segoe UI",sans-serif;color:#59749e}.field{font:12px "Segoe UI",sans-serif;color:#71819a}.match-text{color:#4b5b72;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;padding-right:12px}mark{background:#fff0b7;color:#775c11;padding:1px 2px;border-radius:3px}.row-actions{display:flex;gap:12px;justify-content:flex-end;color:#a0aec0}.selected{background:#edf4ff!important;border-left:3px solid #4779e8;padding-left:15px}.selected .match-text{color:#345582}.selection-bar{height:39px;display:flex;align-items:center;justify-content:space-between;padding:0 18px;border-top:1px solid #dce7f8;background:#f5f8fe;color:#7085a8;font-size:11px}.selection-actions{display:flex;gap:20px;align-items:center}.selection-actions span{display:flex;gap:6px;align-items:center}.footer{height:39px;padding:0 24px;display:flex;align-items:center;justify-content:space-between;font-size:11px;color:#95a2b5}.footer-left{display:flex;gap:14px;align-items:center}.success{display:flex;align-items:center;gap:6px;color:#5b927e}.status-dot{height:6px;width:6px;border-radius:50%;background:#5eaa88}.divider{width:1px;height:12px;background:#d8e0eb}.caption{height:56px;display:flex;align-items:center;justify-content:space-between;color:#7f8da2;font-size:12px;padding:0 5px}.caption strong{color:#52667f;font-size:13px;font-weight:600}.caption .number{font:500 12px "Segoe UI";color:#9caabd;margin-right:10px}.caption-note{font-size:11px;color:#9da8b9}.inspector{width:390px;border:1px solid #e0e7f1;border-radius:9px;background:#fff;overflow:hidden;flex-shrink:0}.inspector-head{height:53px;padding:0 18px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #edf0f5;font-size:14px;font-weight:600}.inspector-head svg{color:#a0adbf}.inspector-body{padding:17px 18px}.meta-line{display:flex;align-items:center;gap:8px;font-size:11px;color:#7d8da6;line-height:25px}.meta-line .value{color:#536984;font-family:"Segoe UI","Microsoft YaHei UI",sans-serif}.meta-label{min-width:40px;color:#a0adbf}.address-badge{background:#eef4ff;border-radius:4px;color:#4d78c8;padding:0 7px;font:600 11px/21px "Segoe UI"}.preview-title{margin-top:20px;font-size:11px;color:#8797ae;display:flex;align-items:center;justify-content:space-between}.preview-text{border:1px solid #e5eaf2;border-radius:7px;background:#fafcff;margin-top:9px;padding:13px 12px;font-size:13px;line-height:2;color:#53657e;min-height:112px}.context-title{margin-top:16px;margin-bottom:10px;color:#8b9ab0;font-size:11px;display:flex;justify-content:space-between}.mini-grid{border:1px solid #e5eaf2;border-radius:6px;overflow:hidden;font:10px "Segoe UI","Microsoft YaHei UI",sans-serif}.mini-row{display:grid;grid-template-columns:34px 72px 92px 1fr;height:31px;align-items:center;border-bottom:1px solid #edf1f6}.mini-row:last-child{border-bottom:0}.mini-row>span{padding:0 7px;border-right:1px solid #edf1f6;height:100%;display:flex;align-items:center;white-space:nowrap;overflow:hidden;color:#7c8da5}.mini-row.head{background:#f7f9fc}.mini-row .focus-cell{background:#e9f1ff;color:#4473ca;box-shadow:inset 0 0 0 1px #779ded}.panel-actions{display:flex;gap:8px;margin-top:17px}.panel-actions .primary{flex:1;height:35px;font-size:12px;min-width:0}.panel-actions .outline{height:35px;font-size:12px;justify-content:center}.preview-result .columns,.preview-result .data-row{grid-template-columns:210px 72px 95px 1fr 25px}.preview-result .group-path{display:none}.preview-result .result-tools .local-search{width:155px}.preview-result .data-row .tree{padding-left:28px}.context-menu{position:absolute;left:440px;top:128px;width:224px;border:1px solid #dce3ee;background:#fff;border-radius:9px;padding:6px;box-shadow:0 12px 35px #243a5c27;z-index:3}.menu-item{display:flex;align-items:center;gap:9px;font-size:11px;height:34px;padding:0 9px;color:#5b6d86}.menu-item.active{background:#edf3ff;color:#3e6ac6;border-radius:5px}.menu-item .shortcut{margin-left:auto;font:10px "Segoe UI";color:#9ca9bd}.menu-sep{height:1px;background:#edf1f6;margin:5px 5px}.gallery-scroll{flex:1;overflow:hidden;padding:0 16px}.gallery-group{display:flex;align-items:center;gap:7px;height:43px;font-size:12px;color:#60748e}.gallery-group .folder-count{color:#a2afc0;font-size:10px;margin-left:auto}.gallery-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:13px}.file-card{border:1px solid #e5eaf2;border-radius:8px;background:#fff;padding:10px;position:relative}.file-card.active{border-color:#7199ec;background:#f7faff;box-shadow:0 0 0 2px #eaf1ff}.art{height:118px;border-radius:5px;position:relative;overflow:hidden;display:flex;align-items:center;justify-content:center;background:linear-gradient(140deg,var(--tone1),var(--tone2))}.art:before{content:"";position:absolute;inset:-20px;background:radial-gradient(ellipse at 70% 25%,#ffffff26,transparent 50%);transform:rotate(20deg)}.art svg{height:75px;width:75px;filter:drop-shadow(0 4px 12px #0004)}.art .spark{position:absolute;background:#fff;opacity:.45;height:3px;width:3px;border-radius:50%;left:22%;top:23%;box-shadow:94px 12px 0 #fff,33px 52px 0 #fff,115px 65px 0 #fff}.file-title{font-size:11px;color:#5f718a;margin-top:10px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.file-info{font:10px "Segoe UI";color:#a1aec0;display:flex;justify-content:space-between;margin-top:6px}.chosen{height:17px;width:17px;position:absolute;right:16px;top:16px;background:#4e7fe5;border-radius:4px;display:grid;place-items:center;color:#fff;z-index:1}.view-toggle{padding:3px;background:#f3f6fa;border-radius:6px;display:flex;gap:3px}.view-toggle span{display:flex;align-items:center;gap:4px;padding:5px 7px;font-size:10px;color:#8799b1;border-radius:4px}.view-toggle .active{background:#fff;color:#527acc;box-shadow:0 1px 3px #24394f13}.big-art{height:215px;margin-bottom:16px}.big-art svg{height:132px;width:132px}.file-heading{font-size:14px;font-weight:600;color:#4f6481;margin-bottom:11px}.image-meta{display:grid;grid-template-columns:1fr 1fr;gap:12px 10px;padding:12px 0;border-bottom:1px solid #edf1f6}.image-meta small{font-size:10px;color:#a0adc0;display:block;margin-bottom:5px}.image-meta span{font:12px "Segoe UI","Microsoft YaHei UI";color:#6d809b}.path-line{font:11px/1.8 "Segoe UI","Microsoft YaHei UI";color:#8293ac;background:#f7f9fc;border:1px solid #e9edf4;border-radius:5px;padding:8px 10px;margin-top:8px;word-break:break-all}.gallery-result .result-sub{margin-top:4px}.preview-hint{margin-top:13px;font-size:10px;color:#a2aec1;line-height:1.7}
'''


def toolbar(mode, term, images=False):
    tabs = ''.join(f'<span class="mode {"active" if label == mode else ""}">{label}</span>' for label in ('工作表', '单元格', '文件', 'JSON 检查'))
    chips = ('<span class="chip">后缀：png、jpg、webp <span class="x">×</span></span>' if images else
             '<span class="chip">仅启用表格 <span class="x">×</span></span><span class="chip">仅启用数据 <span class="x">×</span></span>')
    return f'''<div class="toolbar"><div class="toolbar-heading"><div class="heading">查找配置与文件<span>常用目录已就绪</span></div><div class="aux"><span>{icon('settings',16)} 设置</span><span>{icon('help',16)} 使用说明</span></div></div>
    <div class="folder-row"><span class="label">工作目录</span><div class="folder-input">{icon('folder',17)}<span class="path">D:\\xy\\data\\dev</span>{icon('chevron',15)}</div><button class="outline">选择目录</button><span class="pinned">{icon('star',14)} 已固定</span></div>
    <div class="search-row"><div class="modes">{tabs}</div><div class="search-input">{icon('search',18)}<span class="term">{term}</span><span class="kbd">Enter</span><span class="subtle">{icon('close',14)}</span></div><button class="primary">{icon('search',17)} 检索</button></div>
    <div class="chips"><span class="chips-label">当前条件</span><span class="chip basic">包含匹配 {icon('chevron',12)}</span>{chips}<span class="link">＋ 更多条件</span><span class="reset">重置</span></div></div>'''


DATA = [
    ('SkillConfig.xlsx', '技能表', '24', [
        ('D128', 'SkillDesc', '技能展示：释放技能时播放主表现特效'),
        ('D132', 'SkillDesc', '在技能展示阶段隐藏角色武器'),
        ('F205', 'ShowEffect', '技能展示结束后恢复角色待机状态'),
        ('D216', 'SkillDesc', '技能展示期间镜头缓慢拉近'),
    ]),
    ('SkillPresentation.xlsx', '技能表现', '18', [
        ('B42', 'Name', '主角技能展示表现'),
        ('E56', 'Description', '技能展示镜头衔接与过渡'),
        ('C73', 'EffectName', '技能展示 · 强化效果'),
    ]),
    ('CharacterConfig.xlsx', '角色表', '12', [
        ('G86', 'ShowDesc', '角色详情中的技能展示入口'),
        ('G91', 'ShowDesc', '点击后打开技能展示界面'),
    ]),
]


def result_table(preview=False):
    head = f'''<div class="result-bar"><div><div class="result-name">检索结果 <span class="count">86 条</span></div><div class="result-sub">8 个工作簿 · 12 张工作表</div></div><div class="result-tools"><div class="local-search">{icon('search',14)} 在结果中查找 <span class="kbd">Ctrl F</span></div>{'' if preview else f'<span class="tiny-button density">{icon("table",14)} 紧凑 {icon("chevron",12)}</span>'}<span class="tiny-button">{icon('export',15)} 导出</span></div></div>'''
    body = '<div class="columns"><span>工作簿 / 工作表</span><span>位置</span><span>字段</span><span>命中内容</span><span></span></div>'
    for book, sheet, count, rows in DATA:
        body += f'<div class="group"><span class="arrow">⌄</span><span class="excel">X</span><span class="group-name">{book}</span><span class="group-count">{count} 条</span><span class="group-path">data / combat / skill</span></div><div class="sheet"><span class="arrow">⌄</span><span class="dot-grid"></span>{sheet}</div>'
        for i, (cell, field, value) in enumerate(rows):
            selected = book == 'SkillConfig.xlsx' and i == 0
            highlighted = escape(value).replace('技能展示', '<mark>技能展示</mark>')
            body += f'''<div class="data-row {'selected' if selected else ''}"><span class="tree">↳ <span>匹配单元格</span></span><span class="cell">{cell}</span><span class="field">{field}</span><span class="match-text">{highlighted}</span><span class="row-actions">{icon('copy',14)}{icon('more',14)}</span></div>'''
    if preview:
        body += f'''<div class="context-menu"><div class="menu-item active">{icon('copy',14)} 复制内容 <span class="shortcut">Ctrl+C</span></div><div class="menu-item">{icon('table',14)} 复制单元格地址</div><div class="menu-item">{icon('folder',14)} 复制文件路径</div><div class="menu-sep"></div><div class="menu-item">{icon('open',14)} 在 Excel 中定位 <span class="shortcut">Enter</span></div><div class="menu-item">{icon('folder',14)} 打开所在目录</div><div class="menu-sep"></div><div class="menu-item">{icon('filter',14)} 只看此工作表</div></div>'''
    return f'''<div class="result {'preview-result' if preview else ''}">{head}<div class="table-scroll">{body}</div><div class="selection-bar"><span>已选择 1 个单元格</span><div class="selection-actions"><span>{icon('copy',13)} 复制内容</span><span>{icon('open',13)} 在 Excel 中定位</span></div></div></div>'''


def cell_inspector():
    context = '<div class="mini-row head"><span></span><span>A · ID</span><span>B · 名称</span><span>D · 描述</span></div>'
    for row, ident, name in [(126,'83000','普通攻击'),(127,'83001','蓄力攻击'),(128,'83002','强化技能'),(129,'83003','特殊技能')]:
        context += f'<div class="mini-row"><span>{row}</span><span>{ident}</span><span>{name}</span><span class="{"focus-cell" if row == 128 else ""}">{"技能展示…" if row == 128 else "角色动作…"}</span></div>'
    return f'''<aside class="inspector"><div class="inspector-head"><span>单元格预览</span>{icon('close',16)}</div><div class="inspector-body"><div class="meta-line"><span class="meta-label">工作簿</span><span class="excel">X</span><span class="value">SkillConfig.xlsx</span></div><div class="meta-line"><span class="meta-label">工作表</span><span class="value">技能表</span><span class="subtle">/</span><span class="address-badge">D128</span></div><div class="meta-line"><span class="meta-label">字段</span><span class="value">SkillDesc</span><span class="subtle">· 文本</span></div><div class="preview-title">完整内容<span>{icon('copy',14)} 复制</span></div><div class="preview-text"><mark>技能展示</mark>：释放技能时播放主表现特效。<br>播放期间隐藏武器，并在结束后恢复角色待机状态。</div><div class="context-title"><span>相邻数据</span><span>第 126–129 行</span></div><div class="mini-grid">{context}</div><div class="panel-actions"><button class="primary">{icon('open',14)} 在 Excel 中定位</button><button class="outline">{icon('folder',14)} 目录</button></div><div class="preview-hint">单击预览 · 双击定位 · Ctrl+C 复制内容</div></div></aside>'''


ARTS = {
    'fire': ('#352956','#191f3e','<path d="M52 9c7 18-3 22 4 29 3-9 10-10 12-18 16 23 28 35 23 55-4 16-17 27-36 27-23 0-40-14-40-35 0-22 22-31 26-48 8 7 6 16 4 21 13-9 10-19 7-31z" fill="#ffbb66"/><path d="M55 48c2 15-12 20-12 31 0 11 7 17 16 17 13 0 21-12 16-22-3 6-6 8-8 10 2-14-5-25-12-36z" fill="#fff0bc"/>'),
    'ice': ('#245579','#17324f','<g stroke="#b8eeff" stroke-width="5" stroke-linecap="round"><path d="M60 15v90M21 38l78 44M21 82l78-44M46 25l14 14 14-14M46 95l14-14 14 14M24 53l20-5-5-20M96 67l-20 5 5 20M39 92l5-20-20-5M81 28l-5 20 20 5"/></g>'),
    'lightning': ('#473573','#28234b','<path d="M68 10 29 65h27l-8 45 43-62H65z" fill="#f2de96"/>'),
    'shield': ('#38635c','#1e3a45','<path d="m60 15 35 13v29c0 24-15 42-35 52-20-10-35-28-35-52V28z" fill="#94d0b5"/><path d="m60 29 22 8v20c0 16-9 29-22 37-13-8-22-21-22-37V37z" fill="#466e64"/><path d="m46 60 10 10 19-24" fill="none" stroke="#c8eed9" stroke-width="7" stroke-linecap="round" stroke-linejoin="round"/>'),
}


def art(kind, big=False):
    c1,c2,shape=ARTS[kind]
    return f'<div class="art {"big-art" if big else ""}" style="--tone1:{c1};--tone2:{c2}"><span class="spark"></span><svg viewBox="0 0 120 120">{shape}</svg></div>'


def gallery():
    files = [('技能_火焰.png','fire','128'),('技能_冰霜.png','ice','116'),('技能_雷击.png','lightning','94'),('技能_守护.png','shield','102'),('技能_火焰_02.png','fire','136'),('技能_冰霜_02.png','ice','118'),('技能_雷击_02.png','lightning','96'),('技能_守护_02.png','shield','109')]
    cards=''
    for i,(name,kind,kb) in enumerate(files):
        cards += f'<div class="file-card {"active" if i == 0 else ""}">{f"<span class=chosen>{icon('check',12)}</span>" if i == 0 else ""}{art(kind)}<div class="file-title">{name}</div><div class="file-info"><span>512 × 512</span><span>{kb} KB</span></div></div>'
    head=f'<div class="result-bar"><div><div class="result-name">文件结果 <span class="count">48 项</span></div><div class="result-sub">4 个文件夹 · 图片文件</div></div><div class="result-tools"><div class="local-search">{icon("search",14)} 筛选当前结果</div><div class="view-toggle"><span>{icon("table",13)} 表格</span><span class="active">{icon("grid",13)} 缩略图</span></div></div></div>'
    return f'<div class="result gallery-result">{head}<div class="gallery-scroll"><div class="gallery-group">{icon("folder",15)} effects / skill<span class="folder-count">16 个文件</span></div><div class="gallery-grid">{cards}</div><div class="gallery-group">{icon("folder",15)} ui / skill / icons<span class="folder-count">12 个文件</span></div></div><div class="selection-bar"><span>已选择 1 个文件</span><div class="selection-actions"><span>{icon("copy",13)} 复制路径</span><span>{icon("folder",13)} 打开目录</span></div></div></div>'


def image_inspector():
    return f'''<aside class="inspector"><div class="inspector-head"><span>文件预览</span>{icon('close',16)}</div><div class="inspector-body">{art('fire',True)}<div class="file-heading">技能_火焰.png</div><div class="image-meta"><div><small>文件类型</small><span>PNG 图像</span></div><div><small>图片尺寸</small><span>512 × 512</span></div><div><small>文件大小</small><span>128 KB</span></div><div><small>修改日期</small><span>2026-10-20</span></div></div><div class="preview-title">所在目录<span>{icon('copy',13)} 复制</span></div><div class="path-line">D:\\xy\\data\\dev\\effects\\skill</div><div class="panel-actions"><button class="primary">{icon('open',14)} 打开文件</button><button class="outline">{icon('folder',14)} 打开目录</button></div></div></aside>'''


def page(number, title, description, mode, content, term='技能展示', images=False):
    scanned = '已扫描 126 个文件' if images else '已扫描 437 个工作簿'
    elapsed = '0.4 秒' if images else '3.5 秒'
    footer=f'<div class="footer"><div class="footer-left"><span class="success"><span class="status-dot"></span>检索完成</span><span class="divider"></span><span>{scanned}</span><span>{icon("clock",12)} {elapsed}</span></div><span>仅检索磁盘上已保存的内容</span></div>'
    return f'''<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>{title}</title><style>{CSS}</style></head><body><div class="stage"><div class="window"><div class="titlebar"><div class="brand"><img src="../assets/nailong_search_master.ico">奶龙检索大师</div><div class="title-actions"><span>−</span><span>□</span><span>×</span></div></div>{toolbar(mode,term,images)}<div class="content">{content}</div>{footer}</div><div class="caption"><span><span class="number">{number}</span><strong>{title}</strong><span style="margin-left:16px">{description}</span></span><span class="caption-note">界面方案预览 · 示例数据</span></div></div></body></html>'''


def main():
    pages = [
        ('01-main-search', page('01','主检索界面','生效条件始终可见，结果更紧凑','单元格',result_table())),
        ('02-cell-preview', page('02','单元格预览与快捷操作','先看完整内容，再决定是否打开 Excel','单元格',result_table(True)+cell_inspector())),
        ('03-file-gallery', page('03','图片缩略图模式','按目录浏览，单击即可预览文件','文件',gallery()+image_inspector(),term='技能',images=True)),
    ]
    for stem, html in pages:
        html_path = ROOT / f'{stem}.html'
        image_path = ROOT / f'{stem}.png'
        html_path.write_text(html, encoding='utf-8')
        profile = ROOT.parent / 'build' / f'preview-chrome-{stem}'
        result = subprocess.run([
            str(CHROME), '--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check',
            '--hide-scrollbars', '--force-device-scale-factor=1', '--window-size=1600,1000',
            f'--user-data-dir={profile}', f'--screenshot={image_path}', '--virtual-time-budget=1500', html_path.as_uri(),
        ], capture_output=True, text=True, timeout=35, creationflags=subprocess.CREATE_NO_WINDOW)
        if result.returncode != 0 or not image_path.exists():
            raise RuntimeError(result.stderr[-2500:])
        print(image_path, flush=True)


if __name__ == '__main__':
    main()
