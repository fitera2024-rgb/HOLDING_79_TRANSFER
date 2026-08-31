from __future__ import annotations

import tools.windows_web_app as base


PAGE = r'''<!doctype html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>ФИТЭРА · Перенос 79.2 / 79.3</title>
<style>
:root{
  --green:#197a43;
  --green-dark:#0f6434;
  --green-bright:#31a852;
  --green-soft:#eef7ee;
  --green-line:#b8d6be;
  --navy:#173f67;
  --ink:#17332d;
  --muted:#6e7f79;
  --line:#d6dfd8;
  --bg:#f4f7f4;
  --white:#fff;
  --warn:#9d6b00;
  --red:#b42318;
  --shadow:0 8px 26px rgba(36,72,48,.06);
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);font-family:"Segoe UI",Arial,sans-serif;color:var(--ink)}
body:before{content:"";display:block;height:6px;background:#58b947}
.hero{background:#fff;border-bottom:1px solid #dce4dd}
.hero-inner{max-width:1330px;margin:0 auto;padding:55px 40px 30px;display:flex;justify-content:space-between;gap:38px;align-items:flex-start}
.brand{display:flex;gap:28px;align-items:center;min-width:0}
.mark{position:relative;width:88px;height:88px;flex:0 0 88px}
.mark .disc{position:absolute;width:72px;height:72px;left:0;top:8px;border-radius:50%;background:conic-gradient(from 0deg,#a7e3b9 0 25%,#5dca7c 25% 50%,#2eae59 50% 75%,#d9f0df 75% 100%);clip-path:polygon(0 0,62% 0,62% 42%,100% 42%,100% 100%,0 100%)}
.mark .dot{position:absolute;width:28px;height:28px;left:34px;top:30px;border-radius:50%;background:#0bb13e;border:4px solid #fff;z-index:2}
.mark .arrow{position:absolute;width:0;height:0;border-top:12px solid transparent;border-bottom:12px solid transparent;border-left:40px solid #08a93a;left:56px;top:15px;transform:rotate(-40deg)}
.eyebrow{font-weight:800;color:var(--green-dark);letter-spacing:.05em;font-size:13px;margin-bottom:12px}
h1{margin:0;color:var(--navy);font-size:46px;line-height:1.08;font-weight:800;letter-spacing:-.02em}
.subtitle{margin-top:15px;color:var(--green-dark);font-size:18px;font-weight:750}
.status-box{width:258px;background:#f4faf4;border:1px solid #c4ddc8;border-radius:14px;padding:20px 22px;text-align:center;color:var(--green-dark);box-shadow:0 2px 4px rgba(0,0,0,.01)}
.status-box b{display:block;font-size:16px;margin-bottom:7px}.status-box small{display:block;color:#75827d;margin-bottom:12px}.status-box div{font-weight:800;line-height:1.9}
.main{max-width:1330px;margin:0 auto;padding:26px 40px 48px}
.ready{background:#eaf5e8;border:1px solid #b9d7b9;border-left:6px solid #a4d59f;border-radius:10px;padding:17px 20px;color:#0e6135;font-weight:650;margin-bottom:20px}
.panel{background:#fff;border:1px solid var(--line);border-radius:12px;box-shadow:var(--shadow);padding:24px;margin-bottom:18px}
.panel-title{display:flex;align-items:center;gap:12px;color:var(--green-dark);font-size:20px;font-weight:800;margin-bottom:22px}
.step{width:34px;height:34px;border-radius:4px;background:var(--green);color:#fff;display:grid;place-items:center;font-weight:900;font-size:17px}
.grid{display:grid;grid-template-columns:minmax(0,1fr) 300px;gap:18px}
.fieldbox{border:1.5px dashed #93bfa0;border-radius:10px;padding:22px;background:#fbfdfb}
.fieldbox label{display:block;font-size:17px;font-weight:800;color:#102d22;margin-bottom:8px}
.fieldbox .hint{font-size:12px;color:#728079;margin-bottom:10px}
input{width:100%;min-height:45px;border:1px solid #b9d1be;border-radius:7px;background:#fff;font:inherit;color:#20342c;padding:9px 11px;outline:none}
input[type=file]{padding:7px 9px}
input:focus{border-color:var(--green);box-shadow:0 0 0 3px rgba(25,122,67,.10)}
.actions{display:flex;align-items:center;gap:16px;margin-top:16px;flex-wrap:wrap}
button,a.btn{font:inherit;font-weight:800;border:0;border-radius:7px;text-decoration:none;cursor:pointer}
.primary{background:var(--green);color:#fff;padding:13px 22px;box-shadow:0 4px 10px rgba(25,122,67,.15)}
.primary:hover{background:var(--green-dark)}
.primary:disabled{opacity:.55;cursor:default}
.secondary{background:#f0f7f0;color:var(--green-dark);border:1px solid #bdd6c2;padding:10px 16px}
.status{color:#61716a;font-size:13px}
.progress{display:none;height:7px;background:#e9f0ea;border-radius:99px;overflow:hidden;margin-top:15px}
.progress.on{display:block}.progress span{display:block;height:100%;width:34%;background:var(--green);border-radius:99px;animation:p 1.1s ease-in-out infinite}@keyframes p{from{transform:translateX(-110%)}to{transform:translateX(320%)}}
.msg{display:none;margin-top:14px;padding:13px 15px;border-radius:8px;font-size:13px}.msg.err{display:block;background:#fff2f0;color:var(--red);border:1px solid #ffd2cc}
.result{display:none}.result.on{display:block}
.metrics{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:18px}.metric{border:1px solid #cfe0d2;background:#f8fbf8;border-radius:9px;padding:15px}.metric b{display:block;color:var(--green-dark);font-size:25px;margin-bottom:4px}.metric span{font-size:12px;color:#6e7d76}
.files{border:1px solid var(--line);border-radius:9px;overflow:hidden;background:#fff}.row{display:flex;justify-content:space-between;gap:18px;align-items:center;padding:13px 15px;border-top:1px solid #e2e9e3}.row:first-child{border-top:0}.name{min-width:0}.name b{display:block;font-size:13px;color:#1f342b;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.name small{color:#748179}.badge{display:inline-block;margin-left:8px;padding:3px 7px;border-radius:999px;background:#fff4cc;color:var(--warn);font-size:10px;font-weight:900}.download{background:#eef7ef;color:var(--green-dark);border:1px solid #c6ddcb;padding:8px 12px;white-space:nowrap}.all{display:flex;justify-content:flex-end;margin-top:15px}
.footer{display:flex;justify-content:space-between;align-items:center;gap:15px;color:#849089;font-size:11px;margin-top:18px}
@media(max-width:820px){.hero-inner{padding:32px 18px 24px;flex-direction:column}.main{padding:20px 18px 40px}.brand{align-items:flex-start}.mark{display:none}h1{font-size:33px}.status-box{width:100%}.grid{grid-template-columns:1fr}.metrics{grid-template-columns:repeat(2,1fr)}}
</style>
</head>
<body>
<header class="hero"><div class="hero-inner">
  <div class="brand"><div class="mark"><div class="disc"></div><div class="dot"></div><div class="arrow"></div></div><div>
    <div class="eyebrow">ООО «ФИТЭРА» · HOLDING 79 TRANSFER</div>
    <h1>Перенос остатков 79.2 / 79.3</h1>
    <div class="subtitle">Формирование проводок 79.1 по организациям</div>
  </div></div>
  <div class="status-box"><b>REPORT_ONLY</b><small>Запись в 1С отключена</small><div>Локальная обработка</div><div>Расчёты неизменны</div><div>Файлы по организациям</div></div>
</div></header>
<main class="main">
  <div class="ready">Сервис готов к расчёту. Запись в 1С отключена.</div>
  <form id="f">
    <section class="panel">
      <div class="panel-title"><span class="step">1</span> Исходные данные</div>
      <div class="grid">
        <div class="fieldbox"><label>ОСВ по счёту 79.2 / 79.3</label><div class="hint">Excel-файл .xlsx</div><input name="source" type="file" accept=".xlsx" required></div>
        <div class="fieldbox"><label>Период расчёта</label><div class="hint">Выберите месяц</div><input name="period" type="month" value="{{MONTH}}" required></div>
      </div>
      <div class="actions"><button class="primary" id="go">Сформировать файлы</button><div class="status" id="status">Готово к запуску</div></div>
      <div class="progress" id="prog"><span></span></div><div class="msg" id="msg"></div>
    </section>
  </form>
  <section class="panel result" id="res">
    <div class="panel-title"><span class="step">2</span> Результат расчёта</div>
    <div id="per" style="color:#6d7d75;font-size:13px;margin:-10px 0 18px"></div>
    <div class="metrics"><div class="metric"><b id="src">0</b><span>остатков к переносу</span></div><div class="metric"><b id="post">0</b><span>проводок</span></div><div class="metric"><b id="books">0</b><span>файлов организаций</span></div><div class="metric"><b id="sp">0</b><span>файлов СПОРНО</span></div></div>
    <div class="files" id="files"></div><div class="all"><a class="btn primary" id="zip">Скачать всё ZIP</a></div>
  </section>
  <div class="footer"><span>Тестовый интерфейс ФИТЭРА · main {{SHA}}</span><button class="secondary" id="close" type="button">Закрыть программу</button></div>
</main>
<script>
const f=document.getElementById('f'),go=document.getElementById('go'),prog=document.getElementById('prog'),msg=document.getElementById('msg'),res=document.getElementById('res'),status=document.getElementById('status');
function esc(s){const d=document.createElement('div');d.textContent=s??'';return d.innerHTML}
f.addEventListener('submit',async e=>{e.preventDefault();msg.className='msg';msg.textContent='';res.classList.remove('on');go.disabled=true;prog.classList.add('on');status.textContent='Обработка ОСВ…';try{const r=await fetch('/api/run',{method:'POST',body:new FormData(f)}),d=await r.json();if(!r.ok||!d.ok)throw new Error(d.error||'Не удалось сформировать файлы.');document.getElementById('src').textContent=d.source_rows;document.getElementById('post').textContent=d.posting_rows;document.getElementById('books').textContent=d.workbooks.length;document.getElementById('sp').textContent=d.disputed_workbooks;document.getElementById('per').textContent='Период: '+d.period_label;document.getElementById('zip').href=d.zip_url;document.getElementById('files').innerHTML=d.workbooks.map(x=>`<div class="row"><div class="name"><b>${esc(x.name)}${x.disputed?'<span class="badge">СПОРНО</span>':''}</b><small>${x.rows} проводок</small></div><a class="btn download" href="${x.url}">Скачать</a></div>`).join('');res.classList.add('on');status.textContent='Файлы сформированы';res.scrollIntoView({behavior:'smooth'})}catch(x){msg.textContent=x.message||String(x);msg.className='msg err';status.textContent='Файлы не сформированы'}finally{prog.classList.remove('on');go.disabled=false}});
document.getElementById('close').onclick=async()=>{if(!confirm('Закрыть программу?'))return;try{await fetch('/api/shutdown',{method:'POST'})}catch(_){ }document.body.innerHTML='<div style="font:16px Segoe UI;padding:40px">Программа закрыта. Эту вкладку можно закрыть.</div>'};
</script>
</body></html>'''

base.PAGE = PAGE

if __name__ == "__main__":
    base.main()
