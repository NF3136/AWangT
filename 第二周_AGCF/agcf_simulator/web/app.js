'use strict';
const $ = id => document.getElementById(id);
const names = ['解析输入信令','读取中继配置','查找用户身份','确认注册路由','补充身份头域','处理计费关联','生成输出与下一跳'];
const labels = {real:'真实配置',mock:'测试数据',assumption:'演示约定',protocol:'转发规则',input:'输入报文',validation:'校验失败'};
let bootstrap, result = null, shown = 0, busy = false;
const escapeHTML = value => String(value).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function chip(origin) { return `<span class="chip ${origin}">${labels[origin] || origin}</span>`; }
function notice(message='', success=false) { $('notice').textContent = message; $('notice').hidden = !message; $('notice').classList.toggle('success',success); }
async function post(path, payload) { const res = await fetch(path, {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)}); const data = await res.json(); if (!res.ok) throw Error(data.error || '请求失败'); return data; }
function invalidate() { result=null;shown=0;$('export').disabled=true;$('step').textContent='单步查看';render(); }
function setBusy(value) { busy=value; ['run','step','reset','scenario','invite','profiles'].forEach(id => $(id).disabled=value); }
function resetSample() {
  if (!bootstrap) return;
  const type=$('scenario').value, config=structuredClone(bootstrap.profiles);
  let message=bootstrap.sample;
  const hints={normal:'PAI + PCV，路由至模拟注册的 S-CSCF',replace:'已有 PAI 按人工配置映射替换，观察报文差异',unregistered:'注册记录缺失时停止，不擅自选择其他出口',unknown:'找不到主叫身份时停止，不从 IOI 猜测身份',invalid:'只处理初始 INVITE；其他方法明确停止',pcv:'已有有效 PCV 时按演示策略保留'};
  if(type==='replace') message=message.replace('Content-Length:', 'P-Asserted-Identity: <sip:old-user@ims.example.invalid>\r\nContent-Length:');
  if(type==='unregistered') {config.users[0].registered=false;config.users[0].scscf='';}
  if(type==='unknown') message=message.replaceAll('+862132010323','+862199999999');
  if(type==='invalid') message=message.replaceAll('INVITE','BYE');
  if(type==='pcv') message=message.replace('Content-Length: 0','P-Charging-Vector: icid-value=demo-existing-call;icid-generated-at=192.0.2.20\r\nContent-Length: 0');
  $('invite').value=message; $('profiles').value=JSON.stringify(config,null,2);$('scenario-note').textContent=hints[type]; notice(); invalidate();
}
async function run(all) {
  if(busy || !bootstrap) return;
  notice();
  if(!all && result && shown < result.steps.length) { shown++;render();return; }
  try {
    const profiles=JSON.parse($('profiles').value);
    setBusy(true);
    result=await post('/api/simulate',{invite:$('invite').value,profiles,trunk:'38'});
    shown=all?result.steps.length:1;
    render();
  } catch(err) {invalidate();notice(err instanceof SyntaxError?'测试配置不是有效 JSON，请检查逗号与引号。':err.message);}
  finally {setBusy(false);}
}
function render() {
  const total=result?result.steps.length:names.length;
  const complete=result&&shown>=total;
  $('progress-label').textContent = result?`${shown} / ${total} 步`:'待运行';
  $('progress-fill').style.width=`${shown/total*100}%`;
  $('step').textContent=result&&shown<total?'下一步 →':'单步查看';
  const rows=result?result.steps:names.map(title=>({title}));
  $('flow').innerHTML=rows.map((s,i)=>{
    const visible=result&&i<shown;
    return `<div class="step-row ${visible?s.status:'pending'}"><span class="step-dot">${visible?(s.status==='error'?'!':'✓'):i+1}</span>${visible?`<details class="step-card"><summary><span class="step-title">${escapeHTML(s.title)}${chip(s.origin)}</span><p>${escapeHTML(s.detail)}</p></summary><pre class="step-data">${escapeHTML(JSON.stringify({依据:labels[s.origin],数据:s.data},null,2))}</pre></details>`:`<div class="step-card"><div class="step-title">${escapeHTML(s.title)}</div><p>等待处理</p></div>`}</div>`;
  }).join('');
  $('empty-output').hidden=!!complete;$('result-content').hidden=!complete;
  $('result-status').className='chip '+(complete?(result.ok?'real':'validation'):'neutral');
  $('result-status').textContent=complete?(result.ok?'模拟完成':'已停止'):(result?'逐步查看':'待运行');
  $('export').disabled=!complete;
  if(!complete) return;
  $('next-hop').className='next-hop'+(result.ok?'':' error');
  $('next-hop').innerHTML=result.ok?`<small>下一跳 · 模拟注册记录 ${chip('mock')}</small><strong>${escapeHTML(result.next_hop)}</strong>`:`<small>未生成可转发的输出</small><strong>${escapeHTML(result.error)}</strong>`;
  $('output-code').textContent=result.output || '处理已停止。没有输出 INVITE，也没有执行任何网络发送。';
  $('copy').disabled=!result.ok;
  $('change-count').textContent=result.changes.length?`(${result.changes.length})`:'';
  $('output-changes').innerHTML=(result.ok?'':'<div class="notes">以下仅为停止前的处理中间变化，不是最终输出。</div>')+result.changes.map(c=>`<article class="change"><header><strong>${escapeHTML(c.field)}</strong>${chip(c.origin)}</header><code class="before">− ${escapeHTML(c.before.length?c.before.join('\n'):'未携带')}</code><code class="after">+ ${escapeHTML(c.after.length?c.after.join('\n'):'不携带')}</code><p>${escapeHTML(c.reason)}</p></article>`).join('');
  $('output-notes').innerHTML=`<div class="notes"><strong>本次结果的适用边界</strong><ul>${[...result.assumptions,...result.warnings].map(t=>`<li>${escapeHTML(t)}</li>`).join('')}</ul><strong>真实数据来源</strong><p>${escapeHTML(result.source_file)}</p><p>本版生成 SIP 文本与路由元数据，不执行真实通信。每次新模拟使用新的 ICID 和 Via branch；尚未实现重传事务复用。</p></div>`;
}
function switchTabs(button, attr, panels) {
  document.querySelectorAll(`[${attr}]`).forEach(b=>{b.classList.toggle('active',b===button);b.setAttribute('aria-selected',String(b===button));});
  Object.entries(panels).forEach(([name,id])=>$(id).hidden=button.getAttribute(attr)!==name);
}
async function lookup(event) {
  event?.preventDefault();
  try {
    const data=await post('/api/explore',{number:$('number-query').value.trim()});
    if(!data.ok) {$('lookup-result').textContent=data.error;return;}
    const r=data.rule;
    $('lookup-result').innerHTML=`<strong>最长前缀：${escapeHTML(r.pfx)} · 号长 ${escapeHTML(r.minl)}–${escapeHTML(r.maxl)} · RSC ${escapeHTML(r.rsc)}</strong><div>主叫分析标志：${escapeHTML(r.clraf || '导出未显式给值')}；路由关联 ${data.routes.length} 条。候选子路由顺序 7 → 5，尚未检查运行状态。</div><ul>${data.warnings.map(t=>`<li>${escapeHTML(t)}</li>`).join('')}</ul>`;
  } catch(err) {$('lookup-result').textContent=err.message;}
}
function renderEvidence() {
  const s=bootstrap.snapshot;
  $('source-file').textContent=s.source_file;
  $('provenance').innerHTML=s.provenance.map(p=>`<article><h3>${escapeHTML(p.file)}</h3><p>${Object.entries(p.counts).map(([t,n])=>`${escapeHTML(t)}: ${n}`).join(' · ')}</p><details><summary>文件 SHA-256</summary>${escapeHTML(p.sha256)}</details></article>`).join('')+`<ul>${s.limitations.map(t=>`<li>${escapeHTML(t)}</li>`).join('')}</ul>`;
  $('table-select').innerHTML=Object.entries(s.tables).map(([t,r])=>`<option value="${escapeHTML(t)}">${escapeHTML(t)} (${r.length})</option>`).join('')+'<option value="identity_samples">ASBR 真实身份样本（独立来源）</option>';
  renderTable();
}
function renderTable() {$('evidence-json').textContent=JSON.stringify($('table-select').value==='identity_samples'?bootstrap.snapshot.identity_samples:bootstrap.snapshot.tables[$('table-select').value],null,2);}
document.querySelectorAll('[data-page]').forEach(b=>b.addEventListener('click',()=>{document.querySelectorAll('[data-page]').forEach(n=>n.classList.toggle('active',n===b));$('sim-page').hidden=b.dataset.page!=='sim';$('evidence-page').hidden=b.dataset.page!=='evidence';if(b.dataset.page==='evidence'&&bootstrap)lookup();}));
document.querySelectorAll('[data-input]').forEach(b=>b.addEventListener('click',()=>switchTabs(b,'data-input',{message:'message-input',config:'config-input'})));
document.querySelectorAll('[data-output]').forEach(b=>b.addEventListener('click',()=>switchTabs(b,'data-output',{message:'output-message',changes:'output-changes',notes:'output-notes'})));
$('run').addEventListener('click',()=>run(true));$('step').addEventListener('click',()=>run(false));$('reset').addEventListener('click',resetSample);$('scenario').addEventListener('change',resetSample);
['invite','profiles'].forEach(id=>$(id).addEventListener('input',invalidate));
$('table-select').addEventListener('change',renderTable);$('lookup-form').addEventListener('submit',lookup);
$('copy').addEventListener('click',async()=>{try{await navigator.clipboard.writeText(result.output);$('copy').textContent='已复制';setTimeout(()=>$('copy').textContent='复制报文',1600);}catch{notice('浏览器未允许复制，请选中输出报文手动复制。');}});
$('export').addEventListener('click',async()=>{if(!result)return;$('export').disabled=true;try{const saved=await post('/api/export',result);notice(`结果已保存到：${saved.path}`,true);}catch(err){notice(`导出失败：${err.message}`);}finally{$('export').disabled=false;}});
setBusy(true);render();
fetch('/api/bootstrap').then(r=>{if(!r.ok)throw Error('数据加载失败');return r.json();}).then(data=>{bootstrap=data;resetSample();renderEvidence();}).catch(e=>notice(`无法加载本地数据：${e.message}。请通过启动脚本打开页面。`)).finally(()=>setBusy(false));
