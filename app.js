const sampleData = [
  {office:'AA営業所',member:'MR0001',date:'2026-01-31',target:'A',facility:'施設名0001',doctor:'Dr名0001',amtul:'U',material:'AAA',activity:'最近気づいたけどBBB10mgはAAA50mgより高い。同等であればAAAへ変更したい。',score:15},
  {office:'AA営業所',member:'MR0002',date:'2026-01-31',target:'A',facility:'施設名0002',doctor:'Dr名0002',amtul:'A',material:'AAA',activity:'特に宣伝許可は必要ない。先生が使いたいと言えば申請を上げてもらう。',score:0},
  {office:'BB営業所',member:'MR0003',date:'2026-01-30',target:'A',facility:'施設名0003',doctor:'Dr名0003',amtul:'A',material:'AAA',activity:'新規処方依頼。新患や単発処方可能患者なし。',score:0},
  {office:'BB営業所',member:'MR0004',date:'2026-01-29',target:'A',facility:'施設名0004',doctor:'Dr名0004',amtul:'M',material:'AAA',activity:'他剤と何が違う？ 嘔吐はどうだった？',score:0},
  {office:'BB営業所',member:'MR0003',date:'2026-01-30',target:'A',facility:'施設名0005',doctor:'Dr名0005',amtul:'A',material:'AAA',activity:'御礼。web案内、新規処方提案。本日予定あり面談不可、案内対応。',score:0},
  {office:'BB営業所',member:'MR0005',date:'2026-01-31',target:'A',facility:'施設名0006',doctor:'Dr名0006',amtul:'A',material:'AAA',activity:'面会不可、お手紙。',score:0},
  {office:'DD営業所',member:'MR0007',date:'2026-01-29',target:'A',facility:'施設名0009',doctor:'Dr名0009',amtul:'T',material:'AAA',activity:'発売案内。BBBメイン。他剤との比較、特徴と持ち越し効果を説明。説明会での紹介依頼。',score:5},
  {office:'EE営業所',member:'MR0008',date:'2026-01-31',target:'B',facility:'施設名0010',doctor:'Dr名0010',amtul:'A',material:'AAA',activity:'Webカンファレンス案内。',score:-1},
  {office:'EE営業所',member:'MR0008',date:'2026-01-31',target:'B',facility:'施設名0011',doctor:'Dr名0011',amtul:'A',material:'AAA',activity:'Webカンファレンス案内とAAAの特徴紹介。',score:0}
];

const aliases={office:['営業所名','営業所','支店','所属'],member:['担当MR名','担当MR','担当者','MR名','社員名'],date:['活動日','日付','活動年月日'],target:['ターゲット','ターゲット区分'],facility:['施設名','医療機関名'],doctor:['医師名','医師'],amtul:['AMTUL','区分'],material:['使用資材','資材'],activity:['活動内容','活動報告','内容','コメント'],score:['DDD','スコア','質スコア']};
let allData=[...sampleData];
let activeEvaluationFilter='';
let selectedAdviceOffice='';
let lastGptAnalysisKey='';
const $=id=>document.getElementById(id);
const normalize=v=>String(v??'').replace(/[\s\n\r　]/g,'').toLowerCase();
const escapeHtml=v=>String(v??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));

function excelDate(value){
  if(value instanceof Date&&!isNaN(value)) return value.toISOString().slice(0,10);
  if(typeof value==='number'&&window.XLSX){const d=XLSX.SSF.parse_date_code(value);if(d)return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`}
  const text=String(value??'').trim().replace(/[./]/g,'-');
  const d=new Date(text);return isNaN(d)?'':`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
function mapRows(rows){
  if(!rows.length)return [];
  const headers=Object.keys(rows[0]);const keyMap={};
  for(const [field,names] of Object.entries(aliases)){keyMap[field]=headers.find(h=>names.some(n=>normalize(h)===normalize(n)))||headers.find(h=>names.some(n=>normalize(h).includes(normalize(n))))}
  if(!keyMap.activity&&!keyMap.office)throw new Error('「営業所名」「活動内容」などの見出しを確認してください。');
  return rows.map(r=>({office:String(r[keyMap.office]??'未設定').trim(),member:String(r[keyMap.member]??'未設定').trim(),date:excelDate(r[keyMap.date]),target:String(r[keyMap.target]??''),facility:String(r[keyMap.facility]??''),doctor:String(r[keyMap.doctor]??''),amtul:String(r[keyMap.amtul]??''),material:String(r[keyMap.material]??''),activity:String(r[keyMap.activity]??''),score:Number(r[keyMap.score])||0})).filter(r=>r.office||r.activity);
}
function loadFile(file){
  if(!window.XLSX){showToast('Excel読込ライブラリを読み込めません。インターネット接続を確認してください。');return}
  const reader=new FileReader();reader.onload=e=>{try{const wb=XLSX.read(e.target.result,{type:'array',cellDates:true});let rows=[];for(const name of wb.SheetNames){const candidate=XLSX.utils.sheet_to_json(wb.Sheets[name],{defval:''});if(candidate.length>rows.length)rows=candidate}const mapped=mapRows(rows);if(!mapped.length)throw new Error('データ行が見つかりません。');allData=mapped;initializeFilters();clearGptAnalysis('Excelを読み込みました。必要に応じて追加分析を実行してください。');render();$('dataSource').textContent=file.name;showToast(`${mapped.length}件の活動報告を読み込みました`)}catch(err){showToast(`読込エラー: ${err.message}`)}};reader.readAsArrayBuffer(file);
}
function initializeFilters(){
  const offices=[...new Set(allData.map(x=>x.office).filter(Boolean))].sort();
  $('officeFilter').innerHTML='<option value="">すべて</option>'+offices.map(x=>`<option>${escapeHtml(x)}</option>`).join('');
  updateMemberOptions('');
  const dates=allData.map(x=>x.date).filter(Boolean).sort();$('dateFrom').value=dates[0]||'';$('dateTo').value=dates.at(-1)||'';
  renderDataWarning();
}
function updateMemberOptions(office,selected=''){
  const members=[...new Set(allData.filter(x=>!office||x.office===office).map(x=>x.member).filter(Boolean))].sort();
  $('memberFilter').innerHTML='<option value="">すべて</option>'+members.map(x=>`<option ${x===selected?'selected':''}>${escapeHtml(x)}</option>`).join('');
}
function renderDataWarning(){
  const yearCounts={};allData.forEach(x=>{const year=x.date?.slice(0,4);if(year)yearCounts[year]=(yearCounts[year]||0)+1});
  const years=Object.entries(yearCounts).sort((a,b)=>b[1]-a[1]);const mainYear=years[0]?.[0];
  const outliers=allData.filter(x=>x.date&&mainYear&&x.date.slice(0,4)!==mainYear);const warning=$('dataWarning');
  if(outliers.length&&outliers.length<=Math.max(3,allData.length*.1)){
    warning.hidden=false;$('dataWarningText').textContent=`主要年は${mainYear}年ですが、異なる年の活動日が${outliers.length}件あります（${[...new Set(outliers.map(x=>x.date))].join('、')}）。入力日をご確認ください。`;
  }else warning.hidden=true;
}
function filteredData(){const from=$('dateFrom').value,to=$('dateTo').value,office=$('officeFilter').value,member=$('memberFilter').value;return allData.filter(x=>(!from||x.date>=from)&&(!to||x.date<=to)&&(!office||x.office===office)&&(!member||x.member===member))}
function aggregate(data,key){const result={};data.forEach(x=>{const k=x[key]||'未設定';result[k]??={count:0,total:0};result[k].count++;result[k].total+=x.score});return Object.entries(result).map(([name,v])=>({name,count:v.count,avg:v.total/v.count})).sort((a,b)=>b.count-a.count||b.avg-a.avg)}
const evaluationRules=[
  {id:'needs',label:'医師の困りごと',short:'課題把握',good:true,description:'治療上の課題、懸念、使いにくさ、質問など',patterns:[/困(って|る|り|ら)/i,/ネック|問題|課題|懸念|不安|難し|無理|できない|出来ない|使いにく|面倒|デメリット|リスク|副作用|嘔吐|悪夢|持ち越し|転倒|一包化|粉砕|違いは|何が違|どうなの|どうだ|なぜ|何故|\?|？/]},
  {id:'competitor',label:'競合関連',short:'競合把握',good:true,description:'AAA以外の薬剤名、他剤、切替・比較の記載',patterns:[/BBB|CCC|他剤|他薬|競合|比較|切り替|切替|併用薬|薬剤で/]},
  {id:'patient',label:'患者像',short:'患者理解',good:true,description:'患者属性、症状、処方例・症例などの記載',patterns:[/患者|新患|再診|症例|例ほど|\d+例|寝つき|途中で起き|不眠|認知症|精神科|内科|高齢|老年|錠剤不可|かかりつけ|短期使用|長期処方|単発処方/]},
  {id:'nextAction',label:'次回アクション',short:'行動具体化',good:true,description:'次回確認、日程、説明会、申請、検討事項など',patterns:[/次回|次に|後日|来週|再来週|[0-9０-９]+月|[0-9０-９]+日|アポイント|アポ|日程|予定|説明会|研修会|カンファ|申請|薬審|発注|確認させ|紹介して|伝達|相談|検討して|案内前日|延期日/]},
  {id:'intent',label:'処方意向',short:'意向獲得',good:true,description:'使いたい、処方予定、採用・発注など前向きな意向',patterns:[/使いたい|使ってみ|使うかも|使う可能性|処方してみ|処方予定|処方する|変えていきたい|切り替えたい|採用|発注|承認|薬審通過|導入|試してみ|良い薬|いいかも|AAAに変わ|症例増や|紹介しておいて/]},
  {id:'noReaction',label:'反応なし・非面談',short:'要改善活動',good:false,description:'面会不可、反応なし、資料提供・案内のみなど',patterns:[/面会不可|不面|面会.*断|会えな|多忙.*面会|資料提供のみ|【資料提供】|資料提供$|お手紙|受付で.*配布|案内を配布|案内対応|WEB案内→ふーん|食いつきが無|反応なし|EPPV対応$|社内研修会$|シンポジウム参加$/i]}
];
function evaluationMatches(item,rule){const text=String(item.activity||'').replace(/\s+/g,'');return rule.patterns.some(pattern=>pattern.test(text))}
function rateFor(data,rule){return data.length?data.filter(x=>evaluationMatches(x,rule)).length/data.length*100:0}
function amtulMix(data){const mix={A:0,M:0,T:0,U:0};data.forEach(x=>{const key=String(x.amtul||'').trim();mix[key]=(mix[key]||0)+1});return mix}
function adviceBaseData(){const from=$('dateFrom').value,to=$('dateTo').value;return allData.filter(x=>(!from||x.date>=from)&&(!to||x.date<=to))}
function adviceContext(){
  const base=adviceBaseData(),office=$('officeFilter').value||selectedAdviceOffice,member=$('memberFilter').value;
  if(member&&office){const comparison=base.filter(x=>x.office===office);return {scope:'member',label:member,base:comparison,data:comparison.filter(x=>x.member===member),comparisonLabel:`${office} 担当者平均`,comparisonCount:new Set(comparison.map(x=>x.member)).size||1,office};}
  const offices=[...new Set(base.map(x=>x.office).filter(Boolean))].sort();if(office&&offices.includes(office))selectedAdviceOffice=office;if(!selectedAdviceOffice||!offices.includes(selectedAdviceOffice))selectedAdviceOffice=offices[0]||'';
  return {scope:'office',label:selectedAdviceOffice,base,data:base.filter(x=>x.office===selectedAdviceOffice),comparisonLabel:`全${offices.length}営業所平均`,comparisonCount:offices.length||1,office:selectedAdviceOffice};
}
function adviceContextKey(context=adviceContext()){return `${context.scope}|${context.office||''}|${context.label||''}|${$('dateFrom').value}|${$('dateTo').value}`}
function renderQualityRadar(officeData,base,officeName){
  const axes=evaluationRules.map(rule=>({label:rule.id==='noReaction'?'非面談':rule.label,office:rateFor(officeData,rule),average:rateFor(base,rule)}));
  const cx=230,cy=158,radius=112,labelRadius=139,count=axes.length,point=(index,value,scale=radius)=>{const angle=-Math.PI/2+index*Math.PI*2/count,distance=scale*(value/100);return [cx+Math.cos(angle)*distance,cy+Math.sin(angle)*distance]};
  const polygon=value=>axes.map((_,i)=>point(i,value).join(',')).join(' '),dataPolygon=key=>axes.map((axis,i)=>point(i,axis[key]).join(',')).join(' ');
  const grid=[20,40,60,80,100].map(value=>`<polygon points="${polygon(value)}" fill="none" stroke="${value===100?'#cfd8e5':'#e5eaf1'}" stroke-width="1"/><text x="${cx+4}" y="${cy-radius*value/100+10}" fill="#a0acbb" font-size="7">${value}</text>`).join('');
  const spokes=axes.map((_,i)=>{const [x,y]=point(i,100);return `<line x1="${cx}" y1="${cy}" x2="${x}" y2="${y}" stroke="#dfe5ed" stroke-width="1"/>`}).join('');
  const labels=axes.map((axis,i)=>{const [x,y]=point(i,100,labelRadius),anchor=x<cx-10?'end':x>cx+10?'start':'middle';return `<text x="${x}" y="${y}" text-anchor="${anchor}" dominant-baseline="middle" fill="#45556b" font-size="10" font-weight="700">${escapeHtml(axis.label)}</text>`}).join('');
  const points=axes.map((axis,i)=>{const [x,y]=point(i,axis.office);return `<circle cx="${x}" cy="${y}" r="3.5" fill="#1479e8" stroke="#fff" stroke-width="2"><title>${escapeHtml(axis.label)} ${axis.office.toFixed(0)}%</title></circle>`}).join('');
  $('qualityRadar').innerHTML=`<svg viewBox="0 0 460 315" role="img" aria-label="${escapeHtml(officeName)}と全営業所平均の活動品質レーダーチャート">${grid}${spokes}<polygon points="${dataPolygon('average')}" fill="#f39a3618" stroke="#f39a36" stroke-width="2" stroke-dasharray="5 4"/><polygon points="${dataPolygon('office')}" fill="#1479e82b" stroke="#1479e8" stroke-width="2.5"/>${points}${labels}</svg>`;
  $('radarOfficeName').textContent=officeName;$('radarValues').innerHTML=axes.map(axis=>`<div class="radar-value-row"><span>${escapeHtml(axis.label)}</span><b>${axis.office.toFixed(0)}%</b><small>${axis.average.toFixed(0)}%</small></div>`).join('');
}
function buildGptAnalysisPayload(context=adviceContext()){
  if(!context.label||!context.data.length)throw new Error('分析対象を選択してください。');
  const targetData=context.data,base=context.base,targetMembers=new Set(targetData.map(x=>x.member)).size||1,baseMembers=new Set(base.map(x=>x.member)).size||1,targetMix=amtulMix(targetData),baseMix=amtulMix(base);
  return {period:{from:$('dateFrom').value,to:$('dateTo').value},analysis_scope:context.scope,comparison_offices:context.comparisonCount,volume:{visits:targetData.length,average_visits:Number((base.length/context.comparisonCount).toFixed(1)),visits_per_member:Number((targetData.length/targetMembers).toFixed(1)),average_visits_per_member:Number((base.length/baseMembers).toFixed(1)),active_members:targetMembers},quality_score:{office:Number((targetData.reduce((s,x)=>s+x.score,0)/(targetData.length||1)).toFixed(1)),average:Number((base.reduce((s,x)=>s+x.score,0)/(base.length||1)).toFixed(1))},amtul:['A','M','T','U'].map(code=>({code,office_rate:Number((targetMix[code]/(targetData.length||1)*100).toFixed(1)),average_rate:Number((baseMix[code]/(base.length||1)*100).toFixed(1))})),quality_metrics:evaluationRules.map(rule=>({id:rule.id,label:rule.label,desirable_direction:rule.good?'higher':'lower',office_rate:Number(rateFor(targetData,rule).toFixed(1)),average_rate:Number(rateFor(base,rule).toFixed(1))}))};
}
function clearGptAnalysis(message='営業所または担当者を変更したため、必要に応じて追加分析を再実行してください。'){lastGptAnalysisKey='';$('gptAnalysisResult').hidden=true;$('gptAnalysisResult').innerHTML='';$('gptAnalysisStatus').textContent=message}
function renderGptList(title,items){return `<article class="gpt-result-block"><h4>${escapeHtml(title)}</h4><ul>${items.map(x=>`<li>${escapeHtml(x)}</li>`).join('')}</ul></article>`}
function renderGptAnalysis(analysis,model){
  const questionTitle='所長が確認すべき質問';
  lastGptAnalysisKey=adviceContextKey();
  $('gptAnalysisResult').hidden=false;$('gptAnalysisResult').innerHTML=`<div class="gpt-executive">${escapeHtml(analysis.executive_comment)}</div><div class="gpt-result-grid">${renderGptList('分析上の着眼点',analysis.observations)}${renderGptList('優先課題',analysis.priority_issues)}${renderGptList('推奨する打ち手',analysis.recommended_actions)}${renderGptList(questionTitle,analysis.manager_questions)}</div><article class="gpt-result-block"><h4>来月のKPI提案</h4><div class="gpt-kpis">${analysis.next_month_kpis.map(x=>`<div class="gpt-kpi"><strong>${escapeHtml(x.name)}</strong><span>${escapeHtml(x.target)}</span><small>${escapeHtml(x.reason)}</small></div>`).join('')}</div></article><p class="gpt-caution">${escapeHtml(analysis.caution)} / 使用モデル: ${escapeHtml(model)}</p>`;
}
async function runGptAnalysis(){
  const button=$('runGptAnalysis'),status=$('gptAnalysisStatus'),context=adviceContext();button.disabled=true;button.textContent='分析中...';$('gptAnalysisResult').hidden=true;status.textContent='匿名化した集計値をAIで分析しています。';
  try{const response=await fetch('/api/gpt-analysis',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(buildGptAnalysisPayload({...context,scope:'office'}))});const contentType=response.headers.get('content-type')||'';if(!contentType.includes('application/json'))throw new Error('Claude分析サービスに接続できません。しばらくしてから再度お試しください。');const result=await response.json();if(!response.ok)throw new Error(result.error||'AI分析に失敗しました。');renderGptAnalysis(result.analysis,result.model);status.textContent=`${context.office||context.label}の追加分析が完了しました。`;}
  catch(error){status.textContent=`エラー: ${error.message}`;}
  finally{button.disabled=false;button.textContent='追加分析を実行';}
}
function renderMemberAdvice(analysis,model){
  $('memberAdviceResult').hidden=false;$('memberAdviceResult').innerHTML=`<div class="gpt-executive">${escapeHtml(analysis.executive_comment)}</div><div class="gpt-result-grid">${renderGptList('分析上の着眼点',analysis.observations)}${renderGptList('優先課題',analysis.priority_issues)}${renderGptList('推奨する打ち手',analysis.recommended_actions)}${renderGptList('担当者へ確認すべき質問',analysis.manager_questions)}</div><article class="gpt-result-block"><h4>来月のKPI提案</h4><div class="gpt-kpis">${analysis.next_month_kpis.map(x=>`<div class="gpt-kpi"><strong>${escapeHtml(x.name)}</strong><span>${escapeHtml(x.target)}</span><small>${escapeHtml(x.reason)}</small></div>`).join('')}</div></article><p class="gpt-caution">${escapeHtml(analysis.caution)} / 使用モデル: ${escapeHtml(model)}</p>`;
}
async function runMemberAdvice(){
  const context=adviceContext();if(context.scope!=='member')return;
  const button=$('runMemberAdvice'),status=$('memberAdviceStatus');button.disabled=true;button.textContent='分析中...';$('memberAdviceResult').hidden=false;$('memberAdviceResult').innerHTML=`<div class="gpt-executive">${escapeHtml(context.label)}向けの担当者アドバイスを作成しています...</div>`;status.textContent='匿名化した担当者集計値をAIで分析しています。';
  try{const response=await fetch('/api/gpt-analysis',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(buildGptAnalysisPayload(context))});const contentType=response.headers.get('content-type')||'';if(!contentType.includes('application/json'))throw new Error('追加分析用サーバーに接続できません。');const result=await response.json();if(!response.ok)throw new Error(result.error||'担当者アドバイスに失敗しました。');renderMemberAdvice(result.analysis,result.model);status.textContent=`${context.label}の担当者アドバイスが完了しました。`;}
  catch(error){status.textContent=`エラー: ${error.message}`;}
  finally{button.disabled=false;button.textContent='担当者アドバイス';}
}
function updateMemberAdvicePanel(){
  const context=adviceContext(),show=context.scope==='member';$('memberAdvicePanel').classList.toggle('show',show);
  if(show){$('memberAdviceDescription').textContent=`${context.office} / ${context.label} の匿名化集計を使って担当者向けにAIが助言します。`;$('memberAdviceStatus').textContent='必要に応じて「担当者アドバイス」を実行してください。';}
  else{$('memberAdviceResult').hidden=true;$('memberAdviceResult').innerHTML='';$('memberAdviceStatus').textContent='営業所名と担当者名を選択してください。';}
}
function render(){const data=filteredData();const members=new Set(data.map(x=>x.member));const avg=data.length?data.reduce((s,x)=>s+x.score,0)/data.length:0;const high=data.filter(x=>x.score>=5).length;
  $('totalActivities').textContent=data.length.toLocaleString();$('activeMembers').textContent=members.size;$('qualityAverage').textContent=avg.toFixed(1);$('highQuality').textContent=high;$('memberFoot').textContent=`1人あたり ${members.size?(data.length/members.size).toFixed(1):0}件`;$('highQualityFoot').textContent=`全体の ${data.length?Math.round(high/data.length*100):0}%`;$('recordStatus').textContent=`${data.length}件を表示中`;$('updatedAt').textContent=`更新: ${new Date().toLocaleString('ja-JP',{month:'numeric',day:'numeric',hour:'2-digit',minute:'2-digit'})}`;
  renderEvaluations(data);renderManagerAdvice();renderOffice(data);renderQuality(data);renderMembers(data);renderTrend(data);renderCategories(data);renderKeywords(data);renderTable(data);renderInsight(data,avg);updateMemberAdvicePanel();
}
function renderEvaluations(data){
  $('evaluationMetrics').innerHTML=evaluationRules.map(rule=>{const count=data.filter(x=>evaluationMatches(x,rule)).length;const rate=data.length?Math.round(count/data.length*100):0;return `<button class="evaluation-card ${rule.good?'positive-metric':'risk-metric'} ${activeEvaluationFilter===rule.id?'selected':''}" data-evaluation="${rule.id}" title="${escapeHtml(rule.description)}"><div class="evaluation-card-head"><span>${escapeHtml(rule.label)}</span><b>${rule.good?'PLUS':'CHECK'}</b></div><div class="evaluation-value"><strong>${rate}</strong><small>%</small><em>${count}件</em></div><div class="evaluation-track"><i style="width:${rate}%"></i></div><p>${escapeHtml(rule.description)}</p></button>`}).join('');
  document.querySelectorAll('[data-evaluation]').forEach(button=>button.addEventListener('click',()=>{activeEvaluationFilter=activeEvaluationFilter===button.dataset.evaluation?'':button.dataset.evaluation;$('clearEvaluationFilter').hidden=!activeEvaluationFilter;renderEvaluations(data);renderTable(data);$('detailSection').scrollIntoView({behavior:'smooth',block:'start'})}));
}
function renderManagerAdvice(){
  const baseAll=adviceBaseData();const offices=[...new Set(baseAll.map(x=>x.office).filter(Boolean))].sort();
  if(!offices.length){$('adviceSummary').innerHTML='<div class="empty-state">対象データがありません</div>';return}
  const filterOffice=$('officeFilter').value;if(filterOffice&&offices.includes(filterOffice))selectedAdviceOffice=filterOffice;
  if(!selectedAdviceOffice||!offices.includes(selectedAdviceOffice))selectedAdviceOffice=offices[0];
  $('adviceOffice').innerHTML=offices.map(x=>`<option ${x===selectedAdviceOffice?'selected':''}>${escapeHtml(x)}</option>`).join('');
  const context=adviceContext(),base=context.base,officeData=context.data,isMember=context.scope==='member';if(!officeData.length){$('adviceSummary').innerHTML='<div class="empty-state">対象データがありません</div>';return}
  if(lastGptAnalysisKey&&lastGptAnalysisKey!==adviceContextKey(context))clearGptAnalysis(isMember?'担当者を選んだので、必要なら追加分析を再実行してください。':'営業所または期間を変更したため、必要に応じて追加分析を再実行してください。');
  $('adviceTitle').textContent=isMember?'担当者別 アドバイス':'営業所別 所長アドバイス';$('adviceQuestionTitle').textContent=isMember?'担当者へ確認すべき質問':'所長が確認すべき質問';
  const officeMembers=new Set(officeData.map(x=>x.member)).size||1;const allMembers=new Set(base.map(x=>x.member)).size||1;
  const avgVisits=base.length/context.comparisonCount,officePerMember=officeData.length/officeMembers,avgPerMember=base.length/allMembers;
  const volumeDiff=avgVisits?Math.round((officeData.length/avgVisits-1)*100):0,perMemberDiff=avgPerMember?Math.round((officePerMember/avgPerMember-1)*100):0;
  const officeMix=amtulMix(officeData),baseMix=amtulMix(base),amtulKeys=['A','M','T','U'];
  const amtulStats=amtulKeys.map(key=>({key,rate:officeData.length?officeMix[key]/officeData.length*100:0,avg:base.length?baseMix[key]/base.length*100:0}));
  const biggestMixGap=amtulStats.slice().sort((a,b)=>Math.abs(b.rate-b.avg)-Math.abs(a.rate-a.avg))[0];
  const metricStats=evaluationRules.map(rule=>({rule,rate:rateFor(officeData,rule),avg:rateFor(base,rule)}));
  const positive=metricStats.filter(x=>x.rule.good),strength=positive.slice().sort((a,b)=>(b.rate-b.avg)-(a.rate-a.avg))[0],priority=positive.slice().sort((a,b)=>(a.rate-a.avg)-(b.rate-b.avg))[0],risk=metricStats.find(x=>x.rule.id==='noReaction');
  const qualityAvg=officeData.length?officeData.reduce((s,x)=>s+x.score,0)/officeData.length:0,baseQuality=base.length?base.reduce((s,x)=>s+x.score,0)/base.length:0;
  const tone=volumeDiff>=10&&priority.rate>=priority.avg?'好調':volumeDiff<-10||risk.rate>risk.avg+5?'改善優先':'標準圏';
  $('adviceBenchmark').innerHTML=`<span>比較基準</span><strong>${escapeHtml(context.comparisonLabel)}</strong><i>訪問 ${avgVisits.toFixed(1)}件</i><i>MR1人あたり ${avgPerMember.toFixed(1)}件</i><i>質スコア ${baseQuality.toFixed(1)}pt</i>`;
  $('adviceSummary').innerHTML=`<div class="advice-grade ${tone==='好調'?'grade-good':tone==='改善優先'?'grade-alert':'grade-standard'}">${tone}</div><div><strong>${escapeHtml(context.label)}</strong><p>訪問量は比較基準比 ${volumeDiff>=0?'+':''}${volumeDiff}%、MR1人あたりは ${perMemberDiff>=0?'+':''}${perMemberDiff}%です。質スコアは${qualityAvg.toFixed(1)}pt（基準${baseQuality.toFixed(1)}pt）。</p></div>`;
  renderQualityRadar(officeData,base,context.label);
  const comparison=(value,avg,inverse=false)=>{const diff=value-avg;const favorable=inverse?diff<=0:diff>=0;return `<b class="${favorable?'good-text':'alert-text'}">${diff>=0?'+':''}${diff.toFixed(0)}pt</b>`};
  $('adviceCurrent').innerHTML=`<ul><li>訪問件数 <b>${officeData.length}件</b>（平均${avgVisits.toFixed(1)}件、${volumeDiff>=0?'+':''}${volumeDiff}%）</li><li>強みは「${strength.rule.label}」${strength.rate.toFixed(0)}%で、平均との差は${comparison(strength.rate,strength.avg)}です。</li><li>AMTULは${biggestMixGap.key}が${biggestMixGap.rate.toFixed(0)}%（全体${biggestMixGap.avg.toFixed(0)}%）と最も構成差があります。</li></ul>`;
  const issueItems=[];if(volumeDiff<0)issueItems.push(`訪問量が比較基準を${Math.abs(volumeDiff)}%下回っています。活動可能時間と重点先配分の確認が必要です。`);if(priority.rate<priority.avg)issueItems.push(`「${priority.rule.label}」が${priority.rate.toFixed(0)}%で、基準を${(priority.avg-priority.rate).toFixed(0)}pt下回っています。`);if(risk.rate>risk.avg)issueItems.push(`「${risk.rule.label}」が${risk.rate.toFixed(0)}%で、基準より${(risk.rate-risk.avg).toFixed(0)}pt高い状態です。`);if(!issueItems.length)issueItems.push('主要指標は基準以上です。良い面談パターンの再現性と、処方変化への接続を確認してください。');
  $('adviceIssues').innerHTML=`<ul>${issueItems.map(x=>`<li>${escapeHtml(x)}</li>`).join('')}</ul>`;
  const actionItems=[];if(volumeDiff<0)actionItems.push(isMember?'重点医師を週次で明確にし、面談枠・アポイント・再訪予定を本人と確認する。':'重点医師を週次で明確にし、面談枠・アポイント・再訪予定をMR別に確認する。');actionItems.push(`「${priority.rule.label}」を面談記録の必須項目にし、質問例を使って面談前に準備する。`);if(risk.rate>risk.avg)actionItems.push('非面談活動は再接触日と別チャネルを必ず設定し、資料配布のみで完了させない。');actionItems.push(`基準を上回る「${strength.rule.label}」の好事例を共有し、再現できる型にする。`);
  $('adviceActions').innerHTML=`<ol>${actionItems.slice(0,3).map(x=>`<li>${escapeHtml(x)}</li>`).join('')}</ol>`;
  const questions=[`今月、処方拡大の障壁を具体的に聞けた医師は誰で、障壁は何でしたか？`,`「${priority.rule.label}」が記載されていない面談で、実際には何を確認しましたか？`,risk.rate>risk.avg?'会えなかった医師への次回接触日と代替チャネルは決まっていますか？':isMember?'処方意向を次の処方確認へつなげる日付は決まっていますか？':`処方意向を次の処方確認へつなげる日付と担当者は決まっていますか？`];
  $('adviceQuestions').innerHTML=`<ol>${questions.map(x=>`<li>${escapeHtml(x)}</li>`).join('')}</ol>`;
  const visitTarget=Math.max(officeData.length+1,Math.ceil(Math.min(avgVisits,officeData.length*1.15||avgVisits)));const priorityTarget=Math.min(100,Math.ceil(Math.max(priority.avg,priority.rate+10)/5)*5);const intent=metricStats.find(x=>x.rule.id==='intent');const intentTarget=Math.min(100,Math.ceil(Math.max(intent.avg,intent.rate+10)/5)*5);const riskTarget=Math.max(0,Math.floor(Math.min(risk.avg,risk.rate-5)/5)*5);
  const kpis=[['訪問件数',`${visitTarget}件以上`,`現状 ${officeData.length}件`],[priority.rule.label,`${priorityTarget}%以上`,`現状 ${priority.rate.toFixed(0)}%`],['処方意向',`${intentTarget}%以上`,`現状 ${intent.rate.toFixed(0)}%`],['反応なし・非面談',`${riskTarget}%以下`,`現状 ${risk.rate.toFixed(0)}%`]];
  $('adviceKpis').innerHTML=kpis.map(([name,target,current])=>`<div class="advice-kpi"><span>${escapeHtml(name)}</span><strong>${escapeHtml(target)}</strong><small>${escapeHtml(current)}</small></div>`).join('');
}
function renderOffice(data){const items=aggregate(data,'office').slice(0,7),max=Math.max(1,...items.map(x=>x.count)),scoreMax=Math.max(5,...items.map(x=>x.avg));$('officeChart').innerHTML=items.length?items.map(x=>`<div class="bar-row"><span class="bar-label" title="${escapeHtml(x.name)}">${escapeHtml(x.name)}</span><div class="bar-track"><div class="bar-fill" style="width:${x.count/max*100}%"></div><i class="quality-marker" title="平均 ${x.avg.toFixed(1)}pt" style="left:${Math.max(2,x.avg/scoreMax*100)}%"></i></div><span class="bar-value">${x.count}件</span></div>`).join(''):'<div class="empty-state">該当データがありません</div>'}
function renderQuality(data){const positive=data.filter(x=>x.score>0).length,neutral=data.filter(x=>x.score===0).length,negative=data.filter(x=>x.score<0).length,total=data.length||1,pct=n=>n/total*100;const p=pct(positive),n=pct(neutral);$('positiveRate').textContent=`${Math.round(p)}%`;$('qualityDonut').style.background=`conic-gradient(var(--green) 0 ${p}%, #ffc35b ${p}% ${p+n}%, #ec6d73 ${p+n}% 100%)`;$('qualityLegend').innerHTML=[[positive,'ポジティブ','#13a879'],[neutral,'ニュートラル','#ffc35b'],[negative,'ネガティブ','#ec6d73']].map(([v,l,c])=>`<div class="quality-item"><i style="background:${c}"></i><span>${l}</span><strong>${v}件</strong></div>`).join('')}
function renderMembers(data){const items=aggregate(data,'member').slice(0,5),max=Math.max(1,...items.map(x=>x.count));$('memberRanking').innerHTML=items.length?items.map((x,i)=>`<div class="rank-row"><span class="rank-number">${String(i+1).padStart(2,'0')}</span><div class="rank-info"><div class="rank-name">${escapeHtml(x.name)}<span>${x.count}件</span></div><div class="rank-track"><div class="rank-fill" style="width:${x.count/max*100}%"></div></div></div><div class="rank-score"><strong>${x.avg.toFixed(1)}</strong><span>平均pt</span></div></div>`).join(''):'<div class="empty-state">該当データがありません</div>'}
function renderTrend(data){const grouped=aggregate(data,'date').filter(x=>x.name!=='未設定').sort((a,b)=>a.name.localeCompare(b.name));if(!grouped.length){$('trendChart').innerHTML='<div class="empty-state">日付データがありません</div>';return}const max=Math.max(1,...grouped.map(x=>x.count)),w=500,h=160,pts=grouped.map((x,i)=>`${grouped.length===1?w/2:i/(grouped.length-1)*w},${h-x.count/max*(h-20)}`);const area=`${pts[0].split(',')[0]},${h} ${pts.join(' ')} ${pts.at(-1).split(',')[0]},${h}`;$('trendChart').innerHTML=`<div class="trend-grid"></div><svg class="trend-svg" viewBox="0 0 ${w} ${h}" preserveAspectRatio="none"><defs><linearGradient id="area" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#1479e8" stop-opacity=".28"/><stop offset="1" stop-color="#1479e8" stop-opacity="0"/></linearGradient></defs><polygon points="${area}" fill="url(#area)"/><polyline points="${pts.join(' ')}" fill="none" stroke="#1479e8" stroke-width="3" vector-effect="non-scaling-stroke"/>${pts.map(p=>{const [x,y]=p.split(',');return `<circle cx="${x}" cy="${y}" r="4" fill="#fff" stroke="#1479e8" stroke-width="3" vector-effect="non-scaling-stroke"/>`}).join('')}</svg><div class="trend-labels">${grouped.map(x=>`<span>${x.name.slice(5).replace('-','/')}</span>`).join('')}</div>`}
function renderCategories(data){
  const renderCategory=(id,key,order,colors)=>{const counts={};data.forEach(x=>{const value=String(x[key]||'未設定').trim();counts[value]=(counts[value]||0)+1});const names=[...new Set([...order,...Object.keys(counts)])].filter(x=>counts[x]);const max=Math.max(1,...names.map(x=>counts[x]));$(id).innerHTML=names.length?names.map((name,i)=>`<div class="category-row-item"><span class="category-code" style="--category-color:${colors[i%colors.length]}">${escapeHtml(name)}</span><div class="category-track"><div class="category-fill" style="width:${counts[name]/max*100}%;background:${colors[i%colors.length]}"></div></div><strong>${counts[name]}件</strong><small>${data.length?Math.round(counts[name]/data.length*100):0}%</small></div>`).join(''):'<div class="empty-state">該当データがありません</div>'};
  renderCategory('amtulChart','amtul',['A','M','T','U'],['#1479e8','#13a879','#f39a36','#7357d9']);
  renderCategory('scaleChart','target',['A','B','C','D'],['#18a7c9','#4b7bec','#8c62d8','#eb6b78']);
}
function renderKeywords(data){const stop=['について','として','できる','ください','ました','します','ない','あり','こと','ため','先生','患者','案内','対応'];const words={};data.forEach(x=>(x.activity.match(/[A-Za-zＡ-Ｚａ-ｚ0-9０-９一-龠ァ-ヶー]{2,}/g)||[]).forEach(w=>{const key=w.toUpperCase();if(!stop.some(s=>key.includes(s))&&key.length<24)words[key]=(words[key]||0)+1}));const top=Object.entries(words).sort((a,b)=>b[1]-a[1]).slice(0,10);$('keywords').innerHTML=top.length?top.map(([w,n])=>`<span class="keyword">${escapeHtml(w)} <strong>${n}</strong></span>`).join(''):'<span class="panel-note">抽出できるキーワードがありません</span>'}
function renderTable(data){const rule=evaluationRules.find(x=>x.id===activeEvaluationFilter);const visible=rule?data.filter(x=>evaluationMatches(x,rule)):data;$('detailCount').textContent=rule?`${rule.label}: ${visible.length}件`:`${visible.length}件`;$('detailBody').innerHTML=visible.length?visible.slice().sort((a,b)=>b.date.localeCompare(a.date)).map(x=>`<tr><td>${escapeHtml(x.date)}</td><td>${escapeHtml(x.office)}</td><td>${escapeHtml(x.member)}</td><td>${escapeHtml(x.doctor)}</td><td>${escapeHtml(x.material)}</td><td class="activity-cell">${escapeHtml(x.activity)}</td><td><span class="score-badge ${x.score>0?'score-positive':x.score<0?'score-negative':'score-neutral'}">${x.score}</span></td></tr>`).join(''):'<tr><td colspan="7" class="empty-state">条件に該当する活動はありません</td></tr>'}
function renderInsight(data,avg){if(!data.length){$('insightText').textContent='条件に該当するデータがありません。絞り込み条件を変更してください。';return}const offices=aggregate(data,'office'),members=aggregate(data,'member');const best=members.slice().sort((a,b)=>b.avg-a.avg)[0];const busiest=offices[0];$('insightText').textContent=`${busiest.name}が${busiest.count}件で活動量トップ。${best.name}は質スコア平均${best.avg.toFixed(1)}ptです。全体平均は${avg.toFixed(1)}pt。`}
function showToast(message){const t=$('toast');t.textContent=message;t.classList.add('show');clearTimeout(showToast.timer);showToast.timer=setTimeout(()=>t.classList.remove('show'),3500)}
$('uploadButton').addEventListener('click',()=>{if($('dropZone').hidden){$('dropZone').hidden=false;$('fileInput').click()}else $('fileInput').click()});$('fileInput').addEventListener('change',e=>e.target.files[0]&&loadFile(e.target.files[0]));$('sampleButton').addEventListener('click',()=>{allData=[...sampleData];initializeFilters();render();$('dataSource').textContent='サンプルデータ';showToast('サンプルデータに戻しました')});['dateFrom','dateTo','memberFilter'].forEach(id=>$(id).addEventListener('change',render));$('officeFilter').addEventListener('change',()=>{const office=$('officeFilter').value;selectedAdviceOffice=office;updateMemberOptions(office);render()});$('resetFilters').addEventListener('click',()=>{selectedAdviceOffice='';initializeFilters();render()});document.querySelectorAll('[data-scroll]').forEach(b=>b.addEventListener('click',()=>$(b.dataset.scroll).scrollIntoView({behavior:'smooth'})));
$('clearEvaluationFilter').addEventListener('click',()=>{activeEvaluationFilter='';$('clearEvaluationFilter').hidden=true;render()});
$('adviceOffice').addEventListener('change',e=>{selectedAdviceOffice=e.target.value;$('officeFilter').value=selectedAdviceOffice;updateMemberOptions(selectedAdviceOffice);clearGptAnalysis();render()});
$('runGptAnalysis').addEventListener('click',runGptAnalysis);
$('runMemberAdvice').addEventListener('click',runMemberAdvice);
const dz=$('dropZone');['dragenter','dragover'].forEach(name=>document.addEventListener(name,e=>{e.preventDefault();dz.hidden=false}));dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('dragging')});dz.addEventListener('dragleave',()=>dz.classList.remove('dragging'));dz.addEventListener('drop',e=>{e.preventDefault();dz.classList.remove('dragging');e.dataTransfer.files[0]&&loadFile(e.dataTransfer.files[0])});
initializeFilters();render();

