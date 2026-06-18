const http = require('http');
const fs = require('fs');
const path = require('path');
const Anthropic = require('@anthropic-ai/sdk');

const port = Number(process.env.PORT || 8000);
const root = __dirname;
const modelId = 'claude-opus-4-8';
const mimeTypes = {'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.md':'text/markdown; charset=utf-8','.json':'application/json; charset=utf-8','.png':'image/png','.svg':'image/svg+xml'};

function json(res,status,body){res.writeHead(status,{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'});res.end(JSON.stringify(body))}
function validNumber(value,min=-10000,max=10000){return typeof value==='number'&&Number.isFinite(value)&&value>=min&&value<=max}
function validatePayload(body){
  if(!body||typeof body!=='object'||Array.isArray(body))return false;
  if(!body.period||!/^\d{4}-\d{2}-\d{2}$/.test(body.period.from)||!/^\d{4}-\d{2}-\d{2}$/.test(body.period.to))return false;
  if(body.analysis_scope&&!['office','member'].includes(body.analysis_scope))return false;
  if(!Number.isInteger(body.comparison_offices)||body.comparison_offices<1||body.comparison_offices>1000)return false;
  if(!body.volume||!validNumber(body.volume.visits,0)||!validNumber(body.volume.average_visits,0)||!validNumber(body.volume.visits_per_member,0))return false;
  if(!validNumber(body.quality_score?.office,-100,100)||!validNumber(body.quality_score?.average,-100,100))return false;
  if(!Array.isArray(body.amtul)||body.amtul.length!==4||!Array.isArray(body.quality_metrics)||body.quality_metrics.length!==6)return false;
  const amtulCodes=new Set(body.amtul.map(x=>x.code)),metricIds=new Set(body.quality_metrics.map(x=>x.id));
  const expectedAmtul=['A','M','T','U'],expectedMetrics=['needs','competitor','patient','nextAction','intent','noReaction'];
  return expectedAmtul.every(x=>amtulCodes.has(x))&&expectedMetrics.every(x=>metricIds.has(x))&&body.amtul.every(x=>validNumber(x.office_rate,0,100)&&validNumber(x.average_rate,0,100))&&body.quality_metrics.every(x=>validNumber(x.office_rate,0,100)&&validNumber(x.average_rate,0,100));
}
async function analyze(req,res){
  if(!process.env.ANTHROPIC_API_KEY)return json(res,503,{error:'ANTHROPIC_API_KEYが設定されていません。'});
  let raw='';for await(const chunk of req){raw+=chunk;if(raw.length>30000)return json(res,413,{error:'送信データが大きすぎます。'});}
  let body;try{body=JSON.parse(raw)}catch{return json(res,400,{error:'JSON形式が正しくありません。'})}
  if(!validatePayload(body))return json(res,400,{error:'匿名化集計の形式が正しくありません。'});
  const systemPrompt=[
    'あなたは医薬品営業組織のマネジメント支援アドバイザーです。匿名化された集計値だけを分析してください。',
    'analysis_scopeがofficeの場合は営業所向け、memberの場合は担当者向けの助言として書いてください。',
    'analysis_scopeがmemberの場合、「当該営業所」ではなく「当該担当者」「この担当者」と表現し、比較基準は同じ営業所内の担当者平均として扱ってください。',
    'analysis_scopeがofficeの場合、比較基準は全営業所平均として扱ってください。',
    '活動量だけを増やす提案ではなく、活動量、AMTUL構成、面談品質、非面談率の相互関係を読み取ってください。',
    '相関を因果関係として断定せず、営業所規模や担当市場の違いも考慮した慎重な日本語で回答してください。',
    '現場で翌月に実行できる具体性を持たせ、所長がMRに確認すべき質問を含めてください。',
    '個人、医師、施設、製品の情報は与えられていません。存在を推測したり架空の固有名詞を作らないでください。',
    '',
    '以下のJSON形式のみで回答してください。JSONの前後にコードブロックや説明文は不要です。',
    '{"executive_comment":"string","observations":["string"],"priority_issues":["string"],"recommended_actions":["string"],"manager_questions":["string"],"next_month_kpis":[{"name":"string","target":"string","reason":"string"}],"caution":"string"}'
  ].join('\n');
  try{
    const client=new Anthropic();
    const message=await client.messages.create({model:modelId,max_tokens:2000,thinking:{type:'adaptive'},system:systemPrompt,messages:[{role:'user',content:`分析対象データ: ${JSON.stringify(body)}`}]});
    const text=message.content.find(b=>b.type==='text')?.text||'';
    if(!text)throw new Error('分析結果を取得できませんでした。');
    const analysis=JSON.parse(text);
    return json(res,200,{analysis,model:modelId});
  }catch(error){
    if(error instanceof SyntaxError)return json(res,502,{error:'分析結果のJSON解析に失敗しました。'});
    return json(res,502,{error:error.message||'Claude分析に失敗しました。'});
  }
}
function serveStatic(req,res){
  const pathname=decodeURIComponent(new URL(req.url,'http://localhost').pathname);const requested=pathname==='/'?'index.html':pathname.slice(1);const filePath=path.resolve(root,requested);
  const publicFiles=new Set(['index.html','app.js','style.css']);if(!publicFiles.has(requested)||!filePath.startsWith(root+path.sep))return json(res,404,{error:'Not found'});
  fs.stat(filePath,(error,stat)=>{if(error||!stat.isFile())return json(res,404,{error:'Not found'});res.writeHead(200,{'Content-Type':mimeTypes[path.extname(filePath).toLowerCase()]||'application/octet-stream','Cache-Control':'no-store'});fs.createReadStream(filePath).pipe(res);});
}
const server=http.createServer((req,res)=>{if(req.method==='POST'&&req.url==='/api/gpt-analysis')return analyze(req,res);if(req.method==='GET')return serveStatic(req,res);json(res,405,{error:'Method not allowed'});});
server.listen(port,()=>console.log(`Activity Report Dashboard: http://localhost:${port}`));
