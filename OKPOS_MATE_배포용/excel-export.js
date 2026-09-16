(function(root){
// Export the same filtered records and hourly aggregation used by the dashboard.
function build(records,filter={}){
 if(!records.length)throw Error('내려받을 데이터가 없습니다.');
 const X=root.XLSX,P=root.OKPOS,H=filter.hours||P.hours,wb=X.utils.book_new(),days=[...new Set(records.map(r=>r.day))].sort(),daily=days.map(day=>({day,hours:P.summarize(records.filter(r=>r.day===day),H)}));
 const add=(name,rows,widths)=>{const ws=X.utils.aoa_to_sheet(rows);ws['!cols']=widths.map(wch=>({wch}));for(const key of Object.keys(ws)){if(ws[key]?.t==='n')ws[key].z='#,##0;[Red]-#,##0';}X.utils.book_append_sheet(wb,ws,name);return ws;};
 for(const [key,name,unit] of [['orders','최초주문 건수','건'],['payments','결제 건수','건'],['amount','최초주문 기준 매출','원']]){
  const header=['영업일자',...H.map(h=>P.label(h)),'합계'];
  const rows=[[name+' (단위: '+unit+')'],['매장',filter.store||'전체 매장'],['조회 기간',filter.from||days[0],filter.to||days.at(-1)],[],header];
  for(const d of daily){const nums=d.hours.map(x=>x[key]);rows.push([d.day,...nums,nums.reduce((a,b)=>a+b,0)]);}
  const totals=H.map((_,i)=>daily.reduce((a,d)=>a+d.hours[i][key],0));rows.push(['전체 합계',...totals,totals.reduce((a,b)=>a+b,0)]);
  rows.push([],['집계 기준',key==='orders'?'홀 테이블 수 + 메이트 배달·포장 건수 · 최초주문 기준':key==='payments'?'홀은 결제시각, 메이트 배달·포장은 주문시각 기준':'홀 최초주문 시각 + 메이트 주문시각 · 실매출액 합계 (반품은 원본 부호 반영)'],['시간 범위',H.length===24?'전체 24시간 · 00-01 = 24시 · 메이트 날짜는 파일명 기준':'15:00~다음 날 04:59 · 00-01 = 24시 · 메이트 날짜는 파일명 기준'],['자료 구분','OKPOS: 매장·영업일·포스·영수증 / MATE: 매장·파일 날짜·시간대·배달/포장']);
  const ws=add(name,rows,[21,...H.map(()=>14),16]);ws['!autofilter']={ref:`A5:${X.utils.encode_col(H.length+1)}${5+daily.length}`};ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:H.length+1}},...Array.from({length:3},(_,i)=>({s:{r:rows.length-3+i,c:1},e:{r:rows.length-3+i,c:H.length+1}}))];
 }
 const outside=records.filter(r=>!H.includes(r.order)||!H.includes(r.pay)||r.type!=='매출');
 if(outside.length){add('집계 확인',[
 ['시간 범위 밖·시간 미확인·매출 외 영수증'],['안내','범위 밖 시각은 해당 지표에서 제외합니다. 매출 외 구분은 건수에서 제외하고 금액은 원본 부호로 반영합니다.'],[],
 ['영업일자','매장','포스','영수증번호','구분','최초주문 시','결제 시','실매출액','주문 집계','결제 건수 집계','매출 집계'],
 ...outside.map(r=>[r.day,r.store,r.pos,r.receipt,r.type,r.order??'미확인',r.pay??'미확인',r.amount,r.type==='매출'&&H.includes(r.order)?'포함':'제외',r.type==='매출'&&H.includes(r.pay)?'포함':'제외',H.includes(r.order)?'포함':'제외'])
 ],[16,30,10,16,12,16,16,18,16,18,16]);}
 add('출처별 집계',[['영업일자','구분','주문 건수','결제 건수','주문 기준 매출'],...days.flatMap(day=>[['OKPOS 홀',r=>r.system!=='MATE'],['MATE 배달',r=>r.channel==='배달'],['MATE 포장',r=>r.channel==='포장']].map(([name,predicate])=>{const a=P.summarize(records.filter(r=>r.day===day&&predicate(r)),H);return [day,name,a.reduce((n,x)=>n+x.orders,0),a.reduce((n,x)=>n+x.payments,0),a.reduce((n,x)=>n+x.amount,0)]}))],[18,20,18,18,24]);
 wb.Props={Title:'OKPOS + MATE 일자별 시간대 분석',Subject:'최초주문 및 결제 분리 집계'};return wb;
}
root.OKPOSExcel={build};if(typeof module!=='undefined')module.exports={build};
})(globalThis);
