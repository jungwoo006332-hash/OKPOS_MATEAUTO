(function(root){
const clean=v=>String(v??'').replace(/\s/g,'');
function date(v){if(typeof v==='number'){const d=new Date(Math.round((v-25569)*86400000));return d.toISOString().slice(0,10)} const m=String(v??'').match(/(20\d{2})[.\-/년]\s*(\d{1,2})[.\-/월]\s*(\d{1,2})/);return m?`${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`:null}
function hour(v){if(typeof v==='number')return Math.floor((v%1)*24+1e-7);let m=String(v??'').match(/(?:^|\s)(\d{1,2}):\d{2}/);return m&&+m[1]<24?+m[1]:null}
function parse(rows,file,sheet){
 const hi=rows.findIndex(r=>r.some(v=>clean(v)==='영수증번호')&&r.some(v=>clean(v)==='결제시각'));
 if(hi<0)return null;const h=rows[hi].map(clean);const col=(...names)=>h.findIndex(x=>names.includes(x));
 const c={pos:col('포스번호','POS번호'),receipt:col('영수증번호'),type:col('구분'),table:col('테이블명'),order:col('최초주문','최초주문시간','최초주문시각'),pay:col('결제시각','결제시간'),item:col('상품코드'),amount:col('실매출액'),day:col('영업일자','매출일자','일자'),store:col('매장명')};
 if([c.pos,c.order,c.amount].some(x=>x<0))throw Error('포스번호·최초주문·실매출액 열을 확인해주세요.');
 const meta=rows.slice(0,hi).flat().filter(Boolean).join(' ');const dates=[...meta.matchAll(/20\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2}/g)].map(m=>date(m[0]));
 const base=[...new Set(dates)];if(base.length>1&&c.day<0)throw Error('여러 날짜가 조회되었지만 일자 열이 없습니다. 일자별 파일로 올려주세요.');
 const fallback=base[0];const store=(meta.match(/검색대상\s*:\s*(.+?)(?:\s{2,}|$)/)?.[1]||'매장 미표기').trim();
 let pos='',day=fallback,current=null;const out=[];
 for(const r of rows.slice(hi+1)){
  if(r.some(v=>/^(합계|총합계|소계)$/.test(clean(v)))){current=null;continue}
  if(c.day>=0&&r[c.day])day=date(r[c.day]);if(r[c.pos])pos=String(r[c.pos]);
  if(r[c.receipt]){if(!day)throw Error('영업일자를 확인할 수 없습니다. 조회일자 또는 영업일자 열이 필요합니다.');
   const type=String(r[c.type]||'매출');current={day,store:c.store>=0?String(r[c.store]||store):store,pos,receipt:String(r[c.receipt]),type,table:String(r[c.table]||''),order:hour(r[c.order]),pay:hour(r[c.pay]),amount:0,source:file,sheet,lines:0};
   current.key=[current.store,day,pos,current.receipt,type].join('|');out.push(current);
  }
  if(current&&r[c.amount]!==null&&r[c.amount]!==undefined&&r[c.amount]!==''){const a=Number(String(r[c.amount]).replace(/,/g,''));if(!Number.isFinite(a))throw Error('실매출액에 숫자가 아닌 값이 있습니다.');current.amount+=a;current.lines++}
 }
 return out;
}
const hours=[15,16,17,18,19,20,21,22,23,0,1,2,3,4];
function summarize(records,selectedHours=hours){return selectedHours.map(h=>({hour:h,orders:records.filter(r=>r.order===h&&r.type==='매출').reduce((n,r)=>n+(r.count??1),0),payments:records.filter(r=>r.pay===h&&r.type==='매출').reduce((n,r)=>n+(r.count??1),0),amount:records.filter(r=>r.order===h).reduce((s,r)=>s+r.amount,0)}))}
function parseMate(rows,file,sheet,store){
 const hi=rows.findIndex(r=>r.some(v=>clean(v)==='시간대')&&r.some(v=>clean(v)==='총건수'));
 if(hi<0)return null;
 if(!store?.trim())throw Error('메이트 자료의 매장명을 입력한 후 다시 올려주세요.');
 const m=file.match(/(?:^|[^0-9])(20\d{2}|\d{2})[.\-_](\d{2})[.\-_](\d{2})(?=\D|$)/);
 if(!m)throw Error('날짜가 없는 메이트 자료입니다. 파일명을 D26.09.01.xlsx 형식으로 변경해주세요.');
 const year=m[1].length===2?'20'+m[1]:m[1],day=`${year}-${m[2]}-${m[3]}`;
 if(new Date(day+'T00:00:00Z').toISOString().slice(0,10)!==day)throw Error('파일명의 날짜를 확인해주세요.');
 const group=rows[hi-1]||[],header=rows[hi].map(clean),channels=['배달','포장'];
 const indices=channels.map(channel=>{const c=group.findIndex(v=>clean(v)===channel);if(c<0||header[c]!=='건수'||header[c+1]!=='실매출액')throw Error('메이트 배달·포장 열 구성을 확인해주세요.');return {channel,c};});
 const out=[];for(const r of rows.slice(hi+1)){const t=String(r[0]??'').trim().match(/^(\d{2})\s*[-~]\s*\d{2}$/);if(!t)continue;const h=Number(t[1])===24?0:Number(t[1]);if(h>23)throw Error('잘못된 시간대입니다.');
 for(const {channel,c} of indices){const count=Number(String(r[c]??0).replace(/,/g,'')),amount=Number(String(r[c+1]??0).replace(/,/g,''));if(!Number.isInteger(count)||count<0||!Number.isFinite(amount))throw Error('건수·매출액에 유효하지 않은 값이 있습니다.');
 out.push({key:['MATE',store.trim(),day,h,channel].join('|'),day,store:store.trim(),pos:'MATE',receipt:'시간대 합계',table:channel,type:'매출',order:h,pay:h,amount,count,channel,system:'MATE',source:file,sheet,lines:1});}}
 if(!out.length)throw Error('메이트 시간대별 자료가 없습니다.');return out;
}
function label(h){return h===0?'24시':String(h).padStart(2,'0')+'시'}
const allHours=[...hours,5,6,7,8,9,10,11,12,13,14];
const api={parse,parseMate,hour,date,hours,allHours,label,summarize};if(typeof module!=='undefined')module.exports=api;root.OKPOS=api;
})(globalThis);
