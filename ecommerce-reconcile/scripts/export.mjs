import fs from 'node:fs/promises';
import path from 'node:path';
import { SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const moneyFmt = '#,##0.00;[Red](#,##0.00);0.00';
const font = 'Microsoft YaHei';
const [mode, filename, flag] = process.argv.slice(2);
const obj = JSON.parse((await fs.readFile(filename, 'utf8')).replace(/^\uFEFF/, ''));
const col = n => { let s=''; for(n++; n; n=Math.floor((n-1)/26)) s=String.fromCharCode(65+(n-1)%26)+s; return s; };
let tableCount=0;
const literal = v => typeof v === 'string' && v.startsWith('=') ? "'"+v : v;

function table(wb, name, headers, rows, widths, start=1) {
  const s = wb.worksheets.add(name);
  s.showGridLines = false;
  const matrix = [headers, ...rows].map(r => headers.map((_, i) => literal(r[i] ?? null)));
  s.getRangeByIndexes(start-1, 0, matrix.length, headers.length).values = matrix;
  const area = s.getRangeByIndexes(0, 0, start+matrix.length, headers.length);
  area.format.font = {name:font, size:10, color:'#202B3C'};
  area.format.rowHeight = 26;
  area.format.verticalAlignment = 'center';
  widths.forEach((width, i) => s.getRange(`${col(i)}1:${col(i)}${start+matrix.length}`).format.columnWidth=width);
  const header = s.getRangeByIndexes(start-1, 0, 1, headers.length);
  header.format.fill = '#243A56';
  header.format.font = {name:font, size:10, color:'#FFFFFF', bold:true};
  header.format.horizontalAlignment = 'center';
  if (rows.length) {
    const t = s.tables.add(`A${start}:${col(headers.length-1)}${start+rows.length}`, true, `T${++tableCount}`);

    t.showFilterButton = true;
  }
  s.freezePanes.freezeRows(start);
  return s;
}

async function save(wb, dest, previews=[]) {
  wb.recalculate();
  await fs.mkdir(path.dirname(dest), {recursive:true});
  const x = await SpreadsheetFile.exportXlsx(wb);
  await x.save(dest);
  for (const [sheetName, range, name] of previews) {
    const img = await wb.render({sheetName, range, scale:1.5, format:'png'});
    await fs.writeFile(path.join(path.dirname(dest), name+'.png'), new Uint8Array(await img.arrayBuffer()));
  }
}

if (mode === 'fixtures') {
  for (const book of obj) {
    const w = Workbook.create();
    const s = table(w, book.sheet ?? '数据', book.headers, book.rows, book.headers.map(h=>h.endsWith('号')?26:20));
    for (let c=0; c<book.headers.length; c++) {
      const h = book.headers[c];
      const r = s.getRange(`${col(c)}2:${col(c)}${Math.max(2,book.rows.length+1)}`);
      if (h.endsWith('号')) r.setNumberFormat('@');
      else if (h.includes('金额') || h.includes('成本') || h==='成本冲回') r.setNumberFormat(moneyFmt);
    }
    for (const p of book.formulas ?? []) s.getRange(p.cell).formulas=[[p.formula]];
    if (book.extra_sheet) w.worksheets.add(book.extra_sheet);
    await save(w, book.path, book.preview ? [['数据',`A1:${col(book.headers.length-1)}${Math.min(7,book.rows.length+1)}`, '模板预览']] : []);
  }
  console.log(`Exported ${obj.length} fixture/template workbooks`);
} else if (mode === 'report') {
  const r=obj, w=Workbook.create(), dest=path.join(path.dirname(filename),'经营利润报表.xlsx');
  const yuan=c => c/100; // Display conversion only; every monetary calculation already used integer cents.
  const t=r.totals_cents;
  const summary = [
    ['本次核对状态',r.status,'完整仅指已上传数据通过本版规则校验'],
    ['核对截止日',r.as_of,'订单批次口径，含已提供的关联退款与费用'],
    ['付款日期范围',r.payment_range.join(' 至 '),'此范围不代表自然月权责利润'],
    ['原始订单行数',r.source_order_rows,'重复行仍保留在原始文件和审计记录中'],
    ['不同订单号数量',r.unique_orders,`另有 ${r.missing_id_order_rows} 行订单号缺失或格式无效`],
    ['可计算订单数量',r.eligible_orders,'其他订单须先补齐或核实'],
    ['阻断问题数',r.issue_counts['阻断']??0,'一笔订单可能有多个问题；详见异常清单'],
    ['提醒数',r.issue_counts['提醒']??0,'提醒金额仍按原值计算'],
    ['以下仅为可计算订单小计',r.complete?'已上传批次通过校验':'不能作为全店利润','金额单位：人民币元'],
    ['实收金额',yuan(t.paid),'成功收款，扣退款与平台费用之前'],
    ['减：成功退款',yuan(t.refund),'处理中和失败退款未扣减'],
    ['减：平台费用',yuan(t.fee),'按明确的订单费用记录汇总'],
    ['减：商品成本',yuan(t.goods),'本批次可计算订单的商品总成本'],
    ['减：履约成本',yuan(t.fulfillment),'商家承担的运费、包装等'],
    ['加：成本冲回',yuan(t.recovery),'仅采用商家明确提供的冲回金额'],
    ['订单经营贡献',r.eligible_orders?yuan(t.profit):null,r.eligible_orders?'税费、人工、租金、广告等未纳入':'无可计算订单，利润不可用'],
  ];
  const s = table(w,'经营汇总',['指标','结果','说明'],summary,[30,40,68],4);
  s.getRange('A2').values=[['订单经营核对']];
  s.getRange('A2').format.font={name:font,size:16,bold:true,color:'#243A56'};
  s.getRange('B14:B20').setNumberFormat(moneyFmt);
  s.getRange('A20:C20').format.fill='#E8EEF5';
  s.getRange('A20:C20').format.font={name:font,size:11,bold:true};
  s.getRange('B5').conditionalFormats.add('containsText',{text:'不完整',format:{fill:'#FCE8E6',font:{color:'#9C2924',bold:true}}});
  s.freezePanes.unfreeze();
  const keys=['paid','refund','fee','goods','fulfillment','recovery','profit'];
  const detailRows=r.details.map(d=>[d.order_id,d.source_rows.join(', '),d.status,
    ...keys.map(k=>d.amounts_cents ? yuan(d.amounts_cents[k]) : null),
    d.references.refunds.join(', '),d.references.fees.join(', '),d.references.costs.join(', '),d.blockers.join('；')]);
  const d=table(w,'订单明细',['订单号','订单源行','状态','实收','成功退款','平台费用','商品成本','履约成本','成本冲回','经营贡献','退款源行','费用源行','成本源行','阻断原因'],detailRows,[24,14,12,16,16,16,16,16,16,16,16,16,16,60]);
  if(detailRows.length) {
    d.getRange(`D2:J${detailRows.length+1}`).setNumberFormat(moneyFmt);
    d.getRange(`C2:C${detailRows.length+1}`).conditionalFormats.add('containsText',{text:'待核实',format:{fill:'#FFF0CE',font:{color:'#865A00'}}});
    d.getRange(`J2:J${detailRows.length+1}`).conditionalFormats.add('cellIs',{operator:'lessThan',formula:0,format:{font:{color:'#B42318',bold:true}}});
  }
  d.freezePanes.freezeColumns(1);
  const issueRows=r.issues.map((i,n)=>[n+1,i.severity,i.code,i.file,i.sheet,i.row||'文件/订单级',i.order_id,i.field,i.raw_value,i.detail,i.action]);
  const a=table(w,'异常清单',['序号','级别','规则代码','来源文件','工作表','Excel行号','订单号','字段','原始值','问题说明','处理建议'],issueRows,[8,10,34,22,12,18,24,18,28,64,64]);
  if(issueRows.length) a.getRange(`B2:B${issueRows.length+1}`).conditionalFormats.add('containsText',{text:'阻断',format:{fill:'#FCE8E6',font:{color:'#9C2924',bold:true}}});
  const instructions=[...r.rules,['引擎版本',r.version],['金额复核','审计记录.json 中所有金额为整数分。Excel 仅展示元，报表无可执行公式。'],
    ...Object.values(r.sources).map(x=>[x.file+ ' SHA-256',x.sha256]),
    ...Object.entries(r.row_counts).map(([k,v])=>[k+' 行去向',JSON.stringify(v)]),
    ['行去向说明','计入 + 隔离待核实 + 未成功退款 = 全部非空来源行；空白行另见审计记录。']];
  const e=table(w,'计算说明',['项目','规则 / 证据'],instructions,[34,118]);
  e.getRange(`B2:B${instructions.length+1}`).format.wrapText=true;
  e.getRange(`A2:B${instructions.length+1}`).format.rowHeight=44;
  e.getRange('A1:B1').format.rowHeight=26;
  w.recalculate();
  const inspect=await w.inspect({kind:'table',range:'经营汇总!A5:C20',include:'values,formulas',tableMaxRows:16,tableMaxCols:3,maxChars:3500});
  await fs.writeFile(path.join(path.dirname(filename),'导出检查.jsonl'),inspect.ndjson);
  const previews=flag==='--render' ? [
    ['经营汇总','A1:C20','预览_经营汇总'],
    ['订单明细',`A1:J${Math.min(8,detailRows.length+1)}`,'预览_订单金额'],
    ['订单明细',`K1:N${Math.min(8,detailRows.length+1)}`,'预览_订单追溯'],
    ['异常清单',`A1:I${Math.max(2,Math.min(7,issueRows.length+1))}`,'预览_异常定位'],
    ['异常清单',`J1:K${Math.max(2,Math.min(7,issueRows.length+1))}`,'预览_异常说明'],
    ['计算说明','A1:B10','预览_计算说明']
  ]:[];
  await save(w,dest,previews);
  console.log(dest);
} else throw new Error('Use fixtures or report');



