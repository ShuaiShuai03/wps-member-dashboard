/* ============================================================
   会员客户台账｜看板数据导出（AirScript · 只读）
   放在任意一个战区管理系统的脚本编辑器里直接运行。
   - 按【字段名】读取，字段顺序被调整不影响结果
   - 只读源表：会员客户台账，不依赖任何统计表、合并表
   - 运行后把控制台打印的 JSON 整段复制，粘贴进网页看板
   ============================================================ */

var LEDGER_KEY = '会员客户台账';
var ZONE = '';   // 留空自动从表名推断

var FIELDS = ['订单号','部门','组别','员工','引流渠道','轮次','轮次负责人','客户姓名','客户分类','出单类型',
  '地址区域','转介绍','订单金额','是否对接','是否观课','最新沟通日期','当下客户状态',
  '首轮是否成交','首轮成交金额','二轮是否成交','二轮成交金额','三轮是否成交','三轮成交金额',
  '往期单量','往期金额','总金额','广告费','创建时间'];

function findSheetByKey(key){
  var n = Application.Sheets.Count;
  for (var i = 1; i <= n; i++){
    var s = Application.Sheets.Item(i);
    if (s.Name && s.Name.indexOf(key) >= 0 && s.Name.indexOf('统计') < 0 && s.Name.indexOf('📊') < 0) return s;
  }
  return null;
}
function fieldMap(sheet){
  var m = {};
  sheet.GetFields().forEach(function(f){ m[f.name] = f.id; });
  return m;
}
function cellValue(sheet, row, fmap, name){
  var fid = fmap[name];
  if (!fid) return null;
  try {
    var v = sheet.RecordRange(row, fid).Value;
    if (v && typeof v === 'object'){
      if (Array.isArray(v)) return v.map(function(p){ return (p && (p.nickname||p.name||p.text)) || p; }).join(',');
      if (Array.isArray(v.Value)) return v.Value.map(function(p){ return (p && (p.nickname||p.name||p.text||p.userId)) || p; }).join(',');
      if (v.Value !== undefined) return v.Value;
      return null;
    }
    return v;
  } catch (e) { return null; }
}
function guessZone(name){
  var m = String(name||'').match(/(.{1,4}战区)/);
  return m ? m[1] : '';
}

function main(){
  var sheet = findSheetByKey(LEDGER_KEY);
  if (!sheet){ console.log('⚠️ 未找到包含「' + LEDGER_KEY + '」的数据表'); return; }
  var zone = ZONE || guessZone(sheet.Name) || '本战区';
  var fmap = fieldMap(sheet);
  var total = sheet.RecordRange().Count;
  var detail = [];
  for (var r = 1; r <= total; r++){
    var o = { 战区: zone }, hasData = false;
    for (var i = 0; i < FIELDS.length; i++){
      var v = cellValue(sheet, r, fmap, FIELDS[i]);
      o[FIELDS[i]] = v;
      if (v !== null && v !== '' && v !== 0) hasData = true;
    }
    if (hasData) detail.push(o);
  }
  var payload = {
    meta: { zone: zone, generatedAt: new Date().toISOString(), source: sheet.Name + '(' + detail.length + '条)' },
    detail: detail
  };
  console.log('==== 复制下面这一整段 JSON ====');
  console.log(JSON.stringify(payload));
  console.log('==== 共 ' + detail.length + ' 条记录 ====');
  return payload;
}

main();
