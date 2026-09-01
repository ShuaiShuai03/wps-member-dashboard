/* ============================================================
   会员客户台账｜看板数据导出（AirScript · 只读）
   放在任意一个战区管理系统的脚本编辑器里直接运行。
   - 按【字段名】读取，字段顺序被调整不影响结果
   - 只读源表：会员客户台账，不依赖任何统计表、合并表
   - 输出分段打印（每段 700 字），因为脚本日志面板会把超长的单行截断。
     运行完在日志区全选复制（Cmd/Ctrl+A → Cmd/Ctrl+C），
     整段粘进看板的「导入 AirScript 数据」即可，看板会自动拼接。
   ============================================================ */

var LEDGER_KEY = '会员客户台账';
// ZONE 留空则自动从表名推断
var ZONE = '';

var FIELDS = ['订单号','部门','组别','员工','引流渠道','轮次','轮次负责人','客户姓名','客户分类','出单类型',
  '地址区域','转介绍','订单金额','是否对接','是否观课','最新沟通日期','当下客户状态',
  '首轮是否成交','首轮成交金额','二轮是否成交','二轮成交金额','三轮是否成交','三轮成交金额',
  '往期单量','往期金额','总金额','广告费','创建时间'];

var CHUNK = 700;

// 返回所有匹配的表：单战区文档只有一张，总管理系统里每个战区各一张，全部导出
function findSheetsByKey(key){
  var n = Application.Sheets.Count, out = [];
  for (var i = 1; i <= n; i++){
    var s = Application.Sheets.Item(i);
    if (s.Name && s.Name.indexOf(key) >= 0 && s.Name.indexOf('统计') < 0 && s.Name.indexOf('📊') < 0) out.push(s);
  }
  return out;
}

function fieldMap(sheet){
  var m = {};
  sheet.GetFields().forEach(function(f){ m[f.name] = f.id; });
  return m;
}

// 关联/人员/选项字段的单元格是 DBCellValue：{ Value:[{id,str}], display, ... }
function textOf(p){
  if (p == null) return '';
  if (typeof p !== 'object') return String(p);
  return p.str || p.nickname || p.name || p.text || p.title || '';
}

function cellValue(sheet, row, fmap, name){
  var fid = fmap[name];
  if (!fid) return null;
  try {
    var v = sheet.RecordRange(row, fid).Value;
    if (v == null) return null;
    if (typeof v !== 'object') return v;
    if (Array.isArray(v)){
      return v.map(textOf).filter(function(x){ return x !== ''; }).join(',');
    }
    if (Array.isArray(v.Value)){
      var joined = v.Value.map(textOf).filter(function(x){ return x !== ''; }).join(',');
      return joined || (v.display || '');
    }
    if (v.Value !== undefined && typeof v.Value !== 'object') return v.Value;
    return v.display || '';
  } catch (e) { return null; }
}

function guessZone(name){
  var m = String(name || '').match(/(.{1,4}战区)/);
  return m ? m[1] : '';
}

function emit(str){
  var total = Math.ceil(str.length / CHUNK);
  console.log('==== 数据开始：共 ' + total + ' 段，请在日志区全选复制后粘进看板 ====');
  for (var i = 0; i < total; i++){
    console.log('@@' + (i + 1) + '@@' + str.substr(i * CHUNK, CHUNK) + '@@/@@');
  }
  console.log('==== 数据结束 ====');
}

function main(){
  var sheets = findSheetsByKey(LEDGER_KEY);
  if (!sheets.length){ console.log('未找到包含「' + LEDGER_KEY + '」的数据表'); return; }

  var cols = ['战区'].concat(FIELDS);
  var rows = [], srcs = [], zones = [];
  sheets.forEach(function(sheet){
    var zone = ZONE || guessZone(sheet.Name) || '本战区';
    if (zones.indexOf(zone) < 0) zones.push(zone);
    var fmap = fieldMap(sheet);
    var total = sheet.RecordRange().Count, got = 0;
    for (var r = 1; r <= total; r++){
      var rec = [zone], hasData = false;
      for (var i = 0; i < FIELDS.length; i++){
        var v = cellValue(sheet, r, fmap, FIELDS[i]);
        rec.push(v === undefined ? null : v);
        // 创建时间是系统自动填的，不能当作"这一行有数据"的依据，否则空行也会被导出
        if (FIELDS[i] !== '创建时间' && v !== null && v !== '' && v !== 0) hasData = true;
      }
      if (hasData){ rows.push(rec); got++; }
    }
    srcs.push(sheet.Name + '(' + got + '条)');
  });

  var payload = {
    meta: { zone: zones.join('、') || '本战区', generatedAt: new Date().toISOString(), source: srcs.join(' + ') },
    cols: cols,
    rows: rows
  };
  emit(JSON.stringify(payload));
  console.log('共 ' + rows.length + ' 条记录');
  return payload;
}

main();
