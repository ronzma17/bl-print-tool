/**********************************************************************
 *  Pre-Arrival Notice 抓取工具
 *
 *  挂在一个 Google 表格上运行。打开表格 -> 菜单「提单工具」-> 输入航次号
 *  就会把那一批的 Pre-Arrival Notice 附件抓下来存进 Google Drive,
 *  按提单号重命名。
 *
 *  同一个提单号收到多封时(会话里的回复 / 改单重发),
 *  只下载【邮件时间最新】的那一封。
 *
 *  安装步骤见最下面。
 **********************************************************************/


// ======================= 可调参数 =======================

// 除了航次号之外, 还要加上的固定搜索条件
var BASE_QUERY = 'subject:"Notice of Arrival" has:attachment';

// 提单号格式, 如 CHN3966555 / NGP4259647
var KEY_REGEX = /[A-Z]{2,4}\d{6,}/;

// 一次最多翻多少个会话
var MAX_THREADS = 400;

// =======================================================


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('提单工具')
    .addItem('① 先看看会抓到什么', 'menuPreview')
    .addItem('② 抓下来存到 Drive', 'menuDownload')
    .addToUi();
}


function askVoyage_(title) {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt(
    title,
    '输入这一批的航次号,例如  EXX030W19\n\n' +
    '(想用别的条件, 可以直接粘贴完整 Gmail 搜索语句,\n' +
    ' 比如  from:cma-cgm.com newer_than:3d )',
    ui.ButtonSet.OK_CANCEL);

  if (res.getSelectedButton() !== ui.Button.OK) return null;
  var v = res.getResponseText().trim();
  if (!v) { ui.alert('没输入内容'); return null; }
  return v;
}


function buildQuery_(input) {
  // 带冒号就当成完整搜索语句直接用, 否则当航次号
  if (input.indexOf(':') >= 0) return input;
  return 'subject:' + input + ' ' + BASE_QUERY;
}


/**
 * 按提单号归堆, 每堆按邮件时间从新到旧排好
 */
function collect_(query) {
  var threads = GmailApp.search(query, 0, MAX_THREADS);
  var buckets = {}, noKey = [], total = 0;

  for (var t = 0; t < threads.length; t++) {
    var msgs = threads[t].getMessages();
    for (var m = 0; m < msgs.length; m++) {
      var msg = msgs[m];
      var subject = msg.getSubject() || '';
      total++;
      var hit = subject.toUpperCase().match(KEY_REGEX);
      if (!hit) { noKey.push(subject); continue; }
      var key = hit[0];
      if (!buckets[key]) buckets[key] = [];
      buckets[key].push({ msg: msg, date: msg.getDate(), subject: subject });
    }
  }
  for (var k in buckets) {
    buckets[k].sort(function (a, b) { return b.date - a.date; });
  }
  return { buckets: buckets, noKey: noKey, threads: threads.length, totalMsgs: total };
}


function fmt_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd HH:mm');
}


function summarize_(r) {
  var keys = Object.keys(r.buckets).sort();
  var multi = [];
  for (var i = 0; i < keys.length; i++) {
    var list = r.buckets[keys[i]];
    if (list.length > 1) {
      var times = [];
      for (var j = 0; j < list.length; j++) times.push(fmt_(list[j].date));
      multi.push(keys[i] + '  ' + list.length + ' 封 [' + times.join(' , ') +
                 ']  用最新 ' + fmt_(list[0].date));
    }
  }
  return { keys: keys, multi: multi };
}


// ---------------------------------------------------------------------
// 菜单 ①  只看不抓
// ---------------------------------------------------------------------
function menuPreview() {
  var ui = SpreadsheetApp.getUi();
  var input = askVoyage_('先看看会抓到什么');
  if (!input) return;

  var query = buildQuery_(input);
  var r = collect_(query);
  var s = summarize_(r);

  if (!s.keys.length) {
    ui.alert('一封都没搜到\n\n搜索条件:\n' + query + '\n\n航次号打错了?');
    return;
  }

  writeSheet_(r, s, query, null);

  var txt = '搜索条件\n' + query + '\n\n'
          + '会话 ' + r.threads + ' 个,邮件 ' + r.totalMsgs + ' 封\n'
          + '不重复提单号 ' + s.keys.length + ' 个\n'
          + '同一提单号收到多封 ' + s.multi.length + ' 个\n'
          + '标题里没有提单号 ' + r.noKey.length + ' 封\n\n'
          + '明细已写进表格。数字对的话,再跑菜单 ②。';
  ui.alert('预览结果', txt, ui.ButtonSet.OK);
}


// ---------------------------------------------------------------------
// 菜单 ②  正式抓
// ---------------------------------------------------------------------
function menuDownload() {
  var ui = SpreadsheetApp.getUi();
  var input = askVoyage_('抓下来存到 Drive');
  if (!input) return;

  var query = buildQuery_(input);
  var r = collect_(query);
  var s = summarize_(r);

  if (!s.keys.length) {
    ui.alert('一封都没搜到\n\n搜索条件:\n' + query);
    return;
  }

  var ok = ui.alert(
    '确认',
    '提单号 ' + s.keys.length + ' 个\n'
    + '同一提单号多封 ' + s.multi.length + ' 个(只取最新)\n'
    + '标题无提单号 ' + r.noKey.length + ' 封\n\n'
    + '开始下载?',
    ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;

  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm');
  var safe = input.replace(/[^A-Za-z0-9]/g, '').slice(0, 20) || 'batch';
  var folderName = 'PreArrival_' + safe + '_' + stamp;
  var folder = DriveApp.createFolder(folderName);

  var saved = 0, skipped = 0, noPdf = [];

  for (var i = 0; i < s.keys.length; i++) {
    var key = s.keys[i];
    var list = r.buckets[key];
    var got = false;

    // 从最新往旧找, 第一个带 PDF 的就用它
    for (var j = 0; j < list.length && !got; j++) {
      var atts = list[j].msg.getAttachments({ includeInlineImages: false });
      for (var a = 0; a < atts.length; a++) {
        if (atts[a].getName().toLowerCase().slice(-4) !== '.pdf') continue;
        folder.createFile(atts[a].copyBlob()).setName(key + '.pdf');
        list[j].usedFile = atts[a].getName();
        list[j].used = true;
        saved++;
        skipped += (list.length - 1);
        got = true;
        break;
      }
    }
    if (!got) noPdf.push(key);
  }

  writeSheet_(r, s, query, folderName);

  var txt = '完成\n\n'
          + '存了 ' + saved + ' 个 PDF\n'
          + '因为有更新版本而跳过 ' + skipped + ' 封\n'
          + '一个附件都没有 ' + noPdf.length + ' 个\n'
          + '标题无提单号 ' + r.noKey.length + ' 封\n\n'
          + 'Drive 文件夹:\n' + folderName + '\n\n'
          + '下一步: 打开 Drive,右键这个文件夹 -> 下载,\n'
          + '把 zip 丢进 4_DROP_PREARRIVAL_HERE';
  ui.alert('完成', txt, ui.ButtonSet.OK);
}


// ---------------------------------------------------------------------
// 把明细写进表格
// ---------------------------------------------------------------------
function writeSheet_(r, s, query, folderName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = '结果';
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();

  var rows = [];
  rows.push(['搜索条件', query]);
  rows.push(['时间', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')]);
  rows.push(['Drive 文件夹', folderName || '(只是预览,没有下载)']);
  rows.push([]);
  rows.push(['提单号', '共收到几封', '用了哪封', '下载的文件', '邮件标题']);

  for (var i = 0; i < s.keys.length; i++) {
    var key = s.keys[i];
    var list = r.buckets[key];
    var pick = null;
    for (var j = 0; j < list.length; j++) if (list[j].used) { pick = list[j]; break; }
    if (!pick) pick = list[0];
    rows.push([
      key,
      list.length,
      fmt_(pick.date),
      folderName ? (pick.used ? key + '.pdf' : '没有 PDF 附件') : '',
      pick.subject
    ]);
  }

  if (r.noKey.length) {
    rows.push([]);
    rows.push(['标题里找不到提单号的邮件']);
    for (var n = 0; n < r.noKey.length; n++) rows.push([r.noKey[n]]);
  }

  sh.getRange(1, 1, rows.length, 5).setValues(
    rows.map(function (x) { while (x.length < 5) x.push(''); return x; }));
  sh.getRange(5, 1, 1, 5).setFontWeight('bold');
  sh.setFrozenRows(5);
  for (var c = 1; c <= 5; c++) sh.autoResizeColumn(c);
}


/**********************************************************************
 *  【安装步骤】只做一次
 *
 *  1. 打开 sheets.google.com, 新建一个空白表格,
 *     命名成「提单工具」之类, 收藏起来
 *
 *  2. 表格里点顶部菜单「扩展程序」->「Apps 脚本」
 *
 *  3. 把编辑器里原有内容全删掉, 粘贴本文件全部内容, Ctrl+S 保存
 *
 *  4. 回到表格, 刷新页面(F5)
 *     顶部会多出一个菜单「提单工具」
 *
 *  5. 第一次点菜单会要授权:
 *       需要授权 -> 审核权限 -> 选账号
 *       -> 「Google 尚未验证此应用」-> 左下「高级」
 *       -> 「转至 xxx(不安全)」-> 允许
 *     (脚本是你自己的, 只在你账号里跑, 不外发数据)
 *
 *
 *  【每批怎么用】
 *
 *  1. 打开那个表格
 *  2. 菜单「提单工具」->「① 先看看会抓到什么」-> 输入航次号 EXX030W19
 *  3. 看弹窗数字对不对, 表格「结果」页有每一票的明细
 *  4. 菜单「提单工具」->「② 抓下来存到 Drive」-> 同样输入航次号 -> 确认
 *  5. 打开 Google Drive, 找到 PreArrival_EXX030W19_日期 文件夹
 *  6. 右键文件夹 -> 下载, 得到一个 zip
 *  7. zip 丢进 4_DROP_PREARRIVAL_HERE, 跑 START_HERE.bat 选 4 再选 5
 *
 *  不用改代码, 不用记搜索语法, 每批就是换个航次号。
 **********************************************************************/
