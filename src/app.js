(function(){
'use strict';

/* ================= 常數 ================= */
var CATS = {crime:'犯罪事實', sentence:'科刑資料'};
var CAT_ORDER = ['crime','sentence'];
var PART_TITLE = {
  crime:'有關犯罪事實之證據調查（包括兼為科刑資料調查之部分）',
  sentence:'科刑資料之調查'
};
var DEFAULT_FACT = {crime:'全部犯罪事實。', sentence:'量刑審酌事項。'};
var METHODS = ['record','doc','media','exhibit','witness'];
var METHOD_LABEL = {record:'筆錄', doc:'文書', media:'錄音錄影', exhibit:'證物', witness:'證人'};
var METHOD_TITLE = {
  record:'依國民法官法第74條第3項、第1項，以告以要旨之方式調查筆錄',
  doc:'依國民法官法第74條第3項、第1項，以告以要旨之方式調查其他可為證據之文書',
  media:'依國民法官法第75條第2項，以適當設備顯示聲音、影像之方式，調查錄音、錄影、電磁紀錄',
  exhibit:'依國民法官法第76條第1項，以提示供辨認之方式調查證物',
  witness:'依國民法官法第73條第1項，傳喚證人或鑑定人'
};
var ASK_TITLE = '依刑事訴訟法第163條第1項本文及國民法官法施行細則第155條第1項詢問被告';
var SUB_NUM = ['㈠','㈡','㈢','㈣','㈤','㈥','㈦','㈧','㈨','㈩'];
var PART_NUM = ['一','二','三','四'];
var RULINGS = ['grant','sentonly','reserve','deny'];
var RUL_LABEL = {grant:'准許', sentonly:'僅准作科刑資料', reserve:'保留', deny:'駁回'};
var RUL_BTN = {grant:'准許', sentonly:'僅准科刑', reserve:'保留', deny:'駁回'};
var STORE_KEY = 'evtool.autosave.v1';
var FILE_VERSION = 1;

/* ================= 狀態 ================= */
var S = {items:[], askDef:{crime:true, sentence:true}, fontPt:14, tableCm:16, source:'', locked:false, lockMax:0, lockedAt:'', onlyNew:false};
var TEXT_W = 9070; // A4 扣除左右邊界 2.5 公分後的版心寬度（twips，約 16 公分）
var UI = {filter:'all', search:'', sel:new Set(), tab:'apply', cmpFilter:'all', rulFilter:'all', rulSel:new Set(), nums:{}};
var seq = 1;
function newId(){ return 'e' + (Date.now().toString(36)) + (seq++); }

function makeItem(no, name, opts){
  opts = opts || {};
  var it = {id:newId(), no:no||'', name:name||'', apply:null, cat:null,
    method:suggestMethod(name||''), methodAuto:true, fact:'', factAuto:true,
    lockedNo:null, ruling:null, rulingNote:''};
  if (opts.apply !== undefined) it.apply = opts.apply;
  if (opts.manual) it.manual = true;
  return it;
}
function factOf(it){
  if (!it.factAuto) return it.fact;
  return it.cat ? DEFAULT_FACT[it.cat] : '';
}
function statusOf(it){
  if (it.apply === false) return 'off';
  if (it.apply === true) return it.cat ? it.cat : 'catpend';
  return 'pend';
}

/* ================= 調查方法預選 ================= */
function suggestMethod(name){
  var n = String(name).replace(/[\s　]/g,'');
  if (!n) return 'doc';
  n = normCjk(n);
  var base = n;
  for (var i=0;i<4;i++){
    var b2 = base.replace(/([（(][^（）()]*[)）])$/,'').replace(/[0-9.\/\-]+$/,'');
    if (b2 === base) break;
    base = b2;
  }
  var endsLetter = /函(稿)?$|函[（(]稿[)）]$/.test(n) || /函$/.test(base);
  if (/筆錄/.test(n) && !/(勘驗|相驗|勘察|勘查|搜索|扣押|履勘|解剖|檢驗)筆錄/.test(n) && !endsLetter) return 'record';
  if (/(函|報告|書|表|單|據|狀|證明|資料|結果|紀錄|記錄|明細|清冊|清單|截圖|擷圖|照片|相片|譯文|條|簡表|附件|卷宗|病歷|影本|通知|押票|裁定)$/.test(base)) return 'doc';
  if (/^(證人|鑑定人|鑑定證人)/.test(n)) return 'witness';
  if (/(影像|錄影|錄音|音檔|影片|光碟|電磁紀錄|電子檔|檔案)/.test(n) && !/(截圖|擷圖|影像資料查詢)/.test(n)) return 'media';
  if (/(扣押物|扣案物|證物|扣案)/.test(n)) return 'exhibit';
  return 'doc';
}

/* ================= 讀取 Word ================= */
var W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function wChildren(el, local){
  // 直接子元素（穿透 sdt / customXml / sdtContent 等包裝）
  var out = [];
  (function walk(node){
    for (var c = node.firstElementChild; c; c = c.nextElementSibling){
      if (c.namespaceURI === W_NS && c.localName === local) out.push(c);
      else if (c.namespaceURI === W_NS && (c.localName === 'sdt' || c.localName === 'sdtContent' || c.localName === 'customXml' || c.localName === 'smartTag')) walk(c);
    }
  })(el);
  return out;
}
function cellText(tc){
  var paras = [];
  var cur = '';
  (function walk(node){
    for (var c = node.firstElementChild; c; c = c.nextElementSibling){
      if (c.namespaceURI !== W_NS){ walk(c); continue; }
      var ln = c.localName;
      if (ln === 'del' || ln === 'moveFrom' || ln === 'rPr' || ln === 'pPr' || ln === 'tcPr' || ln === 'instrText' || ln === 'delText') continue;
      if (ln === 'p'){ cur = ''; walk(c); paras.push(cur); cur=''; continue; }
      if (ln === 't'){ cur += c.textContent; continue; }
      if (ln === 'tab'){ cur += ' '; continue; }
      if (ln === 'br' || ln === 'cr'){ cur += '\n'; continue; }
      if (ln === 'noBreakHyphen'){ cur += '-'; continue; }
      walk(c);
    }
  })(tc);
  return paras.map(cleanLine).filter(function(s){return s!=='';}).join('\n');
}
// PDF 常把「林、年、錄」等字存成外觀相同的相容表意字，統一轉回標準字
function normCjk(s){
  return String(s).replace(/[\u2E80-\u2FDF\uF900-\uFAFF]|[\uD87E][\uDC00-\uDE1F]/g, function(ch){ return ch.normalize('NFKC'); });
}
function cleanLine(s){
  return normCjk(s).split('\n').map(function(x){
    return x.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'').replace(/[ \t ]+/g,' ').replace(/^[ 　]+|[ 　]+$/g,'');
  }).filter(function(x){return x!=='';}).join('\n');
}
var NO_RE = /^[^\d\s]{0,4}\d{1,4}(-\d{1,3})?$/;

async function parseDocx(buf){
  var zip = await JSZip.loadAsync(buf);
  var f = zip.file('word/document.xml');
  if (!f) throw new Error('這個檔案不是有效的 Word（.docx）檔。');
  var xml = await f.async('string');
  var dom = new DOMParser().parseFromString(xml, 'application/xml');
  var all = Array.prototype.slice.call(dom.getElementsByTagNameNS(W_NS, 'tbl'));
  var tops = all.filter(function(t){
    for (var p = t.parentNode; p; p = p.parentNode){ if (p.namespaceURI === W_NS && p.localName === 'tbl') return false; }
    return true;
  });
  if (!tops.length) throw new Error('這份 Word 檔裡找不到表格。');
  var tables = tops.map(function(t){
    return wChildren(t,'tr').map(function(tr){ return wChildren(tr,'tc').map(cellText); });
  });
  tables.sort(function(a,b){ return b.length - a.length; });
  return rowsToItems(tables[0]);
}

function rowsToItems(rows){
  rows = rows.filter(function(r){ return r.some(function(c){return c && c.trim();}); });
  if (!rows.length) return [];
  var ncol = Math.max.apply(null, rows.map(function(r){return r.length;}));
  var numCol = -1, nameCol = -1, start = 0;
  var head = rows[0];
  var headLike = head.some(function(c){return /編號|名稱|證據|項目|序號/.test(c);}) && !head.some(function(c){return NO_RE.test(c.trim());});
  if (headLike){
    start = 1;
    head.forEach(function(c,i){
      if (numCol < 0 && /編號|序號|^號/.test(c)) numCol = i;
    });
    head.forEach(function(c,i){
      if (nameCol < 0 && i !== numCol && /名稱|證據|內容|項目/.test(c)) nameCol = i;
    });
  }
  var body = rows.slice(start);
  if (numCol < 0 && ncol > 1){
    var best = -1, bestScore = 0;
    for (var i=0;i<ncol;i++){
      var sc = body.filter(function(r){ return r[i] && NO_RE.test(r[i].trim()); }).length;
      if (sc > bestScore){ bestScore = sc; best = i; }
    }
    if (bestScore >= Math.max(1, body.length * 0.4)) numCol = best;
  }
  if (nameCol < 0){
    var bestLen = -1;
    for (var j=0;j<ncol;j++){
      if (j === numCol) continue;
      var tot = body.reduce(function(s,r){ return s + (r[j] ? r[j].length : 0); }, 0);
      if (tot > bestLen){ bestLen = tot; nameCol = j; }
    }
  }
  var out = [];
  body.forEach(function(r){
    var no = numCol >= 0 ? (r[numCol]||'').trim() : '';
    var name = (r[nameCol]||'').trim();
    if (numCol < 0 && ncol === 1){
      var m = name.match(/^([^\d\s]{0,4}\d{1,4})[\s　.、．]+([\s\S]+)$/);
      if (m){ no = m[1]; name = m[2].trim(); }
    }
    if (!name) return;
    out.push({no:no.replace(/\s+/g,''), name:name});
  });
  return out;
}

/* ================= 讀取 PDF ================= */
function setupPdf(){
  if (!window.pdfjsLib) return;
  // 以主執行緒模式執行，離線、免另開 worker 檔
  try { window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'about:blank'; } catch(e){}
}
async function parsePdf(buf){
  if (!window.pdfjsLib) throw new Error('PDF 讀取元件載入失敗。');
  var doc = await pdfjsLib.getDocument({data:new Uint8Array(buf), isEvalSupported:false, useSystemFonts:true}).promise;
  var lines = [];
  for (var p=1; p<=doc.numPages; p++){
    var page = await doc.getPage(p);
    var tc = await page.getTextContent();
    var its = [];
    tc.items.forEach(function(it){
      if (!it.str || !it.str.trim()) return;
      var h = Math.abs(it.transform[3]) || Math.abs(it.transform[0]) || it.height || 12;
      its.push({s:it.str, x:it.transform[4], y:it.transform[5], w:it.width, h:h});
    });
    its.sort(function(a,b){ return (b.y - a.y) || (a.x - b.x); });
    var pl = [];
    its.forEach(function(it){
      var L = pl.length ? pl[pl.length-1] : null;
      if (L && Math.abs(L.y - it.y) < it.h * 0.45) L.items.push(it);
      else pl.push({y:it.y, h:it.h, items:[it], page:p});
    });
    pl.forEach(function(L){
      L.items.sort(function(a,b){return a.x-b.x;});
      var segs = [];
      L.items.forEach(function(it){
        var S2 = segs.length ? segs[segs.length-1] : null;
        if (S2 && it.x - S2.end < it.h * 0.9) { S2.text += it.s; S2.end = Math.max(S2.end, it.x + it.w); }
        else segs.push({x:it.x, end:it.x + it.w, text:it.s});
      });
      segs.forEach(function(sg){ sg.text = sg.text.replace(/^\s+|\s+$/g,''); });
      L.segs = segs.filter(function(sg){return sg.text;});
      L.x = L.segs.length ? L.segs[0].x : 0;
    });
    lines = lines.concat(pl.filter(function(L){return L.segs.length;}));
  }
  return pdfLinesToItems(lines);
}

function pdfLinesToItems(lines){
  var cand = [];
  lines.forEach(function(L, idx){
    var first = L.segs[0].text, no = null, rest = null, nameX = null;
    if (NO_RE.test(first) && L.segs.length > 1){
      no = first; rest = L.segs.slice(1).map(function(s){return s.text;}).join(''); nameX = L.segs[1].x;
    } else {
      var m = first.match(/^([^\d\s]{0,4}\d{1,4})[\s　]+(.+)$/);
      if (m){ no = m[1]; rest = m[2] + L.segs.slice(1).map(function(s){return s.text;}).join(''); nameX = L.segs[0].x + L.h * 2.2; }
    }
    if (no !== null) cand.push({idx:idx, no:no, rest:rest, x:L.x, nameX:nameX, h:L.h});
  });
  if (!cand.length) throw new Error('在 PDF 中找不到「編號＋證據名稱」形式的清單。若為掃描檔，本程式無法讀取文字。');
  function median(a){ a = a.slice().sort(function(x,y){return x-y;}); return a[Math.floor(a.length/2)]; }
  var numX = median(cand.map(function(c){return c.x;}));
  var nameX = median(cand.map(function(c){return c.nameX;}));
  var hh = median(cand.map(function(c){return c.h;}));
  // 只取編號位於同一欄位、且名稱起點相近的列
  var good = cand.filter(function(c){ return Math.abs(c.x - numX) < hh * 1.5 && Math.abs(c.nameX - nameX) < hh * 3; });
  // 若字頭多數一致，排除字頭明顯不同者（例如「一、」）
  var byIdx = {};
  good.forEach(function(c){ byIdx[c.idx] = c; });
  var out = [], cur = null, lastLine = null;
  for (var i=0;i<lines.length;i++){
    var L = lines[i];
    if (byIdx[i]){
      cur = {no:byIdx[i].no, name:byIdx[i].rest};
      out.push(cur); lastLine = L; continue;
    }
    if (!cur) continue;
    var txt = L.segs.map(function(s){return s.text;}).join('');
    var isPageNo = /^[\d\s\-–—第頁共\/]+$/.test(txt);
    if (isPageNo) continue;
    var aligned = Math.abs(L.x - nameX) < hh * 1.2;
    var near = (L.page === lastLine.page) ? (lastLine.y - L.y) < lastLine.h * 2.6 : true;
    if (aligned && near){
      var a = cur.name.slice(-1), b = txt.charAt(0);
      cur.name += (/[A-Za-z]/.test(a) && /[A-Za-z]/.test(b)) ? ' ' + txt : txt;
      lastLine = L;
    } else if (L.page === lastLine.page || L.x < nameX - hh) {
      cur = null;
    }
  }
  return out.map(function(o){ return {no:normCjk(o.no), name:cleanLine(o.name)}; });
}

/* ================= 匯入 ================= */
async function readFile(file){
  var buf = await file.arrayBuffer();
  var nm = file.name.toLowerCase();
  if (/\.docx$/.test(nm)) return parseDocx(buf);
  if (/\.pdf$/.test(nm)) return parsePdf(buf);
  if (/\.doc$/.test(nm)) throw new Error('「' + file.name + '」是舊版 .doc 格式，請先用 Word 開啟後「另存新檔」為 .docx 再匯入。');
  if (/\.json$/.test(nm)) throw new Error('進度檔請用「載入進度」開啟。');
  throw new Error('「' + file.name + '」不是 Word（.docx）或 PDF 檔。');
}
async function importFiles(files){
  files = Array.prototype.slice.call(files || []);
  if (!files.length) return;
  showBusy(true);
  var got = [], errs = [];
  for (var i=0;i<files.length;i++){
    try { var rs = await readFile(files[i]); got.push({file:files[i].name, rows:rs}); }
    catch(e){ errs.push(e && e.message ? e.message : String(e)); }
  }
  showBusy(false);
  var rows = [];
  got.forEach(function(g){ rows = rows.concat(g.rows); });
  if (errs.length && !rows.length){ alertModal('無法匯入', errs.join('\n')); return; }
  if (!rows.length){ alertModal('無法匯入', '檔案中沒有讀到任何證據項目。'); return; }
  var srcName = got.map(function(g){return g.file;}).join('、');
  var apply = function(mode){
    var newItems = rows.map(function(r){ return makeItem(r.no, r.name); });
    if (mode === 'replace'){ S.items = newItems; S.source = srcName; UI.sel.clear(); UI.rulSel.clear(); S.locked = false; S.lockMax = 0; S.lockedAt = ''; S.onlyNew = false; }
    else { S.items = S.items.concat(newItems); S.source = S.source ? S.source + '、' + srcName : srcName; }
    UI.filter = 'all'; UI.search = ''; document.getElementById('search').value = '';
    renderAll(); save();
    toast('已匯入 ' + newItems.length + ' 項證據' + (errs.length ? '；另有檔案無法讀取：' + errs.join(' ') : ''));
  };
  if (S.items.length){
    choiceModal('匯入 ' + rows.length + ' 項證據', '目前清單已有 ' + S.items.length + ' 項。要取代目前清單，還是接在後面？', [
      {label:'取消'},
      {label:'接在後面', fn:function(){apply('append');}},
      {label:'取代目前清單', primary:true, fn:function(){apply('replace');}}
    ]);
  } else apply('replace');
}

/* ================= 檢證編號 ================= */
// 依輸出順序列出所有已聲請且已選類別的證據，並給定檢證編號。
// 鎖定後：已鎖定者沿用原編號，新增者自最後一號往下接續。
function numberedRows(){
  var out = [];
  CAT_ORDER.forEach(function(cat){
    METHODS.forEach(function(m){
      var g = S.items.filter(function(it){ return it.apply === true && it.cat === cat && it.method === m; });
      if (S.locked){
        g = g.map(function(it, i){ return {it:it, i:i}; }).sort(function(a, b){
          var x = a.it.lockedNo == null ? Infinity : a.it.lockedNo, y = b.it.lockedNo == null ? Infinity : b.it.lockedNo;
          return (x - y) || (a.i - b.i);
        }).map(function(o){ return o.it; });
      }
      g.forEach(function(it){ out.push({it:it, cat:cat, method:m}); });
    });
  });
  var next = S.locked ? (S.lockMax || 0) : 0;
  out.forEach(function(r){
    r.locked = S.locked && r.it.lockedNo != null;
    r.num = r.locked ? r.it.lockedNo : ++next;
  });
  return out;
}
function numMap(){
  var m = {};
  numberedRows().forEach(function(r){ m[r.it.id] = {num:r.num, locked:r.locked}; });
  return m;
}
function newSinceLock(){ return S.locked ? numberedRows().filter(function(r){ return !r.locked; }) : []; }

/* ================= 輸出結構 ================= */
function buildOutline(){
  var onlyNew = S.locked && S.onlyNew;
  var rows = numberedRows();
  if (onlyNew) rows = rows.filter(function(r){ return !r.locked; });
  var parts = [];
  CAT_ORDER.forEach(function(cat){
    var subs = [];
    METHODS.forEach(function(m){
      var g = rows.filter(function(r){ return r.cat === cat && r.method === m; });
      if (g.length) subs.push({kind:'table', method:m, title:METHOD_TITLE[m], rows:g.map(function(r){
        return {id:r.it.id, no:'檢證' + r.num, name:r.it.name, fact:factOf(r.it), locked:r.locked};
      })});
    });
    if (S.askDef[cat]) subs.push({kind:'ask', title:ASK_TITLE});
    if (subs.length) parts.push({cat:cat, title:PART_TITLE[cat], subs:subs});
  });
  parts.forEach(function(p,i){
    p.label = PART_NUM[i] + '、' + p.title;
    p.subs.forEach(function(s,j){
      s.label = (SUB_NUM[j] || '(' + (j+1) + ')') + '、' + s.title + (s.kind === 'ask' ? '。' : '：');
    });
  });
  return parts;
}
function hasTables(parts){ return parts.some(function(p){ return p.subs.some(function(s){return s.kind==='table';}); }); }

/* ================= 畫面：清單 ================= */
var $ = function(id){ return document.getElementById(id); };
function esc(s){ return String(s).replace(/[&<>"']/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; }); }

function visibleItems(){
  var q = UI.search.trim();
  return S.items.filter(function(it){
    var st = statusOf(it);
    if (UI.filter === 'pend' && !(st === 'pend' || st === 'catpend')) return false;
    if (UI.filter === 'crime' && st !== 'crime') return false;
    if (UI.filter === 'sentence' && st !== 'sentence') return false;
    if (UI.filter === 'off' && st !== 'off') return false;
    if (q && (it.name + ' ' + it.no).indexOf(q) < 0) return false;
    return true;
  });
}
function renderTabs(){
  var c = {all:S.items.length, pend:0, crime:0, sentence:0, off:0};
  S.items.forEach(function(it){
    var st = statusOf(it);
    if (st === 'pend' || st === 'catpend') c.pend++; else c[st]++;
  });
  var defs = [['all','全部'],['pend','未決定'],['crime','犯罪事實'],['sentence','科刑資料'],['off','不聲請']];
  $('tabs').innerHTML = defs.map(function(d){
    return '<button class="tab t-' + d[0] + '" data-filter="' + d[0] + '" aria-pressed="' + (UI.filter === d[0]) + '">' + d[1] + '<b>' + c[d[0]] + '</b></button>';
  }).join('');
}
function segBtn(act, v, label, pressed){
  return '<button type="button" data-act="' + act + '" data-v="' + v + '" aria-pressed="' + (pressed ? 'true':'false') + '">' + label + '</button>';
}
function rowHtml(it){
  var st = statusOf(it);
  var h = '<div class="row s-' + st + '" data-id="' + it.id + '">';
  h += '<input type="checkbox" class="sel" aria-label="勾選此項"' + (UI.sel.has(it.id) ? ' checked' : '') + '>';
  h += '<input class="no" data-f="no" value="' + esc(it.no) + '" aria-label="清冊編號" placeholder="編號">';
  h += '<textarea class="name" data-f="name" rows="1" aria-label="證據名稱" placeholder="輸入證據名稱">' + esc(it.name) + '</textarea>';
  h += '<button type="button" class="del" data-act="del">刪除</button>';
  h += '<div class="ctl">';
  var nb = UI.nums[it.id];
  if (nb) h += '<span class="kz' + (nb.locked ? ' lk' : '') + '" title="' + (nb.locked ? '檢證編號已鎖定' : '目前的檢證編號') + '">檢證' + nb.num + (nb.locked ? '・鎖定' : '') + '</span>';
  h += '<div class="seg apply" role="group" aria-label="是否聲請">' + segBtn('apply','1','聲請', it.apply === true) + segBtn('apply','0','不聲請', it.apply === false) + '</div>';
  if (it.apply !== false){
    h += '<div class="seg cat' + (st === 'catpend' ? ' need' : '') + '" role="group" aria-label="證據調查類別">' + segBtn('cat','crime','犯罪事實', it.cat === 'crime') + segBtn('cat','sentence','科刑資料', it.cat === 'sentence') + '</div>';
    h += '<div class="seg method" role="group" aria-label="調查方法">' + METHODS.map(function(m){ return segBtn('method', m, METHOD_LABEL[m], it.method === m); }).join('') + '</div>';
    if (it.methodAuto) h += '<span class="auto" title="依證據名稱自動預選，可點選更改">預選</span>';
  }
  if (it.apply === true){
    h += '<label class="fact"><span>待證事實</span><textarea data-f="fact" rows="1" placeholder="' + (it.cat ? '' : '請先選擇犯罪事實或科刑資料') + '">' + esc(factOf(it)) + '</textarea>';
    if (!it.factAuto && it.cat) h += '<button type="button" class="reset" data-act="factreset">還原預設</button>';
    h += '</label>';
  }
  h += '</div></div>';
  return h;
}
function renderList(){
  var body = $('listBody');
  var head = $('listHead');
  if (!S.items.length){
    head.hidden = true;
    body.innerHTML = '<div class="drop" id="drop"><strong>把證據清冊拖到這裡</strong>' +
      '<p>支援 Word（.docx）與 PDF。Word 檔只讀取其中的表格。</p>' +
      '<p class="note">掃描成圖片的 PDF 無法讀取；舊版 .doc 請先另存為 .docx。</p>' +
      '<div class="acts"><button type="button" data-act="pick">選擇檔案</button><button type="button" class="sec" data-act="add">手動新增一項</button></div>' +
      '<p class="note" style="margin-top:18px">所有檔案只在這台電腦上處理，不會上傳到任何地方。</p></div>';
    return;
  }
  head.hidden = false;
  renderTabs();
  UI.nums = numMap();
  var vis = visibleItems();
  var h = vis.length ? vis.map(rowHtml).join('') : '<div class="noresult">沒有符合條件的項目。</div>';
  h += '<div class="listfoot"><button type="button" data-act="add">新增一項（例如證人）</button></div>';
  body.innerHTML = h;
  body.querySelectorAll('textarea').forEach(autosize);
  renderBulk();
}
function rerenderRow(id){
  var el = document.querySelector('.row[data-id="' + id + '"]');
  var it = findItem(id);
  if (!el || !it) { renderList(); return; }
  UI.nums = numMap();
  var tmp = document.createElement('div');
  tmp.innerHTML = rowHtml(it);
  var nu = tmp.firstChild;
  el.replaceWith(nu);
  nu.querySelectorAll('textarea').forEach(autosize);
  renderTabs();
}
function renderBulk(){
  // 移除已不存在的勾選
  UI.sel.forEach(function(id){ if (!findItem(id)) UI.sel.delete(id); });
  var n = UI.sel.size;
  $('bulk').classList.toggle('show', n > 0);
  $('bulkCount').textContent = '已勾選 ' + n + ' 項：';
  var vis = visibleItems();
  var allSel = vis.length > 0 && vis.every(function(it){ return UI.sel.has(it.id); });
  $('selAll').checked = allSel;
  $('selAll').indeterminate = !allSel && vis.some(function(it){ return UI.sel.has(it.id); });
}
function autosize(t){ t.style.height = 'auto'; t.style.height = (t.scrollHeight + 2) + 'px'; }
function findItem(id){ for (var i=0;i<S.items.length;i++) if (S.items[i].id === id) return S.items[i]; return null; }

/* ================= 畫面：預覽 ================= */
function renderApplyPreview(){
  var parts = buildOutline();
  var paper = $('paper');
  var warns = [];
  var catpend = S.items.filter(function(it){ return statusOf(it) === 'catpend'; }).length;
  var pend = S.items.filter(function(it){ return statusOf(it) === 'pend'; }).length;
  if (catpend) warns.push('<div class="warn">有 ' + catpend + ' 項已選「聲請」但還沒選犯罪事實或科刑資料，尚未列入表格。<button type="button" data-goto="pend">查看</button></div>');
  if (pend) warns.push('<div class="warn">有 ' + pend + ' 項尚未決定是否聲請。<button type="button" data-goto="pend">查看</button></div>');
  $('warns').innerHTML = warns.join('');
  var ok = hasTables(parts);
  $('btnDocx').disabled = !ok;
  $('btnCopy').disabled = !ok;
  if (!ok){
    paper.innerHTML = '<div class="pv-empty">' + (S.locked && S.onlyNew ? '鎖定後還沒有新增的聲請項目。' : '在左側清單選擇「聲請」及類別後，表格會出現在這裡。') + '</div>';
    return;
  }
  var h = '';
  parts.forEach(function(p){
    h += '<p class="h1">' + esc(p.label) + '</p>';
    p.subs.forEach(function(s){
      h += '<p class="h2">' + esc(s.label) + '</p>';
      if (s.kind !== 'table') return;
      h += '<table><thead><tr><th class="c-no">檢證編號</th><th>證據名稱</th><th class="c-fact">待證事實</th><th class="mv"></th></tr></thead><tbody>';
      s.rows.forEach(function(r, i){
        h += '<tr class="pr" data-id="' + r.id + '" title="點一下可跳到左側清單中的這一項"><td class="c-no">' + esc(r.no) + '</td><td class="nl">' + esc(r.name) + '</td><td class="c-fact nl">' + esc(r.fact) + '</td>' +
          '<td class="mv">' + (r.locked ? '<span class="lkmark" title="檢證編號已鎖定，無法移動">已鎖定</span>' :
          '<button type="button" data-mv="up" aria-label="上移"' + (i === 0 || s.rows[i-1].locked ? ' disabled' : '') + '>上</button> <button type="button" data-mv="down" aria-label="下移"' + (i === s.rows.length - 1 ? ' disabled' : '') + '>下</button>') + '</td></tr>';
      });
      h += '</tbody></table>';
    });
  });
  paper.innerHTML = h;
}

function renderPreview(){
  renderApplyPreview();
  renderLock();
  if (UI.tab === 'cmp') renderCmp();
  if (UI.tab === 'rul') renderRul();
}
function moveWithinGroup(id, dir){
  var it = findItem(id); if (!it) return;
  if (S.locked && it.lockedNo != null) return;
  var same = function(x){ return x.apply === true && x.cat === it.cat && x.method === it.method && (!S.locked || x.lockedNo == null); };
  var idx = S.items.indexOf(it);
  var j = idx + (dir === 'up' ? -1 : 1);
  while (j >= 0 && j < S.items.length && !same(S.items[j])) j += (dir === 'up' ? -1 : 1);
  if (j < 0 || j >= S.items.length) return;
  var other = S.items[j];
  S.items[idx] = other; S.items[j] = it;
  renderList(); renderPreview(); save();
}

function renderAll(){
  $('srcName').textContent = S.source ? '清冊：' + S.source : '';
  $('srcName').title = S.source || '';
  $('askCrime').checked = !!S.askDef.crime;
  $('askSent').checked = !!S.askDef.sentence;
  $('fontPt').value = String(S.fontPt);
  $('tableCm').value = (parseFloat(S.tableCm) || 16).toFixed(1);
  renderList(); renderPreview();
  showTab(UI.tab);
}

/* ================= 事件：清單 ================= */
function bindList(){
  var body = $('listBody');
  body.addEventListener('click', function(e){
    var b = e.target.closest('[data-act]');
    if (!b) return;
    var act = b.dataset.act;
    if (act === 'pick'){ $('fileList').click(); return; }
    if (act === 'add'){ addItem(); return; }
    var row = b.closest('.row'); if (!row) return;
    var it = findItem(row.dataset.id); if (!it) return;
    if (act === 'apply'){
      var v = b.dataset.v === '1';
      it.apply = (it.apply === v) ? null : v;
    } else if (act === 'cat'){
      it.cat = (it.cat === b.dataset.v) ? null : b.dataset.v;
      if (it.cat) it.apply = true;
    } else if (act === 'method'){
      it.method = b.dataset.v; it.methodAuto = false;
    } else if (act === 'factreset'){
      it.factAuto = true; it.fact = '';
    } else if (act === 'del'){
      deleteItems([it.id]); return;
    }
    if (act === 'factreset') rerenderRow(it.id); else renderList();
    renderPreview(); save();
    var again = document.querySelector('.row[data-id="' + it.id + '"] [data-act="' + act + '"][data-v="' + (b.dataset.v || '') + '"]');
    if (again) again.focus();
  });
  body.addEventListener('input', function(e){
    var t = e.target, f = t.dataset.f;
    if (!f) return;
    var row = t.closest('.row'); var it = row && findItem(row.dataset.id); if (!it) return;
    if (f === 'name'){
      it.name = t.value;
      if (it.methodAuto){
        var m = suggestMethod(it.name);
        if (m !== it.method){
          it.method = m;
          row.querySelectorAll('.seg.method button').forEach(function(bb){ bb.setAttribute('aria-pressed', bb.dataset.v === m ? 'true':'false'); });
        }
      }
    }
    else if (f === 'no') it.no = t.value;
    else if (f === 'fact'){ it.fact = t.value; it.factAuto = false; }
    if (t.tagName === 'TEXTAREA') autosize(t);
    schedulePreview(); save();
  });
  body.addEventListener('change', function(e){
    var t = e.target;
    if (t.classList.contains('sel')){
      var row = t.closest('.row');
      if (t.checked) UI.sel.add(row.dataset.id); else UI.sel.delete(row.dataset.id);
      renderBulk();
    }
    if (t.dataset.f === 'fact'){ rerenderRow(t.closest('.row').dataset.id); }
  });
  $('tabs').addEventListener('click', function(e){
    var b = e.target.closest('[data-filter]'); if (!b) return;
    UI.filter = b.dataset.filter; renderList();
  });
  $('search').addEventListener('input', function(e){ UI.search = e.target.value; renderList(); });
  $('selAll').addEventListener('change', function(e){
    visibleItems().forEach(function(it){ if (e.target.checked) UI.sel.add(it.id); else UI.sel.delete(it.id); });
    renderList();
  });
  $('bulk').addEventListener('click', function(e){
    var b = e.target.closest('[data-bulk]'); if (!b) return;
    var k = b.dataset.bulk;
    var ids = Array.from(UI.sel);
    if (k === 'none'){ UI.sel.clear(); renderList(); return; }
    if (k === 'del'){ deleteItems(ids); return; }
    ids.forEach(function(id){
      var it = findItem(id); if (!it) return;
      if (k === 'apply1') it.apply = true;
      if (k === 'apply0') it.apply = false;
      if (k === 'crime' || k === 'sentence'){ it.cat = k; it.apply = true; }
    });
    renderList(); renderPreview(); save();
    toast('已設定 ' + ids.length + ' 項');
  });
  $('bulkMethod').addEventListener('change', function(e){
    var m = e.target.value; if (!m) return;
    var ids = Array.from(UI.sel);
    ids.forEach(function(id){ var it = findItem(id); if (it){ it.method = m; it.methodAuto = false; } });
    e.target.value = '';
    renderList(); renderPreview(); save();
    toast('已將 ' + ids.length + ' 項的調查方法設為「' + METHOD_LABEL[m] + '」');
  });
}
function addItem(){
  var it = makeItem('', '', {apply:true, manual:true});
  S.items.push(it);
  UI.filter = 'all'; UI.search = ''; $('search').value = '';
  renderList(); renderPreview(); save();
  var el = document.querySelector('.row[data-id="' + it.id + '"] .name');
  if (el){ el.scrollIntoView({block:'center'}); el.focus(); }
}
var lastDeleted = null;
function deleteItems(ids, confirmed){
  if (!ids.length) return;
  var lockedOnes = S.locked ? ids.map(findItem).filter(function(it){ return it && it.lockedNo != null; }) : [];
  if (lockedOnes.length && !confirmed){
    choiceModal('刪除已鎖定編號的證據', '其中 ' + lockedOnes.length + ' 項已鎖定檢證編號（' + lockedOnes.map(function(it){return '檢證' + it.lockedNo;}).slice(0,6).join('、') + (lockedOnes.length > 6 ? '…' : '') + '）。刪除後這些編號會成為空號，其他編號不變。',
      [{label:'取消'},{label:'刪除', primary:true, fn:function(){ deleteItems(ids, true); }}]);
    return;
  }
  var snapshot = S.items.slice();
  S.items = S.items.filter(function(it){ return ids.indexOf(it.id) < 0; });
  ids.forEach(function(id){ UI.sel.delete(id); });
  lastDeleted = snapshot;
  renderList(); renderPreview(); save();
  toast('已刪除 ' + ids.length + ' 項', function(){ S.items = lastDeleted; lastDeleted = null; renderList(); renderPreview(); save(); });
}

/* ================= 事件：預覽 ================= */
var pvTimer = null;
function schedulePreview(){ clearTimeout(pvTimer); pvTimer = setTimeout(renderPreview, 250); }
function bindPreview(){
  $('paper').addEventListener('click', function(e){
    var mv = e.target.closest('[data-mv]');
    var tr = e.target.closest('tr.pr');
    if (!tr) return;
    if (mv){ moveWithinGroup(tr.dataset.id, mv.dataset.mv); return; }
    gotoItem(tr.dataset.id);
  });
  $('warns').addEventListener('click', function(e){
    var b = e.target.closest('[data-goto]'); if (!b) return;
    UI.filter = b.dataset.goto; renderList();
  });
  $('askCrime').addEventListener('change', function(e){ S.askDef.crime = e.target.checked; renderPreview(); save(); });
  $('askSent').addEventListener('change', function(e){ S.askDef.sentence = e.target.checked; renderPreview(); save(); });
  $('fontPt').addEventListener('change', function(e){ S.fontPt = parseInt(e.target.value,10) || 14; save(); });
  $('tableCm').addEventListener('change', function(e){
    var v = parseFloat(e.target.value);
    if (!(v >= 8 && v <= 20)){ v = 16; toast('表格寬度請輸入 8 至 20 公分'); }
    S.tableCm = Math.round(v * 10) / 10; e.target.value = S.tableCm.toFixed(1); save();
  });
  $('btnDocx').addEventListener('click', downloadDocx);
  $('btnCopy').addEventListener('click', copyTables);
}
function gotoItem(id){
  var el = document.querySelector('.row[data-id="' + id + '"]');
  if (!el){ UI.filter = 'all'; UI.search = ''; $('search').value = ''; renderList(); el = document.querySelector('.row[data-id="' + id + '"]'); }
  if (!el) return;
  el.scrollIntoView({block:'center', behavior:'smooth'});
  el.classList.remove('flash'); void el.offsetWidth; el.classList.add('flash');
  var n = el.querySelector('.name'); if (n) n.focus({preventScroll:true});
}

/* ================= 欄寬（Word 與複製共用） ================= */
function tableTwips(){
  var cm = parseFloat(S.tableCm);
  if (!(cm >= 8 && cm <= 20)) cm = 16;
  return cm === 16 ? TEXT_W : Math.round(cm * 567);
}
function colWidths(pt){
  var cw = pt * 20, pad = 85;
  var total = tableTwips();
  var no = Math.round(cw * 4.3 + pad * 2);
  var fact = Math.round(cw * 7.3 + pad * 2);
  var name = total - no - fact;
  var minName = cw * 6 + pad * 2;
  if (name < minName){ // 寬度太窄時按比例縮
    var k = (total - minName) / (no + fact);
    no = Math.round(no * k); fact = Math.round(fact * k); name = total - no - fact;
  }
  return {total:total, no:no, name:name, fact:fact, pad:pad};
}

/* ================= Word 產生 ================= */
function xesc(s){
  return String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g,'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function runXml(text){
  var lines = String(text).split('\n');
  return lines.map(function(l, i){
    return (i ? '<w:r><w:br/></w:r>' : '') + '<w:r><w:t xml:space="preserve">' + xesc(l) + '</w:t></w:r>';
  }).join('');
}
function buildDocumentXml(parts, pt){
  var cw = pt * 20;                 // 一個全形字寬（twips）
  var mL = 1418, mR = 1418;
  var CW = colWidths(pt);
  var textW = CW.total, pad = CW.pad, wNo = CW.no, wName = CW.name, wFact = CW.fact;
  var bodyLine = Math.round(pt * 20 * 1.5);
  var pBody = function(text, hangChars){
    var hang = hangChars * cw;
    return '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="' + bodyLine + '" w:lineRule="exact"/>' +
      '<w:ind w:left="' + hang + '" w:hanging="' + hang + '"/><w:jc w:val="both"/></w:pPr>' + runXml(text) + '</w:p>';
  };
  var pCell = function(text){
    return '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/></w:pPr>' + runXml(text) + '</w:p>';
  };
  var cell = function(w, text){
    return '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/></w:tcPr>' + pCell(text) + '</w:tc>';
  };
  var bd = function(n){ return '<w:' + n + ' w:val="single" w:sz="4" w:space="0" w:color="000000"/>'; };
  var table = function(rows){
    var x = '<w:tbl><w:tblPr><w:tblW w:w="' + textW + '" w:type="dxa"/><w:tblInd w:w="0" w:type="dxa"/>' +
      '<w:tblBorders>' + bd('top') + bd('left') + bd('bottom') + bd('right') + bd('insideH') + bd('insideV') + '</w:tblBorders>' +
      '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="28" w:type="dxa"/><w:left w:w="' + pad + '" w:type="dxa"/><w:bottom w:w="28" w:type="dxa"/><w:right w:w="' + pad + '" w:type="dxa"/></w:tblCellMar>' +
      '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="' + wNo + '"/><w:gridCol w:w="' + wName + '"/><w:gridCol w:w="' + wFact + '"/></w:tblGrid>';
    x += '<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>' + cell(wNo,'檢證編號') + cell(wName,'證據名稱') + cell(wFact,'待證事實') + '</w:tr>';
    rows.forEach(function(r){
      x += '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + cell(wNo, r.no) + cell(wName, r.name) + cell(wFact, r.fact) + '</w:tr>';
    });
    return x + '</w:tbl>';
  };
  var body = '';
  parts.forEach(function(p){
    body += pBody(p.label, 2);
    p.subs.forEach(function(s){
      body += pBody(s.label, 2);
      if (s.kind === 'table') body += table(s.rows);
    });
  });
  body += '<w:p/>';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="' + W_NS + '" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>' + body +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="' + mR + '" w:bottom="1418" w:left="' + mL + '" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>' +
    '</w:body></w:document>';
}
function buildStylesXml(pt){
  var sz = pt * 2;
  var f = '<w:rFonts w:ascii="標楷體" w:hAnsi="標楷體" w:eastAsia="標楷體" w:cs="標楷體" w:hint="eastAsia"/>';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:styles xmlns:w="' + W_NS + '">' +
    '<w:docDefaults><w:rPrDefault><w:rPr>' + f + '<w:kern w:val="2"/><w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/><w:lang w:val="en-US" w:eastAsia="zh-TW" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl w:val="0"/></w:pPr><w:rPr>' + f + '<w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/></w:rPr></w:style>' +
    '<w:style w:type="table" w:default="1" w:styleId="a1"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
    '</w:styles>';
}
async function makeDocxBlob(docXml){
  var parts = buildOutline();
  var pt = S.fontPt || 14;
  var zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>');
  zip.file('word/settings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="' + W_NS + '"><w:defaultTabStop w:val="480"/><w:characterSpacingControl w:val="compressPunctuation"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>');
  zip.file('word/styles.xml', buildStylesXml(pt));
  zip.file('word/document.xml', docXml || buildDocumentXml(parts, pt));
  return zip.generateAsync({type:'blob', mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
}
function baseName(){
  var s = (S.source || '').split('、')[0].replace(/\.[^.]+$/,'').trim();
  return s ? s.slice(0, 40) : '';
}
async function downloadDocx(){
  if (!hasTables(buildOutline())){ toast('還沒有可輸出的內容'); return; }
  var blob = await makeDocxBlob();
  var b = baseName();
  saveBlob(blob, '聲請調查證據' + (b ? '_' + b : '') + '.docx');
  toast('已下載 Word 檔');
}
function saveBlob(blob, name){
  if (window.navigator && window.navigator.msSaveOrOpenBlob){ window.navigator.msSaveOrOpenBlob(blob, name); return; }
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a); a.click();
  setTimeout(function(){ URL.revokeObjectURL(a.href); a.remove(); }, 1500);
}

/* ================= 複製到剪貼簿 ================= */
function clipHtml(){
  // 仿照 Word 自身放到剪貼簿的 HTML：每一欄都有明確寬度、不帶網頁底色與行高
  var parts = buildOutline();
  var pt = S.fontPt || 14;
  var CW = colWidths(pt);
  var tw2pt = function(t){ return (t / 20).toFixed(1); };
  var tw2px = function(t){ return Math.round(t / 20 * 96 / 72); };
  var fam = 'font-family:標楷體;mso-ascii-font-family:標楷體;mso-hansi-font-family:標楷體;mso-fareast-font-family:標楷體';
  var run = function(t){
    return '<span style="font-size:' + pt + '.0pt;' + fam + '">' + esc(t).replace(/\n/g, '<br>') + '</span>';
  };
  var hang = (pt * 2).toFixed(1);
  var head = function(t){
    return '<p class="MsoNormal" style="margin:0cm;margin-left:' + hang + 'pt;text-indent:-' + hang + 'pt;line-height:' + (pt * 1.5).toFixed(1) + 'pt;mso-line-height-rule:exactly;text-align:justify;font-size:' + pt + '.0pt;' + fam + '">' + run(t) + '</p>';
  };
  var padCss = 'padding:1.4pt 4.25pt 1.4pt 4.25pt';
  var td = function(w, t){
    return '<td width="' + tw2px(w) + '" valign="top" style="width:' + tw2pt(w) + 'pt;border:solid windowtext 1.0pt;mso-border-alt:solid windowtext .5pt;' + padCss + '">' +
      '<p class="MsoNormal" style="margin:0cm;line-height:normal;text-align:left;font-size:' + pt + '.0pt;' + fam + '">' + run(t) + '</p></td>';
  };
  var row = function(a, b, c){ return '<tr style="page-break-inside:avoid">' + td(CW.no, a) + td(CW.name, b) + td(CW.fact, c) + '</tr>'; };
  var h = '';
  parts.forEach(function(pp){
    h += head(pp.label);
    pp.subs.forEach(function(s){
      h += head(s.label);
      if (s.kind !== 'table') return;
      h += '<table class="MsoTableGrid" border="1" cellspacing="0" cellpadding="0" width="' + tw2px(CW.total) + '" style="width:' + tw2pt(CW.total) + 'pt;border-collapse:collapse;border:none;mso-border-alt:solid windowtext .5pt;mso-table-layout-alt:fixed;mso-padding-alt:1.4pt 4.25pt 1.4pt 4.25pt">';
      h += row('檢證編號', '證據名稱', '待證事實');
      s.rows.forEach(function(r){ h += row(r.no, r.name, r.fact); });
      h += '</table>';
    });
  });
  return h;
}
function clipText(){
  var out = [];
  buildOutline().forEach(function(pp){
    out.push(pp.label);
    pp.subs.forEach(function(s){
      out.push(s.label);
      if (s.kind !== 'table') return;
      out.push('檢證編號\t證據名稱\t待證事實');
      s.rows.forEach(function(r){ out.push(r.no + '\t' + r.name.replace(/\n/g,' ') + '\t' + r.fact.replace(/\n/g,' ')); });
    });
  });
  return out.join('\r\n');
}
function copyViaEvent(html, text){
  // 在 copy 事件中直接寫入內容，瀏覽器就不會夾帶網頁本身的底色、行高等樣式
  var ok = false;
  var handler = function(e){
    if (!e.clipboardData) return;
    e.clipboardData.setData('text/html', html);
    e.clipboardData.setData('text/plain', text);
    e.preventDefault();
    ok = true;
  };
  document.addEventListener('copy', handler, true);
  var box = $('clip');
  try {
    box.textContent = '.';
    var r = document.createRange(); r.selectNodeContents(box);
    var sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r);
    document.execCommand('copy');
    sel.removeAllRanges();
  } catch(e){ ok = false; }
  finally { document.removeEventListener('copy', handler, true); box.textContent = ''; }
  return ok;
}
async function copyTables(){
  if (!hasTables(buildOutline())){ toast('還沒有可輸出的內容'); return; }
  var html = clipHtml(), text = clipText();
  var done = copyViaEvent(html, text);
  if (!done && navigator.clipboard && window.ClipboardItem){
    try {
      await navigator.clipboard.write([new ClipboardItem({'text/html':new Blob([html],{type:'text/html'}),'text/plain':new Blob([text],{type:'text/plain'})})]);
      done = true;
    } catch(e){}
  }
  toast(done ? '已複製，可到 Word 或漢書中貼上' : '瀏覽器不允許複製，請改用「下載 Word」');
}

/* ================= 存檔 / 載入 ================= */
function snapshot(){
  return {app:'聲請調查證據表格產生器', version:FILE_VERSION, savedAt:new Date().toISOString(),
    source:S.source, fontPt:S.fontPt, tableCm:S.tableCm, askDef:S.askDef,
    locked:S.locked, lockMax:S.lockMax, lockedAt:S.lockedAt, onlyNew:S.onlyNew, items:S.items};
}
var saveTimer = null;
function save(){
  clearTimeout(saveTimer);
  saveTimer = setTimeout(function(){
    try {
      if (S.items.length) localStorage.setItem(STORE_KEY, JSON.stringify(snapshot()));
      else localStorage.removeItem(STORE_KEY);
    } catch(e){}
  }, 300);
}
function loadSnapshot(d){
  if (!d || !Array.isArray(d.items)) throw new Error('這不是本程式的進度檔。');
  S.items = d.items.map(function(x){
    return {id:newId(), no:String(x.no||''), name:String(x.name||''),
      apply:(x.apply === true || x.apply === false) ? x.apply : null,
      cat:(x.cat === 'crime' || x.cat === 'sentence') ? x.cat : null,
      method:METHODS.indexOf(x.method) >= 0 ? x.method : suggestMethod(x.name||''),
      methodAuto:x.methodAuto !== false, fact:String(x.fact||''), factAuto:x.factAuto !== false,
      manual:!!x.manual,
      lockedNo:(typeof x.lockedNo === 'number' && x.lockedNo > 0) ? x.lockedNo : null,
      ruling:RULINGS.indexOf(x.ruling) >= 0 ? x.ruling : null, rulingNote:String(x.rulingNote||'')};
  });
  S.source = d.source || '';
  S.fontPt = parseInt(d.fontPt,10) || 14;
  S.tableCm = parseFloat(d.tableCm) || 16;
  S.locked = !!d.locked && S.items.some(function(it){ return it.lockedNo != null; });
  S.lockMax = S.locked ? Math.max(parseInt(d.lockMax,10) || 0, Math.max.apply(null, S.items.map(function(it){ return it.lockedNo || 0; }))) : 0;
  S.lockedAt = S.locked ? (d.lockedAt || '') : '';
  S.onlyNew = S.locked && !!d.onlyNew;
  if (!S.locked) S.items.forEach(function(it){ it.lockedNo = null; });
  S.askDef = {crime: !d.askDef || d.askDef.crime !== false, sentence: !d.askDef || d.askDef.sentence !== false};
  UI.sel.clear(); UI.rulSel.clear(); UI.filter = 'all'; UI.search = ''; $('search').value = '';
  renderAll(); save();
}
function stamp(){
  var d = new Date(), z = function(n){ return (n<10?'0':'') + n; };
  return d.getFullYear() + z(d.getMonth()+1) + z(d.getDate()) + '-' + z(d.getHours()) + z(d.getMinutes());
}
function bindTop(){
  $('btnImport').addEventListener('click', function(){ $('fileList').click(); });
  $('fileList').addEventListener('change', function(e){ var fs = e.target.files; importFiles(fs).then(function(){ e.target.value = ''; }); });
  $('btnSave').addEventListener('click', function(){
    if (!S.items.length){ toast('清單是空的，沒有可儲存的進度'); return; }
    var b = baseName();
    saveBlob(new Blob([JSON.stringify(snapshot(), null, 1)], {type:'application/json'}), '證據調查進度' + (b ? '_' + b : '') + '_' + stamp() + '.json');
    toast('已儲存進度檔，之後可用「載入進度」繼續');
  });
  $('btnLoad').addEventListener('click', function(){ $('fileProg').click(); });
  $('fileProg').addEventListener('change', async function(e){
    var f = e.target.files[0]; e.target.value = '';
    if (!f) return;
    try {
      var d = JSON.parse(await f.text());
      var go = function(){ loadSnapshot(d); toast('已載入進度（' + S.items.length + ' 項）'); };
      if (S.items.length) choiceModal('載入進度', '載入後會取代目前清單（' + S.items.length + ' 項）。', [{label:'取消'},{label:'載入',primary:true,fn:go}]);
      else go();
    } catch(err){ alertModal('無法載入', err.message && /本程式/.test(err.message) ? err.message : '這個檔案不是本程式的進度檔。'); }
  });
  $('btnClear').addEventListener('click', function(){
    if (!S.items.length) return;
    choiceModal('清空清單', '會移除目前全部 ' + S.items.length + ' 項。需要的話，請先「儲存進度」。', [{label:'取消'},{label:'清空',primary:true,fn:function(){
      S.items = []; S.source = ''; S.locked = false; S.lockMax = 0; S.lockedAt = ''; S.onlyNew = false;
      UI.sel.clear(); UI.rulSel.clear(); renderAll(); save();
    }}]);
  });
  // 拖放
  var main = $('listPane');
  ['dragenter','dragover'].forEach(function(ev){ main.addEventListener(ev, function(e){ e.preventDefault(); main.classList.add('dragging'); }); });
  ['dragleave','drop'].forEach(function(ev){ main.addEventListener(ev, function(e){ e.preventDefault(); if (ev === 'dragleave' && main.contains(e.relatedTarget)) return; main.classList.remove('dragging'); }); });
  main.addEventListener('drop', function(e){ if (e.dataTransfer && e.dataTransfer.files.length) importFiles(e.dataTransfer.files); });
  window.addEventListener('dragover', function(e){ e.preventDefault(); });
  window.addEventListener('drop', function(e){ e.preventDefault(); });
}
function checkRestore(){
  var raw = null;
  try { raw = localStorage.getItem(STORE_KEY); } catch(e){}
  if (!raw) return;
  var d; try { d = JSON.parse(raw); } catch(e){ return; }
  if (!d || !d.items || !d.items.length) return;
  var t = d.savedAt ? new Date(d.savedAt) : null;
  $('restoreText').textContent = '上次有未完成的工作：' + d.items.length + ' 項' + (d.source ? '（' + d.source + '）' : '') + (t ? '，' + t.toLocaleString('zh-TW') : '') + '。';
  $('restoreBanner').classList.add('show'); $('main').classList.add('has-banner');
  var close = function(){ $('restoreBanner').classList.remove('show'); $('main').classList.remove('has-banner'); };
  $('btnRestore').onclick = function(){ loadSnapshot(d); close(); toast('已恢復上次的工作'); };
  $('btnDiscard').onclick = function(){ close(); try { localStorage.removeItem(STORE_KEY); } catch(e){} };
}

/* ================= 對話框 / 提示 ================= */
function choiceModal(title, text, btns){
  $('modalTitle').textContent = title;
  $('modalText').textContent = text;
  var acts = $('modalActs'); acts.innerHTML = '';
  var close = function(){ $('modal').classList.remove('show'); document.removeEventListener('keydown', onKey); };
  var onKey = function(e){ if (e.key === 'Escape') close(); };
  btns.forEach(function(b){
    var el = document.createElement('button');
    el.type = 'button'; el.textContent = b.label; if (b.primary) el.className = 'primary';
    el.onclick = function(){ close(); if (b.fn) b.fn(); };
    acts.appendChild(el);
  });
  $('modal').classList.add('show');
  document.addEventListener('keydown', onKey);
  var p = acts.querySelector('.primary') || acts.lastChild; if (p) p.focus();
}
function alertModal(title, text){ choiceModal(title, text, [{label:'知道了', primary:true}]); }
var toastTimer = null;
function toast(msg, undo){
  $('toastText').textContent = msg;
  var b = $('toastBtn');
  b.hidden = !undo;
  b.onclick = undo ? function(){ undo(); $('toast').classList.remove('show'); } : null;
  $('toast').classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function(){ $('toast').classList.remove('show'); }, undo ? 7000 : 3000);
}
function showBusy(on){ $('busy').classList.toggle('show', !!on); }

/* ================= 民國日期 ================= */
function rocDate(d){
  d = d || new Date();
  return (d.getFullYear() - 1911) + '年' + (d.getMonth() + 1) + '月' + d.getDate() + '日';
}

/* ================= 鎖定檢證編號 ================= */
function renderLock(){
  var el = $('lockArea'); if (!el) return;
  if (!S.locked){
    var n = numberedRows().length;
    el.innerHTML = n ? '<button type="button" class="lbtn" data-lock="lock" title="書狀送出後按下。鎖定後編號固定，之後新增的聲請從最後一號往下接續">鎖定檢證編號</button><span class="note">書狀送出後按下，避免日後增刪造成編號變動</span>' : '';
    return;
  }
  var nw = newSinceLock().length;
  el.innerHTML = '<span class="lkstat">已鎖定檢證1～' + S.lockMax + (S.lockedAt ? '（' + esc(S.lockedAt) + '）' : '') + '</span>' +
    (nw ? '<button type="button" class="lbtn" data-lock="more" title="補充聲請狀送出後按下，將新增項目的編號也固定">鎖定新增的 ' + nw + ' 項</button>' : '') +
    '<button type="button" class="lbtn" data-lock="unlock">解除鎖定</button>' +
    '<label><input type="checkbox" id="onlyNew"' + (S.onlyNew ? ' checked' : '') + '> 只輸出鎖定後新增的項目（' + nw + ' 項）</label>';
}
function doLock(){
  var rows = numberedRows();
  var mx = S.lockMax || 0;
  rows.forEach(function(r){ if (r.it.lockedNo == null) r.it.lockedNo = r.num; if (r.num > mx) mx = r.num; });
  S.locked = true; S.lockMax = mx; S.lockedAt = rocDate(); S.onlyNew = false;
  renderList(); renderPreview(); save();
}
function bindLock(){
  $('lockArea').addEventListener('click', function(e){
    var b = e.target.closest('[data-lock]'); if (!b) return;
    var k = b.dataset.lock;
    if (k === 'lock'){
      var n = numberedRows().length;
      choiceModal('鎖定檢證編號', '鎖定後，目前的檢證1～' + n + ' 編號固定不變；之後新增的聲請會從檢證' + (n + 1) + ' 往下接續編號，刪除已鎖定的項目則會留下空號。',
        [{label:'取消'},{label:'鎖定', primary:true, fn:function(){ doLock(); toast('已鎖定檢證1～' + S.lockMax); }}]);
    } else if (k === 'more'){
      var nw = newSinceLock().length;
      choiceModal('鎖定新增項目', '將鎖定後新增的 ' + nw + ' 項編號一併固定（補充聲請狀送出後使用）。', [{label:'取消'},{label:'鎖定', primary:true, fn:function(){ doLock(); toast('已鎖定至檢證' + S.lockMax); }}]);
    } else if (k === 'unlock'){
      choiceModal('解除鎖定', '解除後，檢證編號會依目前表格順序重新連續編號，可能與已送出的書狀及法院裁定的編號不一致。', [{label:'取消'},{label:'解除鎖定', primary:true, fn:function(){
        S.items.forEach(function(it){ it.lockedNo = null; });
        S.locked = false; S.lockMax = 0; S.lockedAt = ''; S.onlyNew = false;
        renderList(); renderPreview(); save(); toast('已解除鎖定');
      }}]);
    }
  });
  $('lockArea').addEventListener('change', function(e){
    if (e.target.id === 'onlyNew'){ S.onlyNew = e.target.checked; renderPreview(); save(); }
  });
}

/* ================= 右側分頁 ================= */
function showTab(t){
  UI.tab = t;
  document.querySelectorAll('[data-tab]').forEach(function(b){ b.setAttribute('aria-selected', b.dataset.tab === t ? 'true' : 'false'); });
  document.querySelectorAll('[data-panel]').forEach(function(el){ el.hidden = el.dataset.panel !== t; });
  if (t === 'cmp') renderCmp();
  if (t === 'rul') renderRul();
}
function bindTabs(){
  document.querySelector('.rtab-btns').addEventListener('click', function(e){
    var b = e.target.closest('[data-tab]'); if (b) showTab(b.dataset.tab);
  });
}
function chips(defs, counts, cur, attr){
  return defs.map(function(d){
    return '<button type="button" class="tab t-' + d[0] + '" ' + attr + '="' + d[0] + '" aria-pressed="' + (cur === d[0]) + '">' + d[1] + '<b>' + counts[d[0]] + '</b></button>';
  }).join('');
}

/* ================= 清冊對照表 ================= */
var CMP_DEFS = [['all','全部'],['apply','聲請'],['off','不聲請'],['pend','未決定']];
function cmpRows(){
  var nm = numMap();
  return S.items.map(function(it){
    var st = statusOf(it), n = nm[it.id];
    var key = st === 'off' ? 'off' : (st === 'pend' ? 'pend' : 'apply');
    return {id:it.id, key:key,
      no: it.no || (it.manual ? '（手動新增）' : ''),
      name: it.name,
      apply: key === 'off' ? '不聲請' : (key === 'pend' ? '未決定' : '聲請'),
      kz: n ? '檢證' + n.num : '',
      cat: it.apply === true ? (it.cat ? CATS[it.cat] : '（未選）') : '',
      method: it.apply === true ? METHOD_LABEL[it.method] : ''};
  });
}
function cmpCounts(rows){
  var c = {all:rows.length, apply:0, off:0, pend:0};
  rows.forEach(function(r){ c[r.key]++; });
  return c;
}
function cmpSummary(rows){
  var c = cmpCounts(rows);
  var cr = 0, se = 0, np = 0;
  S.items.forEach(function(it){ if (it.apply === true){ if (it.cat === 'crime') cr++; else if (it.cat === 'sentence') se++; else np++; } });
  return '共 ' + c.all + ' 項：聲請 ' + c.apply + ' 項（犯罪事實 ' + cr + '、科刑資料 ' + se + (np ? '、未選類別 ' + np : '') + '），不聲請 ' + c.off + ' 項，未決定 ' + c.pend + ' 項。';
}
function cmpFiltered(){
  var rows = cmpRows();
  return UI.cmpFilter === 'all' ? rows : rows.filter(function(r){ return r.key === UI.cmpFilter; });
}
function renderCmp(){
  var all = cmpRows();
  $('cmpTabs').innerHTML = chips(CMP_DEFS, cmpCounts(all), UI.cmpFilter, 'data-cf');
  var rows = cmpFiltered();
  $('btnCmpCopy').disabled = $('btnCmpDocx').disabled = !rows.length;
  var paper = $('cmpPaper');
  if (!S.items.length){ paper.innerHTML = '<div class="pv-empty">請先在左側匯入證據清冊。</div>'; return; }
  var h = '<p class="sum">' + esc(cmpSummary(all)) + '</p>';
  if (!rows.length){ paper.innerHTML = h + '<div class="pv-empty">這個分類沒有項目。</div>'; return; }
  h += '<table class="cmp"><thead><tr><th>清冊編號</th><th class="nm">證據名稱</th><th>是否聲請</th><th>檢證編號</th><th>類別</th><th>調查方法</th></tr></thead><tbody>';
  rows.forEach(function(r){
    h += '<tr class="pr" data-id="' + r.id + '" title="點一下可跳到左側清單中的這一項"><td>' + esc(r.no) + '</td><td class="nl">' + esc(r.name) + '</td><td class="k-' + r.key + '">' + r.apply + '</td><td>' + esc(r.kz) + '</td><td>' + esc(r.cat) + '</td><td>' + esc(r.method) + '</td></tr>';
  });
  paper.innerHTML = h + '</tbody></table>';
}
function cmpSpec(){
  var rows = cmpFiltered();
  var lbl = CMP_DEFS.filter(function(d){ return d[0] === UI.cmpFilter; })[0][1];
  return {
    title:'證據清冊與聲請調查證據對照表',
    lines:[(S.source ? '證據清冊：' + S.source + '　' : '') + '分類：' + lbl + '　製表日期：' + rocDate(), cmpSummary(cmpRows())],
    cols:[{label:'清冊編號', chars:4.3},{label:'證據名稱', flex:1},{label:'是否聲請', chars:4.3},{label:'檢證編號', chars:4.3},{label:'類別', chars:4.3},{label:'調查方法', chars:4.3}],
    rows:rows.map(function(r){ return [r.no, r.name, r.apply, r.kz, r.cat, r.method]; }),
    file:'證據清冊對照表_' + lbl
  };
}
function bindCmp(){
  $('cmpTabs').addEventListener('click', function(e){
    var b = e.target.closest('[data-cf]'); if (!b) return;
    UI.cmpFilter = b.dataset.cf; renderCmp();
  });
  $('cmpPaper').addEventListener('click', function(e){
    var tr = e.target.closest('tr.pr'); if (tr) gotoItem(tr.dataset.id);
  });
  $('btnCmpCopy').addEventListener('click', function(){ copyRef(cmpSpec()); });
  $('btnCmpDocx').addEventListener('click', function(){ downloadRef(cmpSpec()); });
}

/* ================= 法院裁定對照表 ================= */
var RUL_DEFS = [['all','全部'],['grant','准許'],['sentonly','僅准作科刑資料'],['reserve','保留'],['deny','駁回'],['none','未填']];
function rulRows(){
  return numberedRows().slice().sort(function(a, b){ return a.num - b.num; }).map(function(r){
    var it = r.it;
    return {id:it.id, it:it, num:r.num, no:'檢證' + r.num, name:it.name, cat:r.cat, method:r.method,
      key: it.ruling || 'none', ruling: it.ruling ? RUL_LABEL[it.ruling] : '', note: it.rulingNote || ''};
  });
}
function rulCounts(rows){
  var c = {all:rows.length, grant:0, sentonly:0, reserve:0, deny:0, none:0};
  rows.forEach(function(r){ c[r.key]++; });
  return c;
}
function rulSummary(rows){
  var c = rulCounts(rows);
  return '已聲請 ' + c.all + ' 項：准許 ' + c.grant + ' 項，僅准作科刑資料 ' + c.sentonly + ' 項，保留 ' + c.reserve + ' 項，駁回 ' + c.deny + ' 項，未填 ' + c.none + ' 項。';
}
function rulFiltered(){
  var rows = rulRows();
  return UI.rulFilter === 'all' ? rows : rows.filter(function(r){ return r.key === UI.rulFilter; });
}
function rulRowHtml(r){
  var it = r.it;
  var h = '<tr class="rr r-' + r.key + '" data-id="' + r.id + '">';
  h += '<td><input type="checkbox" class="rsel" aria-label="勾選' + r.no + '"' + (UI.rulSel.has(r.id) ? ' checked' : '') + '></td>';
  h += '<td class="rno" title="點一下可跳到左側清單中的這一項">' + r.no + '<small>' + CATS[r.cat] + '・' + METHOD_LABEL[r.method] + '</small></td>';
  h += '<td class="rname" title="點一下可跳到左側清單中的這一項">' + esc(r.name) + '</td>';
  h += '<td class="rseg"><div class="seg rul" role="group" aria-label="法院裁定">' + RULINGS.filter(function(k){
      return k !== 'sentonly' || r.cat === 'crime' || it.ruling === 'sentonly';
    }).map(function(k){ return segBtn('rul', k, RUL_BTN[k], it.ruling === k).replace('<button ', '<button title="' + RUL_LABEL[k] + '" '); }).join('') + '</div></td>';
  h += '<td><textarea class="rnote" rows="1" placeholder="裁定日期、理由等" aria-label="備註">' + esc(r.note) + '</textarea></td>';
  return h + '</tr>';
}
function renderRul(){
  var all = rulRows();
  UI.rulSel.forEach(function(id){ if (!all.some(function(r){ return r.id === id; })) UI.rulSel.delete(id); });
  $('rulTabs').innerHTML = chips(RUL_DEFS, rulCounts(all), UI.rulFilter, 'data-rf');
  var rows = rulFiltered();
  $('btnRulCopy').disabled = $('btnRulDocx').disabled = !rows.length;
  $('btnRulFill').disabled = !all.some(function(r){ return r.key === 'none'; });
  var n = UI.rulSel.size;
  $('rulBulk').classList.toggle('show', n > 0);
  $('rulBulkCount').textContent = '已勾選 ' + n + ' 項：';
  var body = $('rulBody');
  if (!all.length){ body.innerHTML = '<div class="pv-empty">還沒有聲請調查的證據。請先在左側清單選擇「聲請」及類別。</div>'; return; }
  var h = '<div class="rsum">' + esc(rulSummary(all)) + '</div>';
  if (!rows.length){ body.innerHTML = h + '<div class="pv-empty">這個分類沒有項目。</div>'; return; }
  var allSel = rows.every(function(r){ return UI.rulSel.has(r.id); });
  h += '<table class="rt"><thead><tr><th style="width:28px"><input type="checkbox" id="rulSelAll" aria-label="全選目前顯示"' + (allSel ? ' checked' : '') + '></th><th>檢證編號</th><th>證據名稱</th><th>法院裁定</th><th>備註</th></tr></thead><tbody>';
  h += rows.map(rulRowHtml).join('') + '</tbody></table>';
  var st = body.scrollTop;
  body.innerHTML = h;
  body.scrollTop = st;
  body.querySelectorAll('.rnote').forEach(autosize);
}
function setRuling(it, k){
  if (k === 'clear'){ it.ruling = null; return true; }
  if (k === 'sentonly' && it.cat !== 'crime') return false;
  it.ruling = k; return true;
}
function rulSpec(){
  var rows = rulFiltered();
  var lbl = RUL_DEFS.filter(function(d){ return d[0] === UI.rulFilter; })[0][1];
  return {
    title:'聲請調查證據法院裁定對照表',
    lines:['分類：' + lbl + '　製表日期：' + rocDate(), rulSummary(rulRows())],
    cols:[{label:'檢證編號', chars:4.3},{label:'證據名稱', flex:3},{label:'聲請類別', chars:4.3},{label:'法院裁定', chars:7.3},{label:'備註', flex:2}],
    rows:rows.map(function(r){ return [r.no, r.name, CATS[r.cat], r.ruling, r.note]; }),
    file:'法院裁定對照表_' + lbl
  };
}
function bindRul(){
  $('rulTabs').addEventListener('click', function(e){
    var b = e.target.closest('[data-rf]'); if (!b) return;
    UI.rulFilter = b.dataset.rf; renderRul();
  });
  var body = $('rulBody');
  body.addEventListener('click', function(e){
    var tr = e.target.closest('tr.rr'); if (!tr) return;
    var it = findItem(tr.dataset.id); if (!it) return;
    var b = e.target.closest('[data-act="rul"]');
    if (b){
      var k = b.dataset.v;
      setRuling(it, it.ruling === k ? 'clear' : k);
      renderRul(); save();
      var again = document.querySelector('#rulBody tr[data-id="' + it.id + '"] [data-act="rul"][data-v="' + k + '"]');
      if (again) again.focus();
      return;
    }
    if (e.target.closest('.rno, .rname')) gotoItem(it.id);
  });
  body.addEventListener('change', function(e){
    if (e.target.id === 'rulSelAll'){
      rulFiltered().forEach(function(r){ if (e.target.checked) UI.rulSel.add(r.id); else UI.rulSel.delete(r.id); });
      renderRul(); return;
    }
    if (e.target.classList.contains('rsel')){
      var id = e.target.closest('tr').dataset.id;
      if (e.target.checked) UI.rulSel.add(id); else UI.rulSel.delete(id);
      var n = UI.rulSel.size;
      $('rulBulk').classList.toggle('show', n > 0);
      $('rulBulkCount').textContent = '已勾選 ' + n + ' 項：';
    }
  });
  body.addEventListener('input', function(e){
    if (!e.target.classList.contains('rnote')) return;
    var it = findItem(e.target.closest('tr').dataset.id); if (!it) return;
    it.rulingNote = e.target.value; autosize(e.target); save();
  });
  $('rulBulk').addEventListener('click', function(e){
    var b = e.target.closest('[data-rb]'); if (!b) return;
    var k = b.dataset.rb;
    if (k === 'none'){ UI.rulSel.clear(); renderRul(); return; }
    var ok = 0, skip = 0;
    UI.rulSel.forEach(function(id){ var it = findItem(id); if (!it) return; if (setRuling(it, k)) ok++; else skip++; });
    UI.rulSel.clear();
    renderRul(); save();
    toast('已設定 ' + ok + ' 項' + (skip ? '；另有 ' + skip + ' 項原本即以科刑資料聲請，未套用「僅准作科刑資料」' : ''));
  });
  $('btnRulFill').addEventListener('click', function(){
    var n = 0;
    rulRows().forEach(function(r){ if (!r.it.ruling){ r.it.ruling = 'grant'; n++; } });
    renderRul(); save();
    toast('已將 ' + n + ' 項設為准許');
  });
  $('btnRulCopy').addEventListener('click', function(){ copyRef(rulSpec()); });
  $('btnRulDocx').addEventListener('click', function(){ downloadRef(rulSpec()); });
}

/* ================= 對照表輸出（Word 與複製共用） ================= */
function refWidths(cols, pt, total){
  var cw = pt * 20, pad = 85;
  var fixed = cols.map(function(c){ return c.chars ? Math.round(c.chars * cw + pad * 2) : 0; });
  var flexW = cols.reduce(function(a, c){ return a + (c.flex || 0); }, 0);
  var minFlex = cols.reduce(function(a, c){ return a + (c.flex ? (cw * 5 * c.flex + pad * 2) : 0); }, 0);
  var sumFixed = fixed.reduce(function(a, b){ return a + b; }, 0);
  if (total - sumFixed < minFlex){
    var k = (total - minFlex) / sumFixed;
    fixed = fixed.map(function(w){ return Math.round(w * k); });
    sumFixed = fixed.reduce(function(a, b){ return a + b; }, 0);
  }
  var rest = total - sumFixed, used = 0, lastFlex = -1;
  var w = cols.map(function(c, i){
    if (!c.flex) return fixed[i];
    lastFlex = i;
    var x = Math.round(rest * c.flex / flexW); used += x; return x;
  });
  if (lastFlex >= 0) w[lastFlex] += rest - used;
  return w;
}
function buildRefDocumentXml(spec, pt){
  var pageW = 16838, pageH = 11906, m = 1134;     // A4 橫式，邊界 2 公分
  var total = pageW - m * 2;
  var W = refWidths(spec.cols, pt, total);
  var para = function(text, extra){
    return '<w:p><w:pPr><w:spacing w:before="0" w:after="60"/>' + (extra && extra.center ? '<w:jc w:val="center"/>' : '') + '</w:pPr>' +
      (extra && extra.big ? runXml(text).replace(/<w:r>/g, '<w:r><w:rPr><w:b/><w:sz w:val="' + (pt * 2 + 4) + '"/><w:szCs w:val="' + (pt * 2 + 4) + '"/></w:rPr>') : runXml(text)) + '</w:p>';
  };
  var cell = function(w, text){
    return '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' + runXml(text) + '</w:p></w:tc>';
  };
  var bd = function(n){ return '<w:' + n + ' w:val="single" w:sz="4" w:space="0" w:color="000000"/>'; };
  var x = '<w:tbl><w:tblPr><w:tblW w:w="' + total + '" w:type="dxa"/><w:tblInd w:w="0" w:type="dxa"/><w:tblBorders>' + bd('top') + bd('left') + bd('bottom') + bd('right') + bd('insideH') + bd('insideV') + '</w:tblBorders>' +
    '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="28" w:type="dxa"/><w:left w:w="85" w:type="dxa"/><w:bottom w:w="28" w:type="dxa"/><w:right w:w="85" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>' +
    W.map(function(w){ return '<w:gridCol w:w="' + w + '"/>'; }).join('') + '</w:tblGrid>';
  x += '<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>' + spec.cols.map(function(c, i){ return cell(W[i], c.label); }).join('') + '</w:tr>';
  spec.rows.forEach(function(r){ x += '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + r.map(function(v, i){ return cell(W[i], v || ''); }).join('') + '</w:tr>'; });
  x += '</w:tbl>';
  var body = para(spec.title, {center:true, big:true}) + spec.lines.map(function(l){ return para(l); }).join('') + x + '<w:p/>';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="' + W_NS + '" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>' + body +
    '<w:sectPr><w:pgSz w:w="' + pageW + '" w:h="' + pageH + '" w:orient="landscape"/><w:pgMar w:top="' + m + '" w:right="' + m + '" w:bottom="' + m + '" w:left="' + m + '" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr></w:body></w:document>';
}
async function downloadRef(spec){
  if (!spec.rows.length){ toast('這個分類沒有項目'); return; }
  var pt = S.fontPt || 14;
  var blob = await makeDocxBlob(buildRefDocumentXml(spec, pt));
  var b = baseName();
  saveBlob(blob, spec.file + (b ? '_' + b : '') + '.docx');
  toast('已下載 Word 檔');
}
function refClipHtml(spec){
  var pt = S.fontPt || 14;
  var W = refWidths(spec.cols, pt, tableTwips());
  var total = W.reduce(function(a, b){ return a + b; }, 0);
  var tw2pt = function(t){ return (t / 20).toFixed(1); };
  var tw2px = function(t){ return Math.round(t / 20 * 96 / 72); };
  var fam = 'font-family:標楷體;mso-ascii-font-family:標楷體;mso-hansi-font-family:標楷體;mso-fareast-font-family:標楷體';
  var run = function(t){ return '<span style="font-size:' + pt + '.0pt;' + fam + '">' + esc(t).replace(/\n/g, '<br>') + '</span>'; };
  var p = function(t, center){ return '<p class="MsoNormal" style="margin:0cm;line-height:normal;' + (center ? 'text-align:center;' : '') + 'font-size:' + pt + '.0pt;' + fam + '">' + run(t) + '</p>'; };
  var td = function(w, t){
    return '<td width="' + tw2px(w) + '" valign="top" style="width:' + tw2pt(w) + 'pt;border:solid windowtext 1.0pt;mso-border-alt:solid windowtext .5pt;padding:1.4pt 4.25pt 1.4pt 4.25pt">' +
      '<p class="MsoNormal" style="margin:0cm;line-height:normal;text-align:left;font-size:' + pt + '.0pt;' + fam + '">' + run(t || '') + '</p></td>';
  };
  var row = function(vals){ return '<tr style="page-break-inside:avoid">' + vals.map(function(v, i){ return td(W[i], v); }).join('') + '</tr>'; };
  var h = p(spec.title, true) + spec.lines.map(function(l){ return p(l); }).join('');
  h += '<table class="MsoTableGrid" border="1" cellspacing="0" cellpadding="0" width="' + tw2px(total) + '" style="width:' + tw2pt(total) + 'pt;border-collapse:collapse;border:none;mso-border-alt:solid windowtext .5pt;mso-table-layout-alt:fixed;mso-padding-alt:1.4pt 4.25pt 1.4pt 4.25pt">';
  h += row(spec.cols.map(function(c){ return c.label; }));
  spec.rows.forEach(function(r){ h += row(r); });
  return h + '</table>';
}
function refClipText(spec){
  var out = [spec.title].concat(spec.lines);
  out.push(spec.cols.map(function(c){ return c.label; }).join('\t'));
  spec.rows.forEach(function(r){ out.push(r.map(function(v){ return String(v || '').replace(/\n/g, ' '); }).join('\t')); });
  return out.join('\r\n');
}
async function copyRef(spec){
  if (!spec.rows.length){ toast('這個分類沒有項目'); return; }
  var html = refClipHtml(spec), text = refClipText(spec);
  var done = copyViaEvent(html, text);
  if (!done && navigator.clipboard && window.ClipboardItem){
    try { await navigator.clipboard.write([new ClipboardItem({'text/html':new Blob([html],{type:'text/html'}),'text/plain':new Blob([text],{type:'text/plain'})})]); done = true; } catch(e){}
  }
  toast(done ? '已複製，可到 Word 或漢書中貼上' : '瀏覽器不允許複製，請改用「下載 Word」');
}

/* ================= 啟動 ================= */
setupPdf();
bindTop(); bindList(); bindPreview(); bindLock(); bindTabs(); bindCmp(); bindRul();
renderAll();
checkRestore();

// 供測試用
window.__evtool = {S:S, UI:UI, numberedRows:numberedRows, cmpSpec:cmpSpec, rulSpec:rulSpec, refClipHtml:refClipHtml, clipHtml:clipHtml, clipText:clipText, suggestMethod:suggestMethod, rowsToItems:rowsToItems, buildOutline:buildOutline, makeDocxBlob:makeDocxBlob, parseDocx:parseDocx, parsePdf:parsePdf, renderAll:renderAll, importFiles:importFiles};
})();
