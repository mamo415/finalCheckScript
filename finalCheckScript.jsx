/*
====================================================================
====================================================================
 -------------------------------------------------------------------
 PURPOSE  :  Automatically detects 14 common pre-press risks inside an
             Illustrator document and visualises them as red frames on a
             dedicated layer while saving a tab-separated text report
             next to the artwork file.
            + Adds processed-item cache, detailed log output, and a
              ScriptUI progress bar.
-------------------------------------------------------------------
*/

#target "Illustrator"
(function () {

/* ==================================================================
   0. CONSTANTS & MESSAGE TABLES
   ==================================================================*/

// しきい値設定 ------------------------------------------------------
var MIN_LINE_MM   = 0.12;                     // No.6 細線判定 (mm) ⭐ 更新: 0.12 mm
var PT_PER_MM     = 2.834645669;              // 1 mm ≒ 2.8346 pt
var MIN_LINE_PT   = MIN_LINE_MM * PT_PER_MM;  // 0.12 mm → pt
var VIS_DIFF_PT   = 0.1;                      // bounds 差分判定 (pt)

// TXT レポート装飾 --------------------------------------------------
var COL_WIDTH = 30;                           // 列幅 30 桁固定
var FATAL_SET = {6:true, 10:true, 14:true, 15:true};   // 重大度マーク対象
var SEP_TOP   = "┏" + repeat("━", 70) + "┓";  // 70 桁
var SEP_BOT   = "┗" + repeat("━", 70) + "┛";  // 70 桁
var SEP_MID   = repeat("=", 60);              // 集計セクション境界

// 推奨対処メッセージ ------------------------------------------------
var FIX_MSG = [
    "", "Blend を分割","シンボルを展開","塗りのみ開パス修正","縦横比を修正",
    "外観を分割・拡張","線幅を0.3pt以上","塗りを追加","パスを閉じる",
    "文字をアウトライン","画像を再リンク","画像をCMYK化",
    "効果を分割･拡張","不透明度100%に","不要オブジェクト削除", 
    "パターンオブジェクトのFIX文章",
];

// エラー種別ラベル --------------------------------------------------
var ERR_LABEL = [
    "", "未展開 Blend","Symbol / Plugin","塗りのみ開パス","細長パス",
    "外観ストロークG","細線","塗りなし線","開いた線",
    "未アウトライン文字","画像リンク切れ","RGB画像",
    "アピアランス効果","不透明","空オブジェクト",
];

// 文字列を n 回繰り返して返すユーティリティ ------------------------
function repeat(ch, n) {
    var s = ""; while (n--) s += ch;
    return s;
}

/* ==================================================================
   0-B) NEW UTILITIES (v1.3.3) + 1-B) CACHE/LOG/PROGRESS (v1.3.4)
   ==================================================================*/

// Safe bounds with visibleBounds fallback --------------------------
function getSafeBounds(it) {
    var gb = ("geometricBounds" in it) ? it.geometricBounds : null;
    if (!gb) return it.visibleBounds;
    var w = Math.abs(gb[2] - gb[0]);
    var h = Math.abs(gb[1] - gb[3]);
    if (w < 0.01 || h < 0.01) return it.visibleBounds;
    return gb;
}

// Unlock all parent containers ------------------------------------
function unlockParents(obj) {
    try {
        while (obj && obj.locked !== undefined) {
            if (obj.locked) obj.locked = false;
            obj = obj.parent;
        }
    } catch (e) { /* ignore */ }
}

// Safe PlacedItem file check with robust try/catch -----------------
function isLinkMissing(it) {
    try {
        var f = null;
        try { f = it.file; } catch (_) { return true; }
        if (!f) return true;
        return (typeof f.exists === 'boolean') ? !f.exists : true;
    } catch (e) { return true; }
}

/* --- ★NEW★: プロセス済みキャッシュ & ログ & 進捗バー変数 --- */
var processedItems = [];   
var logLines       = [];   
var canceled       = false;            // ★NEW★
var PROGRESS_UPDATE_INTERVAL = 20;     // ★NEW v1.3.6★ 進捗バー更新間隔 :contentReference[oaicite:0]{index=0}

// ★Polyfill for Array.indexOf (ExtendScript対応) ★
if (!Array.prototype.indexOf) {
    Array.prototype.indexOf = function(searchElement) {
        for (var i = 0; i < this.length; i++) {
            if (this[i] === searchElement) return i;
        }
        return -1;
    };
}

// ★修正★: ExtendScript 互換の ISO タイムスタンプ生成 ---------------
function timestamp() {
    var d = new Date();
    return d.getFullYear() + '-' +
           pad2(d.getMonth() + 1) + '-' +
           pad2(d.getDate()) + 'T' +
           pad2(d.getHours()) + ':' +
           pad2(d.getMinutes()) + ':' +
           pad2(d.getSeconds());
}
// 2桁ゼロ埋めユーティリティ
function pad2(n) {
    return (n < 10 ? '0' : '') + n;
}

// 全アイテム数を再帰カウント（進捗バー用） --------------------------
function countAllItems(items) {
    var cnt = 0;
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        cnt++;
        if (it.typename === "GroupItem") 
            cnt += countAllItems(it.pageItems);
        else if (it.typename === "CompoundPathItem") 
            cnt += countAllItems(it.pathItems);
    }
    return cnt;
}

// ScriptUI：進捗バー初期化 --------------------------------------------
var totalItems    = 0, processedCount = 0, progressWin = null;
function initProgress() {
    totalItems = countAllItems(doc.pageItems);
    progressWin = new Window('palette', 'Preflight Progress', undefined, {closeButton:false});
    progressWin.orientation = 'column';
    progressWin.add('statictext', undefined, 'Processing... 0 / ' + totalItems);
    var pb = progressWin.add('progressbar', undefined, 0, totalItems);
    pb.preferredSize.width = 300;
    progressWin.pb  = pb;
    progressWin.txt = progressWin.children[0];
    progressWin.show();
}

// ScriptUI：進捗バー更新 ---------------------------------------------
function updateProgress(){
    if(canceled) return;
    processedCount++;
    // 20件ごと or 最終件のみ UI 更新
    var doUpdate = (processedCount % PROGRESS_UPDATE_INTERVAL === 0)
                 || (processedCount === totalItems);
    if(doUpdate){
        try{ progressWin.pb.value = processedCount; }catch(e){}
        // パーセント表示
        var pct = Math.round(processedCount/totalItems*100);
        try{
            progressWin.txt.text = '「'+ currentLayer +'」をチェック中… ' + pct + '%';
            progressWin.update();
        }catch(e){}
    }
    // ログには毎回書き込む
    logLines.push('['+timestamp()+'] Processed '+ processedCount +
                  ' / ' + totalItems +
                  ' → Layer: ' + currentLayer +
                  ', Obj: ' + currentObj);
}

// ログファイル書き出し -----------------------------------------------
function writeLog() {
    var base   = doc.name.replace(/\.[^\.]+$/, '');
    var folder = doc.fullName.path;
    var lf     = new File(folder + '/' + base + '_log.txt');
    lf.encoding = 'UTF-8';
    lf.open('w');
    lf.writeln('=== Preflight Log (' + timestamp() + ') ===');
    for (var i = 0; i < logLines.length; i++) {
        lf.writeln(logLines[i]);
    }
    lf.close();
}

/* ==================================================================
   1. ENVIRONMENT SETUP
   ==================================================================*/

var doc = app.activeDocument;
if (!doc) {
    alert("開いているドキュメントがありません。");
    return;
}
var originalUI = app.userInteractionLevel;

// エラー可視化レイヤー作成 + ロック解除 -----------------------------
var errorLayer = doc.layers.add();
errorLayer.name = "エラービジュアル";
unlockParents(errorLayer);

// UI 非対話モードへ (高速化) ----------------------------------------
try { app.displayDialogs = DialogModes.NO; } catch(e){}
app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

/* ==================================================================
   2. GLOBAL STATE VARIABLES
   ==================================================================*/

var bucketLines = Array(16), countByNo = Array(16);
for(var i=0; i<16; i++){
    bucketLines[i]=[];
    countByNo[i]=0;
}
// 重大／中／軽微 用のオブジェクト配列を用意
var errorsMajorArray  = [];
var errorsMediumArray = [];
var errorsMinorArray  = [];


var errorCount   = 0;
var redRectsBuf  = [];
var unnamedCount = 0;

var currentItem, currentLayer, currentObj;

/* ==================================================================
   3. bucket() 定義
   ==================================================================*/

function bucket(no, desc) {
    var fix  = FIX_MSG[no];
    if(FIX_MSG.length <= fix) {
        alert("FIX_MSGの"+no+"番目のメッセージがありません");
    }
    var mark = FATAL_SET[no] ? "‼ " : "  ";
    var row  = [
        pad(no, 4),
        pad(desc, COL_WIDTH),
        pad(fix, COL_WIDTH),
        currentLayer,
        currentObj
    ].join("\t");
    bucketLines[no].push(mark + row);
    countByNo[no]++;
    errorCount++;
    redRectsBuf.push(getSafeBounds(currentItem));
    // ログにも記録
    logLines.push('[' + timestamp() + '] ERROR No.' + no +
                  ' ' + desc +
                  ' → Layer: ' + currentLayer +
                  ', Obj: ' + currentObj);
                     var record = {
      typeId:    no,
      name:      desc,
      layerPath: currentLayer,
      objectName: currentObj
    };
    // 重大エラーの場合 (たとえばNo.6,10,14,15)
    if (no===6 || no===10 || no===14 || no===15) {
      errorsMajorArray.push(record);
    }
    // 中エラーの場合 (例: No.1~5,11~13)
    else if (no===1||no===2||no===3||no===4||no===5||no===11||no===12||no===13) {
      errorsMediumArray.push(record);
    }
    // 軽微エラーの場合 (例: No.7~8)
    else {
      errorsMinorArray.push(record);
    }
}

/* ==================================================================
   4. MAIN EXECUTION FLOW
   ==================================================================*/

// ★NEW★ 進捗バー初期化 & ログ開始
initProgress();
logLines.push('[' + timestamp() + '] Preflight started for "' + doc.name + '".');

processItems(doc.pageItems);

// 赤枠描画
for (var r = 0; r < redRectsBuf.length; r++) {
    drawRedBox(redRectsBuf[r]);
}
app.userInteractionLevel = originalUI;

if(0 < errorCount) {
    // レポート書き出し
    saveReport();
    // 完了アラート + UI復帰
    alert("Preflight 完了\n検出件数: " + errorCount);
}
else {
    errorLayer.remove();
    alert("🟢 気になるエラーはありませんでした！");
}




/* ==================================================================
   5. FUNCTION DEFINITIONS
   ==================================================================*/

// 階層ロック判定 ------------------------------------------------------
function isLayerLocked(item) {
    try {
        var l = item.layer;
        while (l) {
            if (l.locked) return true;
            l = l.parent;
        }
    } catch(e){}
    return false;
}

// 外観ストロークグループ判定 ----------------------------------------
function hasStrokeAppearance(gr) {
    if (!gr.stroked || gr.strokeWidth <= 0) return false;
    var gb = getSafeBounds(gr);
    var cb = gr.geometricBounds;
    var diff = Math.abs((cb[2] - cb[0]) - (gb[2] - gb[0]))
             + Math.abs((cb[1] - cb[3]) - (gb[1] - cb[3]));
    var childStroke = false;
    for (var i = 0; i < gr.pageItems.length; i++) {
        var c = gr.pageItems[i];
        if (c.stroked && c.strokeWidth > 0) { childStroke = true; break; }
    }
    return !childStroke && diff > VIS_DIFF_PT;
}

// Symbol 内部問題スキャン -------------------------------------------
function countSymbolInternalIssues(sym) {
    var cnt = 0;
    function scan(items) {
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            if (it.typename === "BlendItem" || it.typename === "SymbolItem" ||
                it.typename === "GraphItem" || it.typename === "PluginItem") {
                cnt++;
            } else if (it.typename === "PathItem") {
                var p = it;
                if (!p.stroked && p.filled && !p.closed) cnt++;
                if (p.closed) {
                    var b = getSafeBounds(p);
                    var w = Math.abs(b[2] - b[0]);
                    var h = Math.abs(b[1] - b[3]);
                    if (h > 0 && Math.max(w, h)/Math.min(w, h) > 15) cnt++;
                }
                if (p.stroked && p.strokeWidth < MIN_LINE_PT) cnt++;
                if (!p.filled && p.stroked) cnt++;
                if (p.stroked && !p.closed) cnt++;
            } else if (it.typename === "TextFrame" && it.textRange.characters.length > 0) {
                cnt++;
            } else if (it.typename === "PlacedItem") {
                if (isLinkMissing(it)) cnt++;
                else if (it.imageColorSpace === ImageColorSpace.RGB) cnt++;
            } else if ("appliedEffects" in it && it.appliedEffects.length > 0) {
                cnt++;
            } else if (("opacity" in it) && it.opacity < 100) {
                cnt++;
            }
            if (it.pageItems && it.pageItems.length > 0) scan(it.pageItems);
        }
    }
    try { if (sym.definition && sym.definition.pageItems) scan(sym.definition.pageItems); } catch(e){}
    return cnt;
}

// アイテム再帰処理 (DFS) ---------------------------------------------
function processItems(items) {

    for (var i = 0; i < items.length; i++) {
        var item = items[i];

        // ── 重複チェック ──
        if (processedItems.indexOf(item) >= 0) {
            continue;
        }
        processedItems.push(item);

        // ── 表示・ガイド・ロック・クリッピングはスキップ ──
        if (item.hidden || item.guides) {
            continue;
        }
        if (isLayerLocked(item)) continue;
        if (item.typename === "PathItem" && item.clipping) {
             continue;
        }
        // ── 進捗バー更新 ──
        currentLayer = item.layer ? item.layer.name : "(NoLayer)";
        currentObj   = item.name || "名前なし #" + (++unnamedCount);
        updateProgress();

        currentItem = item;

        // ── 1: 未展開 Blend ──
        if (item.typename === "BlendItem") {
            bucket(1, "未展開 Blend");
        }
        // ── 2: Symbol／Graph／Plugin ──
        if (item.typename === "SymbolItem") {
            var c = countSymbolInternalIssues(item);
            bucket(2, c > 0 ? "Symbol 内部 " + c + "件" : "Symbol");
        }
        else if (item.typename === "GraphItem") {
            bucket(2, "Graph");
        }
        else if (item.typename === "PluginItem") {
            bucket(2, "Plugin");
        }
        // ── 5: 外観ストロークグループ ──
        if (item.typename === "GroupItem" && hasStrokeAppearance(item)) {
            bucket(5, "外観ストロークG");
        }

        // ── 3〜8,14,15: PathItem 系チェック ──
        if (item.typename === "PathItem") {
            var p = item;

            if (p.filled && p.fillColor && p.fillColor.typename === "PatternColor") {
                bucket(15, "未アウトライン化パターン");
                redRectsBuf.push(getSafeBounds(p));
            }

            if (!p.stroked && p.filled && !p.closed) {
                bucket(3, "塗りのみ開パス");
            }
            if (p.closed) {
                var b = getSafeBounds(p),
                    w = Math.abs(b[2] - b[0]),
                    h = Math.abs(b[1] - b[3]);
                if (h > 0 && Math.max(w, h) / Math.min(w, h) > 15) {
                    bucket(4, "閉じた細長パス");
                }
            }
            if (p.stroked && p.strokeWidth < MIN_LINE_PT) {
                bucket(6, "細線");
            }
            if (!p.filled && p.stroked) {
                bucket(7, "塗りなし線");
            }
            if (p.stroked && !p.closed) {
                bucket(8, "開いた線");
            }
            if (p.pathPoints.length <= 1 || (!p.filled && !p.stroked)) {
                bucket(14, "空オブジェクト");
            }
        }

        // ── 9: 未アウトライン文字 ──
        if (item.typename === "TextFrame" && item.textRange.characters.length > 0) {
            bucket(9, "未アウトライン文字");
        }

        // ── 10〜11: 画像リンク／RGB画像 ──
        if (item.typename === "PlacedItem") {
            if (isLinkMissing(item)) {
                bucket(10, "画像リンク切れ");
            } else if (item.imageColorSpace === ImageColorSpace.RGB) {
                bucket(11, "RGB画像");
            }
        }

        // ── 12: アピアランス効果残存 ──
        if ("appliedEffects" in item && item.appliedEffects.length > 0) {
            bucket(12, "アピアランス効果");
        }
        // ── 13: 不透明度 < 100% ──
        if ("opacity" in item && item.opacity < 100) {
            bucket(13, "不透明");
        }
        
        // ── 再帰：グループ／コンパウンドパス ──
        if (item.typename === "GroupItem") {
            processItems(item.pageItems);
        } else if (item.typename === "CompoundPathItem") {
            processItems(item.pathItems);
        }
    }
}

// 赤枠描画 -----------------------------------------------------------
function drawRedBox(bounds) {
    var x1=bounds[0], y1=bounds[1], x2=bounds[2], y2=bounds[3];
    var w=x2-x1, h=y1-y2;
    var rect = errorLayer.pathItems.rectangle(y1, x1, w, h);
    rect.stroked     = true;
    rect.strokeWidth = 0.5;
    rect.strokeColor = makeRGBColor(255, 0, 0);
    rect.filled      = false;
}

// RGBColor作成 ------------------------------------------------------
function makeRGBColor(r,g,b) {
    var c = new RGBColor();
    c.red   = r;
    c.green = g;
    c.blue  = b;
    return c;
}

// レポート書き出し --------------------------------------------------
// —— 新版レポート生成 (v1.3.7) —— 
function saveReport() {
  // 引数がなくても動くように、ここで値を取得します
  var sourceFile = app.activeDocument;  
  var results = {
    major: errorsMajorArray,   // ←ここは既存の「重大エラー配列名」に置き換え
    medium: errorsMediumArray, // ←「中エラー配列名」に
    minor: errorsMinorArray    // ←「軽微エラー配列名」に
  };

  var docName   = sourceFile.name;
  var now       = new Date();
  var timestamp = formatDate(now, "yyyy-MM-dd HH:mm:ss");
  var cMajor    = results.major.length;
  var cMedium   = results.medium.length;
  var cMinor    = results.minor.length;
  var total     = cMajor + cMedium + cMinor;

  var lines = [];
  lines.push("■ 入稿チェックレポート");
  lines.push("・ファイル名：" + docName);
  lines.push("・実行日時：" + timestamp);
  lines.push("・エラー総件数：" + total + "件（重大：" + cMajor + "件／中：" + cMedium + "件／軽微：" + cMinor + "件）");
  lines.push(repeatChar('━',70));
  lines.push("\n");

  function writeSection(title, arr, icon, riskText, actionText) {
    lines.push("■ " + title + "（" + arr.length + "件）");
    if (!arr.length) return;
    lines.push(icon + " [" + arr[0].typeId + "] " + arr[0].name);
    lines.push("  リスク：" + riskText);
    lines.push("  対処：" + actionText);
    lines.push("\n");
    for (var i = 0; i < arr.length; i++) {
      lines.push(repeatChar('-',70));
      lines.push("エラー (" + (i+1) + "/" + arr.length + ")");
      lines.push("  レイヤー：" + arr[i].layerPath);
      lines.push("  オブジェクト名：" + arr[i].objectName);
    }
    lines.push(repeatChar('━',70));
    lines.push("\n");
  }

  writeSection("重大", results.major,  "❗", 
               "フォントが置き換わり、文字化けや欠落が発生する", 
               "文字をアウトライン化する");
  writeSection("中",   results.medium, "⚠️", 
               "CMYK変換時に色味が変化し、仕上がりがイメージと異なる可能性がある", 
               "CMYKに変換する");
  writeSection("軽微", results.minor,  "ℹ️", 
               "不要なオブジェクトが残り、ファイル容量が増える", 
               "不要であれば削除する");

  var reportFile = new File(
    sourceFile.path + "/" + docName.replace(/\.ai$/i,"") + "_errreport.txt"
  );
  reportFile.encoding = "UTF-8";
  reportFile.open('w');
  reportFile.writeln(lines.join("\n"));
  reportFile.close();

  alert("入稿チェックレポートを出力しました：\n" + reportFile.fullName);
}

// 固定幅パッド ------------------------------------------------------
function pad(val, w) {
    var s = val.toString();
    while (s.length < w) s += " ";
    return s;
}

// 日付フォーマット：yyyy-MM-dd HH:mm:ss 形式に整形
function formatDate(dt, fmt) {
    function pad(n){ return (n < 10 ? '0' : '') + n; }
    return fmt.replace(/yyyy/, dt.getFullYear())
              .replace(/MM/,   pad(dt.getMonth() + 1))
              .replace(/dd/,   pad(dt.getDate()))
              .replace(/HH/,   pad(dt.getHours()))
              .replace(/mm/,   pad(dt.getMinutes()))
              .replace(/ss/,   pad(dt.getSeconds()));
}

// 指定文字を count 回繰り返す
function repeatChar(ch, count) {
    var s = "";
    for (var i = 0; i < count; i++) {
        s += ch;
    }
    return s;
}

})(); // <<< END OF IIFE >>>
