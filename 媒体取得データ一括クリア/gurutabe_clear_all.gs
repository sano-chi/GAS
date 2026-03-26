// ============================================================
// 媒体集計フォルダ 一括データ消去スクリプト
//
// 【使い方】
//   1. Google Apps Script (script.google.com) でスタンドアロンプロジェクトを作成
//   2. このコードを貼り付ける
//   3. CONFIG 内の HOT_FOLDER_ID に「媒体集計フォルダ」のIDを入力
//   4. clearAllSheets() を実行
//   5. スプレッドシートが600件ある場合は自動的にバッチ処理されます
//      （5分ごとに一時停止→再実行で続きから処理）
//
// 【削除範囲の変更方法】
//   CONFIG の TARGETS 内の ranges を書き換えるだけでOKです。
// ============================================================


// ============================================================
// ★ 設定（ここだけ変更すれば動作を調整できます）
// ============================================================
const CONFIG = {

  // 媒体集計フォルダのID
  // DriveでフォルダURLを開いたとき末尾の文字列
  // 例: https://drive.google.com/drive/folders/【ここ】
  HOT_FOLDER_ID: '1ZekCvwebZbioTV4sMLN2KfX0QxQnOxO7',

  // 処理対象シート名と削除範囲（書式は残し、値のみ消去）
  TARGETS: [
    {
      name: "ぐるなび集計",
      ranges: ["B3:N16", "B21:N34", "B39:N52", "B56:N69"]
    },
    {
      name: "食べログ集計",
      ranges: ["B2:O15", "B19:O32", "B38:O50", "B55:O67"]
    },
    {
      name: "ホットペッパー集計①",
      ranges: ["A17:L500", "P17:Z500"]
    }
  ]
};


// ============================================================
// メイン処理（ここを実行してください）
// ============================================================
function clearAllSheets() {
  const startTime = new Date();
  const props = PropertiesService.getScriptProperties();

  // 未処理IDリストを取得（前回の続きがあればそこから、なければ初期化）
  let pendingIds = JSON.parse(props.getProperty('pendingIds') || 'null');
  let processedCount = parseInt(props.getProperty('processedCount') || '0');
  let errorLog = JSON.parse(props.getProperty('errorLog') || '[]');

  if (!pendingIds) {
    // ── 初回実行：全スプレッドシートIDを収集 ──
    Logger.log('📂 スプレッドシートを収集中...');
    pendingIds = collectSpreadsheetIds(CONFIG.HOT_FOLDER_ID);
    const total = pendingIds.length;
    Logger.log(`✅ 合計 ${total} 件のスプレッドシートを検出しました`);
    props.setProperty('totalCount', total.toString());
    processedCount = 0;
    errorLog = [];
  }

  const totalCount = parseInt(props.getProperty('totalCount') || pendingIds.length.toString());
  const TIME_LIMIT_MS = 5 * 60 * 1000; // 5分（GASの実行制限6分に対して余裕を持たせる）

  // ── バッチ処理ループ ──
  while (pendingIds.length > 0) {

    // 時間制限チェック
    if (new Date() - startTime > TIME_LIMIT_MS) {
      props.setProperty('pendingIds', JSON.stringify(pendingIds));
      props.setProperty('processedCount', processedCount.toString());
      props.setProperty('errorLog', JSON.stringify(errorLog));
      Logger.log(`\n⏱ 時間制限（5分）に達しました。`);
      Logger.log(`   処理済み: ${processedCount} / ${totalCount} 件`);
      Logger.log(`   残り: ${pendingIds.length} 件`);
      Logger.log(`➡ 「clearAllSheets」をもう一度実行すると続きから処理します。`);
      return;
    }

    const id = pendingIds.shift();
    try {
      const ss = SpreadsheetApp.openById(id);
      CONFIG.TARGETS.forEach(target => {
        const sheet = ss.getSheetByName(target.name);
        if (sheet) {
          target.ranges.forEach(rangeStr => {
            sheet.getRange(rangeStr).clearContent(); // 値のみ消去（書式は維持）
          });
        }
      });
      processedCount++;
      Logger.log(`✅ [${processedCount}/${totalCount}] ${ss.getName()}`);

    } catch (e) {
      errorLog.push({ id: id, error: e.message });
      processedCount++;
      Logger.log(`❌ [${processedCount}/${totalCount}] エラー: ${id} → ${e.message}`);
    }
  }

  // ── 全件処理完了 ──
  props.deleteProperty('pendingIds');
  props.deleteProperty('processedCount');
  props.deleteProperty('totalCount');
  props.deleteProperty('errorLog');

  Logger.log(`\n✨ 全件完了！ ${processedCount} 件を処理しました。`);

  if (errorLog.length > 0) {
    Logger.log(`\n⚠️ エラーが ${errorLog.length} 件ありました：`);
    errorLog.forEach(e => Logger.log(`   - ID: ${e.id}  エラー: ${e.error}`));
  } else {
    Logger.log('エラーはありませんでした。');
  }
}


// ============================================================
// 進捗リセット（最初からやり直したい場合に実行）
// ============================================================
function resetProgress() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  Logger.log('🔄 リセット完了。次に clearAllSheets() を実行すると最初から処理します。');
}


// ============================================================
// テスト用：特定のスプレッドシート1件だけ処理して動作確認
// ============================================================
function testSingleSpreadsheet() {
  const TEST_SPREADSHEET_ID = 'テストしたいスプレッドシートのIDを入力';
  try {
    const ss = SpreadsheetApp.openById(TEST_SPREADSHEET_ID);
    CONFIG.TARGETS.forEach(target => {
      const sheet = ss.getSheetByName(target.name);
      if (sheet) {
        target.ranges.forEach(rangeStr => {
          sheet.getRange(rangeStr).clearContent();
        });
        Logger.log(`✅ シート「${target.name}」の消去完了`);
      } else {
        Logger.log(`⚠️ シート「${target.name}」が見つかりません`);
      }
    });
    Logger.log('テスト完了');
  } catch (e) {
    Logger.log(`❌ エラー: ${e.message}`);
  }
}


// ============================================================
// ユーティリティ：フォルダ内の全スプレッドシートIDを再帰的に収集
// ============================================================
function collectSpreadsheetIds(folderId) {
  const ids = [];
  const folder = DriveApp.getFolderById(folderId);
  collectFromFolder(folder, ids);
  return ids;
}

function collectFromFolder(folder, ids) {
  // このフォルダ内のスプレッドシートを収集
  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) {
    ids.push(files.next().getId());
  }
  // サブフォルダを再帰的に処理（企業フォルダ → 店舗スプレッドシート）
  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    collectFromFolder(subFolders.next(), ids);
  }
}
