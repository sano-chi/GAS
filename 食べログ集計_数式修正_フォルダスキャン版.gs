/**
 * 食べログ集計シート 数式修正スクリプト【フォルダ自動スキャン版】
 *
 * 対象シート：食べログ集計
 * F74 → =VALUE(REGEXREPLACE(TO_TEXT(G71),"[¥￥, 　]",""))
 * I74 → =VALUE(REGEXREPLACE(TO_TEXT(F72),"[¥￥, 　]",""))
 *
 * 実行方法：
 *   1. https://script.google.com を開く
 *   2. 「新しいプロジェクト」を作成
 *   3. このファイルの内容をコード.gsに貼り付けて保存
 *   4. 「fixFormulas」関数を選択して実行
 *   5. 初回実行時はGoogleアカウントの権限許可を求められるので許可する
 *   6. 実行後、「表示」→「ログ」で処理結果を確認できます
 *
 * 応用する場合は下記「設定エリア」の3項目を変更するだけでOKです。
 */

// ============================================================
// ★ 設定エリア（ここだけ変更すればOK）
// ============================================================

// 対象フォルダのID（GoogleドライブURLの folders/ 以降の文字列）
var FOLDER_ID = "1ZekCvwebZbioTV4sMLN2KfX0QxQnOxO7";

// 対象シート名
var TARGET_SHEET = "食べログ集計";

// 書き換える数式（セル: 数式 の形式で複数指定可能）
var FORMULAS = {
  "F74": '=VALUE(REGEXREPLACE(TO_TEXT(G71),"[¥￥, \u3000]",""))',
  "I74": '=VALUE(REGEXREPLACE(TO_TEXT(F72),"[¥￥, \u3000]",""))'
};

// サブフォルダも再帰的にスキャンするか（true=する / false=しない）
var SCAN_SUBFOLDERS = true;

// ============================================================


/**
 * メイン処理：フォルダをスキャンして数式を一括修正する
 */
function fixFormulas() {
  var successCount = 0;
  var skipCount    = 0;
  var errorCount   = 0;
  var logs         = [];

  // フォルダ内のスプレッドシートを全件取得
  Logger.log("フォルダをスキャン中...");
  var files = getSpreadsheetFilesInFolder(FOLDER_ID, SCAN_SUBFOLDERS);
  Logger.log("スプレッドシート検出数: " + files.length + " 件");

  for (var i = 0; i < files.length; i++) {
    var file = files[i];
    try {
      var ss    = SpreadsheetApp.open(file);
      var sheet = ss.getSheetByName(TARGET_SHEET);

      // 対象シートが存在しない場合はスキップ
      if (!sheet) {
        logs.push("[SKIP] シートなし: " + ss.getName());
        skipCount++;
        continue;
      }

      // 数式を書き換え
      for (var cell in FORMULAS) {
        sheet.getRange(cell).setFormula(FORMULAS[cell]);
      }

      logs.push("[OK] " + ss.getName());
      successCount++;

    } catch (e) {
      logs.push("[ERROR] " + file.getName() + " → " + e.message);
      errorCount++;
    }

    // API制限対策：10件ごとに少し待機
    if ((i + 1) % 10 === 0) {
      Utilities.sleep(1000);
    }
  }

  // 結果サマリーをログ出力
  Logger.log("==============================");
  Logger.log("処理完了");
  Logger.log("  成功: "                    + successCount + " 件");
  Logger.log("  スキップ（シートなし）: "  + skipCount    + " 件");
  Logger.log("  エラー: "                  + errorCount   + " 件");
  Logger.log("------------------------------");
  for (var j = 0; j < logs.length; j++) {
    Logger.log(logs[j]);
  }
  Logger.log("==============================");
}


/**
 * 指定フォルダ（およびサブフォルダ）内のスプレッドシートを全件取得する
 *
 * @param {string}  folderId       - 対象フォルダのID
 * @param {boolean} includeSubFolders - サブフォルダを再帰スキャンするか
 * @return {File[]} スプレッドシートのFileオブジェクト配列
 */
function getSpreadsheetFilesInFolder(folderId, includeSubFolders) {
  var result = [];
  var folder = DriveApp.getFolderById(folderId);

  // 直下のスプレッドシートを取得
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) {
    result.push(files.next());
  }

  // サブフォルダを再帰的にスキャン
  if (includeSubFolders) {
    var subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
      var subFolder = subFolders.next();
      var subFiles  = getSpreadsheetFilesInFolder(subFolder.getId(), true);
      result = result.concat(subFiles);
    }
  }

  return result;
}
