/**
 * Google Apps Script - CSV Proxy for Hotel Review System
 *
 * デプロイ手順:
 * 1. https://script.google.com で新しいプロジェクトを作成
 * 2. このコードを貼り付け
 * 3. 「デプロイ」→「新しいデプロイ」
 * 4. 種類: 「ウェブアプリ」
 * 5. 実行ユーザー: 「自分」
 * 6. アクセスできるユーザー: 「全員」
 * 7. デプロイ → URLをコピー
 *
 * 使用方法:
 *   {DEPLOY_URL}?id={SPREADSHEET_ID}&gid={GID}
 *   {DEPLOY_URL}?key={HOTEL_KEY}
 *
 * 例:
 *   https://script.google.com/macros/s/xxxx/exec?key=daiwa_osaki
 *   https://script.google.com/macros/s/xxxx/exec?id=1IIHEn4nAIy9UXzrYptU-RQIiTbKkV0G_CaABF7znVrY&gid=0
 */

var HOTEL_MAP = {
  "daiwa_osaki":              { id: "1IIHEn4nAIy9UXzrYptU-RQIiTbKkV0G_CaABF7znVrY", gid: "0" },
  "chisan":                   { id: "1IWigsWTzbRG-juWtIlg4ZchiuWqRhJFpPPczXdQxG6Y", gid: "0" },
  "hearton":                  { id: "1A25mmVRYSnG3ZB8oa0oZVp-vCP2xMkwX-Zqdkk4BIzI", gid: "0" },
  "keyakigate":               { id: "1srchDxFyv7TJ3IEZXJ19miH04p3jRug5nVtA3BLertQ", gid: "605247000" },
  "richmond_mejiro":          { id: "1XWU6925CpT3GMMonAqy4UENKM11gWloUkJIsgGImUts", gid: "0" },
  "keisei_kinshicho":         { id: "1jUS_HwTfowG1xIHFtwJbCL5dTj7FrhvUe6d32AevZ2g", gid: "0" },
  "daiichi_ikebukuro":        { id: "1X2GgFKxTOs7CuJSlPYrpzigraSnWcKh6cMJLfsXhWlU", gid: "0" },
  "comfort_roppongi":         { id: "1Jtm0rXTigY2OVManNjx1qQ6G9EKQEXuPs_T1BdlOvls", gid: "0" },
  "comfort_suites_tokyobay":  { id: "1zCFAmzRqvSDbjwvK7qI4cYBHlrmBifTPm0Y-g0rruyE", gid: "0" },
  "comfort_era_higashikanda": { id: "1H9jmOVQR4UdEQ5hsxZ2Xz44BT72RJDwNa6BOKFhXxRg", gid: "0" },
  "comfort_yokohama_kannai":  { id: "1rnQOsyUXuSzBKdqPN_ey_4Iw5VtYWTgSR5Z4nh-1zd4", gid: "0" },
  "comfort_narita":           { id: "1lQ3FRDuE75dkByQRFd0i0F2xcHnl-3-UAOJwhIt3jAU", gid: "0" },
  "apa_kamata":               { id: "16xuhAdNzdeyAKu-LhU8ATgR8_kZ1JXfa9lT51tAB1Nw", gid: "0" },
  "apa_sagamihara":           { id: "1E2ZQJyE6pOJ3jr6GyB56KcYnVVq54m6dO_6h_SQy39A", gid: "0" },
  "court_shinyokohama":       { id: "1Qm5lPPc8m7yutyIH3Pf03YUnF2KpnWjn0SecMzq0CjY", gid: "0" },
  "comment_yokohama":         { id: "1cVH7khdgh8bDN-wtAw2KVakJqHILo58VOBu0SKmBFrU", gid: "0" },
  "kawasaki_nikko":           { id: "1aQ2MaKJmOz7eT53oqszCDO9Fa3UEbfhFSgXfVmVpO9A", gid: "0" },
  "henn_na_haneda":           { id: "18DkZLJ8UDQ2-4MBrh7B4y28tHaYnoWIQqEoFkvFDNKg", gid: "2026949334" },
  "comfort_hakata":           { id: "1_7xoyIiq1llfO0I2328ZQlB6sD0lMsnpRMp1rMGNPcg", gid: "0" }
};

function doGet(e) {
  var spreadsheetId, gid;

  // key パラメータでホテルを指定
  if (e.parameter.key && HOTEL_MAP[e.parameter.key]) {
    var hotel = HOTEL_MAP[e.parameter.key];
    spreadsheetId = hotel.id;
    gid = hotel.gid;
  } else if (e.parameter.id) {
    spreadsheetId = e.parameter.id;
    gid = e.parameter.gid || "0";
  } else {
    // キー一覧を返す
    return ContentService.createTextOutput(JSON.stringify(Object.keys(HOTEL_MAP)))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = null;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId().toString() === gid) {
        sheet = sheets[i];
        break;
      }
    }
    if (!sheet) sheet = sheets[0];

    var data = sheet.getDataRange().getValues();
    var csv = data.map(function(row) {
      return row.map(function(cell) {
        if (cell instanceof Date) {
          var y = cell.getFullYear();
          var m = ("0" + (cell.getMonth() + 1)).slice(-2);
          var d = ("0" + cell.getDate()).slice(-2);
          cell = y + "-" + m + "-" + d;
        }
        var str = String(cell);
        if (str.indexOf(",") >= 0 || str.indexOf('"') >= 0 || str.indexOf("\n") >= 0) {
          return '"' + str.replace(/"/g, '""') + '"';
        }
        return str;
      }).join(",");
    }).join("\n");

    return ContentService.createTextOutput(csv)
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput("ERROR: " + err.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}
