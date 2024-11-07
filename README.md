# question_spreadsheet

1. Google Driveのフォルダにシートを追加したいSpreadsheetファイルを置く。

2. 追加したいシートを開く

3. 拡張機能 → Apps Scriptを選択

4. 開いたページに下のコードをコピーして貼り付けし、保存する

```
function sheetCopyAnotherFile() {
  try {
    SpreadsheetApp.getActiveSpreadsheet()
    var folder_name = Browser.inputBox("Google DriveのフォルダURLを入れてください");
    var folder_id = folder_name.replace("https://drive.google.com/drive/u/0/folders/", "");

    folder = DriveApp.getFolderById(folder_id);
    files = folder.getFiles();
        
    let copysheet = Browser.inputBox("フォルダ内のファイル全てにコピーするシート名を入力してください。\\n例：シート1など。");
    if(copysheet == "cancel"){
      Browser.msgBox("実行をキャンセルします。");
      return;
    }

    while (files.hasNext()) {
      var buff = files.next();
      ss = SpreadsheetApp.getActive();
      sheet = ss.getSheetByName(copysheet);
      destSpreadsheet = SpreadsheetApp.openByUrl(buff.getUrl());
      newCopySheet = sheet.copyTo(destSpreadsheet);
      newCopySheet.setName(copysheet + '_copy');
    }

  } catch (e) {
    Browser.msgBox(e+"\\n実行をキャンセルします。");
    return false;
  }
}

function deleteAllSheet() {
  try {
    SpreadsheetApp.getActiveSpreadsheet()
    var folder_name = Browser.inputBox("Google DriveのフォルダURLを入れてください");
    var folder_id = folder_name.replace("https://drive.google.com/drive/u/0/folders/", "");

    folder = DriveApp.getFolderById(folder_id);
    files = folder.getFiles();
        
    let deletesheet = Browser.inputBox("削除するシート名を入力してください。\\n例：シート1_copyなど。");
    if(deletesheet == "cancel"){
      Browser.msgBox("実行をキャンセルします。");
      return;
    }

    while (files.hasNext()) {
      var buff = files.next();
      ss = SpreadsheetApp.getActive();
      destSpreadsheet = SpreadsheetApp.openByUrl(buff.getUrl());
      let delSheet = destSpreadsheet.getSheetByName(deletesheet);
      destSpreadsheet.deleteSheet(delSheet);
    }

  } catch (e) {
    Browser.msgBox(e+"\\n実行をキャンセルします。");
    return false;
  }
}

//スプレッドシートを開いた際に呼び出される関数で、メニューを追加
function onOpen() {
 SpreadsheetApp
   .getActiveSpreadsheet()
   .addMenu("シート一括追加", [
     { name: "Driveフォルダ内のファイルにシート追加", functionName: "sheetCopyAnotherFile" },
     { name: "Driveフォルダ内のファイルからシート削除", functionName: "deleteAllSheet" },
   ]);
}
```

5. 上のメニューの"実行" "デバッグ" の右の所(最初はsheetCopyAnotherFileになってるはず)をonOpenに変更。

6. 実行ボタンを押す。(1回目は権限の許可が必要。"このアプリはGoogleで確認されていません"と出たら左下の"詳細"から(安全ではないページ)に移動をクリック)。

7. 元のスプレッドシートに「シート一括追加」タブが増えている。
