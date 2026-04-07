

// 1. スプレッドシートを開いたときに上部に専用メニューを追加する
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🌟名刺管理')
    .addItem('📥 新着スキャン画像の読み込み', 'processFolderScans')
    .addToUi();
}

function processFolderScans() {
  const ui = SpreadsheetApp.getUi();
  
  // APIキーとフォルダIDは 000_global.gs の共通設定を使用
  const apiKey = GEMINI_API_KEY; 
  const targetFolderId = TARGET_FOLDER_ID;
  // ★ =============================================== ★
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("名刺データ");
  if (!sheet) {
    ui.alert("エラー", "「名刺データ」という名前のシートが見つかりません。", ui.ButtonSet.OK);
    return;
  }
  
  // 指定のフォルダを取得
  let folder;
  try {
    folder = DriveApp.getFolderById(targetFolderId);
  } catch(e) {
    ui.alert("エラー", "フォルダIDが正しくありません。\n設定を確認してください。", ui.ButtonSet.OK);
    return;
  }
  
  const files = folder.getFiles();
  let processedCount = 0;
  
  // 見出し行のデータ（列のインデックス）を取得
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // ここからフォルダ内の全画像を1つずつチェック
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    
    // ▼ 今回の目玉ポイント：二重処理の防止 ▼
    // 画像の名前の最後に「_表.jpg」や「_表.jpeg」などが付いている場合は「処理済」とみなして無視する
    if (fileName.match(/_表\.[a-zA-Z]+$/)) {
      continue;
    }
    
    // 画像ファイル以外（PDFなど）は無視する
    const mimeType = file.getMimeType();
    if (!mimeType.startsWith('image/')) {
      continue;
    }

    try {
      // 1. 画像データをBase64に変換
      const base64Data = Utilities.base64Encode(file.getBlob().getBytes());
      
      // 2. Gemini APIの呼び出し（安定版の gemini-2.5-flash を使用）
      const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey.trim();
      const payload = {
        "contents": [{
          "parts": [
            {
              "text": "以下の日本の名刺画像から情報を抽出してJSON形式で出力してください。見つからない場合は空文字にしてください。キーは必ず以下を使用すること: company, department, name, phone, mobile, fax, email1, email2, zip, address。余計な説明や```jsonなどの表記は絶対に入れず、JSONの中身だけを返してください。"
            },
            {
              "inlineData": { 
                "mimeType": mimeType,
                "data": base64Data
              }
            }
          ]
        }]
      };
      
      const options = {
         "method": "post",
         "contentType": "application/json",
         "payload": JSON.stringify(payload),
         "muteHttpExceptions": true 
      };
      
      const response = UrlFetchApp.fetch(url, options);
      if(response.getResponseCode() !== 200) {
         console.error("Gemini APIエラー (" + fileName + "): " + response.getContentText());
         continue; 
      }
      
      const result = JSON.parse(response.getContentText());
      const jsonText = result.candidates[0].content.parts[0].text.trim().replace(/^```(?:json)?|```$/g, "").trim();
      const ocrData = JSON.parse(jsonText);
      
      // 3. ファイル名の変更（リネーム）
      const cleanCompany = (ocrData.company || '会社不明').replace(/[/\\?%*:|"<>]/g, '-');
      const cleanName = (ocrData.name || '氏名不明').replace(/[/\\?%*:|"<>]/g, '-');
      // 元の拡張子を保持する（.jpg, .png など）
      const extension = fileName.split('.').pop();
      const newFileName = `${cleanCompany}_${cleanName}_表.${extension}`;
      file.setName(newFileName);
      
      // 4. 新しい行としてスプレッドシートに追加するデータの準備
      const newRowData = new Array(headers.length).fill("");
      
      // 項目とヘッダーの紐付け
      const fieldMapping = {
        "名刺（表面）": file.getUrl(), // AppSheetを使わないならURLを直接入れるのが便利
        "会社名": ocrData.company || "",
        "氏名": ocrData.name || "",
        "部署・役職": ocrData.department || "",
        "代表電話": ocrData.phone || "",
        "携帯電話": ocrData.mobile || "",
        "ファックス": ocrData.fax || "",
        "メールアドレス1": ocrData.email1 || "",
        "メールアドレス2": ocrData.email2 || "",
        "郵便番号": ocrData.zip || "",
        "住所": ocrData.address || ""
      };
      
      // ユニークIDが必要なら自動生成
      if (headers.indexOf("ID") >= 0) {
        fieldMapping["ID"] = Utilities.getUuid();
      }
      
      // 設定した値をヘッダーの列位置に流し込み
      for (const [headerName, value] of Object.entries(fieldMapping)) {
        const idx = headers.indexOf(headerName);
        if (idx !== -1) {
          newRowData[idx] = value;
        }
      }
      
      // スプレッドシートの一番下に行ごと追加
      sheet.appendRow(newRowData);
      processedCount++;
      
      // 連続でAPIを叩きすぎないように、1回ごとに5秒（5000ミリ秒）のお休みを入れる（スピード違反対策強化）
      Utilities.sleep(5000);
      
    } catch(e) {
      console.error("エラー (" + fileName + "): " + e.toString());
    }
  }
  
  if (processedCount > 0) {
    ui.alert("処理完了", `${processedCount} 件の新しい名刺画像を読み込み、スプレッドシートに追加しました！`, ui.ButtonSet.OK);
  } else {
    ui.alert("お知らせ", "新しく読み込む画像は見つかりませんでした。\n（名前が「_表」で終わっていない画像が対象です）", ui.ButtonSet.OK);
  }
}
