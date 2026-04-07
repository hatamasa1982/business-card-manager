

function checkMyModels() {
  const apiKey = GEMINI_API_KEY; 
  const url = "https://generativelanguage.googleapis.com/v1beta/models?key=" + apiKey;
  const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  console.log(res.getContentText());
}

function setupTrigger() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger("processNewCards").forSpreadsheet(sheet).onChange().create();
}

function processNewCards() {
  // グローバル定数のAPIキーを使用
  const apiKey = GEMINI_API_KEY;
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("名刺データ");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const imgIdx = headers.indexOf("名刺（表面）");
  const companyIdx = headers.indexOf("会社名");
  const nameIdx = headers.indexOf("氏名");
  
  for(let i = 1; i < data.length; i++) {
    const row = data[i];
    const imgPath = row[imgIdx];
    const company = row[companyIdx];
    
    if(imgPath && !company) {
      try {
        const sheetFolder = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents().next();
        const imagesFolderItr = sheetFolder.getFoldersByName("名刺データ_Images");
        if(!imagesFolderItr.hasNext()) continue;
        const imagesFolder = imagesFolderItr.next();
        
        const fileName = imgPath.split('/').pop();
        const filesItr = imagesFolder.getFilesByName(fileName);
        if(!filesItr.hasNext()) continue;
        const file = filesItr.next();
        
        const base64Data = Utilities.base64Encode(file.getBlob().getBytes());
        const mimeType = file.getMimeType();
        
        // 大容量無料枠を持つ最速のLiteモデル（gemini-flash-lite-latest）を使用
        const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-lite-latest:generateContent?key=" + apiKey.trim();
        const payload = {
          "contents": [{
            "parts": [
              {
                "text": "以下の日本の名刺画像から情報を抽出してJSON形式で出力してください。見つからない場合は空文字にしてください。キーは必ず以下を使用すること: company, department, name, phone, mobile, fax, email1, email2, zip, address。余計な説明や```jsonなどの表記は絶対に入れず、JSONの中身だけを返してください。"
              },
              {
                "inlineData": { // ★プロパティ名を修正
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
           "muteHttpExceptions": true // ★エラーの詳細を見るための設定
        };
        
        const response = UrlFetchApp.fetch(url, options);
        // エラーかどうかの判定
        if(response.getResponseCode() !== 200) {
           console.error("Gemini APIエラー: " + response.getContentText());
           continue; // エラーの場合はスキップ
        }
        
        const result = JSON.parse(response.getContentText());
        const jsonText = result.candidates[0].content.parts[0].text.trim().replace(/^```(?:json)?|```$/g, "").trim();
        const ocrData = JSON.parse(jsonText);
        
        const cleanCompany = (ocrData.company || '会社不明').replace(/[/\\?%*:|"<>]/g, '-');
        const cleanName = (ocrData.name || '氏名不明').replace(/[/\\?%*:|"<>]/g, '-');
        const newFileName = `${cleanCompany}_${cleanName}_表.jpg`;
        file.setName(newFileName);
        
        const newImgPath = "名刺データ_Images/" + newFileName;
        sheet.getRange(i+1, imgIdx+1).setValue(newImgPath);
        
        if(companyIdx >= 0) sheet.getRange(i+1, companyIdx+1).setValue(ocrData.company || "");
        if(nameIdx >= 0) sheet.getRange(i+1, nameIdx+1).setValue(ocrData.name || "");
        
        const fields = [
          {key: "department", header: "部署・役職"},
          {key: "phone", header: "代表電話"},
          {key: "mobile", header: "携帯電話"},
          {key: "fax", header: "ファックス"},
          {key: "email1", header: "メールアドレス1"},
          {key: "email2", header: "メールアドレス2"},
          {key: "zip", header: "郵便番号"},
          {key: "address", header: "住所"}
        ];
        
        fields.forEach(f => {
           let idx = headers.indexOf(f.header);
           if(idx >= 0) {
             sheet.getRange(i+1, idx+1).setValue(ocrData[f.key] || "");
           }
        });
        console.log("読み取り成功：" + newFileName);
      } catch(e) {
        console.error("エラーが発生しました: " + e.toString());
      }
    }
  }
}
