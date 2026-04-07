// ==========================================
// 全スクリプト共通のグローバル設定 (000_global.gs)
// ==========================================

// 1. スクリプトプロパティからAPIキーを取得
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");

// 2. スキャン画像を保存しておくフォルダのID（ver.2・ver.3の手動読み込み時に使用）
// 下記にご自身のフォルダID（アルファベットと数字の羅列）を貼り付けてください
const TARGET_FOLDER_ID = "YOUR_FOLDER_ID_HERE";
