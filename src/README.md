## 使い方（Google Apps Script 側）
- スプレッドシートを開き、拡張機能 → Apps Script を開く
- GASプロジェクトに下記ファイルを作り、リポの `src/` 各ファイルの内容を貼り付け
  - menu.gs / csv_import.gs / bubble_upload_dev.gs / pivot_a.gs / pivot_b.gs / tab_tools.gs / time_trigger.gs
- 必要に応じてシート名やログシート名を環境に合わせて調整
- メニュー（onOpen）から各機能を実行可能
- 定期実行したい場合は、トリガーで `runUploadAndPivotPipeline_`（または `runCsvImportPipeline_`）を時間主導で設定

## 構成
src/
menu.gs # onOpenメニュー作成
csv_import.gs # CSV取り込みとユーザーへの反映
bubble_upload_dev.gs # Bubble側へのアップロード処理
pivot_A.gs # ピボットA作成
pivot_B.gs # ピボットB作成
tab_tools.gs # タブ生成/整形などの補助ツール
time_trigger.gs # パイプライン（定時実行エントリ）

## エントリポイント
- 手動：スプレッドシートのメニューから実行
- 定期：`runUploadAndPivotPipeline_` をトリガー登録（毎日 / 毎時など）
