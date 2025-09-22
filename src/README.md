# Bubble → GAS 自動化（CSV取込／Bubbleアップロード／Pivot A/B）
> 応募提出用のミニ実務プロジェクト。**APIキー等の秘密情報は一切含みません。**

## 概要
- Bubble の Data API から／へデータ連携し、Google スプレッドシート上で整形・集計（Pivot）までを自動化します。
- コードは **Apps Script (GAS)** で構成。

## リポ構成
src/
- menu.gs # onOpen メニュー作成（各機能を呼び出す）
- csv_import.gs # CSV取り込みとユーザーシート反映（認証必須）
- bubble_upload_dev.gs # Bubble へのアップロード（認証必須 / 開発版エンドポイント推奨）
- pivot_A.gs # ピボット A の作成
- pivot_B.gs # ピボット B の作成
- tab_tools.gs # タブ生成・整形などの補助
- time_trigger.gs # パイプライン（定時実行のエントリ）
  
## セキュリティ方針
- このリポジトリは**コードのみ**公開します。**APIキー／トークン／認証付きURL／実データ**は一切含みません。
- **鍵は共有しません。** 動作確認する方は、各自のテスト用ワークスペースと **ご自身のAPIキー** をご用意ください。  
  - Bubble を用いる場合、**開発版エンドポイント `/version-test/`** とテスト用データベースの使用を推奨します。
  - 不要になったキーは**即時無効化（Revoke）**可能な体制で運用してください。

> 画面共有によるデモや動画での動作提示も可能です（秘密情報は非公開のまま実施できます）。

## 前提・要件
- Google アカウント／Google スプレッドシート
- Bubble（評価者側のテスト用ワークスペースがあるとベター）
- ネットワークから Bubble Data API に到達できること

## セットアップ（GAS 側）
1. **スプレッドシートを用意**（例：`【プロジェクト名】`）。  
2. **拡張機能 → Apps Script** を開き、`src/` 内の各ファイルを **同名** で作成し、内容を貼り付け。  
   - `menu.gs` / `csv_import.gs` / `bubble_upload_dev.gs` / `pivot_A.gs` / `pivot_B.gs` / `tab_tools.gs` / `time_trigger.gs`
3. **スクリプトのプロパティ** を設定（GAS エディタ右上の歯車 → プロジェクトのプロパティ → スクリプトのプロパティ）。

### スクリプトプロパティ
- | KEY | 例 | 説明 |
- | `BUBBLE_TOKEN` | a1b2c3d4exxxxxxxxxxxxxxx7y8z9 | Bubble の Data API トークン |

## 使い方
### 手動実行（メニュー）
- シートを開き直すと `onOpen` によりメニューが表示されます。  
  - **CSV Import**：`csv_import.gs` の処理を実行（外部CSVなら認証必須）  
  - **Bubble Upload**：`bubble_upload_dev.gs` の処理を実行（認証必須）  
  - **Pivot_user / Pivot_面談報告**：ピボット作成（認証不要）
  - **TimeTrigger**：フル実行の時刻設定スケジューラ（認証不要）  
  - **シートタブ整理**：タブ生成・整形など（認証不要）
  - **シート削除**：不要シート削除（認証不要）

### 定期実行（TimeTrigger）
- Apps Script の **トリガー** で時間主導を設定します。代表例：  
  - `runUploadAndPivotPipeline_` … **外部連携を含む**フルパイプライン（**トークン必須**）  
  - `runCsvImportPipeline_` … CSV→Pivot 等のパイプライン（外部CSVを取得する構成ならトークン必須）

## 認証の要否（早見表）
| 処理 | ファイル | 外部アクセス | トークン |
| CSV 取り込み | `csv_import.gs` | **あり** | **必要** |
| Bubble へアップロード | `bubble_upload_dev.gs` | **あり** | **必要** |
| Pivot A/B | `pivot_A.gs` / `pivot_B.gs` | なし | 不要 |
| まとめ実行（フル） | `time_trigger.gs` | **あり** | **必要** |

 **鍵が無い場合に実行できるのは**：Pivot／TabTools など **外部アクセスの無い処理のみ**。
