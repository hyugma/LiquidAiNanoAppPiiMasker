# LiquidAiNanoAppPiiMasker

https://huggingface.co/LiquidAI/LFM2-350M-ENJP-MT のGGUF版モデルを利用したサンプル/デモ アプリケーションです。

ローカル LLM（LLamaSharp + GGUF）で PII（個人情報5種）候補を抽出し、Office 文書やテキストファイル内の該当箇所をマスクする Windows デスクトップアプリ（WPF）です。タスクトレイ常駐＋ドラッグ＆ドロップ UI により、ファイル／フォルダーを簡単に一括処理できます。

操作動画はこちら↓
https://www.youtube.com/watch?v=FzbfzWWb3WY

- 対応 OS: Windows 10/11（x64）
- 対応形式: .docx / .pptx / .xlsx / .txt / .md / .rtf
- モデル: GGUF 形式の日本語向け PII 抽出モデル（初回自動ダウンロード）
- 処理: すべてローカルで実行（モデルのダウンロード時のみネットワーク接続を使用）

注: 現状 UI は WPF のため Windows 専用です。macOS 向けに配布するには Avalonia などクロスプラットフォーム UI への移行、もしくは CLI 版の追加が必要です（後述の「将来計画」を参照）。

---

## 特長

- ドラッグ＆ドロップで簡単マスキング
  - 左クリックで小さなウィンドウを表示 → ファイル／フォルダーを DnD
  - 右クリックメニューから「Mask File…」「Mask Folder…」も選択可能
- Office Open XML 形式の直接編集
  - Word（.docx）／PowerPoint（.pptx）／Excel（.xlsx）に対し、テキストを上書きマスク
- テキスト系はブロックマスク
  - .txt / .md / .rtf は '█' ブロック文字による置換
- ローカル LLM 推論
  - インターネットに送信せずローカルで PII 抽出（初回のみモデルを自動取得）
- Auto-open 設定
  - マスク完了後に結果ファイルを自動で開く設定のオン／オフ切替
- ログ出力
  - 問題解析用ログを %LOCALAPPDATA%\LiquidAiNanoAppPiiMasker\ に保存

---

## 仕組み（アーキテクチャ概要）

- PII 抽出（Services/Core.cs, LlamaInference）
  - LLamaSharp（StatelessExecutor）でプロンプトを生成し、モデル出力から JSON を抽出・デシリアライズして候補を収集
  - 対象タグ: address, company_name, email_address, human_name, phone_number
- 文書テキストの抽出・置換（Services/Core.cs, OfficeUtils）
  - .docx: 本文中の Text 要素を列挙し、PII 候補を '█' ブロック文字列へ置換
  - .pptx: スライド中の A.Text を置換
  - .xlsx: セル文字列を抽出し、該当セルを "[MASKED]" に置換（SharedString／InlineString を考慮）
  - .txt / .md / .rtf: ファイル全体で該当文字列を '█' ブロックに置換
- アプリ UI（WPF + タスクトレイ）
  - App.xaml.cs: タスクトレイアイコン（WinForms NotifyIcon）とコンテキストメニュー
  - MainWindow.xaml/.cs: ドラッグ＆ドロップ受け入れウィンドウ
- モデル管理（App.xaml.cs）
  - 初回起動時、既定モデル（GGUF）を %LOCALAPPDATA%\LiquidAiNanoAppPiiMasker\models\ にダウンロード
  - 既定の URL（Hugging Face）とファイル名はコード内の定数で管理（`ModelRepoUrl`, `DefaultModelFileName`）

---

## 動作要件

- OS: Windows 10/11（x64）
- ランタイム: .NET Desktop Runtime 8.0（x64）
- CPU: AVX2 対応を推奨（LLamaSharp CPU バックエンドを利用）
- メモリ: 8 GB 以上推奨
- ストレージ: 500 MB 以上（モデルダウンロードにより追加が必要）
- ネットワーク: 初回モデルダウンロード時のみ必要（以降はオフライン可）

詳細な手順は docs/INSTALL_ja.md を参照してください。

---

## インストール（配布物からの利用）

配布 ZIP からのセットアップ手順は、以下の日本語ドキュメントにまとめています。

- docs/INSTALL_ja.md

要点:
- ZIP を任意のフォルダーに展開して exe を実行（ポータブル配布）
- 初回起動時にモデルの自動ダウンロードを実施
- SmartScreen の警告が出た場合は回避手順を参照
- 詳細なトラブルシューティングは上記ドキュメントを参照

---

## 使い方

- タスクトレイ常駐
  - 右クリックメニュー
    - Mask File…: ファイルを選択しマスク
    - Mask Folder…: フォルダー配下を再帰的に一括マスク
    - Auto-open file after masking: マスク完了後に自動オープンのオン／オフ
    - Show Debug Logs…: ログファイルを開く
    - Exit: アプリ終了
  - 左クリック
    - ドラッグ＆ドロップ用の小ウィンドウを表示
    - 対応拡張子: .docx / .pptx / .xlsx / .txt / .md / .rtf
- 出力
  - 入力ファイルと同一フォルダーに「xxxx_masked.拡張子」を生成
  - Auto-open がオンの場合は処理後に自動で開く

---

## モデル（GGUF）

- 既定のモデル
  - リポジトリ: https://huggingface.co/LiquidAI/LFM2-350M-PII-Extract-JP-GGUF
  - 既定ファイル名: `LFM2-350M-PII-Extract-JP-Q4_K_M.gguf`
- 配置場所
  - `%LOCALAPPDATA%\LiquidAiNanoAppPiiMasker\models\`
- 置き換え
  - 同名で差し替えるか、App.xaml.cs の `DefaultModelFileName`／`ModelRepoUrl` を変更してビルドし直すことで別モデルの使用が可能

注意:
- ログには推論の「生出力（raw）」が出力される場合があります（PII を含む可能性）。ログの取り扱いには十分注意してください。

---

## ビルド方法（開発者向け）

前提:
- .NET SDK 8.x（Windows）
- Visual Studio 2022 もしくは VS Code + C# 拡張

コマンド例（リポジトリ直下）:

```bat
dotnet restore "LiquidAiNanoAppPiiMasker.sln"
dotnet build   "LiquidAiNanoAppPiiMasker.sln" -c Release -t:Rebuild
```

成果物:
- フレームワーク依存（既定）
  - `LiquidAiNanoAppPiiMasker\bin\Release\net8.0-windows\LiquidAiNanoAppPiiMasker.exe`
- 自己完結（self-contained）配布を作る場合（サイズは大きくなります）:

```bat
dotnet publish "LiquidAiNanoAppPiiMasker/LiquidAiNanoAppPiiMasker.csproj" ^
  -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

出力:
- `LiquidAiNanoAppPiiMasker\bin\Release\net8.0-windows\win-x64\publish\`

---

## プロジェクト構成（主要ファイル）

- LiquidAiNanoAppPiiMasker/
  - App.xaml / App.xaml.cs … タスクトレイ常駐、メニュー、モデル管理、ログ
  - MainWindow.xaml / MainWindow.xaml.cs … DnD ウィンドウ、拡張子フィルタ
  - Services/Core.cs … PII 抽出、文書処理（OpenXML）
  - Assets/ … アイコンやロゴの候補（存在する場合は PII テキストオーバレイを合成）

---

## 制限事項・注意点

- 誤検出／漏れ
  - モデルベースの抽出のため、誤検知や見落としがあり得ます。重要文書は人手レビューを推奨
- 書式の変化
  - 置換処理により、レイアウトや書式（特に PowerPoint/Excel）が一部変わる可能性
- Excel のマスキング
  - 対応は主に文字列セルです。式／数値のみのセルや特殊な埋め込みは対象外となる場合あり
- 画像内文字・埋め込みオブジェクト
  - 画像内テキストやオブジェクトは対象外（OCR 等は未実装）
- PDF
  - ソースに PDF のユーティリティ例（コメントアウト）が含まれますが、現行ビルドでは PDF マスクを提供していません
- ログの取り扱い
  - 「モデルの生出力」がログへ記録されるため、ログが機微情報を含む可能性があります。保管／共有時は注意

---

## トラブルシューティング（抜粋）

- 起動エラー（0xc0000135 等）
  - .NET Desktop Runtime 8.0（x64）が未インストールの可能性
- ZIP から直接実行して失敗
  - 必ず「すべて展開」したフォルダー内の exe を実行
- モデルの自動ダウンロード失敗
  - ネットワークを確認。ログ（Show Debug Logs…）で詳細を参照
- 権限エラー
  - Program Files 等ではなくユーザーフォルダー配下へ展開

詳細は docs/INSTALL_ja.md を参照。

---

## ライセンス・サードパーティ

本アプリは以下の OSS を利用します（一部）:
- LLamaSharp / LLamaSharp.Backend.Cpu
- DocumentFormat.OpenXml
- Serilog / Serilog.Sinks.File
- Microsoft.Windows.Compatibility

各ライブラリのライセンスは公式リポジトリをご確認ください。本リポジトリのライセンスは `LICENSE` を参照（未設定の場合は追加予定）。

---

## 将来計画（macOS 対応の検討）

- 現状 WPF のため Windows 専用
- macOS 対応案
  - Avalonia UI: WPF から移行しやすい XAML ベースのクロスプラットフォーム UI（Win/macOS/Linux）
  - .NET MAUI: デスクトップ＋モバイルを含むが、Mac でのビルド／サインが必要
  - CLI 版: GUI なしで macOS バイナリを提供（最短で提供可能）
- 配布上の注意
  - macOS の公開配布には codesign／notarization（Apple Developer アカウント／Mac 環境）が必要

---

## コントリビュート

- Issue／PR 歓迎
- 方針提案（モデル更新、抽出精度の改善、PDF／画像の対応、クロスプラットフォーム化等）も歓迎

---

## 免責

本ソフトウェアの利用により生じたいかなる損害についても作者は責任を負いません。重要文書のマスキング結果は必ず人手で確認し、組織の情報セキュリティポリシーに従って取り扱ってください。
