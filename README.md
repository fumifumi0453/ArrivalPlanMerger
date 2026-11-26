# ArrivalPlanMerger

受信したExcelから入荷予定ファイルを自動更新するツールです。

1. Excel を CSV にする
2. 必要な列だけに加工する
3. オーダー番号でファイルを分ける
4. オーダー別フォルダに移動する
5. 各フォルダの入荷予定ファイルにマージする（キー列が同じ行は上書き、無ければ追加）

## CSharp/ (.NET 8 コンソールアプリ)

- Excel 不要（ClosedXML で直接読み込み）
- 設定は `settings.json`
- 発行（フォルダー）すると、ランタイム不要の単一 exe（win-x64）ができます
- 実行先には `ArrivalPlanMerger.exe` と `settings.json` をコピー

## PowerShell/ (Windows PowerShell 5.1)

- Excel 本体が必要（COM で CSV 保存）
- 設定はスクリプト冒頭の `$Config`
- `実行.bat` をダブルクリック、またはタスクスケジューラで `実行.bat /nopause`

## 共通

- CSV は Shift_JIS / CRLF
- マージ前に `入荷予定.csv.bak` を作成
- ログは `logs\yyyyMMdd.log`
