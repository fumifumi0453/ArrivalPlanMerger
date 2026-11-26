<#
.SYNOPSIS
  入荷予定ファイル自動更新スクリプト

  1. Excel を CSV にする
  2. 必要な列だけに加工する
  3. オーダー番号でファイルを分ける
  4. オーダー別フォルダに移動する
  5. 各フォルダの入荷予定ファイルに新しい内容をマージする
     （キー列が同じ行があれば上書き、無ければ追加）

  動作環境: Windows PowerShell 5.1 / Excel がインストールされていること
#>

# =====================================================================
# 設定（ここを実際の環境に合わせて書き換えてください）
# =====================================================================
$Config = @{
    # 処理対象の Excel を置くフォルダ
    InputFolder      = 'C:\Work\入荷\受信'
    # 処理が終わった Excel の移動先（'' なら移動しない）
    ProcessedFolder  = 'C:\Work\入荷\受信\処理済'
    # 読み込むシート名（'' なら先頭シート）
    SheetName        = ''
    # 見出し行の行番号（1始まり。見出しより上の行は捨てる）
    HeaderRow        = 1

    # CSV に残す列（この順番で出力）
    OutputColumns    = @('オーダー番号', '品番', '品名', '数量', '入荷予定日')
    # ファイル分割に使う列
    OrderNoColumn    = 'オーダー番号'

    # オーダー別フォルダの親フォルダ（この下に「オーダー番号」名のフォルダ）
    OrderRootFolder  = 'C:\Work\入荷\オーダー'
    # オーダー別フォルダが無いとき作るか（$false ならスキップして警告）
    CreateOrderFolderIfMissing = $true
    # 分割CSVを各フォルダに保存するか（履歴として残す）
    SaveSplitFile    = $true

    # 各オーダーフォルダ内の入荷予定ファイル名
    PlanFileName     = '入荷予定.csv'
    # 同じ行とみなす列（@() なら全列一致で判定）
    KeyColumns       = @('オーダー番号', '品番')

    # マージ前に入荷予定ファイルの .bak を作るか
    BackupPlanFile   = $true
}
# =====================================================================

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# CSV の文字コード（Shift_JIS）。Windows PowerShell 5.1 の 'Default' は日本語環境で Shift_JIS
$CsvEncoding = 'Default'

$LogFolder = Join-Path $PSScriptRoot 'logs'
New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
$LogFile = Join-Path $LogFolder ((Get-Date -Format 'yyyyMMdd') + '.log')

function Write-Log([string]$Message) {
    $line = '{0} {1}' -f (Get-Date -Format 'yyyy/MM/dd HH:mm:ss'), $Message
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

function Get-SafeName([string]$Name) {
    foreach ($c in [IO.Path]::GetInvalidFileNameChars()) { $Name = $Name.Replace($c, '_') }
    return $Name
}

# ---------------------------------------------------------------------
# ① Excel → CSV（Excel 本体で「名前を付けて保存」する）
# ---------------------------------------------------------------------
function Convert-ExcelToCsv([string]$ExcelPath, [string]$CsvPath) {
    $excel = $null; $wb = $null; $ws = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        # 読み取り専用で開く（元ファイルは変更しない）
        $wb = $excel.Workbooks.Open($ExcelPath, 0, $true)
        if ($Config.SheetName) { $ws = $wb.Worksheets.Item($Config.SheetName) }
        else                   { $ws = $wb.Worksheets.Item(1) }
        $ws.Activate()

        # 見出しより上の行を削除（メモリ上だけ）
        if ($Config.HeaderRow -gt 1) {
            $ws.Rows("1:$($Config.HeaderRow - 1)").Delete() | Out-Null
        }

        # 6 = xlCSV（日本語環境では Shift_JIS、表示されている書式のまま出力される）
        $wb.SaveAs($CsvPath, 6)
    }
    finally {
        if ($wb)    { $wb.Close($false) }
        if ($excel) { $excel.Quit() }
        foreach ($o in @($ws, $wb, $excel)) {
            if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# ---------------------------------------------------------------------
# CSV の書き込み（一時ファイル経由で置き換えるので途中で落ちても元が壊れない）
# ---------------------------------------------------------------------
function Save-Csv($Rows, [string[]]$Columns, [string]$Path) {
    $tmp = "$Path.tmp"
    if (@($Rows).Count -eq 0) {
        # 行が無いときも見出しだけは書く
        ($Columns | ForEach-Object { '"' + $_.Replace('"', '""') + '"' }) -join ',' |
            Set-Content -Path $tmp -Encoding $CsvEncoding
    }
    else {
        $Rows | Select-Object -Property $Columns |
            Export-Csv -Path $tmp -NoTypeInformation -Encoding $CsvEncoding
    }
    Move-Item -Path $tmp -Destination $Path -Force
}

# 入荷予定ファイルの見出し（行が0件でも取れるように1行目から読む）
function Get-CsvHeader([string]$Path) {
    $first = Get-Content -Path $Path -Encoding $CsvEncoding -TotalCount 1
    if (-not $first) { return @() }
    return @($first -split ',' | ForEach-Object { $_.Trim().Trim('"') })
}

function Get-RowKey($Row, [string[]]$KeyCols) {
    return ($KeyCols | ForEach-Object { ([string]$Row.$_).Trim() }) -join [char]0x1F
}

# ---------------------------------------------------------------------
# ⑤ マージ
# ---------------------------------------------------------------------
function Merge-Plan([string]$PlanPath, $NewRows) {
    if (Test-Path $PlanPath) {
        $planCols = Get-CsvHeader $PlanPath
        $planRows = New-Object System.Collections.ArrayList
        foreach ($r in @(Import-Csv -Path $PlanPath -Encoding $CsvEncoding)) { [void]$planRows.Add($r) }
        if ($Config.BackupPlanFile) { Copy-Item $PlanPath "$PlanPath.bak" -Force }
    }
    else {
        $planCols = $Config.OutputColumns
        $planRows = New-Object System.Collections.ArrayList
        Write-Log "  入荷予定ファイルが無いため新規作成: $PlanPath"
    }

    $missing = @($Config.OutputColumns | Where-Object { $planCols -notcontains $_ })
    if ($missing.Count -gt 0) {
        Write-Log ('  [警告] 入荷予定ファイルに無い列は取り込みません: ' + ($missing -join ', '))
    }

    $keyCols = if (@($Config.KeyColumns).Count -gt 0) { $Config.KeyColumns } else { $planCols }
    foreach ($k in $keyCols) {
        if ($planCols -notcontains $k) { throw "入荷予定ファイルにキー列「$k」がありません: $PlanPath" }
    }

    # 既存行のキー → 行番号
    $index = @{}
    for ($i = 0; $i -lt $planRows.Count; $i++) { $index[(Get-RowKey $planRows[$i] $keyCols)] = $i }

    $added = 0; $updated = 0
    foreach ($n in $NewRows) {
        $key = Get-RowKey $n $keyCols

        if ($index.ContainsKey($key)) {
            # 上書き（新しいファイルに無い列は元の値を残す）
            $old = $planRows[$index[$key]]
            $changed = $false
            foreach ($c in $planCols) {
                if ($Config.OutputColumns -contains $c) {
                    if ([string]$old.$c -ne [string]$n.$c) { $old.$c = $n.$c; $changed = $true }
                }
            }
            if ($changed) { $updated++ }
        }
        else {
            # 追加（入荷予定ファイルの列順に合わせる）
            $obj = [ordered]@{}
            foreach ($c in $planCols) {
                $obj[$c] = if ($Config.OutputColumns -contains $c) { $n.$c } else { '' }
            }
            $index[$key] = $planRows.Add([pscustomobject]$obj)
            $added++
        }
    }

    Save-Csv $planRows $planCols $PlanPath
    return @($added, $updated)
}

# ---------------------------------------------------------------------
# 1ファイル分の処理
# ---------------------------------------------------------------------
function Invoke-OneExcel([string]$ExcelPath) {
    Write-Log "--- $(Split-Path $ExcelPath -Leaf) ---"

    # ① Excel → CSV
    $tmpCsv = Join-Path $env:TEMP ('arrival_' + [guid]::NewGuid().ToString('N') + '.csv')
    try {
        Convert-ExcelToCsv $ExcelPath $tmpCsv
        $rows = @(Import-Csv -Path $tmpCsv -Encoding $CsvEncoding)
    }
    finally {
        if (Test-Path $tmpCsv) { Remove-Item $tmpCsv -Force }
    }

    if ($rows.Count -eq 0) { Write-Log '  データ行がありません。'; return }

    # ② 必要な列だけに加工する
    $srcCols = @($rows[0].PSObject.Properties.Name)
    foreach ($c in $Config.OutputColumns) {
        if ($srcCols -notcontains $c) { throw "Excel に列「$c」が見つかりません。" }
    }
    if ($Config.OutputColumns -notcontains $Config.OrderNoColumn) {
        throw "オーダー番号列「$($Config.OrderNoColumn)」が OutputColumns に含まれていません。"
    }
    $rows = @($rows | Select-Object -Property $Config.OutputColumns |
              Where-Object { -not [string]::IsNullOrWhiteSpace($_.($Config.OrderNoColumn)) })

    # ③ オーダー番号でファイルを分ける
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $groups = $rows | Group-Object -Property { ([string]$_.($Config.OrderNoColumn)).Trim() }

    foreach ($g in $groups) {
        $order  = $g.Name
        $folder = Join-Path $Config.OrderRootFolder (Get-SafeName $order)

        # ④ オーダー別フォルダに移動する
        if (-not (Test-Path $folder)) {
            if (-not $Config.CreateOrderFolderIfMissing) {
                Write-Log "  [警告] フォルダが無いためスキップ: $folder"
                continue
            }
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
            Write-Log "  フォルダ作成: $folder"
        }

        if ($Config.SaveSplitFile) {
            $splitPath = Join-Path $folder ("{0}_{1}.csv" -f (Get-SafeName $order), $timestamp)
            Save-Csv $g.Group $Config.OutputColumns $splitPath
        }

        # ⑤ マージ
        $planPath = Join-Path $folder $Config.PlanFileName
        $result = Merge-Plan $planPath $g.Group
        Write-Log ("  オーダー {0}: 追加 {1} 行 / 上書き {2} 行" -f $order, $result[0], $result[1])
    }
}

# =====================================================================
# メイン
# =====================================================================
$files = @(Get-ChildItem -Path $Config.InputFolder -Filter '*.xlsx' -File |
           Where-Object { -not $_.Name.StartsWith('~$') } |   # Excel のロックファイルを除外
           Sort-Object Name)

if ($files.Count -eq 0) {
    Write-Log '処理対象の Excel がありません。'
    exit 0
}

$errors = 0
foreach ($f in $files) {
    try {
        Invoke-OneExcel $f.FullName

        if ($Config.ProcessedFolder) {
            New-Item -ItemType Directory -Path $Config.ProcessedFolder -Force | Out-Null
            $dest = Join-Path $Config.ProcessedFolder ('{0}_{1}{2}' -f $f.BaseName, (Get-Date -Format 'yyyyMMdd_HHmmss'), $f.Extension)
            Move-Item -Path $f.FullName -Destination $dest
        }
    }
    catch {
        $errors++
        Write-Log "[エラー] $($f.Name): $($_.Exception.Message)"
    }
}

Write-Log ("完了: {0} 件成功 / {1} 件失敗" -f ($files.Count - $errors), $errors)
if ($errors -gt 0) { exit 1 } else { exit 0 }
