<#
.SYNOPSIS
  入荷予定ファイル自動更新スクリプト

  1. Excel を CSV にする
  2. 必要な列だけに加工する
  3. 必要な行だけ抜き出す
  4. オーダー番号でファイルを分ける
  5. オーダー別フォルダに移動する
  6. 各フォルダの入荷予定ファイル（Excel）に新しい内容をマージする
     （キー列が同じ行があれば上書き、無ければ追加。追加・上書きした行は最終更新日を現在時刻にする）
  7. マージに使った新しいファイルをオールドフォルダに移動する

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

    # 残す行の条件（すべて満たす行だけ残す。@() なら全行）
    # Operator: Equals / NotEquals / Contains / NotContains / In / NotIn / Empty / NotEmpty / Regex
    # 判定は元の全列で行うので、OutputColumns に無い列でも絞り込める
    RowFilters       = @(
        @{ Column = '数量'; Operator = 'NotEmpty' }
        @{ Column = '区分'; Operator = 'In'; Values = @('入荷', '追加') }
    )

    # オーダー別フォルダの親フォルダ（この下に「オーダー番号」名のフォルダ）
    OrderRootFolder  = 'C:\Work\入荷\オーダー'
    # オーダー別フォルダが無いとき作るか（$false ならスキップして警告）
    CreateOrderFolderIfMissing = $true
    # マージに使った分割CSVの移動先（相対パスなら各オーダーフォルダの中）
    OldFolder        = 'old'

    # 各オーダーフォルダ内の入荷予定ファイル名（Excel）
    # 同じ名前の .csv が既にあれば、初回にその内容を引き継ぐ
    PlanFileName     = '入荷予定.xlsx'
    # 入荷予定ファイルを新規作成するときのシート名
    PlanSheetName    = '入荷予定'
    # 最後の列に追加する最終更新日（行を追加・上書きしたときに現在時刻を入れる）
    LastUpdatedColumn = '最終更新日'
    LastUpdatedFormat = 'yyyy/MM/dd HH:mm:ss'
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
# Excel 本体（起動は最初の1回だけ。全ファイルで使い回す）
# ---------------------------------------------------------------------
$script:Excel = $null

function Release-Com($o) {
    if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
}

function Start-Excel {
    $script:Excel = New-Object -ComObject Excel.Application
    $script:Excel.Visible = $false
    $script:Excel.DisplayAlerts = $false
}

function Stop-Excel {
    if ($script:Excel) {
        $script:Excel.Quit()
        Release-Com $script:Excel
        $script:Excel = $null
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

# ---------------------------------------------------------------------
# ① Excel → CSV（Excel 本体で「名前を付けて保存」する）
# ---------------------------------------------------------------------
function Convert-ExcelToCsv([string]$ExcelPath, [string]$CsvPath) {
    $wb = $null; $ws = $null
    try {
        # 読み取り専用で開く（元ファイルは変更しない）
        $wb = $script:Excel.Workbooks.Open($ExcelPath, 0, $true)
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
        if ($wb) { $wb.Close($false) }
        Release-Com $ws; Release-Com $wb
    }
}

# ---------------------------------------------------------------------
# 入荷予定ファイル（Excel）の読み書き
# ---------------------------------------------------------------------
function Read-PlanXlsx([string]$Path) {
    $wb = $null; $ws = $null; $range = $null
    try {
        $wb = $script:Excel.Workbooks.Open($Path, 0, $true)
        $ws = $wb.Worksheets.Item(1)
        $ur = $ws.UsedRange
        $lastRow = $ur.Row + $ur.Rows.Count - 1
        $lastCol = $ur.Column + $ur.Columns.Count - 1
        Release-Com $ur

        $range = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($lastRow, $lastCol))
        $v = $range.Value2
        if ($v -isnot [Array]) {
            # セルが1つしか無いとき
            $one = New-Object 'object[,]' 1, 1
            $one[0, 0] = $v
            $v = $one
            $base = 0
        } else {
            $base = 1   # Excel から返る配列は 1 始まり
        }
        $nRows = $v.GetLength(0); $nCols = $v.GetLength(1)

        $cols = @()
        for ($c = 0; $c -lt $nCols; $c++) {
            $cols += ([string]$v[$base, ($c + $base)]).Trim()
        }
        # 末尾の空の見出しは捨てる
        while ($cols.Count -gt 0 -and $cols[-1] -eq '') { $cols = @($cols[0..($cols.Count - 2)]) }

        $rows = New-Object System.Collections.ArrayList
        for ($r = 1; $r -lt $nRows; $r++) {
            $obj = [ordered]@{}
            $empty = $true
            for ($c = 0; $c -lt $cols.Count; $c++) {
                $s = [string]$v[($r + $base), ($c + $base)]
                if ($s.Trim() -ne '') { $empty = $false }
                $obj[$cols[$c]] = $s
            }
            if (-not $empty) { [void]$rows.Add([pscustomobject]$obj) }
        }
        return @{ Cols = $cols; Rows = $rows }
    }
    finally {
        if ($wb) { $wb.Close($false) }
        Release-Com $range; Release-Com $ws; Release-Com $wb
    }
}

function Save-PlanXlsx([string]$Path, [string[]]$Cols, $Rows) {
    $wb = $null; $ws = $null; $range = $null
    $isNew = -not (Test-Path $Path)
    try {
        if ($isNew) {
            $wb = $script:Excel.Workbooks.Add()
            $ws = $wb.Worksheets.Item(1)
            $ws.Name = $Config.PlanSheetName
        } else {
            # 既存ファイルは値だけ書き換える（列幅や色などの書式は残る）
            $wb = $script:Excel.Workbooks.Open($Path)
            $ws = $wb.Worksheets.Item(1)
        }

        $rowList = @($Rows)
        $n = $rowList.Count + 1
        $arr = New-Object 'object[,]' $n, $Cols.Count
        for ($c = 0; $c -lt $Cols.Count; $c++) { $arr[0, $c] = $Cols[$c] }
        for ($r = 0; $r -lt $rowList.Count; $r++) {
            for ($c = 0; $c -lt $Cols.Count; $c++) { $arr[($r + 1), $c] = [string]$rowList[$r].($Cols[$c]) }
        }

        $range = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($n, $Cols.Count))
        # 先頭の 0 が消えたり日付に化けたりしないよう、すべて文字列で書く
        $range.NumberFormat = '@'
        $range.Value2 = $arr

        if ($isNew) {
            $ws.Rows.Item(1).Font.Bold = $true
            [void]$range.Columns.AutoFit()
        }

        # 一時ファイルに保存してから置き換える（途中で落ちても元ファイルが壊れない）
        $tmp = "$Path.tmp.xlsx"
        if (Test-Path $tmp) { Remove-Item $tmp -Force }
        $wb.SaveAs($tmp, 51)   # 51 = xlOpenXMLWorkbook (.xlsx)
        $wb.Close($false); $wb = $null
        Move-Item -Path $tmp -Destination $Path -Force
    }
    finally {
        if ($wb) { $wb.Close($false) }
        Release-Com $range; Release-Com $ws; Release-Com $wb
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

function Test-RowFilter($Row, $Filter) {
    $cell = ([string]$Row.($Filter.Column)).Trim()
    $value = if ($Filter.ContainsKey('Value')) { [string]$Filter.Value } else { '' }
    $values = if ($Filter.ContainsKey('Values')) { @($Filter.Values) } else { @() }
    switch ($Filter.Operator) {
        # 大文字・小文字は区別しない（-eq / -contains / -match は既定で区別しない）
        'Equals'      { return $cell -eq $value }
        'NotEquals'   { return $cell -ne $value }
        'Contains'    { return $cell.IndexOf($value, [StringComparison]::OrdinalIgnoreCase) -ge 0 }
        'NotContains' { return $cell.IndexOf($value, [StringComparison]::OrdinalIgnoreCase) -lt 0 }
        'In'          { return $values -contains $cell }
        'NotIn'       { return $values -notcontains $cell }
        'Empty'       { return $cell.Length -eq 0 }
        'NotEmpty'    { return $cell.Length -gt 0 }
        'Regex'       { return $cell -match $value }
        default       { throw "RowFilters の Operator「$($Filter.Operator)」は使えません。" }
    }
}

function Get-RowKey($Row, [string[]]$KeyCols) {
    return ($KeyCols | ForEach-Object { ([string]$Row.$_).Trim() }) -join [char]0x1F
}

# ---------------------------------------------------------------------
# ⑥ マージ
# ---------------------------------------------------------------------
function Merge-Plan([string]$PlanPath, $NewRows) {
    $csvPath = [IO.Path]::ChangeExtension($PlanPath, '.csv')
    $planRows = New-Object System.Collections.ArrayList

    if (Test-Path $PlanPath) {
        $plan = Read-PlanXlsx $PlanPath
        $planCols = @($plan.Cols)
        foreach ($r in $plan.Rows) { [void]$planRows.Add($r) }
        if ($Config.BackupPlanFile) { Copy-Item $PlanPath "$PlanPath.bak" -Force }
    }
    elseif (Test-Path $csvPath) {
        # 以前の CSV 版の入荷予定ファイルがあれば、その内容を引き継いで Excel にする
        $planCols = @(Get-CsvHeader $csvPath)
        foreach ($r in @(Import-Csv -Path $csvPath -Encoding $CsvEncoding)) { [void]$planRows.Add($r) }
        Write-Log "  CSV の入荷予定ファイルを Excel に移行: $csvPath → $PlanPath"
    }
    else {
        $planCols = @($Config.OutputColumns)
        Write-Log "  入荷予定ファイルが無いため新規作成: $PlanPath"
    }

    # 空のファイルなら新しいファイルの列で作る
    if ($planCols.Count -eq 0) { $planCols = @($Config.OutputColumns) }

    # 最終更新日の列が無ければ最後の列に追加する
    $lastCol = $Config.LastUpdatedColumn
    if ($planCols -notcontains $lastCol) {
        $planCols += $lastCol
        foreach ($r in $planRows) { $r | Add-Member -NotePropertyName $lastCol -NotePropertyValue '' -Force }
    }

    $missing = @($Config.OutputColumns | Where-Object { $planCols -notcontains $_ })
    if ($missing.Count -gt 0) {
        Write-Log ('  [警告] 入荷予定ファイルに無い列は取り込みません: ' + ($missing -join ', '))
    }

    # 取り込む列（入荷予定ファイルにあって、新しいファイルにもある列。最終更新日は除く）
    $copyCols = @($planCols | Where-Object { $_ -ne $lastCol -and $Config.OutputColumns -contains $_ })

    $keyCols = if (@($Config.KeyColumns).Count -gt 0) { $Config.KeyColumns }
               else { @($planCols | Where-Object { $_ -ne $lastCol }) }
    foreach ($k in $keyCols) {
        if ($planCols -notcontains $k) { throw "入荷予定ファイルにキー列「$k」がありません: $PlanPath" }
    }

    # 既存行のキー → 行番号（PowerShell のハッシュテーブルは大文字・小文字を区別しない）
    $index = @{}
    for ($i = 0; $i -lt $planRows.Count; $i++) { $index[(Get-RowKey $planRows[$i] $keyCols)] = $i }

    $now = Get-Date -Format $Config.LastUpdatedFormat
    $added = 0; $updated = 0
    foreach ($n in $NewRows) {
        $key = Get-RowKey $n $keyCols

        if ($index.ContainsKey($key)) {
            # 内容が同じでも上書きし、最終更新日を更新する（新しいファイルに無い列は元の値を残す）
            $old = $planRows[$index[$key]]
            foreach ($c in $copyCols) { $old.$c = $n.$c }
            $old.$lastCol = $now
            $updated++
        }
        else {
            # 追加（入荷予定ファイルの列順に合わせる）
            $obj = [ordered]@{}
            foreach ($c in $planCols) {
                $obj[$c] = if ($copyCols -contains $c) { $n.$c } else { '' }
            }
            $obj[$lastCol] = $now
            $index[$key] = $planRows.Add([pscustomobject]$obj)
            $added++
        }
    }

    Save-PlanXlsx $PlanPath $planCols $planRows
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

    $srcCols = @($rows[0].PSObject.Properties.Name)
    foreach ($c in $Config.OutputColumns) {
        if ($srcCols -notcontains $c) { throw "Excel に列「$c」が見つかりません。" }
    }

    # ③ 必要な行だけ抜き出す（元の全列で判定）
    foreach ($f in @($Config.RowFilters)) {
        if ($srcCols -notcontains $f.Column) { throw "抽出条件の列「$($f.Column)」が Excel に見つかりません。" }
    }
    $before = $rows.Count
    foreach ($f in @($Config.RowFilters)) {
        $rows = @($rows | Where-Object { Test-RowFilter $_ $f })
    }
    Write-Log "  行の抽出: $before 行 → $($rows.Count) 行"
    if ($rows.Count -eq 0) { return }

    # ② 必要な列だけに加工する
    if ($Config.OutputColumns -notcontains $Config.OrderNoColumn) {
        throw "オーダー番号列「$($Config.OrderNoColumn)」が OutputColumns に含まれていません。"
    }
    $rows = @($rows | Select-Object -Property $Config.OutputColumns |
              Where-Object { -not [string]::IsNullOrWhiteSpace($_.($Config.OrderNoColumn)) })

    # ④ オーダー番号でファイルを分ける
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $groups = $rows | Group-Object -Property { ([string]$_.($Config.OrderNoColumn)).Trim() }

    foreach ($g in $groups) {
        $order  = $g.Name
        $folder = Join-Path $Config.OrderRootFolder (Get-SafeName $order)

        # ⑤ オーダー別フォルダに移動する
        if (-not (Test-Path $folder)) {
            if (-not $Config.CreateOrderFolderIfMissing) {
                Write-Log "  [警告] フォルダが無いためスキップ: $folder"
                continue
            }
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
            Write-Log "  フォルダ作成: $folder"
        }

        $splitName = "{0}_{1}.csv" -f (Get-SafeName $order), $timestamp
        $splitPath = Join-Path $folder $splitName
        Save-Csv $g.Group $Config.OutputColumns $splitPath

        # ⑥ マージ
        $planPath = Join-Path $folder $Config.PlanFileName
        $result = Merge-Plan $planPath $g.Group
        Write-Log ("  オーダー {0}: 追加 {1} 行 / 上書き {2} 行" -f $order, $result[0], $result[1])

        # ⑦ マージに使った新しいファイルをオールドフォルダに移動する
        if ([IO.Path]::IsPathRooted($Config.OldFolder)) {
            $oldFolder = Join-Path $Config.OldFolder (Get-SafeName $order)
        } else {
            $oldFolder = Join-Path $folder $Config.OldFolder
        }
        New-Item -ItemType Directory -Path $oldFolder -Force | Out-Null
        Move-Item -Path $splitPath -Destination (Join-Path $oldFolder $splitName) -Force
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
Start-Excel
try {
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
}
finally {
    Stop-Excel
}

Write-Log ("完了: {0} 件成功 / {1} 件失敗" -f ($files.Count - $errors), $errors)
if ($errors -gt 0) { exit 1 } else { exit 0 }
