using System.Text;
using System.Text.Json;
using ClosedXML.Excel;

namespace ArrivalPlanMerger;

/// <summary>設定ファイル(settings.json)の内容</summary>
public sealed class Settings
{
    /// <summary>処理対象のExcelを置くフォルダ</summary>
    public string InputFolder { get; set; } = "";
    /// <summary>処理が終わったExcelの移動先（空なら移動しない）</summary>
    public string ProcessedFolder { get; set; } = "";
    /// <summary>読み込むシート名（空なら先頭シート）</summary>
    public string SheetName { get; set; } = "";
    /// <summary>見出し行の行番号（1始まり）</summary>
    public int HeaderRow { get; set; } = 1;

    /// <summary>CSVに残す列（この順番で出力）</summary>
    public List<string> OutputColumns { get; set; } = new();
    /// <summary>ファイル分割に使うオーダー番号の列名</summary>
    public string OrderNoColumn { get; set; } = "";

    /// <summary>残す行の条件（すべて満たす行だけ残す。空なら全行）</summary>
    public List<RowFilter> RowFilters { get; set; } = new();

    /// <summary>オーダー別フォルダの親フォルダ</summary>
    public string OrderRootFolder { get; set; } = "";
    /// <summary>オーダー別フォルダが無いとき作るか（falseならスキップして警告）</summary>
    public bool CreateOrderFolderIfMissing { get; set; } = true;
    /// <summary>分割CSVのファイル名。{order} と {timestamp} が使える</summary>
    public string SplitFileNameFormat { get; set; } = "{order}_{timestamp}.csv";
    /// <summary>マージに使った分割CSVの移動先。フォルダ名だけなら各オーダーフォルダの中、パス指定ならカレントフォルダ基準</summary>
    public string OldFolder { get; set; } = "old";

    /// <summary>各オーダーフォルダ内の入荷予定ファイル名（Excel）</summary>
    public string PlanFileName { get; set; } = "入荷予定.xlsx";
    /// <summary>入荷予定ファイルを新規作成するときのシート名</summary>
    public string PlanSheetName { get; set; } = "入荷予定";
    /// <summary>最終更新日の列名（無ければ最後の列に追加）</summary>
    public string LastUpdatedColumn { get; set; } = "最終更新日";
    /// <summary>最終更新日の書式</summary>
    public string LastUpdatedFormat { get; set; } = "yyyy/MM/dd HH:mm:ss";
    /// <summary>状態の列名（無ければ最後の列に追加）</summary>
    public string StatusColumn { get; set; } = "状態";
    /// <summary>今回のマージで更新しなかった行の状態の列に入れる文字</summary>
    public string CompletedText { get; set; } = "完納";
    /// <summary>今回のマージで更新しなかった行の背景色（"#RRGGBB"。空なら色を付けない）</summary>
    public string UntouchedRowColor { get; set; } = "#D9D9D9";
    /// <summary>同じ行とみなす列。空なら全列一致で判定</summary>
    public List<string> KeyColumns { get; set; } = new();

    /// <summary>CSVの文字コード（shift_jis / utf-8 / utf-8-bom）</summary>
    public string CsvEncoding { get; set; } = "shift_jis";
    /// <summary>マージ前に入荷予定ファイルの .bak を作るか</summary>
    public bool BackupPlanFile { get; set; } = true;
}

/// <summary>行の抽出条件</summary>
public sealed class RowFilter
{
    /// <summary>判定する列名（OutputColumns に無い列も指定可）</summary>
    public string Column { get; set; } = "";
    /// <summary>Equals / NotEquals / Contains / NotContains / In / NotIn / Empty / NotEmpty / Regex</summary>
    public string Operator { get; set; } = "Equals";
    /// <summary>比較する値（In / NotIn は Values を使う）</summary>
    public string Value { get; set; } = "";
    public List<string> Values { get; set; } = new();

    private const StringComparison IgnoreCase = StringComparison.OrdinalIgnoreCase;

    public bool IsMatch(string cell)
    {
        cell = (cell ?? "").Trim();
        return Operator.ToLowerInvariant() switch
        {
            // 大文字・小文字は区別しない
            "equals"      => string.Equals(cell, Value, IgnoreCase),
            "notequals"   => !string.Equals(cell, Value, IgnoreCase),
            "contains"    => cell.Contains(Value, IgnoreCase),
            "notcontains" => !cell.Contains(Value, IgnoreCase),
            "in"          => Values.Any(v => string.Equals(cell, v.Trim(), IgnoreCase)),
            "notin"       => !Values.Any(v => string.Equals(cell, v.Trim(), IgnoreCase)),
            "empty"       => cell.Length == 0,
            "notempty"    => cell.Length > 0,
            "regex"       => System.Text.RegularExpressions.Regex.IsMatch(cell, Value,
                                 System.Text.RegularExpressions.RegexOptions.IgnoreCase),
            _ => throw new InvalidOperationException($"RowFilters の Operator「{Operator}」は使えません。"),
        };
    }
}

/// <summary>見出し付きの表データ</summary>
public sealed class Table
{
    public List<string> Header { get; } = new();
    public List<List<string>> Rows { get; } = new();

    public int IndexOf(string column) => Header.FindIndex(h => h == column);
}

internal static class Program
{
    private static StreamWriter? _log;

    private static int Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        string settingsPath = args.Length > 0 ? args[0] : Path.Combine(AppContext.BaseDirectory, "settings.json");
        string logPath = Path.Combine(AppContext.BaseDirectory, "logs", $"{DateTime.Now:yyyyMMdd}.log");
        Directory.CreateDirectory(Path.GetDirectoryName(logPath)!);
        _log = new StreamWriter(logPath, append: true, new UTF8Encoding(false)) { AutoFlush = true, NewLine = "\r\n" };

        try
        {
            var settings = LoadSettings(settingsPath);
            var encoding = GetEncoding(settings.CsvEncoding);

            var excelFiles = Directory.GetFiles(settings.InputFolder, "*.xlsx")
                .Where(f => !Path.GetFileName(f).StartsWith("~$")) // Excelが開いているときのロックファイルを除外
                .OrderBy(f => f)
                .ToList();

            if (excelFiles.Count == 0)
            {
                Log("処理対象のExcelがありません。");
                return 0;
            }

            int errorCount = 0;
            foreach (var file in excelFiles)
            {
                try
                {
                    ProcessExcel(file, settings, encoding);
                    MoveToProcessed(file, settings);
                }
                catch (Exception ex)
                {
                    errorCount++;
                    Log($"[エラー] {Path.GetFileName(file)}: {ex.Message}");
                }
            }

            Log($"完了: {excelFiles.Count - errorCount} 件成功 / {errorCount} 件失敗");
            return errorCount == 0 ? 0 : 1;
        }
        catch (Exception ex)
        {
            Log($"[致命的エラー] {ex}");
            return 2;
        }
        finally
        {
            _log?.Dispose();
        }
    }

    // ---------------------------------------------------------------
    // 1ファイル分の処理
    // ---------------------------------------------------------------
    private static void ProcessExcel(string excelPath, Settings s, Encoding enc)
    {
        Log($"--- {Path.GetFileName(excelPath)} ---");

        // ① Excel を読み込む（CSV化）
        var source = ReadExcel(excelPath, s.SheetName, s.HeaderRow);

        // ② 必要な列だけに加工する
        // ③ 必要な行だけ抜き出す（条件は元の全列で判定するので、残さない列でも絞り込める）
        var filtered = FilterRows(source, s.RowFilters);
        Log($"行の抽出: {source.Rows.Count} 行 → {filtered.Rows.Count} 行");
        var trimmed = SelectColumns(filtered, s.OutputColumns);

        // ④ オーダー番号でファイルを分ける
        int orderIdx = trimmed.IndexOf(s.OrderNoColumn);
        if (orderIdx < 0)
            throw new InvalidOperationException($"オーダー番号列「{s.OrderNoColumn}」が OutputColumns に含まれていません。");

        var groups = trimmed.Rows
            .Where(r => !string.IsNullOrWhiteSpace(r[orderIdx]))
            .GroupBy(r => r[orderIdx].Trim(), StringComparer.OrdinalIgnoreCase); // Windowsのフォルダ名と同じく大文字・小文字を区別しない

        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

        foreach (var g in groups)
        {
            string order = g.Key;
            string folder = Path.Combine(s.OrderRootFolder, SanitizeFileName(order));

            // ⑤ オーダー別フォルダに移動する
            if (!Directory.Exists(folder))
            {
                if (!s.CreateOrderFolderIfMissing)
                {
                    Log($"[警告] フォルダがないためスキップ: {folder}");
                    continue;
                }
                Directory.CreateDirectory(folder);
                Log($"フォルダ作成: {folder}");
            }

            var part = new Table();
            part.Header.AddRange(trimmed.Header);
            part.Rows.AddRange(g);

            string splitName = s.SplitFileNameFormat
                .Replace("{order}", SanitizeFileName(order))
                .Replace("{timestamp}", timestamp);
            string splitPath = Path.Combine(folder, splitName);
            WriteCsv(splitPath, part, enc);

            // ⑥ 入荷予定ファイルにマージする
            string planPath = Path.Combine(folder, s.PlanFileName);
            var (added, updated) = MergeIntoPlan(planPath, part, s, enc);
            Log($"オーダー {order}: 追加 {added} 行 / 上書き {updated} 行");

            // ⑦ マージに使った新しいファイルをオールドフォルダに移動する
            string oldFolder = Path.IsPathRooted(s.OldFolder)
                ? Path.Combine(s.OldFolder, SanitizeFileName(order))
                : Path.Combine(folder, s.OldFolder);
            Directory.CreateDirectory(oldFolder);
            File.Move(splitPath, Path.Combine(oldFolder, splitName), overwrite: true);
        }
    }

    // ---------------------------------------------------------------
    // Excel 読み込み
    // ---------------------------------------------------------------
    private static Table ReadExcel(string path, string sheetName, int headerRow)
    {
        // Excelで開いたままでも読めるように共有モードで開く
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var wb = new XLWorkbook(fs);
        var ws = string.IsNullOrWhiteSpace(sheetName) ? wb.Worksheet(1) : wb.Worksheet(sheetName);

        var table = new Table();
        var used = ws.RangeUsed();
        if (used == null) return table;

        int lastCol = used.LastColumn().ColumnNumber();
        int lastRow = used.LastRow().RowNumber();

        for (int c = 1; c <= lastCol; c++)
            table.Header.Add(ws.Cell(headerRow, c).GetFormattedString().Trim());

        for (int r = headerRow + 1; r <= lastRow; r++)
        {
            var row = new List<string>(lastCol);
            for (int c = 1; c <= lastCol; c++)
                row.Add(ws.Cell(r, c).GetFormattedString()); // 画面に見えている表示形式のまま取得
            if (row.All(string.IsNullOrWhiteSpace)) continue; // 空行は無視
            table.Rows.Add(row);
        }
        return table;
    }

    private static Table FilterRows(Table src, List<RowFilter> filters)
    {
        if (filters.Count == 0) return src;

        var idx = filters.Select(f =>
        {
            int i = src.IndexOf(f.Column);
            if (i < 0) throw new InvalidOperationException($"抽出条件の列「{f.Column}」がExcelに見つかりません。");
            return i;
        }).ToArray();

        var dst = new Table();
        dst.Header.AddRange(src.Header);
        foreach (var row in src.Rows)
        {
            bool ok = true;
            for (int k = 0; k < filters.Count && ok; k++)
                ok = filters[k].IsMatch(idx[k] < row.Count ? row[idx[k]] : "");
            if (ok) dst.Rows.Add(row);
        }
        return dst;
    }

    private static Table SelectColumns(Table src, List<string> columns)
    {
        var indexes = new List<int>();
        foreach (var col in columns)
        {
            int idx = src.IndexOf(col);
            if (idx < 0) throw new InvalidOperationException($"Excelに列「{col}」が見つかりません。");
            indexes.Add(idx);
        }

        var dst = new Table();
        dst.Header.AddRange(columns);
        foreach (var row in src.Rows)
            dst.Rows.Add(indexes.Select(i => i < row.Count ? row[i] : "").ToList());
        return dst;
    }

    // ---------------------------------------------------------------
    // マージ: 同じキーの行は上書き、無ければ末尾に追加
    // ---------------------------------------------------------------
    private static (int added, int updated) MergeIntoPlan(string planPath, Table incoming, Settings s, Encoding enc)
    {
        Table plan;
        string csvPath = Path.ChangeExtension(planPath, ".csv");
        if (File.Exists(planPath))
        {
            plan = ReadPlanXlsx(planPath);
            if (s.BackupPlanFile)
                File.Copy(planPath, planPath + ".bak", overwrite: true);
        }
        else if (File.Exists(csvPath))
        {
            // 以前の CSV 版の入荷予定ファイルがあれば、その内容を引き継いで Excel にする
            plan = ReadCsv(csvPath, enc);
            Log($"CSV の入荷予定ファイルを Excel に移行: {csvPath} → {planPath}");
        }
        else
        {
            plan = new Table();
            plan.Header.AddRange(incoming.Header);
            Log($"入荷予定ファイルが無いため新規作成: {planPath}");
        }

        // 空のファイルなら新しいファイルの列で作る
        if (plan.Header.Count == 0)
            plan.Header.AddRange(incoming.Header);

        // 最終更新日の列が無ければ最後の列に追加する
        string lastUpdatedCol = s.LastUpdatedColumn;
        if (plan.IndexOf(lastUpdatedCol) < 0)
            plan.Header.Add(lastUpdatedCol);
        int lastUpdatedIdx = plan.IndexOf(lastUpdatedCol);

        // 状態の列（今回更新しなかった行に「完納」を入れる）が無ければ最後の列に追加する
        string statusCol = s.StatusColumn;
        if (plan.IndexOf(statusCol) < 0)
            plan.Header.Add(statusCol);
        int statusIdx = plan.IndexOf(statusCol);

        foreach (var r in plan.Rows)
            while (r.Count < plan.Header.Count) r.Add("");

        // 入荷予定ファイルの列順に合わせるための対応表（最終更新日・状態は取り込み対象外）
        var map = plan.Header.Select((h, i) =>
            i == lastUpdatedIdx || i == statusIdx ? -1 : incoming.IndexOf(h)).ToArray();
        var missing = incoming.Header.Where(h => plan.IndexOf(h) < 0).ToList();
        if (missing.Count > 0)
            Log($"[警告] 入荷予定ファイルに無い列は取り込みません: {string.Join(", ", missing)}");

        var keyCols = s.KeyColumns.Count > 0
            ? s.KeyColumns
            : plan.Header.Where(h => h != lastUpdatedCol && h != statusCol).ToList();
        var keyIdx = keyCols.Select(k =>
        {
            int i = plan.IndexOf(k);
            if (i < 0) throw new InvalidOperationException($"入荷予定ファイルにキー列「{k}」がありません: {planPath}");
            return i;
        }).ToArray();

        string KeyOf(List<string> row) =>
            string.Join("\u001F", keyIdx.Select(i => i < row.Count ? row[i].Trim() : ""));

        // 既存行のキー → 行番号（大文字・小文字は区別しない）
        var index = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < plan.Rows.Count; i++)
            index[KeyOf(plan.Rows[i])] = i;

        string now = DateTime.Now.ToString(s.LastUpdatedFormat);
        var touched = new HashSet<int>();   // 今回追加・上書きした行
        int added = 0, updated = 0;
        foreach (var inRow in incoming.Rows)
        {
            var row = map.Select(m => m >= 0 ? inRow[m] : "").ToList();
            string key = KeyOf(row);

            if (index.TryGetValue(key, out int existing))
            {
                // 新しいファイルに無い列（最終更新日を含む）は元の値を残す
                var old = plan.Rows[existing];
                for (int c = 0; c < row.Count; c++)
                    if (map[c] < 0) row[c] = old[c];

                // 内容が同じでも上書きし、最終更新日を更新する
                row[lastUpdatedIdx] = now;
                plan.Rows[existing] = row;
                touched.Add(existing);
                updated++;
            }
            else
            {
                row[lastUpdatedIdx] = now;
                plan.Rows.Add(row);
                index[key] = plan.Rows.Count - 1;
                touched.Add(plan.Rows.Count - 1);
                added++;
            }
        }

        // 今回更新しなかった行は「完納」、更新した行は空にする
        for (int i = 0; i < plan.Rows.Count; i++)
            plan.Rows[i][statusIdx] = touched.Contains(i) ? "" : s.CompletedText;

        WritePlanXlsx(planPath, plan, s.PlanSheetName, touched, s.UntouchedRowColor);
        return (added, updated);
    }

    // ---------------------------------------------------------------
    // 入荷予定ファイル（Excel）の読み書き
    // ---------------------------------------------------------------
    private static Table ReadPlanXlsx(string path)
    {
        var table = new Table();
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var wb = new XLWorkbook(fs);
        var ws = wb.Worksheet(1);

        var lastCell = ws.LastCellUsed();
        if (lastCell == null) return table;
        int lastCol = ws.Row(1).LastCellUsed()?.Address.ColumnNumber ?? 0;
        int lastRow = lastCell.Address.RowNumber;

        for (int c = 1; c <= lastCol; c++)
            table.Header.Add(ws.Cell(1, c).GetFormattedString().Trim());

        for (int r = 2; r <= lastRow; r++)
        {
            var row = new List<string>(lastCol);
            for (int c = 1; c <= lastCol; c++)
                row.Add(ws.Cell(r, c).GetFormattedString());
            if (row.All(string.IsNullOrWhiteSpace)) continue;
            table.Rows.Add(row);
        }
        return table;
    }

    private static void WritePlanXlsx(string path, Table table, string sheetName,
                                      HashSet<int> touched, string untouchedColor)
    {
        // 既存ファイルがあれば開いて値だけ書き換える（列幅や色などの書式は残る）
        bool isNew = !File.Exists(path);
        using var wb = isNew ? new XLWorkbook() : OpenWorkbook(path);
        var ws = isNew ? wb.Worksheets.Add(sheetName) : wb.Worksheet(1);

        int cols = table.Header.Count;
        int rows = table.Rows.Count + 1;

        // 先頭の 0 が消えたり日付に化けたりしないよう、すべて文字列で書く
        ws.Range(1, 1, rows, cols).Style.NumberFormat.Format = "@";

        for (int c = 0; c < cols; c++)
            ws.Cell(1, c + 1).Value = table.Header[c];
        for (int r = 0; r < table.Rows.Count; r++)
        {
            var row = table.Rows[r];
            for (int c = 0; c < cols; c++)
                ws.Cell(r + 2, c + 1).Value = c < row.Count ? row[c] : "";
        }

        // 今回更新しなかった行に色を付ける（更新した行は色を消す）
        if (!string.IsNullOrWhiteSpace(untouchedColor))
        {
            var color = XLColor.FromHtml(untouchedColor);
            for (int r = 0; r < table.Rows.Count; r++)
                ws.Range(r + 2, 1, r + 2, cols).Style.Fill.BackgroundColor =
                    touched.Contains(r) ? XLColor.NoColor : color;
        }

        if (isNew)
        {
            ws.Row(1).Style.Font.Bold = true;
            ws.SheetView.FreezeRows(1);
            ws.Columns(1, cols).AdjustToContents();
        }

        // 一時ファイルに保存してから置き換える（途中で落ちても元ファイルが壊れない）
        string tmp = path + ".tmp.xlsx";
        wb.SaveAs(tmp);
        File.Move(tmp, path, overwrite: true);
    }

    private static XLWorkbook OpenWorkbook(string path)
    {
        // using で FileStream を閉じると保存できなくなるため、メモリに読み込んでから開く
        var ms = new MemoryStream(File.ReadAllBytes(path));
        return new XLWorkbook(ms);
    }

    // ---------------------------------------------------------------
    // CSV 読み書き（ダブルクォート・改行入りセル対応）
    // ---------------------------------------------------------------
    private static Table ReadCsv(string path, Encoding enc)
    {
        var table = new Table();
        string text;
        using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var sr = new StreamReader(fs, enc, detectEncodingFromByteOrderMarks: true))
            text = sr.ReadToEnd();

        var records = ParseCsv(text);
        if (records.Count == 0) return table;

        table.Header.AddRange(records[0].Select(h => h.Trim()));
        foreach (var rec in records.Skip(1))
        {
            if (rec.All(string.IsNullOrWhiteSpace)) continue;
            while (rec.Count < table.Header.Count) rec.Add("");
            table.Rows.Add(rec);
        }
        return table;
    }

    private static List<List<string>> ParseCsv(string text)
    {
        var result = new List<List<string>>();
        var row = new List<string>();
        var field = new StringBuilder();
        bool inQuotes = false;

        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            if (inQuotes)
            {
                if (ch == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"') { field.Append('"'); i++; }
                    else inQuotes = false;
                }
                else field.Append(ch);
            }
            else
            {
                switch (ch)
                {
                    case '"': inQuotes = true; break;
                    case ',': row.Add(field.ToString()); field.Clear(); break;
                    case '\r': break;
                    case '\n':
                        row.Add(field.ToString()); field.Clear();
                        result.Add(row); row = new List<string>();
                        break;
                    default: field.Append(ch); break;
                }
            }
        }
        if (field.Length > 0 || row.Count > 0)
        {
            row.Add(field.ToString());
            result.Add(row);
        }
        return result;
    }

    private static void WriteCsv(string path, Table table, Encoding enc)
    {
        // 一時ファイルに書いてから置き換える（途中で落ちても元ファイルが壊れない）
        string tmp = path + ".tmp";
        using (var sw = new StreamWriter(tmp, false, enc) { NewLine = "\r\n" })
        {
            sw.WriteLine(string.Join(",", table.Header.Select(Escape)));
            foreach (var row in table.Rows)
                sw.WriteLine(string.Join(",", row.Select(Escape)));
        }
        File.Move(tmp, path, overwrite: true);
    }

    private static string Escape(string v)
    {
        if (v.IndexOfAny(new[] { ',', '"', '\r', '\n' }) < 0) return v;
        return "\"" + v.Replace("\"", "\"\"") + "\"";
    }

    // ---------------------------------------------------------------
    // その他
    // ---------------------------------------------------------------
    private static void MoveToProcessed(string file, Settings s)
    {
        if (string.IsNullOrWhiteSpace(s.ProcessedFolder)) return;
        Directory.CreateDirectory(s.ProcessedFolder);
        string dest = Path.Combine(s.ProcessedFolder,
            $"{Path.GetFileNameWithoutExtension(file)}_{DateTime.Now:yyyyMMdd_HHmmss}{Path.GetExtension(file)}");
        File.Move(file, dest);
    }

    private static Settings LoadSettings(string path)
    {
        var json = File.ReadAllText(path, Encoding.UTF8);
        var opt = new JsonSerializerOptions { ReadCommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true };
        var s = JsonSerializer.Deserialize<Settings>(json, opt) ?? throw new InvalidOperationException("settings.json が読めません。");
        if (s.OutputColumns.Count == 0) throw new InvalidOperationException("OutputColumns が空です。");

        // 相対パスは exe を実行したときのカレントフォルダを基準に絶対パスへ変換する
        string baseDir = Environment.CurrentDirectory;
        Log($"カレントフォルダ: {baseDir}");
        s.InputFolder     = ToFullPath(s.InputFolder, baseDir);
        s.ProcessedFolder = ToFullPath(s.ProcessedFolder, baseDir);
        s.OrderRootFolder = ToFullPath(s.OrderRootFolder, baseDir);
        // OldFolder は「フォルダ名だけ」なら各オーダーフォルダの中、
        // それ以外の相対パス（.\old や ..\old など）はカレントフォルダ基準
        if (!string.IsNullOrWhiteSpace(s.OldFolder) && s.OldFolder.IndexOfAny(new[] { '\\', '/' }) >= 0)
            s.OldFolder = ToFullPath(s.OldFolder, baseDir);
        return s;
    }

    private static string ToFullPath(string path, string baseDir) =>
        string.IsNullOrWhiteSpace(path) ? path : Path.GetFullPath(path, baseDir);

    private static Encoding GetEncoding(string name) => name.ToLowerInvariant() switch
    {
        "utf-8" or "utf8" => new UTF8Encoding(false),
        "utf-8-bom" or "utf8bom" => new UTF8Encoding(true),
        _ => Encoding.GetEncoding(name),
    };

    private static string SanitizeFileName(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
        return name;
    }

    private static void Log(string message)
    {
        string line = $"{DateTime.Now:yyyy/MM/dd HH:mm:ss} {message}";
        Console.WriteLine(line);
        _log?.WriteLine(line);
    }
}
