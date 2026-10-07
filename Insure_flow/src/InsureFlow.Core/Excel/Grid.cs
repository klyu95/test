using System.Globalization;
using System.Text;
using NPOI.SS.UserModel;

namespace InsureFlow.Core.Excel;

/// <summary>엑셀 첫 시트를 값 격자(0-based)로 읽어 둔 것. .xls/.xlsx 공통.</summary>
public sealed class Grid
{
    private readonly object?[][] _rows;

    public string SheetName { get; }
    public int RowCount => _rows.Length;

    private Grid(string sheetName, object?[][] rows)
    {
        SheetName = sheetName;
        _rows = rows;
    }

    public int ColCount(int r) => _rows[r].Length;

    public string Text(int r, int c)
    {
        if (r < 0 || r >= _rows.Length || c < 0 || c >= _rows[r].Length) return "";
        return _rows[r][c] switch
        {
            null => "",
            decimal d => d.ToString("0.############", CultureInfo.InvariantCulture),
            var o => (o.ToString() ?? "").Trim(),
        };
    }

    public decimal? Num(int r, int c)
    {
        if (r < 0 || r >= _rows.Length || c < 0 || c >= _rows[r].Length) return null;
        if (_rows[r][c] is decimal d) return d;
        return ParseMoney(Text(r, c));
    }

    public static decimal? ParseMoney(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        var t = s.Replace(",", "").Replace(" ", "").Trim();
        return decimal.TryParse(t, NumberStyles.Number | NumberStyles.AllowLeadingSign,
            CultureInfo.InvariantCulture, out var v) ? v : null;
    }

    public static Grid Load(string path)
    {
        try
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var wb = WorkbookFactory.Create(fs);
            var sheet = wb.GetSheetAt(0);
            var last = sheet.LastRowNum;
            var rows = new object?[last + 1][];
            for (var r = 0; r <= last; r++)
            {
                var row = sheet.GetRow(r);
                if (row == null) { rows[r] = Array.Empty<object?>(); continue; }
                var n = Math.Max(row.LastCellNum, (short)0);
                var vals = new object?[n];
                for (var c = 0; c < n; c++) vals[c] = CellValue(row.GetCell(c));
                rows[r] = vals;
            }
            return new Grid(sheet.SheetName, rows);
        }
        catch (IOException ex)
        {
            throw new InsureFlowException($"파일을 열 수 없습니다. 엑셀 등에서 열려 있으면 닫은 뒤 다시 시도하세요.\n{Path.GetFileName(path)} ({ex.Message})");
        }
        catch (Exception ex) when (ex is not InsureFlowException)
        {
            throw new InsureFlowException($"엑셀 파일 형식을 읽을 수 없습니다: {Path.GetFileName(path)} ({ex.Message})");
        }
    }

    internal static object? CellValue(ICell? cell)
    {
        if (cell == null) return null;
        var type = cell.CellType == CellType.Formula ? cell.CachedFormulaResultType : cell.CellType;
        return type switch
        {
            CellType.Numeric => (decimal)cell.NumericCellValue,
            CellType.String => cell.StringCellValue,
            CellType.Boolean => cell.BooleanCellValue ? "TRUE" : "FALSE",
            _ => null,
        };
    }

    /// <summary>공백(NBSP·전각 포함) 제거 + NFC 정규화한 이름 키.</summary>
    public static string NameKey(string? name)
    {
        if (string.IsNullOrEmpty(name)) return "";
        var sb = new StringBuilder();
        foreach (var ch in name.Normalize(NormalizationForm.FormC))
            if (!char.IsWhiteSpace(ch) && ch != ' ' && ch != '　') sb.Append(ch);
        return sb.ToString();
    }

    public static string RrnKey(string? rrn)
    {
        if (string.IsNullOrEmpty(rrn)) return "";
        return new string(rrn.Where(char.IsDigit).ToArray());
    }

    /// <summary>화면·보고서용 마스킹 (생년월일+성별 1자리만 노출).</summary>
    public static string MaskRrn(string rrnKey) =>
        rrnKey.Length >= 7 ? $"{rrnKey[..6]}-{rrnKey[6]}******" : "(없음)";
}

public sealed class InsureFlowException : Exception
{
    public InsureFlowException(string message) : base(message) { }
}
