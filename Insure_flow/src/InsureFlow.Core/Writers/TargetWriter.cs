using InsureFlow.Core.Excel;
using InsureFlow.Core.Model;
using InsureFlow.Core.Processing;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace InsureFlow.Core.Writers;

public static class TargetWriter
{
    /// <summary>대상 .xls를 열어 입력 칸만 채운 뒤 새 파일로 저장한다. 원본 파일은 변경하지 않는다.</summary>
    /// <returns>실제로 값을 쓴 칸(행, 열) 목록.</returns>
    public static List<(int Row, int Col, decimal Value)> Save(string sourcePath, string outPath, ProcessResult result)
    {
        if (string.Equals(Path.GetFullPath(sourcePath), Path.GetFullPath(outPath), StringComparison.OrdinalIgnoreCase))
            throw new InsureFlowException("원본 파일과 같은 경로에는 저장할 수 없습니다. 다른 이름을 지정하세요.");
        if (!string.Equals(Path.GetExtension(outPath), ".xls", StringComparison.OrdinalIgnoreCase))
            throw new InsureFlowException("저장 파일 확장자는 .xls여야 합니다.");

        HSSFWorkbook wb;
        try
        {
            using var fs = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            wb = new HSSFWorkbook(fs);
        }
        catch (IOException ex)
        {
            throw new InsureFlowException($"대상 양식을 열 수 없습니다. 엑셀에서 열려 있으면 닫아 주세요. ({ex.Message})");
        }
        catch (Exception ex) when (ex is not InsureFlowException)
        {
            throw new InsureFlowException($"대상 양식은 .xls(엑셀 97-2003) 형식이어야 합니다. ({ex.Message})");
        }

        var written = new List<(int, int, decimal)>();
        using (wb)
        {
            var sheet = wb.GetSheetAt(0);
            foreach (var emp in result.Employees)
            {
                var row = sheet.GetRow(emp.Target.SheetRow) ?? sheet.CreateRow(emp.Target.SheetRow);
                foreach (var f in FieldInfo.All)
                {
                    if (!emp.Fields.TryGetValue(f, out var fr) || fr.Value is not { } value) continue;
                    var col = result.Target.Columns[f];
                    var cell = row.GetCell(col);
                    if (cell == null)
                    {
                        cell = row.CreateCell(col);
                        var style = NeighborStyle(sheet, emp.Target.SheetRow, col);
                        if (style != null) cell.CellStyle = style;
                    }
                    cell.SetCellValue((double)value);
                    written.Add((emp.Target.SheetRow, col, value));
                }
            }

            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outPath))!);
            try
            {
                using var os = new FileStream(outPath, FileMode.Create, FileAccess.Write, FileShare.None);
                wb.Write(os, false);
            }
            catch (IOException ex)
            {
                throw new InsureFlowException($"파일을 저장할 수 없습니다. 같은 이름의 파일이 열려 있는지 확인하세요. ({ex.Message})");
            }
        }
        return written;
    }

    private static ICellStyle? NeighborStyle(ISheet sheet, int rowIdx, int col)
    {
        var colStyle = sheet.GetColumnStyle(col);
        if (colStyle != null) return colStyle;
        for (var d = 1; d < sheet.LastRowNum; d++)
            foreach (var r in new[] { rowIdx - d, rowIdx + d })
                if (r >= 0 && sheet.GetRow(r)?.GetCell(col) is { } c) return c.CellStyle;
        return null;
    }
}

/// <summary>저장된 파일을 다시 열어 원본 양식과 비교한다. 입력한 칸 외에는 값·서식 차이가 없어야 한다.</summary>
public static class TargetVerifier
{
    public static List<string> Compare(string sourcePath, string outPath, IEnumerable<(int Row, int Col, decimal Value)> written)
    {
        var diffs = new List<string>();
        var expected = written.ToDictionary(w => (w.Row, w.Col), w => w.Value);

        using var fa = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var fb = new FileStream(outPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var a = new HSSFWorkbook(fa);
        using var b = new HSSFWorkbook(fb);

        if (a.NumberOfSheets != b.NumberOfSheets) diffs.Add($"시트 수 다름: {a.NumberOfSheets} → {b.NumberOfSheets}");
        for (var s = 0; s < Math.Min(a.NumberOfSheets, b.NumberOfSheets); s++)
        {
            var sa = a.GetSheetAt(s); var sb = b.GetSheetAt(s);
            if (sa.SheetName != sb.SheetName) diffs.Add($"시트 이름 다름: {sa.SheetName} → {sb.SheetName}");
            if (sa.NumMergedRegions != sb.NumMergedRegions) diffs.Add($"[{sa.SheetName}] 병합 영역 수 다름");
            if (sa.LastRowNum != sb.LastRowNum) diffs.Add($"[{sa.SheetName}] 마지막 행 다름: {sa.LastRowNum} → {sb.LastRowNum}");

            var maxCol = 0;
            for (var r = 0; r <= sa.LastRowNum; r++) maxCol = Math.Max(maxCol, sa.GetRow(r)?.LastCellNum ?? 0);
            for (var c = 0; c < maxCol; c++)
                if (sa.GetColumnWidth(c) != sb.GetColumnWidth(c)) diffs.Add($"[{sa.SheetName}] 열 {c + 1} 너비 다름");

            for (var r = 0; r <= sa.LastRowNum; r++)
            {
                var ra = sa.GetRow(r); var rb = sb.GetRow(r);
                if ((ra == null) != (rb == null)) { diffs.Add($"행 {r + 1} 존재 여부 다름"); continue; }
                if (ra == null) continue;
                if (ra.HeightInPoints != rb!.HeightInPoints) diffs.Add($"행 {r + 1} 높이 다름");
                var n = Math.Max(ra.LastCellNum, rb.LastCellNum);
                for (var c = 0; c < n; c++)
                {
                    var ca = ra.GetCell(c); var cb = rb.GetCell(c);
                    var va = Grid.CellValue(ca); var vb = Grid.CellValue(cb);
                    var addr = $"{ColName(c)}{r + 1}";
                    if (s == 0 && expected.TryGetValue((r, c), out var ev))
                    {
                        if (vb is not decimal d || d != ev) diffs.Add($"{addr}: 입력값 불일치 (기대 {ev}, 실제 {vb})");
                    }
                    else if (!Equals(va, vb)) diffs.Add($"{addr}: 값 변경됨 ({va} → {vb})");

                    var sta = StyleKey(ca); var stb = StyleKey(cb);
                    if (sta != stb && !(s == 0 && expected.ContainsKey((r, c)) && ca == null))
                        diffs.Add($"{addr}: 서식 변경됨");
                }
            }
        }
        return diffs;
    }

    private static string ColName(int c)
    {
        var s = "";
        for (c++; c > 0; c = (c - 1) / 26) s = (char)('A' + (c - 1) % 26) + s;
        return s;
    }

    private static string StyleKey(ICell? c)
    {
        if (c == null) return "-";
        var st = c.CellStyle;
        if (st == null) return "none";
        var f = st.GetFont(c.Sheet.Workbook);
        return string.Join("|", st.GetDataFormatString(), st.Alignment, st.VerticalAlignment, st.BorderTop, st.BorderBottom,
            st.BorderLeft, st.BorderRight, st.FillPattern, st.FillForegroundColor, st.WrapText, st.IsLocked,
            f.FontName, f.FontHeight, f.IsBold, f.IsItalic, f.Color);
    }
}
