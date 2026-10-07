using InsureFlow.Core.Model;
using InsureFlow.Core.Processing;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace InsureFlow.Core.Writers;

/// <summary>보험별 합계 검증 보고서(.xlsx). 주민번호는 기록하지 않는다.</summary>
public static class ReportWriter
{
    public static void Save(string path, ProcessResult r, string outputXlsName, DateTime now, IEnumerable<string>? verifyDiffs = null)
    {
        using var wb = new XSSFWorkbook();
        var head = wb.CreateCellStyle();
        var bold = wb.CreateFont(); bold.IsBold = true; head.SetFont(bold);
        head.FillForegroundColor = IndexedColors.Grey25Percent.Index;
        head.FillPattern = FillPattern.SolidForeground;
        var money = wb.CreateCellStyle();
        money.DataFormat = wb.CreateDataFormat().GetFormat("#,##0;[Red]-#,##0");
        var title = wb.CreateCellStyle();
        var tf = wb.CreateFont(); tf.IsBold = true; tf.FontHeightInPoints = 14; title.SetFont(tf);

        // 1) 합계 검증
        var s1 = wb.CreateSheet("합계 검증");
        var row = 0;
        Text(s1, row++, 0, "InsureFlow 보험별 합계 검증 보고서", title);
        Text(s1, row++, 0, $"생성: {now:yyyy-MM-dd HH:mm}   양식: {r.Target.SheetName}   저장 파일: {outputXlsName}");
        Text(s1, row++, 0, r.ReconOk ? "결과: 모든 보험 합계가 일치합니다." : "결과: 차이가 있는 항목이 있습니다. 아래 사유를 확인하세요.");
        foreach (var w in r.Warnings) Text(s1, row++, 0, "경고: " + w);
        row++;
        Header(s1, row++, head, "보험", "구분", "원본 근로자 부담 합계", "양식 입력 합계", "대상 외 직원분", "미반영(확인 필요)분", "차이", "결과");
        foreach (var l in r.Recon)
        {
            var rr = s1.CreateRow(row++);
            rr.CreateCell(0).SetCellValue(FieldInfo.Label(FieldInfo.Ins(l.Field)));
            rr.CreateCell(1).SetCellValue(FieldInfo.IsAdjust(l.Field) ? "정산분" : "당월분");
            Num(rr, 2, l.SourceTotal, money); Num(rr, 3, l.Entered, money); Num(rr, 4, l.OutOfTarget, money);
            Num(rr, 5, l.Unreflected, money); Num(rr, 6, l.Diff, money);
            rr.CreateCell(7).SetCellValue(l.IsOk ? "일치" : "차이");
        }
        Text(s1, row++, 0, "원본 합계 = 입력 합계 + 대상 외 직원분 + 미반영분 (차이 0이면 일치)");
        row++;
        if (r.References.Count > 0)
        {
            Header(s1, row++, head, "참고 대조", "", "집계값", "원본 기준값", "", "", "", "결과");
            foreach (var x in r.References)
            {
                var rr = s1.CreateRow(row++);
                rr.CreateCell(0).SetCellValue(x.Label);
                Num(rr, 2, x.Actual, money);
                if (x.Expected is { } e) Num(rr, 3, e, money);
                if (x.IsOk is { } ok) rr.CreateCell(7).SetCellValue(ok ? "일치" : "차이");
            }
        }
        if (verifyDiffs != null)
        {
            row++;
            var list = verifyDiffs.ToList();
            Text(s1, row++, 0, list.Count == 0 ? "저장 파일 재검증: 입력 칸 외 값·서식 변경 없음" : $"저장 파일 재검증: 차이 {list.Count}건");
            foreach (var d in list.Take(50)) Text(s1, row++, 0, "  " + d);
        }
        s1.SetColumnWidth(0, 52 * 256); s1.SetColumnWidth(1, 10 * 256);
        for (var c = 2; c <= 7; c++) s1.SetColumnWidth(c, 18 * 256);

        // 2) 확인 대상
        var s2 = wb.CreateSheet("확인 대상");
        Header(s2, 0, head, "사번", "성명", "사유");
        var n = 1;
        foreach (var e in r.Employees.Where(x => x.NeedsCheck))
        {
            var reasons = e.Issues.Count > 0 ? e.Issues : new List<string> { "확인 필요" };
            foreach (var reason in reasons)
            {
                var rr = s2.CreateRow(n++);
                rr.CreateCell(0).SetCellValue(e.Target.EmpNo);
                rr.CreateCell(1).SetCellValue(e.Target.Name);
                rr.CreateCell(2).SetCellValue(reason);
            }
        }
        s2.SetColumnWidth(0, 14 * 256); s2.SetColumnWidth(1, 12 * 256); s2.SetColumnWidth(2, 80 * 256);

        // 3) 대상 외 원본 직원
        var s3 = wb.CreateSheet("대상 외·미반영 원본");
        Header(s3, 0, head, "원본", "성명", "원본 행", "사유");
        n = 1;
        foreach (var x in r.Extras)
        {
            var rr = s3.CreateRow(n++);
            rr.CreateCell(0).SetCellValue(x.Source);
            rr.CreateCell(1).SetCellValue(x.Name);
            rr.CreateCell(2).SetCellValue(x.SheetRow);
            rr.CreateCell(3).SetCellValue(x.Reason);
        }
        s3.SetColumnWidth(0, 16 * 256); s3.SetColumnWidth(1, 12 * 256); s3.SetColumnWidth(3, 50 * 256);

        // 4) 입력 내역
        var s4 = wb.CreateSheet("입력 내역");
        var cols = new List<string> { "사번", "성명" };
        cols.AddRange(FieldInfo.All.Select(FieldInfo.Label));
        cols.Add("확인");
        Header(s4, 0, head, cols.ToArray());
        n = 1;
        foreach (var e in r.Employees)
        {
            var rr = s4.CreateRow(n++);
            rr.CreateCell(0).SetCellValue(e.Target.EmpNo);
            rr.CreateCell(1).SetCellValue(e.Target.Name);
            var c = 2;
            foreach (var f in FieldInfo.All)
            {
                if (e.Fields.TryGetValue(f, out var fr) && fr.Value is { } v) Num(rr, c, v, money);
                c++;
            }
            rr.CreateCell(c).SetCellValue(e.NeedsCheck ? "확인 대상" : "");
        }
        for (var c = 0; c < cols.Count; c++) s4.SetColumnWidth(c, (c < 2 ? 12 : 14) * 256);

        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
        wb.Write(fs, false);
    }

    private static void Text(ISheet s, int r, int c, string text, ICellStyle? style = null)
    {
        var cell = (s.GetRow(r) ?? s.CreateRow(r)).CreateCell(c);
        cell.SetCellValue(text);
        if (style != null) cell.CellStyle = style;
    }

    private static void Header(ISheet s, int r, ICellStyle style, params string[] names)
    {
        var row = s.CreateRow(r);
        for (var i = 0; i < names.Length; i++)
        {
            var c = row.CreateCell(i);
            c.SetCellValue(names[i]);
            c.CellStyle = style;
        }
    }

    private static void Num(IRow row, int c, decimal v, ICellStyle style)
    {
        var cell = row.CreateCell(c);
        cell.SetCellValue((double)v);
        cell.CellStyle = style;
    }
}
