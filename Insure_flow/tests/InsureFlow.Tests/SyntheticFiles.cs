using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace InsureFlow.Tests;

/// <summary>개인정보 없는 합성 입력 파일 생성기(실제 EDI 파일과 같은 머리글 구조).</summary>
internal sealed class SyntheticFiles : IDisposable
{
    public string Dir { get; } = Path.Combine(Path.GetTempPath(), "insureflow_" + Guid.NewGuid().ToString("N"));
    public SyntheticFiles() => Directory.CreateDirectory(Dir);
    public void Dispose() { try { Directory.Delete(Dir, true); } catch { } }

    public static readonly string[] TargetHeaders =
    {
        "사번", "성명", "국민연금[03]", "건강보험[04]", "장기요양보험[22]", "고용보험[05]", "소득세[01]", "지방소득세[02]",
        "근로소득소득세정산[07]", "근로소득지방소득세정산[08]", "국민연금정산[09]", "건강보험정산[10]", "장기요양보험정산[23]",
        "고용보험정산[25]", "선급금공제[30]", "학자금공제[31]", "기부금공제[32]", "특수공제[06]", "농특세[11]", "근로소득농특세정산[24]",
    };

    /// <summary>employees: (사번, 성명). extra: (행 인덱스, 열 인덱스) → 이미 들어 있는 값.</summary>
    public string Target(IEnumerable<(string No, string Name)> employees, Dictionary<(int Row, int Col), double>? extra = null)
    {
        var path = Path.Combine(Dir, "target.xls");
        using var wb = new HSSFWorkbook();
        var sh = wb.CreateSheet("(2026-09) 공제 근로자별 변동자료");
        var money = wb.CreateCellStyle();
        money.DataFormat = wb.CreateDataFormat().GetFormat("#,##0");
        var h = sh.CreateRow(0);
        for (var c = 0; c < TargetHeaders.Length; c++) h.CreateCell(c).SetCellValue(TargetHeaders[c]);
        var r = 1;
        foreach (var (no, name) in employees)
        {
            var row = sh.CreateRow(r++);
            row.CreateCell(0).SetCellValue(no);
            row.CreateCell(1).SetCellValue(name);
            for (var c = 2; c < TargetHeaders.Length; c++) row.CreateCell(c).CellStyle = money;
        }
        if (extra != null)
            foreach (var ((rr, cc), v) in extra) sh.GetRow(rr).GetCell(cc).SetCellValue(v);
        sh.SetColumnWidth(1, 5000);
        using var fs = File.Create(path);
        wb.Write(fs, false);
        return path;
    }

    /// <summary>(성명, 주민번호, 당월 본인, 소급 본인, 상실일)</summary>
    public string Np(IEnumerable<(string Name, string Rrn, double? Cur, double? Retro, string Loss)> rows, string ym = "2026-09")
    {
        var path = Path.Combine(Dir, "np.xlsx");
        using var wb = new XSSFWorkbook();
        var sh = wb.CreateSheet("sheet");
        sh.CreateRow(0).CreateCell(0).SetCellValue("2차결정내역통보서");
        var hdr = new[] { "고지년월", "사업장관리번호", "성명", "주민등록번호", "취득일", "상실일", "당월분_기준소득월액(원)", "당월분_월보험료(원)",
            "당월분_(사용자부담금)(원)", "당월분_(본인기여금(원)", "소급분_변동사유", "소급분_해당기간", "소급분_월수", "소급분_전월이전보험료(원)",
            "소급분_(사용자부담금)(원)", "소급분_(본인기여금)", "이자분(0원)", "총부담금계_(사용자부담금)(원)", "총부담금계_(본인기여금)(원)" };
        var h = sh.CreateRow(2);
        for (var c = 0; c < hdr.Length; c++) h.CreateCell(c).SetCellValue(hdr[c]);
        var r = 3;
        foreach (var x in rows)
        {
            var row = sh.CreateRow(r++);
            row.CreateCell(0).SetCellValue(ym);
            row.CreateCell(2).SetCellValue(x.Name);
            row.CreateCell(3).SetCellValue(x.Rrn);
            row.CreateCell(5).SetCellValue(x.Loss);
            if (x.Cur is { } cur) row.CreateCell(9).SetCellValue(cur);
            if (x.Retro is { } rt) row.CreateCell(15).SetCellValue(rt);
        }
        using var fs = File.Create(path);
        wb.Write(fs, false);
        return path;
    }

    /// <summary>(성명, 주민번호, ①, ②, ③) — 근로자 실업급여만. 합계행 포함.</summary>
    public string Ei(IEnumerable<(string Name, string Rrn, double Cur, double Recalc, double Settle)> rows)
    {
        var path = Path.Combine(Dir, "ei.xlsx");
        using var wb = new XSSFWorkbook();
        var sh = wb.CreateSheet("sheet");
        var g = new[] { "순번", "근로자 구분", "근로자명", "주민등록번호", "근로자원부번호", "고용일", "고용종료일", "휴직자월평", "월평균보수금액",
            "산정보험료(해당월①)", "산정보험료(해당월①)", "재산정보험료(해당년도②)", "재산정보험료(해당년도②)", "정산보수총액", "정산보험료(③)", "정산보험료(③)",
            "보험료합계(①+②+③)", "보험료합계(①+②+③)" };
        var s = new[] { "순번", "근로자 구분", "근로자명", "주민등록번호", "근로자원부번호", "고용일", "고용종료일", "휴직자월평", "월평균보수금액",
            "근로자실업 급여보험료", "사업주실업 급여보험료", "근로자실업 급여보험료", "사업주실업 급여보험료", "정산보수총액", "근로자실업 급여보험료", "사업주실업 급여보험료",
            "근로자실업 급여보험료", "사업주실업 급여보험료" };
        var h1 = sh.CreateRow(0); var h2 = sh.CreateRow(1);
        for (var c = 0; c < g.Length; c++) { h1.CreateCell(c).SetCellValue(g[c]); h2.CreateCell(c).SetCellValue(s[c]); }
        var r = 2; double a = 0, b = 0, d = 0; var n = 1;
        foreach (var x in rows)
        {
            var row = sh.CreateRow(r++);
            row.CreateCell(0).SetCellValue(n++);
            row.CreateCell(2).SetCellValue(x.Name);
            row.CreateCell(3).SetCellValue(x.Rrn);
            row.CreateCell(9).SetCellValue(x.Cur); row.CreateCell(10).SetCellValue(x.Cur);
            row.CreateCell(11).SetCellValue(x.Recalc); row.CreateCell(12).SetCellValue(0);
            row.CreateCell(14).SetCellValue(x.Settle); row.CreateCell(15).SetCellValue(0);
            a += x.Cur; b += x.Recalc; d += x.Settle;
        }
        var t = sh.CreateRow(r);
        t.CreateCell(0).SetCellValue("합계"); t.CreateCell(1).SetCellValue("합계"); t.CreateCell(2).SetCellValue("합계");
        t.CreateCell(9).SetCellValue(a); t.CreateCell(11).SetCellValue(b); t.CreateCell(14).SetCellValue(d);
        using var fs = File.Create(path);
        wb.Write(fs, false);
        return path;
    }

    /// <summary>(성명, 주민번호, 건강산출, 건강정산, 건강연말, 요양산출, 요양정산, 요양연말) — 총액 기준.</summary>
    public string Hi(IEnumerable<(string Name, string Rrn, double HCur, double HSet, double HYe, double LCur, double LSet, double LYe)> rows, string ym = "202609")
    {
        var path = Path.Combine(Dir, "hi.xls");
        using var wb = new HSSFWorkbook();
        var sh = wb.CreateSheet("Sheet1");
        var hdr = new[] { "고지년월", "사업장관리번호", "성명", "주민등록번호", "산출보험료", "정산금액", "고지금액", "연말정산",
            "요양산출보험료", "요양정산보험료", "요양고지보험료", "요양연말정산보험료", "산출보험료계(건강+요양)", "정산보험료계(건강+요양)" };
        var h = sh.CreateRow(0);
        for (var c = 0; c < hdr.Length; c++) h.CreateCell(c).SetCellValue(hdr[c]);
        var r = 1;
        foreach (var x in rows)
        {
            var row = sh.CreateRow(r++);
            row.CreateCell(0).SetCellValue(ym);
            row.CreateCell(2).SetCellValue(x.Name);
            row.CreateCell(3).SetCellValue(x.Rrn);
            row.CreateCell(4).SetCellValue(x.HCur); row.CreateCell(5).SetCellValue(x.HSet);
            row.CreateCell(6).SetCellValue(x.HCur + x.HSet); row.CreateCell(7).SetCellValue(x.HYe);
            row.CreateCell(8).SetCellValue(x.LCur); row.CreateCell(9).SetCellValue(x.LSet);
            row.CreateCell(10).SetCellValue(x.LCur + x.LSet); row.CreateCell(11).SetCellValue(x.LYe);
        }
        using var fs = File.Create(path);
        wb.Write(fs, false);
        return path;
    }
}
