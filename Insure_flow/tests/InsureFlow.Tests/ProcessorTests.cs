using InsureFlow.Core;
using InsureFlow.Core.Calc;
using InsureFlow.Core.Excel;
using InsureFlow.Core.Model;
using InsureFlow.Core.Processing;
using InsureFlow.Core.Readers;
using InsureFlow.Core.Writers;
using NPOI.HSSF.UserModel;
using Xunit;

namespace InsureFlow.Tests;

public class ShareCalculatorTests
{
    [Theory]
    [InlineData(1799990, 899990)]   // 899,995 → 10원 미만 절사
    [InlineData(236520, 118260)]
    [InlineData(0, 0)]
    [InlineData(-70960, -35480)]
    [InlineData(-11850, -5920)]     // -5,925 → 0 방향 절사
    [InlineData(8880, 4440)]
    [InlineData(15, 0)]
    public void Employee_share_truncates_toward_zero(int total, int expected) =>
        Assert.Equal(expected, ShareCalculator.EmployeeShare(total));

    [Fact]
    public void Employer_plus_employee_equals_total() =>
        Assert.Equal(1799990m, ShareCalculator.EmployeeShare(1799990m) + ShareCalculator.EmployerShare(1799990m));
}

public class ProcessorTests : IDisposable
{
    private readonly SyntheticFiles _f = new();
    public void Dispose() => _f.Dispose();

    private const string R1 = "900101-1000001", R2 = "900202-2000002", R3 = "900303-1000003";

    private Processor Make(
        (string, string)[] target,
        (string, string, double?, double?, string)[] np,
        (string, string, double, double, double)[] ei,
        (string, string, double, double, double, double, double, double)[] hi,
        Dictionary<(int, int), double>? existing = null)
    {
        var t = _f.Target(target, existing);
        return Processor.Load(t, _f.Np(np), _f.Ei(ei), _f.Hi(hi));
    }

    [Fact]
    public void Fills_eight_fields_with_employee_share_only()
    {
        var p = Make(
            new[] { ("001", "가나다") },
            new[] { ("가나다", R1, (double?)100000, (double?)-5000, "") },
            new[] { ("가나다", R1, 40000.0, 0.0, 1500.0) },
            new[] { ("가나다", R1, 1799990.0, 0.0, 8880.0, 236520.0, -70960.0, 1080.0) });
        var r = p.Run();
        var e = Assert.Single(r.Employees);
        Assert.Equal(MatchState.Matched, e.Match);
        Assert.Equal(100000m, e.Fields[FieldId.NpCurrent].Value);
        Assert.Equal(-5000m, e.Fields[FieldId.NpAdjust].Value);
        Assert.Equal(40000m, e.Fields[FieldId.EiCurrent].Value);
        Assert.Equal(1500m, e.Fields[FieldId.EiAdjust].Value);
        Assert.Equal(899990m, e.Fields[FieldId.HiCurrent].Value);
        Assert.Equal(4440m, e.Fields[FieldId.HiAdjust].Value);        // (0 + 8,880) ÷ 2
        Assert.Equal(118260m, e.Fields[FieldId.LtcCurrent].Value);
        Assert.Equal(-34940m, e.Fields[FieldId.LtcAdjust].Value);     // (-70,960 + 1,080) = -69,880 ÷ 2 → -34,940
        Assert.False(e.NeedsCheck);
        Assert.True(r.ReconOk);
    }

    [Fact]
    public void Employee_missing_in_a_source_is_blank_and_flagged()
    {
        var p = Make(
            new[] { ("001", "가나다") },
            new[] { ("가나다", R1, (double?)100000, (double?)null, "") },
            Array.Empty<(string, string, double, double, double)>(),
            new[] { ("가나다", R1, 1000.0, 0.0, 0.0, 1000.0, 0.0, 0.0) });
        var e = Assert.Single(p.Run().Employees);
        Assert.Null(e.Fields[FieldId.EiCurrent].Value);
        Assert.Equal(CellState.NeedsCheck, e.Fields[FieldId.EiCurrent].State);
        Assert.True(e.NeedsCheck);
        Assert.Equal(CellState.Empty, e.Fields[FieldId.NpAdjust].State);   // 소급 없음은 확인 대상 아님
    }

    [Fact]
    public void Source_only_employee_is_not_added_but_reported()
    {
        var p = Make(
            new[] { ("001", "가나다") },
            new[] { ("가나다", R1, (double?)100, (double?)null, ""), ("라마바", R2, (double?)200, (double?)null, "") },
            new[] { ("가나다", R1, 1.0, 0.0, 0.0) },
            new[] { ("가나다", R1, 10.0, 0.0, 0.0, 10.0, 0.0, 0.0) });
        var r = p.Run();
        Assert.Single(r.Employees);
        Assert.Contains(r.Extras, x => x.Name == "라마바" && x.Source == "국민연금");
        var np = r.Recon.Single(l => l.Field == FieldId.NpCurrent);
        Assert.Equal(300m, np.SourceTotal);
        Assert.Equal(100m, np.Entered);
        Assert.Equal(200m, np.OutOfTarget);
        Assert.True(np.IsOk);
    }

    [Fact]
    public void Duplicate_name_in_source_is_not_auto_filled_until_user_chooses()
    {
        var p = Make(
            new[] { ("001", "동명") },
            new[] { ("동명", R1, (double?)100, (double?)null, ""), ("동명", R2, (double?)300, (double?)null, "") },
            new[] { ("동명", R1, 1.0, 0.0, 0.0), ("동명", R2, 3.0, 0.0, 0.0) },
            new[] { ("동명", R1, 20.0, 0.0, 0.0, 20.0, 0.0, 0.0), ("동명", R2, 60.0, 0.0, 0.0, 60.0, 0.0, 0.0) });
        var e = Assert.Single(p.Run().Employees);
        Assert.Equal(MatchState.Ambiguous, e.Match);
        Assert.All(e.Fields.Values, f => Assert.Null(f.Value));
        Assert.Equal(2, e.Candidates.Count);
        Assert.DoesNotContain(e.Candidates, c => c.Masked.Contains("0000002") || c.Masked.EndsWith("2000002"));

        var chosen = Grid.RrnKey(R2);
        var r2 = p.Run(new Dictionary<int, string> { [e.Target.SheetRow] = chosen });
        var e2 = Assert.Single(r2.Employees);
        Assert.Equal(MatchState.Matched, e2.Match);
        Assert.Equal(300m, e2.Fields[FieldId.NpCurrent].Value);
        Assert.Equal(3m, e2.Fields[FieldId.EiCurrent].Value);
        Assert.Equal(30m, e2.Fields[FieldId.HiCurrent].Value);
        // 고르지 않은 동명이인(R1)은 미반영으로 보고
        Assert.Contains(r2.Extras, x => x.Name == "동명");
        Assert.True(r2.ReconOk);
    }

    [Fact]
    public void Same_name_twice_in_target_is_flagged()
    {
        var p = Make(
            new[] { ("001", "동명"), ("002", "동명") },
            new[] { ("동명", R1, (double?)100, (double?)null, "") },
            new[] { ("동명", R1, 1.0, 0.0, 0.0) },
            new[] { ("동명", R1, 20.0, 0.0, 0.0, 20.0, 0.0, 0.0) });
        var r = p.Run();
        Assert.All(r.Employees, e => Assert.Equal(MatchState.Ambiguous, e.Match));
    }

    [Fact]
    public void Rrn_mismatch_between_sources_is_flagged()
    {
        var p = Make(
            new[] { ("001", "가나다") },
            new[] { ("가나다", R1, (double?)100, (double?)null, "") },
            new[] { ("가나다", R2, 1.0, 0.0, 0.0) },
            new[] { ("가나다", R1, 20.0, 0.0, 0.0, 20.0, 0.0, 0.0) });
        var e = Assert.Single(p.Run().Employees);
        Assert.Equal(MatchState.Ambiguous, e.Match);
    }

    [Fact]
    public void Terminated_member_with_blank_pension_is_flagged_and_settlement_only_in_adjust()
    {
        var p = Make(
            new[] { ("001", "퇴사자") },
            new[] { ("퇴사자", R3, (double?)null, (double?)null, "20260901") },
            new[] { ("퇴사자", R3, 0.0, 0.0, 60930.0) },
            new[] { ("퇴사자", R3, 0.0, 86460.0, 0.0, 0.0, 11320.0, 0.0) });
        var e = Assert.Single(p.Run().Employees);
        Assert.True(e.NeedsCheck);
        Assert.Null(e.Fields[FieldId.NpCurrent].Value);
        Assert.Equal(0m, e.Fields[FieldId.HiCurrent].Value);
        Assert.Equal(43230m, e.Fields[FieldId.HiAdjust].Value);
        Assert.Equal(5660m, e.Fields[FieldId.LtcAdjust].Value);   // 11,320 ÷ 2 = 5,660
        Assert.Equal(60930m, e.Fields[FieldId.EiAdjust].Value);
    }

    [Fact]
    public void Existing_value_in_input_cell_is_flagged_when_overwritten()
    {
        var p = Make(
            new[] { ("001", "가나다") },
            new[] { ("가나다", R1, (double?)100, (double?)null, "") },
            new[] { ("가나다", R1, 1.0, 0.0, 0.0) },
            new[] { ("가나다", R1, 20.0, 0.0, 0.0, 20.0, 0.0, 0.0) },
            new Dictionary<(int, int), double> { [(1, 2)] = 777 });   // 국민연금[03] 열에 기존 값
        var e = Assert.Single(p.Run().Employees);
        Assert.True(e.NeedsCheck);
        Assert.Equal(100m, e.Fields[FieldId.NpCurrent].Value);
    }

    [Fact]
    public void Month_mismatch_produces_warning()
    {
        var t = _f.Target(new[] { ("001", "가나다") });
        var np = _f.Np(new[] { ("가나다", R1, (double?)1, (double?)null, "") }, ym: "2026-08");
        var p = Processor.Load(t, np, _f.Ei(new[] { ("가나다", R1, 1.0, 0.0, 0.0) }),
            _f.Hi(new[] { ("가나다", R1, 10.0, 0.0, 0.0, 10.0, 0.0, 0.0) }));
        Assert.Contains(p.Run().Warnings, w => w.Contains("국민연금"));
    }

    [Fact]
    public void File_detector_identifies_each_kind_and_rejects_wrong_slot()
    {
        var t = _f.Target(new[] { ("001", "가나다") });
        var np = _f.Np(new[] { ("가나다", R1, (double?)1, (double?)null, "") });
        var ei = _f.Ei(new[] { ("가나다", R1, 1.0, 0.0, 0.0) });
        var hi = _f.Hi(new[] { ("가나다", R1, 10.0, 0.0, 0.0, 10.0, 0.0, 0.0) });
        Assert.Equal(FileKind.Target, FileDetector.Detect(Grid.Load(t)));
        Assert.Equal(FileKind.NationalPension, FileDetector.Detect(Grid.Load(np)));
        Assert.Equal(FileKind.Employment, FileDetector.Detect(Grid.Load(ei)));
        Assert.Equal(FileKind.Health, FileDetector.Detect(Grid.Load(hi)));
        Assert.Throws<InsureFlowException>(() => NpReader.Read(hi, Grid.Load(hi)));
    }

    [Fact]
    public void Writer_keeps_other_columns_and_refuses_overwriting_source()
    {
        var existing = new Dictionary<(int, int), double> { [(1, 14)] = 12345, [(1, 6)] = 54321 };   // 선급금공제·소득세
        var p = Make(
            new[] { ("001", "가나다") },
            new[] { ("가나다", R1, (double?)100, (double?)null, "") },
            new[] { ("가나다", R1, 1.0, 0.0, 0.0) },
            new[] { ("가나다", R1, 20.0, 0.0, 0.0, 20.0, 0.0, 0.0) },
            existing);
        var res = p.Run();
        var src = p.Target.Path;
        Assert.Throws<InsureFlowException>(() => TargetWriter.Save(src, src, res));

        var outPath = Path.Combine(_f.Dir, "out.xls");
        var written = TargetWriter.Save(src, outPath, res);
        Assert.Empty(TargetVerifier.Compare(src, outPath, written));

        using var fs = File.OpenRead(outPath);
        using var wb = new HSSFWorkbook(fs);
        var row = wb.GetSheetAt(0).GetRow(1);
        Assert.Equal(12345, row.GetCell(14).NumericCellValue);
        Assert.Equal(54321, row.GetCell(6).NumericCellValue);
        Assert.Equal(100, row.GetCell(2).NumericCellValue);
        Assert.Equal("가나다", row.GetCell(1).StringCellValue);
        Assert.Equal("001", row.GetCell(0).StringCellValue);
    }

    [Fact]
    public void Report_is_written()
    {
        var p = Make(
            new[] { ("001", "가나다") },
            new[] { ("가나다", R1, (double?)100, (double?)null, "") },
            new[] { ("가나다", R1, 1.0, 0.0, 0.0) },
            new[] { ("가나다", R1, 20.0, 0.0, 0.0, 20.0, 0.0, 0.0) });
        var path = Path.Combine(_f.Dir, "report.xlsx");
        ReportWriter.Save(path, p.Run(), "out.xls", new DateTime(2026, 10, 7, 9, 0, 0));
        Assert.True(new FileInfo(path).Length > 1000);
    }
}
