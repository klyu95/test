using System.Globalization;
using InsureFlow.Core.Calc;
using InsureFlow.Core.Excel;
using InsureFlow.Core.Model;
using InsureFlow.Core.Readers;

namespace InsureFlow.Core.Processing;

/// <summary>읽어 둔 4개 자료를 이름으로 매칭하고 8개 칸의 입력값과 합계 검증을 만든다. 순수 계산이라 사용자 선택이 바뀔 때마다 다시 실행해도 된다.</summary>
public sealed class Processor
{
    private const string NpLabel = "국민연금", EiLabel = "고용보험", HiLabel = "건강·요양보험";

    private readonly TargetData _target;
    private readonly SourceData<NpRecord> _np;
    private readonly SourceData<EiRecord> _ei;
    private readonly EiTotals? _eiTotals;
    private readonly SourceData<HiRecord> _hi;

    public Processor(TargetData target, SourceData<NpRecord> np, SourceData<EiRecord> ei, EiTotals? eiTotals, SourceData<HiRecord> hi)
    {
        _target = target; _np = np; _ei = ei; _eiTotals = eiTotals; _hi = hi;
    }

    /// <param name="rrnChoices">동명이인 등 확인이 필요한 직원에 대해 사용자가 고른 주민번호 키(대상 행 번호 → 키).</param>
    public ProcessResult Run(IReadOnlyDictionary<int, string>? rrnChoices = null)
    {
        foreach (var r in _np.Rows) r.Consumed = false;
        foreach (var r in _ei.Rows) r.Consumed = false;
        foreach (var r in _hi.Rows) r.Consumed = false;

        var result = new ProcessResult { Target = _target };
        AddMonthWarnings(result);

        var npBy = _np.Rows.ToLookup(r => r.NameKey);
        var eiBy = _ei.Rows.ToLookup(r => r.NameKey);
        var hiBy = _hi.Rows.ToLookup(r => r.NameKey);
        var targetNameCount = _target.Employees.GroupBy(e => e.NameKey).ToDictionary(g => g.Key, g => g.Count());

        foreach (var t in _target.Employees)
        {
            var emp = new EmployeeResult { Target = t };
            result.Employees.Add(emp);

            var np = npBy[t.NameKey].ToList();
            var ei = eiBy[t.NameKey].ToList();
            var hi = hiBy[t.NameKey].ToList();

            string? chosen = null;
            rrnChoices?.TryGetValue(t.SheetRow, out chosen);
            if (chosen != null)
            {
                np = np.Where(r => r.RrnKey == chosen).ToList();
                ei = ei.Where(r => r.RrnKey == chosen).ToList();
                hi = hi.Where(r => r.RrnKey == chosen).ToList();
                emp.ChosenRrn = chosen;
            }

            var dupTarget = targetNameCount[t.NameKey] > 1;
            var multi = np.Count > 1 || ei.Count > 1 || hi.Count > 1;
            var rrns = np.Select(r => r.RrnKey).Concat(ei.Select(r => r.RrnKey)).Concat(hi.Select(r => r.RrnKey))
                .Where(k => k.Length > 0).Distinct().ToList();
            var mismatch = rrns.Count > 1;

            if (chosen == null && (dupTarget || multi || mismatch))
            {
                emp.Match = MatchState.Ambiguous;
                emp.Issues.Add(dupTarget ? "대상 양식에 같은 이름이 2명 이상 있습니다(동명이인)."
                    : multi ? "원본에 같은 이름이 2건 이상 있습니다(동명이인)."
                    : "원본 간 주민번호가 서로 달라 동일인인지 확인이 필요합니다.");
                BuildCandidates(emp, npBy[t.NameKey], eiBy[t.NameKey], hiBy[t.NameKey]);
                foreach (var f in FieldInfo.All)
                    emp.Fields[f] = new FieldResult { State = CellState.NeedsCheck, Note = "동명이인/불일치 — 자동 입력 안 함" };
                continue;
            }

            emp.Match = MatchState.Matched;
            if (chosen != null)
                BuildCandidates(emp, npBy[t.NameKey], eiBy[t.NameKey], hiBy[t.NameKey]);

            FillNp(emp, Take(np));
            FillEi(emp, Take(ei));
            FillHi(emp, Take(hi));
            WarnOverwrite(emp);
        }

        BuildRecon(result);
        return result;
    }

    private static T? Take<T>(List<T> rows) where T : SourceRecord
    {
        var r = rows.FirstOrDefault(x => !x.Consumed);
        if (r != null) r.Consumed = true;
        return r;
    }

    // ---------- 칸 채우기 ----------

    private static FieldResult Ok(decimal v, string note) => new() { Value = v, State = CellState.Ok, Note = note };
    private static FieldResult Empty(string note) => new() { State = CellState.Empty, Note = note };
    private static FieldResult Check(string note) => new() { State = CellState.NeedsCheck, Note = note };

    private static void Missing(EmployeeResult emp, string source, params FieldId[] fields)
    {
        var msg = $"{source} 자료에 없음";
        emp.Issues.Add(msg);
        foreach (var f in fields) emp.Fields[f] = Check(msg);
    }

    private static void FillNp(EmployeeResult emp, NpRecord? r)
    {
        if (r == null) { Missing(emp, NpLabel, FieldId.NpCurrent, FieldId.NpAdjust); return; }
        if (r.Current is { } cur)
            emp.Fields[FieldId.NpCurrent] = Ok(cur, $"당월분 본인기여금 {cur:N0}");
        else
        {
            var msg = r.LossDate.Length > 0 ? $"{NpLabel} 금액 공란(상실일 {r.LossDate})" : $"{NpLabel} 당월분 금액 공란";
            emp.Issues.Add(msg);
            emp.Fields[FieldId.NpCurrent] = Check(msg);
        }
        emp.Fields[FieldId.NpAdjust] = r.Retro is { } retro
            ? Ok(retro, $"소급분 본인기여금 {retro:N0}")
            : Empty("소급분 없음");
    }

    private static void FillEi(EmployeeResult emp, EiRecord? r)
    {
        if (r == null) { Missing(emp, EiLabel, FieldId.EiCurrent, FieldId.EiAdjust); return; }
        emp.Fields[FieldId.EiCurrent] = Ok(r.Current, $"해당월① 근로자 {r.Current:N0}");
        var adj = r.Recalc + r.Settle;
        emp.Fields[FieldId.EiAdjust] = Ok(adj, $"재산정② {r.Recalc:N0} + 정산③ {r.Settle:N0}");
    }

    private static void FillHi(EmployeeResult emp, HiRecord? r)
    {
        if (r == null)
        {
            Missing(emp, HiLabel, FieldId.HiCurrent, FieldId.HiAdjust, FieldId.LtcCurrent, FieldId.LtcAdjust);
            return;
        }
        FieldResult Share(decimal total, string what)
        {
            var s = ShareCalculator.EmployeeShare(total);
            return Ok(s, $"{what} 총액 {total:N0} ÷2 → 10원 미만 절사 {s:N0}");
        }
        emp.Fields[FieldId.HiCurrent] = Share(r.HealthCurrent, "건강 산출");
        emp.Fields[FieldId.HiAdjust] = Share(r.HealthAdjust, "건강 정산+연말정산");
        emp.Fields[FieldId.LtcCurrent] = Share(r.LtcCurrent, "요양 산출");
        emp.Fields[FieldId.LtcAdjust] = Share(r.LtcAdjust, "요양 정산+연말정산");
    }

    private static void WarnOverwrite(EmployeeResult emp)
    {
        foreach (var f in FieldInfo.All)
        {
            if (!emp.Fields.TryGetValue(f, out var fr) || fr.Value == null) continue;
            if (!emp.Target.Existing.TryGetValue(f, out var old) || old == fr.Value) continue;
            var msg = $"{FieldInfo.Label(f)}: 양식에 있던 값 {old:N0}을(를) {fr.Value:N0}(으)로 덮어씀";
            emp.Issues.Add(msg);
            emp.Fields[f] = new FieldResult { Value = fr.Value, State = CellState.NeedsCheck, Note = msg };
        }
    }

    private static void BuildCandidates(EmployeeResult emp, IEnumerable<NpRecord> np, IEnumerable<EiRecord> ei, IEnumerable<HiRecord> hi)
    {
        var map = new Dictionary<string, List<string>>();
        void Add(string key, string src)
        {
            if (!map.TryGetValue(key, out var l)) map[key] = l = new List<string>();
            if (!l.Contains(src)) l.Add(src);
        }
        foreach (var r in np) Add(r.RrnKey, NpLabel);
        foreach (var r in ei) Add(r.RrnKey, EiLabel);
        foreach (var r in hi) Add(r.RrnKey, HiLabel);
        emp.Candidates.Clear();
        foreach (var (key, srcs) in map.Where(kv => kv.Key.Length > 0))
            emp.Candidates.Add(new RrnCandidate { RrnKey = key, Masked = Grid.MaskRrn(key), Sources = string.Join(", ", srcs) });
    }

    // ---------- 합계 검증 ----------

    private static decimal Amount(FieldId f, SourceRecord r) => (f, r) switch
    {
        (FieldId.NpCurrent, NpRecord n) => n.Current ?? 0,
        (FieldId.NpAdjust, NpRecord n) => n.Retro ?? 0,
        (FieldId.EiCurrent, EiRecord e) => e.Current,
        (FieldId.EiAdjust, EiRecord e) => e.Recalc + e.Settle,
        (FieldId.HiCurrent, HiRecord h) => ShareCalculator.EmployeeShare(h.HealthCurrent),
        (FieldId.HiAdjust, HiRecord h) => ShareCalculator.EmployeeShare(h.HealthAdjust),
        (FieldId.LtcCurrent, HiRecord h) => ShareCalculator.EmployeeShare(h.LtcCurrent),
        (FieldId.LtcAdjust, HiRecord h) => ShareCalculator.EmployeeShare(h.LtcAdjust),
        _ => 0,
    };

    private IEnumerable<SourceRecord> RowsOf(FieldId f) => FieldInfo.Ins(f) switch
    {
        Insurance.NationalPension => _np.Rows,
        Insurance.Employment => _ei.Rows,
        _ => _hi.Rows,
    };

    private void BuildRecon(ProcessResult result)
    {
        var targetKeys = _target.Employees.Select(e => e.NameKey).ToHashSet();

        foreach (var f in FieldInfo.All)
        {
            decimal source = 0, outOf = 0, unref = 0;
            foreach (var r in RowsOf(f))
            {
                var a = Amount(f, r);
                source += a;
                if (r.Consumed) continue;
                if (targetKeys.Contains(r.NameKey)) unref += a; else outOf += a;
            }
            var entered = result.Employees.Sum(e => e.Fields.TryGetValue(f, out var fr) ? fr.Value ?? 0 : 0);
            result.Recon.Add(new ReconLine
            {
                Field = f, SourceTotal = source, Entered = entered, OutOfTarget = outOf, Unreflected = unref,
            });
        }

        // 대상 외 / 미반영 원본 직원 목록
        void Extras<T>(string label, SourceData<T> d) where T : SourceRecord
        {
            foreach (var r in d.Rows.Where(x => !x.Consumed))
                result.Extras.Add(new ExtraRecord
                {
                    Source = label, Name = r.Name, SheetRow = r.SheetRow,
                    Reason = targetKeys.Contains(r.NameKey) ? "미반영(동명이인 등 확정 필요)" : "대상 양식에 없는 직원(추가하지 않음)",
                });
        }
        Extras(NpLabel, _np); Extras(EiLabel, _ei); Extras(HiLabel, _hi);

        // 원본 자체 합계 대조(고용보험 '합계' 행)
        if (_eiTotals != null)
        {
            void Ref(string label, decimal actual, decimal? expected) =>
                result.References.Add(new ReferenceLine { Label = label, Actual = actual, Expected = expected });
            Ref("고용보험 원본: 해당월① 근로자 (행 합산 vs '합계' 행)", _ei.Rows.Sum(r => r.Current), _eiTotals.Current);
            Ref("고용보험 원본: 재산정② 근로자 (행 합산 vs '합계' 행)", _ei.Rows.Sum(r => r.Recalc), _eiTotals.Recalc);
            Ref("고용보험 원본: 정산③ 근로자 (행 합산 vs '합계' 행)", _ei.Rows.Sum(r => r.Settle), _eiTotals.Settle);
        }

        // 건강·요양: 총액과 근로자/사업주 몫 정보
        void Info(string label, decimal total, Func<HiRecord, decimal> sel)
        {
            var emp = _hi.Rows.Sum(r => ShareCalculator.EmployeeShare(sel(r)));
            result.References.Add(new ReferenceLine { Label = $"{label} 총액 합계(원본)", Actual = total });
            result.References.Add(new ReferenceLine { Label = $"{label} 근로자 몫 합계(÷2, 10원 미만 절사)", Actual = emp });
            result.References.Add(new ReferenceLine { Label = $"{label} 사업주 몫 합계", Actual = total - emp });
        }
        Info("건강 당월분", _hi.Rows.Sum(r => r.HealthCurrent), r => r.HealthCurrent);
        Info("건강 정산분", _hi.Rows.Sum(r => r.HealthAdjust), r => r.HealthAdjust);
        Info("요양 당월분", _hi.Rows.Sum(r => r.LtcCurrent), r => r.LtcCurrent);
        Info("요양 정산분", _hi.Rows.Sum(r => r.LtcAdjust), r => r.LtcAdjust);
    }

    private void AddMonthWarnings(ProcessResult result)
    {
        void Check(string label, string? ym)
        {
            if (_target.YearMonth != null && ym != null && ym != _target.YearMonth)
                result.Warnings.Add($"{label} 고지년월({Fmt(ym)})이 대상 양식 월({Fmt(_target.YearMonth)})과 다릅니다.");
        }
        Check(NpLabel, _np.YearMonth);
        Check(HiLabel, _hi.YearMonth);
    }

    private static string Fmt(string ym) => $"{ym[..4]}-{ym[4..]}";

    // ---------- 편의: 파일에서 한 번에 읽기 ----------

    public static Processor Load(string targetPath, string npPath, string eiPath, string hiPath)
    {
        var t = TargetReader.Read(targetPath, Grid.Load(targetPath));
        var np = NpReader.Read(npPath, Grid.Load(npPath));
        var ei = EiReader.Read(eiPath, Grid.Load(eiPath));
        var hi = HiReader.Read(hiPath, Grid.Load(hiPath));
        return new Processor(t, np, ei.Data, ei.Totals, hi);
    }

    public SourceData<NpRecord> Np => _np;
    public SourceData<EiRecord> Ei => _ei;
    public SourceData<HiRecord> Hi => _hi;
    public TargetData Target => _target;
}
