using InsureFlow.Core.Model;

namespace InsureFlow.Core.Processing;

public enum CellState
{
    Ok,          // 정상 입력
    Empty,       // 입력할 값 없음(해당 없음) — 확인 대상 아님
    NeedsCheck,  // 확인 대상(자료 누락, 동명이인, 기존값 덮어쓰기 등)
}

public sealed class FieldResult
{
    public decimal? Value { get; init; }
    public CellState State { get; init; }
    public string Note { get; init; } = "";
}

public enum MatchState { Matched, Ambiguous }

public sealed class RrnCandidate
{
    public string RrnKey { get; init; } = "";
    public string Masked { get; init; } = "";
    public string Sources { get; init; } = "";
    public override string ToString() => $"{Masked} ({Sources})";
}

public sealed class EmployeeResult
{
    public TargetEmployee Target { get; init; } = null!;
    public MatchState Match { get; set; }
    public Dictionary<FieldId, FieldResult> Fields { get; } = new();
    public List<string> Issues { get; } = new();
    public List<RrnCandidate> Candidates { get; } = new();
    public string? ChosenRrn { get; set; }
    public bool NeedsCheck => Match != MatchState.Matched || Fields.Values.Any(f => f.State == CellState.NeedsCheck);
}

public sealed class ReconLine
{
    public FieldId Field { get; init; }
    public decimal SourceTotal { get; init; }      // 원본 전체 근로자 부담 합계
    public decimal Entered { get; init; }          // 양식에 입력된 합계
    public decimal OutOfTarget { get; init; }      // 대상 양식에 없는 직원분
    public decimal Unreflected { get; init; }      // 대상 직원이지만 미반영(동명이인 미확정 등)
    public decimal Diff => SourceTotal - Entered - OutOfTarget - Unreflected;
    public bool IsOk => Diff == 0;
}

public sealed class ReferenceLine
{
    public string Label { get; init; } = "";
    public decimal Actual { get; init; }
    public decimal? Expected { get; init; }
    public bool? IsOk => Expected == null ? null : Actual == Expected;
}

public sealed class ExtraRecord
{
    public string Source { get; init; } = "";
    public string Name { get; init; } = "";
    public int SheetRow { get; init; }
    public string Reason { get; init; } = "";
}

public sealed class ProcessResult
{
    public TargetData Target { get; init; } = null!;
    public List<EmployeeResult> Employees { get; } = new();
    public List<ReconLine> Recon { get; } = new();
    public List<ReferenceLine> References { get; } = new();
    public List<ExtraRecord> Extras { get; } = new();
    public List<string> Warnings { get; } = new();

    public int CheckCount => Employees.Count(e => e.NeedsCheck);
    public bool ReconOk => Recon.All(r => r.IsOk);
}
