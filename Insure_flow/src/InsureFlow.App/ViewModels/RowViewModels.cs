using System.Collections.ObjectModel;
using InsureFlow.Core.Model;
using InsureFlow.Core.Processing;

namespace InsureFlow.App.ViewModels;

public sealed class StepVm : ObservableBase
{
    private bool _current, _done;
    public int Number { get; init; }
    public string Title { get; init; } = "";
    public bool IsCurrent { get => _current; set => Set(ref _current, value); }
    public bool IsDone { get => _done; set => Set(ref _done, value); }
}

public sealed class CellVm
{
    public string Text { get; init; } = "";
    public CellState State { get; init; }
    public bool IsNegative { get; init; }
    public string Note { get; init; } = "";
}

public sealed class RowVm
{
    public EmployeeResult Source { get; }
    public string EmpNo => Source.Target.EmpNo;
    public string Name => Source.Target.Name;
    public bool NeedsCheck => Source.NeedsCheck;
    public string StatusText => Source.Match == MatchState.Ambiguous ? "동명이인 확인"
        : Source.NeedsCheck ? "확인 대상" : "정상";
    public CellVm[] Cells { get; }

    public RowVm(EmployeeResult e)
    {
        Source = e;
        Cells = FieldInfo.All.Select(f =>
        {
            e.Fields.TryGetValue(f, out var fr);
            var text = fr?.Value is { } v ? v.ToString("N0") : fr?.State == CellState.NeedsCheck ? "확인" : "";
            return new CellVm
            {
                Text = text,
                State = fr?.State ?? CellState.Empty,
                IsNegative = fr?.Value < 0,
                Note = fr?.Note ?? "",
            };
        }).ToArray();
    }

    public IEnumerable<string> DetailLines =>
        FieldInfo.All.Select((f, i) => $"{FieldInfo.Label(f)}: {(string.IsNullOrEmpty(Cells[i].Note) ? "-" : Cells[i].Note)}");
}

public sealed class ReconVm
{
    public string Insurance { get; init; } = "";
    public string Period { get; init; } = "";
    public string SourceTotal { get; init; } = "";
    public string Entered { get; init; } = "";
    public string OutOfTarget { get; init; } = "";
    public string Unreflected { get; init; } = "";
    public string Diff { get; init; } = "";
    public string Result { get; init; } = "";
    public bool IsOk { get; init; }

    public static ReconVm From(ReconLine l) => new()
    {
        Insurance = FieldInfo.Label(FieldInfo.Ins(l.Field)),
        Period = FieldInfo.IsAdjust(l.Field) ? "정산분" : "당월분",
        SourceTotal = l.SourceTotal.ToString("N0"), Entered = l.Entered.ToString("N0"),
        OutOfTarget = l.OutOfTarget.ToString("N0"), Unreflected = l.Unreflected.ToString("N0"),
        Diff = l.Diff.ToString("N0"), Result = l.IsOk ? "✔ 일치" : "✖ 차이", IsOk = l.IsOk,
    };
}

public sealed class RefVm
{
    public string Label { get; init; } = "";
    public string Actual { get; init; } = "";
    public string Expected { get; init; } = "";
    public string Result { get; init; } = "";

    public static RefVm From(ReferenceLine r) => new()
    {
        Label = r.Label, Actual = r.Actual.ToString("N0"),
        Expected = r.Expected?.ToString("N0") ?? "",
        Result = r.IsOk is { } ok ? (ok ? "✔ 일치" : "✖ 차이") : "",
    };
}
