using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Data;
using InsureFlow.Core;
using InsureFlow.Core.Excel;
using InsureFlow.Core.Model;
using InsureFlow.Core.Processing;
using InsureFlow.Core.Readers;
using InsureFlow.Core.Writers;
using Microsoft.Win32;

namespace InsureFlow.App.ViewModels;

public sealed class FileSlotVm : ObservableBase
{
    private string _path = "", _status = "파일을 선택하세요.";
    private bool _ok;

    public FileKind Kind { get; init; }
    public string Title { get; init; } = "";
    public string Hint { get; init; } = "";
    public string Path { get => _path; set => Set(ref _path, value); }
    public string Status { get => _status; set => Set(ref _status, value); }
    public bool IsOk { get => _ok; set => Set(ref _ok, value); }
    public RelayCommand BrowseCommand { get; set; } = null!;
}

public sealed class MainViewModel : ObservableBase
{
    private readonly RelayCommand _next, _back, _save;
    private Processor? _processor;
    private ProcessResult? _result;
    private readonly Dictionary<int, string> _choices = new();

    private int _step;
    private bool _busy, _onlyCheck, _confirmChecked, _saved;
    private string _error = "", _summary = "", _warnings = "", _outputPath = "", _saveMessage = "", _reportPath = "";
    private RowVm? _selected;
    private RrnCandidate? _selectedCandidate;

    public ObservableCollection<StepVm> Steps { get; } = new()
    {
        new StepVm { Number = 1, Title = "파일 선택" },
        new StepVm { Number = 2, Title = "입력 미리보기" },
        new StepVm { Number = 3, Title = "확인" },
        new StepVm { Number = 4, Title = "저장" },
    };

    public ObservableCollection<FileSlotVm> Slots { get; } = new();
    public ObservableCollection<RowVm> Rows { get; } = new();
    public ICollectionView RowsView { get; }
    public ObservableCollection<ReconVm> Recon { get; } = new();
    public ObservableCollection<RefVm> References { get; } = new();
    public ObservableCollection<RrnCandidate> Candidates { get; } = new();

    public RelayCommand NextCommand => _next;
    public RelayCommand BackCommand => _back;
    public RelayCommand SaveCommand => _save;
    public RelayCommand ChooseCandidateCommand { get; }
    public RelayCommand ClearChoiceCommand { get; }
    public RelayCommand BrowseOutputCommand { get; }
    public RelayCommand OpenFolderCommand { get; }
    public RelayCommand OpenReportCommand { get; }

    public MainViewModel()
    {
        Slots.Add(MakeSlot(FileKind.Target, "① 대상 양식 (공제 입력 양식)", "매월 내려받는 재직자 기준 최종 입력 양식(.xls)"));
        Slots.Add(MakeSlot(FileKind.NationalPension, "② 국민연금", "2차결정내역통보서 (.xlsx)"));
        Slots.Add(MakeSlot(FileKind.Employment, "③ 고용보험", "당월보험료부과내역조회(고용) (.xlsx)"));
        Slots.Add(MakeSlot(FileKind.Health, "④ 건강·장기요양보험", "보험료 고지(산출) 내역서 (.xls)"));

        RowsView = CollectionViewSource.GetDefaultView(Rows);
        RowsView.Filter = o => !OnlyCheck || (o is RowVm r && r.NeedsCheck);

        _next = new RelayCommand(OnNext, () => !Busy && Step < 3 && (Step != 2 || CanLeaveConfirm));
        _back = new RelayCommand(() => { Step--; Refresh(); }, () => !Busy && Step > 0);
        _save = new RelayCommand(OnSave, () => !Busy && Step == 3 && OutputPath.Length > 0);
        ChooseCandidateCommand = new RelayCommand(OnChoose, () => Selected?.Source.Candidates.Count > 0 && SelectedCandidate != null);
        ClearChoiceCommand = new RelayCommand(OnClearChoice, () => Selected != null && _choices.ContainsKey(Selected.Source.Target.SheetRow));
        BrowseOutputCommand = new RelayCommand(OnBrowseOutput);
        OpenFolderCommand = new RelayCommand(() => OpenPath(Path.GetDirectoryName(OutputPath)), () => _saved);
        OpenReportCommand = new RelayCommand(() => OpenPath(_reportPath), () => _saved);
        UpdateSteps();
    }

    private FileSlotVm MakeSlot(FileKind kind, string title, string hint)
    {
        var slot = new FileSlotVm { Kind = kind, Title = title, Hint = hint };
        slot.BrowseCommand = new RelayCommand(() => Browse(slot));
        return slot;
    }

    // ---------- 바인딩 속성 ----------

    public int Step { get => _step; private set { if (Set(ref _step, value)) { Raise(nameof(IsStep0)); Raise(nameof(IsStep1)); Raise(nameof(IsStep2)); Raise(nameof(IsStep3)); Raise(nameof(NextText)); UpdateSteps(); } } }
    public bool IsStep0 => Step == 0;
    public bool IsStep1 => Step == 1;
    public bool IsStep2 => Step == 2;
    public bool IsStep3 => Step == 3;
    public string NextText => Step == 2 ? "저장 단계로 ▶" : "다음 ▶";

    public bool Busy { get => _busy; private set { if (Set(ref _busy, value)) Refresh(); } }
    public string Error { get => _error; private set => Set(ref _error, value); }
    public string Summary { get => _summary; private set => Set(ref _summary, value); }
    public string Warnings { get => _warnings; private set { Set(ref _warnings, value); Raise(nameof(HasWarnings)); } }
    public bool HasWarnings => Warnings.Length > 0;

    public bool OnlyCheck { get => _onlyCheck; set { if (Set(ref _onlyCheck, value)) RowsView.Refresh(); } }

    public RowVm? Selected
    {
        get => _selected;
        set
        {
            if (!Set(ref _selected, value)) return;
            Candidates.Clear();
            if (value != null) foreach (var c in value.Source.Candidates) Candidates.Add(c);
            SelectedCandidate = Candidates.FirstOrDefault(c => c.RrnKey == value?.Source.ChosenRrn) ?? Candidates.FirstOrDefault();
            Raise(nameof(HasCandidates)); Raise(nameof(DetailTitle));
            ChooseCandidateCommand.Refresh(); ClearChoiceCommand.Refresh();
        }
    }
    public bool HasCandidates => Candidates.Count > 0;
    public string DetailTitle => Selected == null ? "직원을 선택하면 산출 근거가 표시됩니다."
        : $"{Selected.Name} ({Selected.EmpNo})";
    public RrnCandidate? SelectedCandidate { get => _selectedCandidate; set { Set(ref _selectedCandidate, value); ChooseCandidateCommand.Refresh(); } }

    public bool NeedsConfirm => (_result?.CheckCount ?? 0) > 0;
    public bool ConfirmChecked { get => _confirmChecked; set { if (Set(ref _confirmChecked, value)) Refresh(); } }
    private bool CanLeaveConfirm => !NeedsConfirm || ConfirmChecked;
    public string ConfirmText => $"확인 대상 {_result?.CheckCount ?? 0}건을 검토했으며, 해당 칸은 비워 둔 채(또는 표시된 값으로) 저장하는 데 동의합니다.";

    public string OutputPath { get => _outputPath; set { if (Set(ref _outputPath, value)) Refresh(); } }
    public string SaveMessage { get => _saveMessage; private set => Set(ref _saveMessage, value); }

    // ---------- 1단계: 파일 선택 ----------

    private void Browse(FileSlotVm slot)
    {
        var dlg = new OpenFileDialog
        {
            Title = slot.Title + " 선택",
            Filter = "엑셀 파일 (*.xls;*.xlsx)|*.xls;*.xlsx|모든 파일 (*.*)|*.*",
            CheckFileExists = true,
        };
        if (dlg.ShowDialog() == true) Assign(slot, dlg.FileName);
    }

    /// <summary>끌어다 놓은 파일들을 머리글로 종류를 판별해 알맞은 칸에 넣는다.</summary>
    public void HandleDrop(IEnumerable<string> paths)
    {
        foreach (var p in paths)
        {
            try
            {
                var kind = FileDetector.Detect(Grid.Load(p));
                var slot = Slots.FirstOrDefault(s => s.Kind == kind);
                if (slot == null) { Error = $"'{Path.GetFileName(p)}'은(는) 어떤 자료인지 알 수 없습니다."; continue; }
                Assign(slot, p);
            }
            catch (InsureFlowException ex) { Error = ex.Message; }
        }
    }

    private void Assign(FileSlotVm slot, string path)
    {
        Error = "";
        try
        {
            var kind = FileDetector.Detect(Grid.Load(path));
            if (kind != slot.Kind)
            {
                Error = kind == FileKind.Unknown
                    ? $"'{Path.GetFileName(path)}'은(는) 인식할 수 없는 형식입니다."
                    : $"'{Path.GetFileName(path)}'은(는) '{kind.Label()}' 자료로 보입니다. '{slot.Title}' 칸에는 맞지 않습니다.";
                return;
            }
            slot.Path = path;
            slot.Status = "파일 형식 확인됨";
            slot.IsOk = true;
            _processor = null; _result = null;
            Refresh();
        }
        catch (InsureFlowException ex) { Error = ex.Message; }
    }

    // ---------- 단계 이동 ----------

    private async void OnNext()
    {
        Error = "";
        try
        {
            if (Step == 0) await LoadAndPreview();
            else if (Step == 1) { BuildConfirm(); Step = 2; }
            else if (Step == 2)
            {
                OutputPath = DefaultOutputPath();
                SaveMessage = ""; _saved = false;
                Step = 3;
            }
        }
        catch (InsureFlowException ex) { Error = ex.Message; }
        Refresh();
    }

    private async Task LoadAndPreview()
    {
        var missing = Slots.FirstOrDefault(s => !s.IsOk);
        if (missing != null) { Error = $"'{missing.Title}' 파일을 먼저 선택하세요."; return; }

        Busy = true;
        try
        {
            var paths = Slots.Select(s => s.Path).ToArray();
            _processor = await Task.Run(() => Processor.Load(paths[0], paths[1], paths[2], paths[3]));
            _choices.Clear();
            Recompute();
            ConfirmChecked = false;
            Step = 1;
        }
        finally { Busy = false; }
    }

    private void Recompute()
    {
        var keep = Selected?.Source.Target.SheetRow;
        _result = _processor!.Run(_choices);
        Rows.Clear();
        foreach (var e in _result.Employees) Rows.Add(new RowVm(e));
        Selected = keep == null ? null : Rows.FirstOrDefault(r => r.Source.Target.SheetRow == keep);

        var c = _result;
        var amb = c.Employees.Count(e => e.Match == MatchState.Ambiguous);
        Summary = $"직원 {c.Employees.Count}명 중 정상 {c.Employees.Count - c.CheckCount}명, 확인 대상 {c.CheckCount}명"
                  + (amb > 0 ? $" (동명이인 확정 필요 {amb}명)" : "")
                  + $" · 대상 외·미반영 원본 {c.Extras.Count}건";
        Warnings = string.Join("\n", c.Warnings);
        Raise(nameof(NeedsConfirm)); Raise(nameof(ConfirmText));
    }

    private void BuildConfirm()
    {
        Recon.Clear(); References.Clear();
        foreach (var l in _result!.Recon) Recon.Add(ReconVm.From(l));
        foreach (var r in _result.References) References.Add(RefVm.From(r));
    }

    private void OnChoose()
    {
        if (Selected == null || SelectedCandidate == null) return;
        _choices[Selected.Source.Target.SheetRow] = SelectedCandidate.RrnKey;
        Recompute();
    }

    private void OnClearChoice()
    {
        if (Selected == null) return;
        _choices.Remove(Selected.Source.Target.SheetRow);
        Recompute();
    }

    // ---------- 4단계: 저장 ----------

    private string DefaultOutputPath()
    {
        var src = Slots[0].Path;
        var dir = Path.GetDirectoryName(src) ?? "";
        return Path.Combine(dir, $"{Path.GetFileNameWithoutExtension(src)}_InsureFlow_{DateTime.Now:yyyyMMdd_HHmm}.xls");
    }

    private void OnBrowseOutput()
    {
        var dlg = new SaveFileDialog
        {
            Title = "저장할 .xls 파일 이름",
            Filter = "엑셀 97-2003 (*.xls)|*.xls",
            FileName = Path.GetFileName(OutputPath),
            InitialDirectory = Path.GetDirectoryName(OutputPath),
            AddExtension = true, DefaultExt = ".xls",
        };
        if (dlg.ShowDialog() == true) OutputPath = dlg.FileName;
    }

    private async void OnSave()
    {
        if (_result == null) return;
        Busy = true; SaveMessage = ""; Error = "";
        try
        {
            var src = Slots[0].Path; var outPath = OutputPath; var result = _result;
            var report = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(outPath))!,
                Path.GetFileNameWithoutExtension(outPath) + "_검증보고서.xlsx");
            var diffs = await Task.Run(() =>
            {
                var written = TargetWriter.Save(src, outPath, result);
                var d = TargetVerifier.Compare(src, outPath, written);
                ReportWriter.Save(report, result, Path.GetFileName(outPath), DateTime.Now, d);
                return d;
            });
            _reportPath = report; _saved = true;
            SaveMessage = diffs.Count == 0
                ? $"저장 완료\n• 양식: {outPath}\n• 검증 보고서: {report}\n• 재검증: 입력 칸 외 값·서식 변경 없음"
                : $"저장했지만 재검증에서 차이 {diffs.Count}건이 발견되었습니다. 검증 보고서를 확인하세요.\n• 양식: {outPath}\n• 보고서: {report}";
        }
        catch (InsureFlowException ex) { Error = ex.Message; }
        catch (Exception ex) { Error = "저장 중 오류: " + ex.Message; }
        finally { Busy = false; OpenFolderCommand.Refresh(); OpenReportCommand.Refresh(); }
    }

    private static void OpenPath(string? path)
    {
        if (string.IsNullOrEmpty(path)) return;
        try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); } catch { /* 무시 */ }
    }

    private void Refresh()
    {
        _next.Refresh(); _back.Refresh(); _save.Refresh();
        ChooseCandidateCommand.Refresh(); ClearChoiceCommand.Refresh();
    }

    private void UpdateSteps()
    {
        for (var i = 0; i < Steps.Count; i++)
        {
            Steps[i].IsCurrent = i == Step;
            Steps[i].IsDone = i < Step;
        }
    }
}
