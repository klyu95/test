using InsureFlow.Core.Model;
using InsureFlow.Core.Processing;
using InsureFlow.Core.Writers;
using Xunit;
using Xunit.Abstractions;

namespace InsureFlow.Tests;

public class SampleIntegrationTests
{
    private readonly ITestOutputHelper _out;
    public SampleIntegrationTests(ITestOutputHelper o) => _out = o;

    [SkippableFact]
    public void Sample_files_process_and_reconcile()
    {
        Skip.IfNot(SampleFiles.Available, "docs 샘플 없음");
        var p = Processor.Load(SampleFiles.Target!, SampleFiles.Np!, SampleFiles.Ei!, SampleFiles.Hi!);
        var res = p.Run();

        _out.WriteLine($"target={res.Employees.Count} np={p.Np.Rows.Count} ei={p.Ei.Rows.Count} hi={p.Hi.Rows.Count}");
        foreach (var w in res.Warnings) _out.WriteLine("WARN " + w);
        foreach (var l in res.Recon)
            _out.WriteLine($"{l.Field,-10} src={l.SourceTotal,12:N0} entered={l.Entered,12:N0} out={l.OutOfTarget,10:N0} unref={l.Unreflected,10:N0} diff={l.Diff}");
        foreach (var rf in res.References) _out.WriteLine($"REF {rf.Label}: {rf.Actual:N0} / {rf.Expected:N0} {rf.IsOk}");
        foreach (var e in res.Employees.Where(x => x.NeedsCheck)) _out.WriteLine($"CHECK {e.Target.Name}: {string.Join("; ", e.Issues)}");
        foreach (var x in res.Extras) _out.WriteLine($"EXTRA {x.Source} {x.Name} {x.Reason}");

        Assert.Equal(58, res.Employees.Count);
        Assert.All(res.Employees, e => Assert.Equal(MatchState.Matched, e.Match));
        Assert.True(res.ReconOk);
    }

    [SkippableFact]
    public void Sample_output_keeps_everything_except_input_cells()
    {
        Skip.IfNot(SampleFiles.Available, "docs 샘플 없음");
        var p = Processor.Load(SampleFiles.Target!, SampleFiles.Np!, SampleFiles.Ei!, SampleFiles.Hi!);
        var res = p.Run();
        var outPath = Path.Combine(Path.GetTempPath(), "insureflow_test_" + Guid.NewGuid().ToString("N") + ".xls");
        try
        {
            var written = TargetWriter.Save(SampleFiles.Target!, outPath, res);
            var diffs = TargetVerifier.Compare(SampleFiles.Target!, outPath, written);
            foreach (var d in diffs.Take(30)) _out.WriteLine(d);
            Assert.Empty(diffs);
            Assert.True(written.Count > 300);
        }
        finally { File.Delete(outPath); }
    }
}
