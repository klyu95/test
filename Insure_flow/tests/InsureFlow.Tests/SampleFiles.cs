namespace InsureFlow.Tests;

/// <summary>저장소 docs\의 실제 샘플(개인정보 포함, 커밋 제외). 없으면 관련 테스트는 건너뛴다.</summary>
internal static class SampleFiles
{
    public static string? Dir { get; } = FindDir();
    public static string? Target => Pick("1번.xls");
    public static string? Np => Pick("2차결정내역통보서*.xlsx");
    public static string? Ei => Pick("당월보험료부과내역조회(고용)*.xlsx");
    public static string? Hi => Pick("보험료_고지(산출)_내역서*.xls");
    public static bool Available => Target != null && Np != null && Ei != null && Hi != null;

    private static string? Pick(string pattern) =>
        Dir == null ? null : Directory.GetFiles(Dir, pattern).FirstOrDefault(f => !Path.GetFileName(f).StartsWith("~$"));

    private static string? FindDir()
    {
        var d = new DirectoryInfo(AppContext.BaseDirectory);
        while (d != null)
        {
            var docs = Path.Combine(d.FullName, "docs");
            if (Directory.Exists(docs) && Directory.GetFiles(docs, "1번.xls").Length > 0) return docs;
            d = d.Parent;
        }
        return null;
    }
}
