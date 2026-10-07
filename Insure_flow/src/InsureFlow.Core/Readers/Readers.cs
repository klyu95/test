using System.Text.RegularExpressions;
using InsureFlow.Core.Excel;
using InsureFlow.Core.Model;

namespace InsureFlow.Core.Readers;

public enum FileKind { Unknown, Target, NationalPension, Employment, Health }

public static class FileKindExtensions
{
    public static string Label(this FileKind k) => k switch
    {
        FileKind.Target => "대상 양식", FileKind.NationalPension => "국민연금",
        FileKind.Employment => "고용보험", FileKind.Health => "건강·장기요양보험", _ => "알 수 없음",
    };
}

internal static class HeaderUtil
{
    public static string Norm(string s) => Regex.Replace(s ?? "", @"\s+", "");

    /// <summary>앞쪽 maxRows 행 안에서 pred를 만족하는 첫 행 번호(0-based), 없으면 -1.</summary>
    public static int FindRow(Grid g, Func<string[], bool> pred, int maxRows = 12)
    {
        for (var r = 0; r < Math.Min(g.RowCount, maxRows); r++)
        {
            var cells = Enumerable.Range(0, g.ColCount(r)).Select(c => Norm(g.Text(r, c))).ToArray();
            if (pred(cells)) return r;
        }
        return -1;
    }

    public static int FindCol(Grid g, int headerRow, Func<string, bool> pred)
    {
        for (var c = 0; c < g.ColCount(headerRow); c++)
            if (pred(Norm(g.Text(headerRow, c)))) return c;
        return -1;
    }

    public static int Require(int col, string what, string file)
    {
        if (col < 0) throw new InsureFlowException($"'{file}'에서 '{what}' 열을 찾을 수 없습니다. EDI 양식이 변경되었는지 확인하세요.");
        return col;
    }

    public static string? YearMonth(string? text)
    {
        var digits = new string((text ?? "").Where(char.IsDigit).ToArray());
        return digits.Length >= 6 ? digits[..6] : null;
    }
}

public static class FileDetector
{
    public static FileKind Detect(Grid g)
    {
        bool Any(Func<string[], bool> p) => HeaderUtil.FindRow(g, p) >= 0;
        if (Any(c => c.Contains("사번") && c.Contains("성명") && c.Any(x => x.Contains("[03]")))) return FileKind.Target;
        if (Any(c => c.Contains("성명") && c.Contains("요양산출보험료"))) return FileKind.Health;
        if (Any(c => c.Contains("근로자명"))) return FileKind.Employment;
        if (Any(c => c.Any(x => x.Contains("소급분")) && c.Any(x => x.Contains("본인기여금")) || c.Any(x => x.Contains("당월분")) && c.Any(x => x.Contains("본인기여금")))) return FileKind.NationalPension;
        return FileKind.Unknown;
    }
}

public static class TargetReader
{
    public static TargetData Read(string path, Grid g)
    {
        var file = Path.GetFileName(path);
        var hr = HeaderUtil.FindRow(g, c => c.Contains("사번") && c.Contains("성명"));
        if (hr < 0) throw new InsureFlowException($"'{file}'은(는) 대상 양식이 아닙니다. ('사번', '성명' 머리글 없음)");

        var nameCol = HeaderUtil.Require(HeaderUtil.FindCol(g, hr, s => s == "성명"), "성명", file);
        var noCol = HeaderUtil.Require(HeaderUtil.FindCol(g, hr, s => s == "사번"), "사번", file);

        var m = Regex.Match(g.SheetName, @"\((\d{4})-(\d{2})\)");
        var data = new TargetData
        {
            Path = path,
            SheetName = g.SheetName,
            YearMonth = m.Success ? m.Groups[1].Value + m.Groups[2].Value : null,
        };

        foreach (var f in FieldInfo.All)
        {
            var code = "[" + FieldInfo.Code(f) + "]";
            var col = HeaderUtil.FindCol(g, hr, s => s.EndsWith(code));
            data.Columns[f] = HeaderUtil.Require(col, $"{FieldInfo.Label(f)}{code}", file);
        }

        for (var r = hr + 1; r < g.RowCount; r++)
        {
            var name = g.Text(r, nameCol);
            if (name.Length == 0) continue;
            var e = new TargetEmployee
            {
                SheetRow = r,
                EmpNo = g.Text(r, noCol),
                Name = name,
                NameKey = Grid.NameKey(name),
            };
            foreach (var f in FieldInfo.All)
                if (g.Num(r, data.Columns[f]) is { } v) e.Existing[f] = v;
            data.Employees.Add(e);
        }

        if (data.Employees.Count == 0) throw new InsureFlowException($"'{file}'에서 직원 행을 찾을 수 없습니다.");
        return data;
    }
}

public static class NpReader
{
    public static SourceData<NpRecord> Read(string path, Grid g)
    {
        var file = Path.GetFileName(path);
        var hr = HeaderUtil.FindRow(g, c => c.Contains("성명") && c.Any(x => x.Contains("본인기여금")));
        if (hr < 0) throw new InsureFlowException($"'{file}'은(는) 국민연금 결정내역통보서가 아닙니다.");

        var nameCol = HeaderUtil.Require(HeaderUtil.FindCol(g, hr, s => s == "성명"), "성명", file);
        var rrnCol = HeaderUtil.FindCol(g, hr, s => s.Contains("주민등록번호"));
        var curCol = HeaderUtil.Require(HeaderUtil.FindCol(g, hr, s => s.Contains("당월분") && s.Contains("본인기여금")), "당월분 본인기여금", file);
        var retroCol = HeaderUtil.Require(HeaderUtil.FindCol(g, hr, s => s.Contains("소급분") && s.Contains("본인기여금")), "소급분 본인기여금", file);
        var lossCol = HeaderUtil.FindCol(g, hr, s => s.StartsWith("상실일"));
        var ymCol = HeaderUtil.FindCol(g, hr, s => s == "고지년월");

        var rows = new List<NpRecord>();
        string? ym = null;
        for (var r = hr + 1; r < g.RowCount; r++)
        {
            var name = g.Text(r, nameCol);
            if (name.Length == 0 || name == "합계") continue;
            if (ym == null && ymCol >= 0) ym = HeaderUtil.YearMonth(g.Text(r, ymCol));
            rows.Add(new NpRecord
            {
                SheetRow = r + 1,
                Name = name,
                NameKey = Grid.NameKey(name),
                RrnKey = rrnCol >= 0 ? Grid.RrnKey(g.Text(r, rrnCol)) : "",
                Current = g.Num(r, curCol),
                Retro = g.Num(r, retroCol),
                LossDate = lossCol >= 0 ? g.Text(r, lossCol) : "",
            });
        }
        var data = new SourceData<NpRecord> { Path = path, YearMonth = ym };
        data.Rows.AddRange(rows);
        return data;
    }
}

public sealed class EiReadResult
{
    public SourceData<EiRecord> Data { get; init; } = new();
    public EiTotals? Totals { get; init; }
}

public static class EiReader
{
    public static EiReadResult Read(string path, Grid g)
    {
        var file = Path.GetFileName(path);
        var h1 = HeaderUtil.FindRow(g, c => c.Contains("근로자명"));
        if (h1 < 0) throw new InsureFlowException($"'{file}'은(는) 고용보험 부과내역 조회 파일이 아닙니다.");
        var h2 = h1 + 1;

        // 1행(그룹) 머리글은 병합된 경우를 대비해 빈칸을 앞 값으로 채운다.
        var cols = Math.Max(g.ColCount(h1), g.ColCount(h2));
        var group = new string[cols];
        var sub = new string[cols];
        var last = "";
        for (var c = 0; c < cols; c++)
        {
            var t = HeaderUtil.Norm(g.Text(h1, c));
            if (t.Length > 0) last = t;
            group[c] = last;
            sub[c] = HeaderUtil.Norm(g.Text(h2, c));
        }

        int Find(Func<int, bool> p)
        {
            for (var c = 0; c < cols; c++) if (p(c)) return c;
            return -1;
        }

        bool Worker(int c) => sub[c].Contains("근로자") && sub[c].Contains("실업");
        bool Marker(int c, string m) => group[c].Contains(m) && !group[c].Contains("합계") && !group[c].Contains('+');

        var nameCol = HeaderUtil.Require(Find(c => HeaderUtil.Norm(g.Text(h1, c)) == "근로자명"), "근로자명", file);
        var rrnCol = Find(c => HeaderUtil.Norm(g.Text(h1, c)).Contains("주민등록번호"));
        var cur = HeaderUtil.Require(Find(c => Marker(c, "①") && Worker(c)), "산정보험료(해당월①) 근로자 실업급여보험료", file);
        var rec = HeaderUtil.Require(Find(c => Marker(c, "②") && Worker(c)), "재산정보험료(②) 근로자 실업급여보험료", file);
        var set = HeaderUtil.Require(Find(c => Marker(c, "③") && Worker(c)), "정산보험료(③) 근로자 실업급여보험료", file);

        var data = new SourceData<EiRecord>();
        EiTotals? totals = null;
        for (var r = h2 + 1; r < g.RowCount; r++)
        {
            var name = g.Text(r, nameCol);
            var isTotal = Enumerable.Range(0, Math.Min(cols, 6)).Any(c => g.Text(r, c) == "합계");
            if (isTotal)
            {
                totals = new EiTotals { Current = g.Num(r, cur), Recalc = g.Num(r, rec), Settle = g.Num(r, set) };
                continue;
            }
            if (name.Length == 0) continue;
            data.Rows.Add(new EiRecord
            {
                SheetRow = r + 1,
                Name = name,
                NameKey = Grid.NameKey(name),
                RrnKey = rrnCol >= 0 ? Grid.RrnKey(g.Text(r, rrnCol)) : "",
                Current = g.Num(r, cur) ?? 0,
                Recalc = g.Num(r, rec) ?? 0,
                Settle = g.Num(r, set) ?? 0,
            });
        }
        // 고용보험 조회 파일에는 고지년월 열이 없다(조회 조건). YearMonth는 알 수 없음.
        var result = new SourceData<EiRecord> { Path = path };
        result.Rows.AddRange(data.Rows);
        return new EiReadResult { Data = result, Totals = totals };
    }
}

public static class HiReader
{
    public static SourceData<HiRecord> Read(string path, Grid g)
    {
        var file = Path.GetFileName(path);
        var hr = HeaderUtil.FindRow(g, c => c.Contains("성명") && c.Contains("요양산출보험료"));
        if (hr < 0) throw new InsureFlowException($"'{file}'은(는) 건강보험 고지(산출) 내역서가 아닙니다.");

        int Exact(string h) => HeaderUtil.Require(HeaderUtil.FindCol(g, hr, s => s == h), h, file);
        var nameCol = Exact("성명");
        var rrnCol = HeaderUtil.FindCol(g, hr, s => s.Contains("주민등록번호"));
        var ymCol = HeaderUtil.FindCol(g, hr, s => s == "고지년월");
        var hCur = Exact("산출보험료");
        var hSet = Exact("정산금액");
        var hYe = Exact("연말정산");
        var lCur = Exact("요양산출보험료");
        var lSet = Exact("요양정산보험료");
        var lYe = Exact("요양연말정산보험료");

        var rows = new List<HiRecord>();
        string? ym = null;
        for (var r = hr + 1; r < g.RowCount; r++)
        {
            var name = g.Text(r, nameCol);
            if (name.Length == 0 || name == "합계") continue;
            if (ym == null && ymCol >= 0) ym = HeaderUtil.YearMonth(g.Text(r, ymCol));
            rows.Add(new HiRecord
            {
                SheetRow = r + 1,
                Name = name,
                NameKey = Grid.NameKey(name),
                RrnKey = rrnCol >= 0 ? Grid.RrnKey(g.Text(r, rrnCol)) : "",
                HealthCurrent = g.Num(r, hCur) ?? 0,
                HealthAdjust = (g.Num(r, hSet) ?? 0) + (g.Num(r, hYe) ?? 0),
                LtcCurrent = g.Num(r, lCur) ?? 0,
                LtcAdjust = (g.Num(r, lSet) ?? 0) + (g.Num(r, lYe) ?? 0),
            });
        }
        var data = new SourceData<HiRecord> { Path = path, YearMonth = ym };
        data.Rows.AddRange(rows);
        return data;
    }
}
