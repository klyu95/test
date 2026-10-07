namespace InsureFlow.Core.Model;

/// <summary>대상 양식에 입력하는 8개 칸. 양식 헤더의 [코드]로 열을 찾는다.</summary>
public enum FieldId
{
    NpCurrent,   // 국민연금[03]
    NpAdjust,    // 국민연금정산[09]
    HiCurrent,   // 건강보험[04]
    HiAdjust,    // 건강보험정산[10]
    LtcCurrent,  // 장기요양보험[22]
    LtcAdjust,   // 장기요양보험정산[23]
    EiCurrent,   // 고용보험[05]
    EiAdjust,    // 고용보험정산[25]
}

public enum Insurance { NationalPension, Health, LongTermCare, Employment }

public static class FieldInfo
{
    public static readonly FieldId[] All = Enum.GetValues<FieldId>();

    public static string Code(FieldId f) => f switch
    {
        FieldId.NpCurrent => "03", FieldId.NpAdjust => "09",
        FieldId.HiCurrent => "04", FieldId.HiAdjust => "10",
        FieldId.LtcCurrent => "22", FieldId.LtcAdjust => "23",
        FieldId.EiCurrent => "05", FieldId.EiAdjust => "25",
        _ => throw new ArgumentOutOfRangeException(nameof(f)),
    };

    public static string Label(FieldId f) => f switch
    {
        FieldId.NpCurrent => "국민연금", FieldId.NpAdjust => "국민연금정산",
        FieldId.HiCurrent => "건강보험", FieldId.HiAdjust => "건강보험정산",
        FieldId.LtcCurrent => "장기요양", FieldId.LtcAdjust => "장기요양정산",
        FieldId.EiCurrent => "고용보험", FieldId.EiAdjust => "고용보험정산",
        _ => "",
    };

    public static Insurance Ins(FieldId f) => f switch
    {
        FieldId.NpCurrent or FieldId.NpAdjust => Insurance.NationalPension,
        FieldId.HiCurrent or FieldId.HiAdjust => Insurance.Health,
        FieldId.LtcCurrent or FieldId.LtcAdjust => Insurance.LongTermCare,
        _ => Insurance.Employment,
    };

    public static bool IsAdjust(FieldId f) => f is FieldId.NpAdjust or FieldId.HiAdjust or FieldId.LtcAdjust or FieldId.EiAdjust;

    public static string Label(Insurance i) => i switch
    {
        Insurance.NationalPension => "국민연금", Insurance.Health => "건강보험",
        Insurance.LongTermCare => "장기요양보험", _ => "고용보험",
    };
}

// ---------- 대상 양식 ----------

public sealed class TargetEmployee
{
    public int SheetRow { get; init; }          // 0-based 시트 행 번호
    public string EmpNo { get; init; } = "";
    public string Name { get; init; } = "";
    public string NameKey { get; init; } = "";
    /// <summary>입력 칸에 이미 들어 있는 값(없으면 null).</summary>
    public Dictionary<FieldId, decimal> Existing { get; } = new();
}

public sealed class TargetData
{
    public string Path { get; init; } = "";
    public string SheetName { get; init; } = "";
    public string? YearMonth { get; init; }   // yyyyMM
    public Dictionary<FieldId, int> Columns { get; } = new();   // 0-based 열
    public List<TargetEmployee> Employees { get; } = new();
}

// ---------- 원본 레코드 ----------

public abstract class SourceRecord
{
    public int SheetRow { get; init; }   // 1-based (사용자 표시용)
    public string Name { get; init; } = "";
    public string NameKey { get; init; } = "";
    public string RrnKey { get; init; } = "";
    public bool Consumed { get; set; }
}

public sealed class NpRecord : SourceRecord
{
    public decimal? Current { get; init; }   // 당월분 본인기여금
    public decimal? Retro { get; init; }     // 소급분 본인기여금
    public string LossDate { get; init; } = "";
}

public sealed class EiRecord : SourceRecord
{
    public decimal Current { get; init; }    // ① 근로자 실업급여
    public decimal Recalc { get; init; }     // ② 재산정
    public decimal Settle { get; init; }     // ③ 정산
}

public sealed class HiRecord : SourceRecord
{
    public decimal HealthCurrent { get; init; }  // 산출보험료(총액)
    public decimal HealthAdjust { get; init; }   // 정산금액 + 연말정산(총액)
    public decimal LtcCurrent { get; init; }     // 요양산출보험료(총액)
    public decimal LtcAdjust { get; init; }      // 요양정산 + 요양연말정산(총액)
}

public sealed class SourceData<T> where T : SourceRecord
{
    public string Path { get; init; } = "";
    public string? YearMonth { get; init; }
    public List<T> Rows { get; } = new();
}

/// <summary>고용보험 원본의 '합계' 행 (대조용).</summary>
public sealed class EiTotals
{
    public decimal? Current { get; init; }
    public decimal? Recalc { get; init; }
    public decimal? Settle { get; init; }
}
