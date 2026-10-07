namespace InsureFlow.Core.Calc;

/// <summary>건강·장기요양보험처럼 총액(사업주+근로자)만 제공되는 보험의 근로자 몫 계산.</summary>
public static class ShareCalculator
{
    /// <summary>총액 ÷ 2 후 10원 미만 절사(0 방향). 음수(환급)도 0 방향으로 절사한다.</summary>
    public static decimal EmployeeShare(decimal total) =>
        Math.Truncate(total / 2m / 10m) * 10m;

    public static decimal EmployerShare(decimal total) => total - EmployeeShare(total);
}
