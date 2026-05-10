/// <summary>Снимок графика для сохранения в JSON (после явного подтверждения пользователя).</summary>
internal sealed record DutyScheduleSnapshot(
    int ScheduleYear,
    int ScheduleMonth,
    string SourceTelegramFileName,
    string SourceLocalPath,
    DateTimeOffset? SavedAtUtc,
    Dictionary<int, bool>? RestDayByDayFill,
    List<DutyPerson> People,
    List<DutyShiftAssignment> Shifts,
    List<string> Warnings);

internal static class DutyScheduleSnapshotFactory
{
    internal static DutyScheduleSnapshot? TryFromParse(
        DutyScheduleParseResult parsed,
        string telegramFileName,
        string localPath)
    {
        if (parsed.ScheduleMonth is not (>= 1 and <= 12))
            return null;

        Dictionary<int, bool>? rest = null;
        if (parsed.RestDayByDayFill.Count > 0)
            rest = new Dictionary<int, bool>(parsed.RestDayByDayFill);

        return new DutyScheduleSnapshot(
            parsed.ScheduleYear,
            parsed.ScheduleMonth.Value,
            string.IsNullOrEmpty(telegramFileName) ? "—" : telegramFileName,
            localPath,
            null,
            rest,
            [..parsed.People],
            [..parsed.Shifts],
            [..parsed.Warnings]);
    }
}
