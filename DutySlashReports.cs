using System.Globalization;
using System.Text;

/// <summary>
/// <b>Будний столбец:</b> дневное дежурство 09:00–18:00, ночное 18:00 до утреннего порога следующего столбца (08:00);
/// между 08:00 и 09:00 на буднем — перерыв без дежурной смены.
/// <b>Выходной по заливке Excel:</b> день 09:00–21:00, ночь 21:00–09:00.
/// Конец утренней части ночи — 09:00, если столбец следующего утра помечен выходным, иначе 08:00 (при отсутствии заливки — 08:00).
/// Начало ночи относится к <b>столбцу N</b> этого же календарного дня.
/// </summary>
internal static class DutyShiftClock
{
    internal const int DayDutyStartHour = 9;

    internal readonly record struct Pointer(int Year, int Month, int DayNumber, DutyShift Shift);

    internal static TimeSpan MorningCutFor(DateOnly calendarMorning, DutyScheduleSnapshot? snap)
    {
        if (snap?.RestDayByDayFill is not { Count: > 0 } rests ||
            calendarMorning.Year != snap.ScheduleYear ||
            calendarMorning.Month != snap.ScheduleMonth)
            return TimeSpan.FromHours(8);

        return rests.GetValueOrDefault(calendarMorning.Day)
            ? TimeSpan.FromHours(9)
            : TimeSpan.FromHours(8);
    }

    /// <summary>Время перехода день→ночь в столбце календарного дня.</summary>
    internal static TimeSpan EveningStartFor(DateOnly calendarDay, DutyScheduleSnapshot? snap)
    {
        if (snap?.RestDayByDayFill is not { Count: > 0 } rests ||
            calendarDay.Year != snap.ScheduleYear ||
            calendarDay.Month != snap.ScheduleMonth)
            return TimeSpan.FromHours(18);

        return rests.GetValueOrDefault(calendarDay.Day)
            ? TimeSpan.FromHours(21)
            : TimeSpan.FromHours(18);
    }

    internal static Pointer? TryResolveActiveDuty(
        DateTime localNow,
        Func<DateOnly, TimeSpan> morningCutGetter,
        Func<DateOnly, DutyScheduleSnapshot?> snapForCalendarDay)
    {
        var date = DateOnly.FromDateTime(localNow);
        var tod = localNow.TimeOfDay;

        var cutToday = morningCutGetter(date);

        // Хвост ночной смены столбца вчера — до утреннего порога столбца сегодня
        if (tod < cutToday)
        {
            var nightColumn = date.AddDays(-1);
            return new Pointer(nightColumn.Year, nightColumn.Month, nightColumn.Day, DutyShift.Night);
        }

        var morningDayStart = TimeSpan.FromHours(DayDutyStartHour);
        // Буднее утро после 08:00 (или дефолт без заливки), но дневная смена с 09:00
        if (cutToday < morningDayStart && tod < morningDayStart)
            return null;

        var eveningHere = EveningStartFor(date, snapForCalendarDay(date));
        if (tod < eveningHere)
            return new Pointer(date.Year, date.Month, date.Day, DutyShift.Day);

        return new Pointer(date.Year, date.Month, date.Day, DutyShift.Night);
    }
}

internal static class DutySlashReports
{
    private static readonly CultureInfo Ru = CultureInfo.GetCultureInfo("ru-RU");

    internal readonly record struct ActiveDutyAnnKey(DateOnly ColumnDate, DutyShift Shift);

    internal readonly record struct ShiftStartAnnouncement(ActiveDutyAnnKey Key, string Text);

    internal static bool SameDutyAnnouncementKey(ActiveDutyAnnKey? a, ActiveDutyAnnKey? b)
    {
        if (!a.HasValue && !b.HasValue)
            return true;
        if (a.HasValue != b.HasValue)
            return false;
        return a.GetValueOrDefault() == b.GetValueOrDefault();
    }

    private sealed record ResolvedActiveDuty(
        DateOnly PickDate,
        int Dom,
        DutyShift Shift,
        DutyScheduleSnapshot SnapPick,
        DutyScheduleSnapshot SnapCross);

    private static DutyShiftAssignment? Find(IReadOnlyList<DutyShiftAssignment> list, int day, DutyShift shift) =>
        list.FirstOrDefault(s => s.Day == day && s.Shift == shift);

    private static string Who(DutyShiftAssignment? a) =>
        a is null
            ? "— (нет в графике)"
            : string.IsNullOrWhiteSpace(a.MatchedPersonName)
                ? (string.IsNullOrWhiteSpace(a.RawCell) ? "—" : a.RawCell.Trim())
                : a.MatchedPersonName;

    private static string NoFile(int y, int m) =>
        $"Нет утверждённого расписания на {ScheduleLocaleFormatting.RussianMonthYear(y, m)}.\n"
        + "Отправьте Excel и подтвердите сохранение кнопкой «Да, взять в работу».";

    private static async Task<DutyScheduleSnapshot?> LoadOrNull(int y, int m, CancellationToken ct) =>
        await ApprovedScheduleStore.TryLoadMonthAsync(y, m, ct).ConfigureAwait(false);

    private static bool MatchesScheduleMonth(DateOnly d, DutyScheduleSnapshot? z) =>
        z != null && d.Year == z.ScheduleYear && d.Month == z.ScheduleMonth;

    private static async Task<(DutyScheduleSnapshot? snapToday, DutyScheduleSnapshot? snapTomorrow)> LoadTodayTomorrowSnapsAsync(
        DateOnly today,
        CancellationToken ct)
    {
        var snapToday = await LoadOrNull(today.Year, today.Month, ct).ConfigureAwait(false);
        var tomorrow = today.AddDays(1);
        var snapTomorrow =
            today.Month == tomorrow.Month && today.Year == tomorrow.Year
                ? snapToday
                : await LoadOrNull(tomorrow.Year, tomorrow.Month, ct).ConfigureAwait(false);
        return (snapToday, snapTomorrow);
    }

    private static TimeSpan MorningCutCombined(
        DateOnly d,
        DutyScheduleSnapshot? snapToday,
        DutyScheduleSnapshot? snapTomorrow)
    {
        if (MatchesScheduleMonth(d, snapToday))
            return DutyShiftClock.MorningCutFor(d, snapToday);

        return MatchesScheduleMonth(d, snapTomorrow)
            ? DutyShiftClock.MorningCutFor(d, snapTomorrow!)
            : TimeSpan.FromHours(8);
    }

    private static async Task<ResolvedActiveDuty?> FinishResolvedActiveDutyAsync(
        DutyShiftClock.Pointer ptr,
        DutyScheduleSnapshot? snapToday,
        DutyScheduleSnapshot? snapTomorrow,
        CancellationToken ct)
    {
        var (py, pm, pdom, activeShift) = ptr;
        var pickDate = new DateOnly(py, pm, pdom);

        DutyScheduleSnapshot? snapPick =
            MatchesScheduleMonth(pickDate, snapToday)
                ? snapToday
                : MatchesScheduleMonth(pickDate, snapTomorrow)
                    ? snapTomorrow
                    : null;

        snapPick ??= await LoadOrNull(py, pm, ct).ConfigureAwait(false);
        if (snapPick is null)
            return null;

        var nextCal = pickDate.AddDays(1);
        var snapCross =
            nextCal.Year == pickDate.Year && nextCal.Month == pickDate.Month
                ? snapPick
                : await LoadOrNull(nextCal.Year, nextCal.Month, ct).ConfigureAwait(false);

        return new ResolvedActiveDuty(pickDate, pdom, activeShift, snapPick, snapCross ?? snapPick);
    }

    internal static async Task<ShiftStartAnnouncement?> TryGetShiftStartAnnouncementAsync(DateTime now, CancellationToken ct)
    {
        var today = DateOnly.FromDateTime(now);
        var (snapToday, snapTomorrow) = await LoadTodayTomorrowSnapsAsync(today, ct).ConfigureAwait(false);

        TimeSpan MorningCutDeleg(DateOnly d) => MorningCutCombined(d, snapToday, snapTomorrow);

        DutyScheduleSnapshot? SnapForDate(DateOnly d)
        {
            if (MatchesScheduleMonth(d, snapToday))
                return snapToday;
            return MatchesScheduleMonth(d, snapTomorrow) ? snapTomorrow : null;
        }

        var ptr = DutyShiftClock.TryResolveActiveDuty(now, MorningCutDeleg, SnapForDate);
        if (ptr is null)
            return null;

        var resolved = await FinishResolvedActiveDutyAsync(ptr.Value, snapToday, snapTomorrow, ct).ConfigureAwait(false);
        if (resolved is null)
            return null;

        var sb = new StringBuilder();
        sb.AppendLine("\U0001F6CE Начало дежурной смены");
        AppendActiveShiftSummary(sb, resolved);
        return new ShiftStartAnnouncement(new ActiveDutyAnnKey(resolved.PickDate, resolved.Shift), sb.ToString().TrimEnd());
    }

    /// <summary>Снимок месяца дня <paramref name="d"/> и (при смене месяца) следующего дня — для границы ночи.</summary>
    private static async Task<(DutyScheduleSnapshot? snapD, DutyScheduleSnapshot? snapCross)> SnapPairForAsync(
        DateOnly d,
        CancellationToken ct)
    {
        var snapD = await LoadOrNull(d.Year, d.Month, ct).ConfigureAwait(false);
        var next = d.AddDays(1);
        if (next.Year == d.Year && next.Month == d.Month)
            return (snapD, snapD);

        return (snapD, await LoadOrNull(next.Year, next.Month, ct).ConfigureAwait(false));
    }

    private static TimeSpan MorningCutForCalendarMorning(
        DateOnly morning,
        DutyScheduleSnapshot? snapSameAsFirst,
        DutyScheduleSnapshot? snapSecond)
    {
        if (snapSameAsFirst is not null &&
            morning.Year == snapSameAsFirst.ScheduleYear &&
            morning.Month == snapSameAsFirst.ScheduleMonth)
            return DutyShiftClock.MorningCutFor(morning, snapSameAsFirst);

        if (snapSecond is not null &&
            morning.Year == snapSecond.ScheduleYear &&
            morning.Month == snapSecond.ScheduleMonth)
            return DutyShiftClock.MorningCutFor(morning, snapSecond);

        return TimeSpan.FromHours(8);
    }

    /// <returns>Строки «выходной» / «будний» для подписи дня месяца в отчётах slash.</returns>
    private static string SlashDayKindWord(DutyScheduleSnapshot snap, int dayOfMonth)
    {
        if (snap.RestDayByDayFill is null or { Count: 0 })
            return "тип дня не сохранён";
        return snap.RestDayByDayFill.GetValueOrDefault(dayOfMonth) ? "выходной" : "будний";
    }

    private static string SlashDayKindParen(DutyScheduleSnapshot snap, int dayOfMonth)
    {
        if (snap.RestDayByDayFill is null or { Count: 0 })
            return "(тип дня не сохранён)";
        return snap.RestDayByDayFill.GetValueOrDefault(dayOfMonth) ? "(выходной)" : "(будний)";
    }

    private static string SlashMonthDayHeading(DateOnly d, DutyScheduleSnapshot snap)
    {
        var left = $"{d.ToString("dd MMMM", Ru)} - {SlashDayKindWord(snap, d.Day)}";
        return Ru.TextInfo.ToLower(left);
    }

    private static string FormatHmNoLeadingHour(TimeSpan t) => $"{t.Hours}:{t.Minutes:D2}";

    private static string DayDutyParen(bool restByFill) =>
        restByFill ? "9:00 - 21:00" : "9:00 - 18:00";

    private static string FormatNightInterval(TimeSpan eveningStart, TimeSpan morningEndNext) =>
        $"{FormatHmNoLeadingHour(eveningStart)} - {FormatHmNoLeadingHour(morningEndNext)}";

    private static void AppendActiveShiftSummary(StringBuilder sb, ResolvedActiveDuty r)
    {
        var isRest = r.SnapPick.RestDayByDayFill is { Count: > 0 } &&
            r.SnapPick.RestDayByDayFill.GetValueOrDefault(r.Dom);
        var assignment = Find(r.SnapPick.Shifts, r.Dom, r.Shift);
        var nextCal = r.PickDate.AddDays(1);

        sb.AppendLine($"📆 {FormatDateLong(r.PickDate)} {SlashDayKindParen(r.SnapPick, r.Dom)}");

        if (r.Shift == DutyShift.Day)
        {
            sb.AppendLine($"☀️ День: {Who(assignment)} ({DayDutyParen(isRest)})");
        }
        else
        {
            var eve = DutyShiftClock.EveningStartFor(r.PickDate, r.SnapPick);
            var nightEnd = MorningCutForCalendarMorning(nextCal, r.SnapPick, r.SnapCross);
            sb.AppendLine($"🌙 Ночь: {Who(assignment)} ({FormatNightInterval(eve, nightEnd)})");
        }
    }

    private static async Task<string> FormatOneCalendarDayAsync(DateOnly d, CancellationToken ct)
    {
        var (snap, snapCross) = await SnapPairForAsync(d, ct).ConfigureAwait(false);
        if (snap is null)
            return $"📆 {FormatDateLong(d)}\n" + NoFile(d.Year, d.Month);

        var dayAssign = Find(snap.Shifts, d.Day, DutyShift.Day);
        var nightAssign = Find(snap.Shifts, d.Day, DutyShift.Night);
        var isRest = snap.RestDayByDayFill is { Count: > 0 } && snap.RestDayByDayFill.GetValueOrDefault(d.Day);

        var eve = DutyShiftClock.EveningStartFor(d, snap);
        var nightEndMorning = MorningCutForCalendarMorning(d.AddDays(1), snap, snapCross ?? snap);

        var sb = new StringBuilder();
        sb.AppendLine($"📆 {FormatDateLong(d)} {SlashDayKindParen(snap, d.Day)}");
        sb.AppendLine($"☀️ День: {Who(dayAssign)} ({DayDutyParen(isRest)})");
        sb.AppendLine($"🌙 Ночь: {Who(nightAssign)} ({FormatNightInterval(eve, nightEndMorning)})");
        return sb.ToString().TrimEnd();
    }

    private static string FormatDateLong(DateOnly d) =>
        d.ToString("dddd, d MMMM yyyy", Ru);

    private static string FormatDateShort(DateOnly d) =>
        d.ToString("ddd d.MM", Ru);

    internal static async Task<string> DutyNowAsync(CancellationToken ct)
    {
        var now = DateTime.Now;

        var today = DateOnly.FromDateTime(now);
        var (snapToday, snapTomorrow) = await LoadTodayTomorrowSnapsAsync(today, ct).ConfigureAwait(false);

        TimeSpan MorningCutDeleg(DateOnly d) => MorningCutCombined(d, snapToday, snapTomorrow);

        DutyScheduleSnapshot? SnapForDate(DateOnly d)
        {
            if (MatchesScheduleMonth(d, snapToday))
                return snapToday;
            return MatchesScheduleMonth(d, snapTomorrow) ? snapTomorrow : null;
        }

        var ptr = DutyShiftClock.TryResolveActiveDuty(now, MorningCutDeleg, SnapForDate);
        if (ptr is null)
        {
            var calendarToday = DateOnly.FromDateTime(now);
            var morningCut = MorningCutDeleg(calendarToday);
            return "⏱ Сейчас (местное время): " + now.ToString("dd.MM.yyyy HH:mm", Ru)
                + "\n\n"
                + $"С {FormatHmNoLeadingHour(morningCut)} до {DutyShiftClock.DayDutyStartHour}:00 — по графику нет активной смены "
                + "(окно между окончанием ночи и началом дня).";
        }

        var resolved = await FinishResolvedActiveDutyAsync(ptr.Value, snapToday, snapTomorrow, ct).ConfigureAwait(false);
        if (resolved is null)
        {
            var pInfo = ptr.Value;
            return "⏱ " + now.ToString("dd.MM.yyyy HH:mm", Ru) + "\n\n" + NoFile(pInfo.Year, pInfo.Month);
        }

        var sb = new StringBuilder();
        sb.AppendLine($"⏱ Сейчас (местное время): {now.ToString("dd.MM.yyyy HH:mm", Ru)}");
        sb.AppendLine();
        AppendActiveShiftSummary(sb, resolved);
        return sb.ToString().TrimEnd();
    }

    internal static Task<string> DutyTodayAsync(CancellationToken ct) =>
        FormatOneCalendarDayAsync(DateOnly.FromDateTime(DateTime.Now), ct);

    internal static async Task<string> DutyTwoDaysAsync(CancellationToken ct)
    {
        var t = DateOnly.FromDateTime(DateTime.Now);
        var a = await FormatOneCalendarDayAsync(t, ct).ConfigureAwait(false);
        var b = await FormatOneCalendarDayAsync(t.AddDays(1), ct).ConfigureAwait(false);
        return "📋 Два дня (с сегодня)\n\n" + a + "\n\n" + "────────\n\n" + b;
    }

    internal static async Task<string> DutyWeekAsync(CancellationToken ct)
    {
        var start = DateOnly.FromDateTime(DateTime.Now);
        var sb = new StringBuilder();
        sb.AppendLine($"📅 7 дней с {FormatDateShort(start)}\n");

        for (var i = 0; i < 7; i++)
        {
            var d = start.AddDays(i);
            sb.AppendLine(await FormatOneCalendarDayAsync(d, ct).ConfigureAwait(false));
            if (i < 6)
            {
                sb.AppendLine();
                sb.AppendLine("────────");
                sb.AppendLine();
            }
        }

        return sb.ToString().TrimEnd();
    }

    internal static async Task<string> DutyMonthAsync(CancellationToken ct)
    {
        var today = DateOnly.FromDateTime(DateTime.Now);
        var y = today.Year;
        var m = today.Month;
        var snap = await LoadOrNull(y, m, ct).ConfigureAwait(false);
        if (snap is null)
            return NoFile(y, m);

        var days = DateTime.DaysInMonth(y, m);
        var lastDom = new DateOnly(y, m, days);
        var dayAfterLast = lastDom.AddDays(1);
        var snapNext =
            dayAfterLast.Month == m && dayAfterLast.Year == y
                ? snap
                : await LoadOrNull(dayAfterLast.Year, dayAfterLast.Month, ct).ConfigureAwait(false);

        var sb = new StringBuilder();
        sb.AppendLine($"📆 {ScheduleLocaleFormatting.RussianMonthYear(y, m)} — весь месяц\n");

        for (var dom = 1; dom <= days; dom++)
        {
            var cal = new DateOnly(y, m, dom);
            var cross = cal.AddDays(1).Month == m ? snap : snapNext ?? snap;

            sb.AppendLine(SlashMonthDayHeading(cal, snap));

            var isRest = snap.RestDayByDayFill is { Count: > 0 } && snap.RestDayByDayFill.GetValueOrDefault(dom);
            var eve = DutyShiftClock.EveningStartFor(cal, snap);
            var nightEndMorning = MorningCutForCalendarMorning(cal.AddDays(1), snap, cross);

            var dayAssign = Find(snap.Shifts, dom, DutyShift.Day);
            var nightAssign = Find(snap.Shifts, dom, DutyShift.Night);
            sb.AppendLine($"🌙 Ночь: {Who(nightAssign)} ({FormatNightInterval(eve, nightEndMorning)})");
            sb.AppendLine($"☀️ День: {Who(dayAssign)} ({DayDutyParen(isRest)})");
            sb.AppendLine();
        }

        return sb.ToString().TrimEnd();
    }
}
