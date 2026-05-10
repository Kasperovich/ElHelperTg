using System.Globalization;

internal static class ScheduleLocaleFormatting
{
    private static readonly CultureInfo RuCulture = CultureInfo.GetCultureInfo("ru-RU");

    internal static string RussianMonthYear(int year, int month)
    {
        var fragment = new DateTime(year, month, 1).ToString("MMMM yyyy", RuCulture);
        return RuCulture.TextInfo.ToTitleCase(fragment);
    }
}

/// <summary>Смена в графике: для горизонтального макета — ночь и день; для старой вертикальной колонки — одна ячейка на день.</summary>
internal enum DutyShift
{
    Unspecified,
    Night,
    Day,
}

/// <summary>Человек из справочника дежурных + буква в графике (первая буква фамилии из ФИО). Телефон — текст ячейки без нормализации.</summary>
internal sealed record DutyPerson(
    string Name,
    string Phone,
    string ScheduleLetter,
    string? TabNumber = null,
    string? Position = null);

/// <summary>Одна ячейка графика: день месяца, смена, сырое значение, сопоставление со справочником.</summary>
internal sealed record DutyShiftAssignment(int Day, DutyShift Shift, string RawCell, string? MatchedPersonName);

internal sealed class DutyScheduleParseResult
{
    // Эмодзи в тексте сообщений Telegram (режим текста без разметки — UTF‑8 допустим).
    private const string EmMonth = "\uD83D\uDCC6";    // 📆
    private const string EmDays = "\uD83D\uDCC5";    // 📅 календарных дней
    private const string EmWeekend = "\uD83C\uDFDD"; // 🏝 выходные
    private const string EmWork = "\uD83D\uDCBC";     // 💼
    private const string EmPeople = "\uD83D\uDC65";    // 👥
    private const string EmSchedule = "\uD83D\uDCCB"; // 📋
    private const string EmNight = "\uD83C\uDF19";    // 🌙
    private const string EmDayShift = "\u2600\uFE0F"; // ☀️
    private const string EmDayGeneric = "\uD83D\uDCCD"; // 📍 одна ячейка на день
    private const string EmWarn = "\u26A0\uFE0F";     // ⚠️
    private const string EmCard = "\uD83D\uDCC8";     // 📈
    private const string EmDoc = "\uD83D\uDCC4";     // 📄
    private const string EmBell = "\uD83D\uDD14";     // 🔔 счётчик предупреждений в шапке

    /// <summary>Год календаря для подсчёта рабочих/выходных (из имени файла или листа, иначе текущий год).</summary>
    public int ScheduleYear { get; set; } = DateTime.Now.Year;

    /// <summary>Номер месяца 1–12, если удалось найти название месяца на видимом листе или в имени файла.</summary>
    public int? ScheduleMonth { get; set; }

    public List<DutyPerson> People { get; } = new();
    public List<DutyShiftAssignment> Shifts { get; } = new();
    public List<string> Warnings { get; } = new();

    /// <summary>Выходные дни по красноватой заливке шапки дня в Excel (ключ — номер дня месяца в графике).</summary>
    public Dictionary<int, bool> RestDayByDayFill { get; } = new();

    private static readonly System.Globalization.CultureInfo RuCulture =
        System.Globalization.CultureInfo.GetCultureInfo("ru-RU");

    private int DaysCoveredByShiftsInMonth()
    {
        var dom = Shifts.Select(s => s.Day).DefaultIfEmpty(0).Max();
        if (dom <= 0)
            return DateTime.DaysInMonth(ScheduleYear, ScheduleMonth ?? 1);
        return Math.Min(dom, ScheduleMonth is { } mm ? DateTime.DaysInMonth(ScheduleYear, mm) : dom);
    }

    private static string ShiftCellWho(DutyShiftAssignment? a) =>
        a is null
            ? "—"
            : string.IsNullOrWhiteSpace(a.MatchedPersonName)
                ? (string.IsNullOrWhiteSpace(a.RawCell) ? "—" : a.RawCell.Trim())
                : a.MatchedPersonName;

    /// <summary>Строка «01 мая - выходной» и т.п.</summary>
    private string DayKindTagLine(DateTime calendarDay)
    {
        var dom = calendarDay.Day;
        string kind;
        if (RestDayByDayFill.Count == 0)
            kind = "тип дня не сохранён";
        else
            kind = RestDayByDayFill.GetValueOrDefault(dom) ? "выходной" : "будний";

        var left = $"{calendarDay.ToString("dd MMMM", RuCulture)} - {kind}";
        return RuCulture.TextInfo.ToLower(left);
    }

    /// <summary>Первое сообщение после загрузки Excel: сводка разбора + календарь/выходные (без списка дежурных и без расписания по дням).</summary>
    public string BuildUploadInfoFirstMessage(string? fileHint = null)
    {
        var dayCount = Shifts.Select(s => s.Day).Distinct().Count();
        var matchedShifts = Shifts.Count(s => !string.IsNullOrEmpty(s.MatchedPersonName));
        var hint = string.IsNullOrEmpty(fileHint) ? "" : $"{EmDoc} {fileHint}\n\n";
        var period = ScheduleMonth is >= 1 and <= 12
            ? $"{EmMonth} {ScheduleLocaleFormatting.RussianMonthYear(ScheduleYear, ScheduleMonth.Value)}\n"
            : $"{EmMonth} месяц не определён автоматически\n";

        var sb = new System.Text.StringBuilder();
        sb.Append($"{hint}");
        sb.Append($"{period}\n");
        sb.AppendLine($"{EmCard} Что удалось считать:");
        sb.AppendLine($"{EmPeople} Контактов в справочнике: {People.Count.ToString(RuCulture)}");
        sb.AppendLine($"{EmDays} Дней с записями в графике: {dayCount.ToString(RuCulture)}");
        sb.AppendLine(
            $"{EmSchedule} Всего ячеек смен: {Shifts.Count.ToString(RuCulture)} (сопоставлено с ФИО: {matchedShifts.ToString(RuCulture)})");
        sb.AppendLine($"{EmBell} Предупреждений разбора: {Warnings.Count.ToString(RuCulture)}");

        if (ScheduleMonth is >= 1 and <= 12)
        {
            var calDays = DateTime.DaysInMonth(ScheduleYear, ScheduleMonth.Value);
            sb.AppendLine();
            sb.AppendLine($"{EmMonth} Месяц: {ScheduleLocaleFormatting.RussianMonthYear(ScheduleYear, ScheduleMonth.Value)}");
            sb.AppendLine($"{EmDays} Календарных дней: {calDays.ToString(RuCulture)}");

            var spanDays = DaysCoveredByShiftsInMonth();
            var markCount = RestDayByDayFill.Count;
            if (markCount > 0)
            {
                var rests = Enumerable.Range(1, spanDays).Count(d => RestDayByDayFill.GetValueOrDefault(d));
                sb.AppendLine($"{EmWeekend} Выходные: {rests.ToString(RuCulture)}");
                sb.AppendLine($"{EmWork} Будние: {(spanDays - rests).ToString(RuCulture)}");
            }
            else
            {
                sb.AppendLine($"{EmWeekend} Выходные: —");
                sb.AppendLine($"{EmWork} Будние: —");
            }
        }
        else
        {
            sb.AppendLine();
            sb.AppendLine(
                $"{EmMonth} Месяц: не удалось определить (нет названия месяца на видимом листе или в имени файла).");
            sb.AppendLine($"{EmDays} Календарных дней: —");
            sb.AppendLine($"{EmWeekend} Выходные: —");
            sb.AppendLine($"{EmWork} Будние: —");
        }

        return sb.ToString().TrimEnd();
    }

    /// <summary>Второе сообщение: дежурные и расписание по дням.</summary>
    public string BuildUploadDutyAndScheduleMessage(int maxLength = int.MaxValue)
    {
        var sb = new System.Text.StringBuilder();

        sb.AppendLine($"{EmPeople} Дежурные:");
        foreach (var p in People.OrderBy(x => x.Name, StringComparer.CurrentCultureIgnoreCase))
        {
            var posPart = string.IsNullOrWhiteSpace(p.Position) ? "" : $" {p.Position.Trim()};";
            sb.AppendLine($"   \u2728 {p.Name};{posPart} — {p.Phone}");
        }

        sb.AppendLine();
        sb.AppendLine($"{EmSchedule} Расписание:");

        if (ScheduleMonth is not (>= 1 and <= 12))
        {
            sb.AppendLine("— месяц графика не определён, список по дням недоступен.");
        }
        else
        {
            foreach (var g in Shifts.GroupBy(s => s.Day).OrderBy(x => x.Key))
            {
                var dom = g.Key;
                var cal = new DateTime(ScheduleYear, ScheduleMonth!.Value, dom);
                sb.AppendLine(DayKindTagLine(cal));

                var night = g.FirstOrDefault(x => x.Shift == DutyShift.Night);
                var dayShift = g.FirstOrDefault(x => x.Shift == DutyShift.Day);
                var single = g.FirstOrDefault(x => x.Shift == DutyShift.Unspecified);

                if (night != null || dayShift != null)
                {
                    if (night != null)
                        sb.AppendLine($"{EmNight} ночь: {ShiftCellWho(night)}");
                    if (dayShift != null)
                        sb.AppendLine($"{EmDayShift} день: {ShiftCellWho(dayShift)}");
                }
                else if (single != null)
                {
                    sb.AppendLine($"{EmDayGeneric} дежурный: {ShiftCellWho(single)}");
                }

                sb.AppendLine();
            }
        }

        if (Warnings.Count > 0)
        {
            sb.AppendLine($"{EmWarn} Предупреждения:");
            foreach (var w in Warnings.Take(12))
                sb.AppendLine($"   {EmWarn} {w}");
            if (Warnings.Count > 12)
                sb.AppendLine($"… ещё предупреждений: {Warnings.Count - 12}");
        }

        var s = sb.ToString().TrimEnd();
        return s.Length <= maxLength ? s : s[..maxLength] + "\n…(обрезано)";
    }
}
