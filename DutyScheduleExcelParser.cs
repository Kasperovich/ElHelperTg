using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

/// <summary>
/// Разбор графика дежурств: справочник (имя + телефон в строке) + таблица по дням.
/// Поддерживается макет «дни в строке, под ними строка ночи и строка дня», а также запасной вариант «дни в колонке».
/// </summary>
internal static partial class DutyScheduleExcelParser
{
    private sealed record HorizontalGridExtract(
        List<DutyShiftAssignment> Shifts,
        bool LooksLikeStandardDutyCorners,
        Dictionary<int, bool> RestByDayFill);

    private sealed record VerticalGridExtract(
        List<DutyShiftAssignment> Shifts,
        Dictionary<int, bool> RestByDayFill);

    private sealed record PickedDutySegments(
        List<DutyShiftAssignment> Shifts,
        Dictionary<int, bool> RestMarks,
        bool DerivedFromVertical);

    private const int MinDaySequenceLength = 21;

    private static readonly CultureInfo RuCulture = CultureInfo.GetCultureInfo("ru-RU");

    /// <summary>Строка заголовка таблицы: Таб №, Ф.И.О, должность, телефон, …</summary>
    private readonly record struct PersonnelColumns(int HeaderRow, int? TabCol, int FioCol, int? PositionCol, int PhoneCol, int? WorkDaysCol);

    [GeneratedRegex(@"\s+", RegexOptions.None)]
    private static partial Regex Whitespace();

    [GeneratedRegex(@"(?<!\d)(20[0-9]{2})(?!\d)", RegexOptions.None)]
    private static partial Regex ScheduleFourDigitYear();

    /// Соответствие подстроки (нормализованное значение уже в нижнем регистре) номеру месяца — порядок: длинные варианты раньше.
    private static readonly (string Marker, int MonthNumber)[] RussianMonthMarkerToNumber =
    [
        ("сентября", 9), ("сентябрь", 9),
        ("октября", 10), ("октябрь", 10),
        ("ноября", 11), ("ноябрь", 11),
        ("декабря", 12), ("декабрь", 12),
        ("января", 1), ("январь", 1),
        ("февраля", 2), ("февраль", 2),
        ("апреля", 4), ("апрель", 4),
        ("августа", 8), ("август", 8),
        ("марта", 3), ("март", 3),
        ("июля", 7), ("июль", 7),
        ("июня", 6), ("июнь", 6),
        ("мая", 5), ("май", 5),
    ];

    /// <summary>
    /// Подстроки для сопоставления имени месяца (длиннее раньше, чтобы не перепутать «март» и «мая», «июнь» и т.п.).
    /// </summary>
    private static readonly string[] RussianMonthMarkersByLengthDescending =
    [
        "сентября", "сентябрь",
        "октября", "октябрь",
        "ноября", "ноябрь",
        "декабря", "декабрь",
        "января", "январь",
        "февраля", "февраль",
        "апреля", "апрель",
        "августа", "август",
        "марта", "март",
        "июля", "июль",
        "июня", "июнь",
        "мая", "май"
    ];

    /// Скрытые и очень скрытые листы не разбираем — только <see cref="XLWorksheetVisibility.Visible"/>.
    private static IEnumerable<IXLWorksheet> VisibleWorksheets(XLWorkbook wb) =>
        wb.Worksheets.Where(ws => ws.Visibility == XLWorksheetVisibility.Visible);

    /// <summary>
    /// Раньше всегда брался первый лист книги — если «Март» первой вкладкой, а «Май» второй, разбирался март.
    /// Учитываются только видимые листы. Сначала пытаемся сопоставить месяц из имени файла с именем листа,
    /// иначе — видимый лист с самым узнаваемым графиком.
    /// </summary>
    private static IXLWorksheet SelectDutyWorksheet(XLWorkbook wb, string excelPathOnDisk)
    {
        var fileStem =
            HeaderNormalize(Path.GetFileNameWithoutExtension(excelPathOnDisk) ?? "");
        var visible = VisibleWorksheets(wb).ToList();
        if (visible.Count == 0)
            return wb.Worksheets.Worksheet(1);

        foreach (var month in RussianMonthMarkersByLengthDescending)
        {
            if (!MonthHintInText(fileStem, month))
                continue;

            foreach (var ws in visible)
            {
                if (MonthHintInText(HeaderNormalize(ws.Name), month))
                    return ws;
            }

            foreach (var ws in visible)
                if (!string.IsNullOrWhiteSpace(ws.Name) && SheetNameRoughlyMatchesFileStem(ws.Name, fileStem))
                    return ws;
        }

        var bestWs = PreferWorksheetWithDutyScores(wb);
        return bestWs ?? visible[0];
    }

    private static bool MonthHintInText(string normalizedHaystack, string monthMarkerNormalized)
    {
        // marker уже нижним регистром в массиве
        var m = monthMarkerNormalized.Trim();
        return normalizedHaystack.Contains(m, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Имя листа и имя файла явно связаны (общие слова длиннее 5 символов).</summary>
    private static bool SheetNameRoughlyMatchesFileStem(string sheetName, string fileStemNormalized)
    {
        if (fileStemNormalized.Length < 5)
            return false;

        var sn = HeaderNormalize(sheetName);
        foreach (var part in fileStemNormalized.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (part.Length < 5)
                continue;
            if (sn.Contains(part, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        foreach (var part in sn.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (part.Length < 5)
                continue;
            if (fileStemNormalized.Contains(part, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }

    /// <summary>Если вкладку по месяцу не нашли — берём ту, где сильнее похож блок графика 1…31.</summary>
    private static IXLWorksheet? PreferWorksheetWithDutyScores(XLWorkbook wb)
    {
        IXLWorksheet? best = null;
        var bestScore = int.MinValue;

        foreach (var ws in VisibleWorksheets(wb))
        {
            var rng = ws.RangeUsed();
            if (rng == null)
                continue;

            var mr = rng.LastRow().RowNumber();
            var mc = rng.LastColumn().ColumnNumber();
            var hzScore = TryExtractHorizontalDayGrid(ws, mr, mc);
            var vzScore = TryExtractVerticalDayGrid(ws, mr, mc);

            var horizontalCells = hzScore?.Shifts;
            var hoStrong = hzScore?.LooksLikeStandardDutyCorners ?? false;
            var verticalCells = vzScore?.Shifts;

            var hDaysCount = horizontalCells?.Select(static s => s.Day).Distinct().Count() ?? 0;
            var vDaysCount = verticalCells?.Select(static s => s.Day).Distinct().Count() ?? 0;

            var score =
                (hoStrong ? 120_000 : 0)
                + hDaysCount * 900
                + (horizontalCells?.Count ?? 0) * 12
                + vDaysCount * 40;

            if (score > bestScore)
            {
                bestScore = score;
                best = ws;
            }
        }

        return best;
    }

    private static (int Year, int? MonthNumber) InferScheduleYearMonth(string excelPathOnDisk, string sheetName)
    {
        var stemRaw = Path.GetFileNameWithoutExtension(excelPathOnDisk) ?? "";
        var stemNorm = HeaderNormalize(stemRaw);
        var sheetNorm = HeaderNormalize(sheetName);

        var year = DateTime.Now.Year;
        foreach (var src in new[] { stemRaw, sheetName })
        {
            var m = ScheduleFourDigitYear().Match(src);
            if (m.Success && int.TryParse(m.ValueSpan, CultureInfo.InvariantCulture, out var y) && y is >= 2000 and <= 2099)
            {
                year = y;
                break;
            }
        }

        int? monthNumber = null;
        foreach (var (marker, num) in RussianMonthMarkerToNumber)
        {
            if (MonthHintInText(sheetNorm, marker))
            {
                monthNumber = num;
                break;
            }
        }

        if (monthNumber == null)
        {
            foreach (var (marker, num) in RussianMonthMarkerToNumber)
            {
                if (MonthHintInText(stemNorm, marker))
                {
                    monthNumber = num;
                    break;
                }
            }
        }

        return (year, monthNumber);
    }

    private static PickedDutySegments? PickDutyGrid(HorizontalGridExtract? hz, VerticalGridExtract? vz)
    {
        var horizontal = hz?.Shifts;
        var horizontalLooksLikeDutyTable = hz?.LooksLikeStandardDutyCorners ?? false;
        var vertical = vz?.Shifts;

        var hDays = horizontal?.Select(static s => s.Day).Distinct().Count() ?? 0;
        var vDays = vertical?.Where(static s => s.Shift == DutyShift.Unspecified)
                          .Select(static s => s.Day).Distinct().Count()
                       ?? 0;

        if (horizontal != null &&
            horizontalLooksLikeDutyTable &&
            hDays >= MinDaySequenceLength)
            return new PickedDutySegments(horizontal, hz!.RestByDayFill, DerivedFromVertical: false);

        if (horizontal != null &&
            hDays >= MinDaySequenceLength &&
            (vertical == null || hDays >= vDays))
            return new PickedDutySegments(horizontal, hz!.RestByDayFill, DerivedFromVertical: false);

        if (vertical != null && vDays >= MinDaySequenceLength)
            return new PickedDutySegments(vertical, vz!.RestByDayFill, DerivedFromVertical: true);

        if (horizontal != null && hDays > 0)
            return new PickedDutySegments(horizontal, hz!.RestByDayFill, DerivedFromVertical: false);

        if (vertical != null && vDays > 0)
            return new PickedDutySegments(vertical, vz!.RestByDayFill, DerivedFromVertical: true);

        return null;
    }

    public static DutyScheduleParseResult Parse(string path)
    {
        var result = new DutyScheduleParseResult();
        try
        {
            using var wb = new XLWorkbook(path);
            var ws = SelectDutyWorksheet(wb, path);

            var (scheduleYear, scheduleMonth) = InferScheduleYearMonth(path, ws.Name);
            result.ScheduleYear = scheduleYear;
            result.ScheduleMonth = scheduleMonth;

            result.Warnings.Add($"Разбор книги Excel: используется лист «{ws.Name}».");

            var range = ws.RangeUsed();
            if (range == null)
            {
                result.Warnings.Add("Лист пуст или не содержит данных.");
                return result;
            }

            CollectPeople(ws, range, result);

            var maxRow = range.LastRow().RowNumber();
            var maxCol = range.LastColumn().ColumnNumber();
            var hz = TryExtractHorizontalDayGrid(ws, maxRow, maxCol);
            var vz = TryExtractVerticalDayGrid(ws, maxRow, maxCol);

            if (PickDutyGrid(hz, vz) is { } picked)
            {
                result.Shifts.AddRange(picked.Shifts);

                foreach (var kv in picked.RestMarks)
                    result.RestDayByDayFill[kv.Key] = kv.Value;

                if (picked.DerivedFromVertical && picked.RestMarks.Count == 0)
                    result.Warnings.Add(
                        "График по вертикали — заливка ячеек с датами не учитывалась (ожидается горизонтальная шапка дней).");
            }

            var replaced = new List<DutyShiftAssignment>();
            foreach (var d in result.Shifts.OrderBy(x => x.Day).ThenBy(x => x.Shift))
            {
                var token = ExtractScheduleLetterToken(d.RawCell);
                var match = string.IsNullOrEmpty(token)
                    ? null
                    : result.People.FirstOrDefault(p =>
                        string.Equals(p.ScheduleLetter, token, StringComparison.OrdinalIgnoreCase));

                replaced.Add(match != null
                    ? d with { MatchedPersonName = match.Name }
                    : d);

                if (match == null && !string.IsNullOrWhiteSpace(d.RawCell))
                {
                    var shiftLabel = d.Shift switch
                    {
                        DutyShift.Night => "ночь",
                        DutyShift.Day => "день",
                        _ => "смена"
                    };
                    result.Warnings.Add(
                        $"День {d.Day} ({shiftLabel}): «{d.RawCell}» не сопоставлено со справочником.");
                }
            }

            result.Shifts.Clear();
            result.Shifts.AddRange(replaced);
        }
        catch (Exception ex)
        {
            result.Warnings.Add($"Ошибка чтения Excel: {ex.Message}");
        }

        if (result.People.Count == 0)
            result.Warnings.Add(
                "Не найдено ни одной строки справочника: нужны ФИО и непустая ячейка в колонке «телефон».");
        if (result.Shifts.Count == 0)
            result.Warnings.Add("Не найдена строка дней 1…N (горизонтально или вертикально) — возможно объединённые ячейки.");

        return result;
    }

    /// <summary>
    /// Дни в одной строке (часто со 2-й колонки: «День/ночь» | 1 | 2 | …);
    /// следующая строка — ночь, ниже — день.
    /// </summary>
    private static HorizontalGridExtract? TryExtractHorizontalDayGrid(IXLWorksheet ws, int maxRow, int maxCol)
    {
        (int HeaderRow, int StartCol, int Len, bool SwapNightAndDayRows, int Score)? best = null;

        for (var headerRow = 1; headerRow <= maxRow - 2; headerRow++)
        {
            for (var startCol = 1; startCol <= maxCol; startCol++)
            {
                if (!TryGetDayNumber(ws.Cell(headerRow, startCol), out var first) || first != 1)
                    continue;

                var length = 1;
                while (startCol + length <= maxCol &&
                       TryGetDayNumber(ws.Cell(headerRow, startCol + length), out var next) &&
                       next == length + 1)
                    length++;

                if (length < MinDaySequenceLength)
                    continue;

                var swap = ShouldSwapNightAndDayRows(ws, headerRow, startCol);
                var score = HorizontalDutyCandidateScore(ws, headerRow, startCol, length);
                if (!best.HasValue || score > best.Value.Score ||
                    (score == best.Value.Score && length > best.Value.Len))
                    best = (headerRow, startCol, length, swap, score);
            }
        }

        if (best == null)
            return null;

        var (hr, sc, len, swapNightDay, chosenScore) = best.Value;

        var hasCornerMarkers = HorizontalDutyLooksLikeAnnotatedTable(ws, hr, sc);
        var labeledStrong = chosenScore >= 4000;

        var nightRowOffset = swapNightDay ? 2 : 1;
        var dayRowOffset = swapNightDay ? 1 : 2;

        var rests = new Dictionary<int, bool>(len);
        for (var day = 1; day <= len; day++)
            rests[day] = ExcelDutyDayFillAnalyzer.CellLooksLikeRestDayFill(ws.Cell(hr, sc + day - 1));

        var list = new List<DutyShiftAssignment>(len * 2);
        for (var day = 1; day <= len; day++)
        {
            var col = sc + day - 1;
            var nightRaw = CellText(ws.Cell(hr + nightRowOffset, col));
            var dayRaw = CellText(ws.Cell(hr + dayRowOffset, col));
            list.Add(new DutyShiftAssignment(day, DutyShift.Night, nightRaw, null));
            list.Add(new DutyShiftAssignment(day, DutyShift.Day, dayRaw, null));
        }

        return new HorizontalGridExtract(list, hasCornerMarkers || labeledStrong, rests);
    }

    /// <summary>Бонус, если рядом с первым днём стоят подписи «День/ночь», «Ночь», «День».</summary>
    private static bool HorizontalDutyLooksLikeAnnotatedTable(IXLWorksheet ws, int headerRow, int dayStartCol) =>
        dayStartCol > 1
        && (LooksLikeDutyDayNightHeader(CellText(ws.Cell(headerRow, dayStartCol - 1)))
            || HasNightAndDayBandLabels(ws, headerRow, dayStartCol));

    private static bool LooksLikeDutyDayNightHeader(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return false;
        var kn = HeaderNormalize(raw);
        return kn.Contains("ноч") && kn.Contains("день");
    }

    private static bool HasNightAndDayBandLabels(IXLWorksheet ws, int headerRow, int dayStartCol)
    {
        var c = dayStartCol - 1;
        var a = HeaderNormalize(CellText(ws.Cell(headerRow + 1, c)));
        var b = HeaderNormalize(CellText(ws.Cell(headerRow + 2, c)));
        return (LooksLikeNightRowLabel(a) && LooksLikeDayRowLabel(b))
               || (LooksLikeNightRowLabel(b) && LooksLikeDayRowLabel(a));
    }

    /// <returns>Чем выше тем предпочтительнее блок (перебиваем ложную «вертикаль»).</returns>
    private static int HorizontalDutyCandidateScore(
        IXLWorksheet ws,
        int headerRow,
        int dayStartCol,
        int consecutiveDays)
    {
        var score = consecutiveDays * 100;

        if (HorizontalDutyLooksLikeAnnotatedTable(ws, headerRow, dayStartCol))
            score += 5000;

        if (dayStartCol > 2)
            score += 10;

        score += consecutiveDays switch
        {
            >= 28 and <= 31 => 120,
            _ => 0
        };

        return score;
    }

    /// <summary>По подписям слева от блока дней определяем, не перепутаны ли строки ночь/день.</summary>
    /// <returns>true, если под строкой с «1» сначала идёт день, потом ночь — поменять строки местами при чтении.</returns>
    private static bool ShouldSwapNightAndDayRows(IXLWorksheet ws, int headerRow, int startCol)
    {
        if (startCol <= 1)
            return false;

        var labelCol = startCol - 1;
        var labelRow1 = CellText(ws.Cell(headerRow + 1, labelCol));
        var labelRow2 = CellText(ws.Cell(headerRow + 2, labelCol));
        return LooksLikeDayRowLabel(labelRow1) && LooksLikeNightRowLabel(labelRow2);
    }

    private static bool LooksLikeNightRowLabel(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;
        return text.Contains("ноч", StringComparison.OrdinalIgnoreCase);
    }

    private static bool LooksLikeDayRowLabel(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;
        if (text.Contains("день", StringComparison.OrdinalIgnoreCase))
            return true;
        return text.StartsWith("дн", StringComparison.OrdinalIgnoreCase)
               && !text.Contains("ноч", StringComparison.OrdinalIgnoreCase);
    }

    private static VerticalGridExtract? TryExtractVerticalDayGrid(IXLWorksheet ws, int maxRow, int maxCol)
    {
        (int StartRow, int DayCol, int Length, int DutyCol)? best = null;

        for (var dayCol = 1; dayCol <= maxCol; dayCol++)
        {
            for (var r = 1; r <= maxRow - MinDaySequenceLength; r++)
            {
                if (!TryGetDayNumber(ws.Cell(r, dayCol), out var start) || start != 1)
                    continue;

                var length = 1;
                while (r + length <= maxRow &&
                       TryGetDayNumber(ws.Cell(r + length, dayCol), out var next) &&
                       next == length + 1)
                    length++;

                if (length < MinDaySequenceLength)
                    continue;

                var dutyCol = GuessDutyColumn(ws, r, dayCol, length);
                if (dutyCol == 0)
                    continue;

                if (!best.HasValue || length > best.Value.Length)
                    best = (r, dayCol, length, dutyCol);
            }
        }

        if (best == null)
            return null;

        var (startRow, dCol, len, dutyColumnIndex) = best.Value;

        var rests = new Dictionary<int, bool>(len);
        for (var day = 1; day <= len; day++)
            rests[day] = ExcelDutyDayFillAnalyzer.CellLooksLikeRestDayFill(ws.Cell(startRow + day - 1, dCol));

        var list = new List<DutyShiftAssignment>(len);
        for (var day = 1; day <= len; day++)
        {
            var row = startRow + day - 1;
            var dutyText = CellText(ws.Cell(row, dutyColumnIndex));
            list.Add(new DutyShiftAssignment(day, DutyShift.Unspecified, dutyText, null));
        }

        return new VerticalGridExtract(list, rests);
    }

    private static void CollectPeople(IXLWorksheet ws, IXLRange range, DutyScheduleParseResult result)
    {
        var maxRowFromRange = range.LastRow().RowNumber();
        var lastRowOnSheet = ws.LastRowUsed()?.RowNumber();
        var maxRow = lastRowOnSheet.HasValue
            ? Math.Max(maxRowFromRange, lastRowOnSheet.Value)
            : maxRowFromRange;

        if (TryDetectPersonnelColumns(ws, range) is { } map)
            CollectPeopleFromPersonnelTable(ws, map, maxRow, result);
        else
            CollectPeopleHeuristic(ws, range, maxRow, result);
    }

    private static PersonnelColumns? TryDetectPersonnelColumns(IXLWorksheet ws, IXLRange range)
    {
        var firstRow = range.FirstRow().RowNumber();
        var lastSearchRow = Math.Min(range.LastRow().RowNumber(), firstRow + 80);
        var maxCol = range.LastColumn().ColumnNumber();

        for (var r = firstRow; r <= lastSearchRow; r++)
        {
            if (TryParsePersonnelHeaderRow(ws, r, maxCol, out var cols))
                return cols;
        }

        return null;
    }

    private static bool TryParsePersonnelHeaderRow(
        IXLWorksheet ws, int headerRow, int maxCol,
        out PersonnelColumns columns)
    {
        int? tab = null;
        var fio = (int?)null;
        int? position = null;
        var phone = (int?)null;
        int? workDays = null;

        for (var c = 1; c <= maxCol; c++)
        {
            var key = HeaderNormalize(CellText(ws.Cell(headerRow, c)));
            if (string.IsNullOrEmpty(key))
                continue;

            if (IsFioColumnHeader(key))
            {
                fio ??= c;
                continue;
            }

            if (IsPhoneColumnHeader(key))
            {
                phone ??= c;
                continue;
            }

            if (IsTabColumnHeader(key))
            {
                tab ??= c;
                continue;
            }

            if (IsPositionColumnHeader(key))
            {
                position ??= c;
                continue;
            }

            if (IsWorkDaysColumnHeader(key))
            {
                workDays ??= c;
                continue;
            }
        }

        if (fio is not { } fioCol || phone is not { } phoneCol)
        {
            columns = default;
            return false;
        }

        columns = new PersonnelColumns(headerRow, tab, fioCol, position, phoneCol, workDays);
        return true;
    }

    private static string HeaderNormalize(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return string.Empty;

        var s = Whitespace().Replace(raw.Trim().ToLower(RuCulture), " ");
        return s.Replace('ё', 'е');
    }

    private static bool IsFioColumnHeader(string key) =>
        key.Replace(".", "").Contains("фио", StringComparison.Ordinal);

    private static bool IsPhoneColumnHeader(string key) =>
        key.Contains("телефон", StringComparison.Ordinal)
        || key.Contains("№ тел", StringComparison.Ordinal);

    private static bool IsTabColumnHeader(string key) =>
        key.Contains("таб", StringComparison.Ordinal);

    private static bool IsPositionColumnHeader(string key) =>
        key.Contains("должност", StringComparison.Ordinal);

    private static bool IsWorkDaysColumnHeader(string key) =>
        (key.Contains("рабоч", StringComparison.Ordinal) && key.Contains("дн", StringComparison.Ordinal))
        || key.Contains("кол-во", StringComparison.Ordinal)
        || key.Contains("количеств", StringComparison.Ordinal);

    private const int PersonnelTableMaxScanRows = 2000;
    private const int PersonnelEmptyRowAbort = 75;

    /// <summary>Отсекаем фрагменты таблицы (цифры в колонке ФИО, «к.», «л»): не поднимаем предупреждение, просто пропускаем строку.</summary>
    private static bool IsPlausiblePersonnelName(string rawName)
    {
        if (string.IsNullOrWhiteSpace(rawName))
            return false;

        var n = Whitespace().Replace(rawName.Trim(), " ");
        var letterCount = n.Count(char.IsLetter);
        if (letterCount < 3)
            return false;

        var digitCount = n.Count(char.IsDigit);
        if (digitCount > letterCount)
            return false;

        return true;
    }

    private static void CollectPeopleFromPersonnelTable(
        IXLWorksheet ws, PersonnelColumns cols, int maxRow, DutyScheduleParseResult result)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var emptyStreak = 0;
        var scanTo = Math.Min(maxRow, cols.HeaderRow + PersonnelTableMaxScanRows);

        for (var r = cols.HeaderRow + 1; r <= scanTo; r++)
        {
            var name = CellText(ws.Cell(r, cols.FioCol));
            var phoneRaw = CellText(ws.Cell(r, cols.PhoneCol));
            var tabRaw = cols.TabCol is int tc ? CellText(ws.Cell(r, tc)).Trim() : null;
            var posRaw = cols.PositionCol is int pc ? CellText(ws.Cell(r, pc)).Trim() : null;
            var workRaw = cols.WorkDaysCol is int wc ? CellText(ws.Cell(r, wc)).Trim() : null;

            var rowEmpty =
                string.IsNullOrWhiteSpace(name)
                && string.IsNullOrWhiteSpace(phoneRaw)
                && string.IsNullOrWhiteSpace(tabRaw)
                && string.IsNullOrWhiteSpace(posRaw)
                && string.IsNullOrWhiteSpace(workRaw);
            if (rowEmpty)
            {
                emptyStreak++;
                if (emptyStreak >= PersonnelEmptyRowAbort)
                    break;
                continue;
            }

            emptyStreak = 0;

            if (string.IsNullOrWhiteSpace(name))
            {
                if (!string.IsNullOrWhiteSpace(phoneRaw))
                {
                    result.Warnings.Add(
                        $"Справочник, строка {r}: колонка ФИО пуста, но есть телефон «{phoneRaw}» — строка пропущена (проверьте объединение ячеек).");
                }

                continue;
            }

            if (!IsPlausiblePersonnelName(name))
                continue;

            if (string.IsNullOrWhiteSpace(phoneRaw))
            {
                result.Warnings.Add(
                    $"Справочник, строка {r}: есть ФИО «{name}», но ячейка телефона пуста.");
                continue;
            }

            var phoneStored = phoneRaw.Trim();
            var dedupe =
                $"{tabRaw}|{phoneStored}|{name.Trim()}";
            if (!seen.Add(dedupe))
            {
                result.Warnings.Add(
                    $"Справочник, строка {r}: дубликат записи (таб/телефон/ФИО как у другой строки) — «{name}» пропущено.");
                continue;
            }

            var letter = FirstLetterOfSurname(name);
            string? tab = string.IsNullOrWhiteSpace(tabRaw) ? null : tabRaw;
            string? pos = string.IsNullOrWhiteSpace(posRaw) ? null : posRaw;
            result.People.Add(new DutyPerson(name, phoneStored, letter, tab, pos));
        }
    }

    private static void CollectPeopleHeuristic(
        IXLWorksheet ws, IXLRange range, int sheetLastRowInclusive, DutyScheduleParseResult result)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var maxRowFromRange = range.LastRow().RowNumber();
        var maxRow = Math.Max(maxRowFromRange, sheetLastRowInclusive);
        var maxCol = range.LastColumn().ColumnNumber();

        for (var r = range.FirstRow().RowNumber(); r <= maxRow; r++)
        {
            string? phone = null;
            var phoneCol = 0;

            for (var c = 1; c <= maxCol; c++)
            {
                var text = CellText(ws.Cell(r, c));
                if (!string.IsNullOrWhiteSpace(text) && text.Count(char.IsDigit) >= 5)
                {
                    phone = text.Trim();
                    phoneCol = c;
                    break;
                }
            }

            if (phone == null)
                continue;

            string? bestName = null;
            var bestScore = 0;

            for (var c = 1; c <= maxCol; c++)
            {
                if (c == phoneCol)
                    continue;
                var candidate = CellText(ws.Cell(r, c));
                if (string.IsNullOrWhiteSpace(candidate))
                    continue;
                var score = NameHeuristic(candidate);
                if (score > bestScore)
                {
                    bestScore = score;
                    bestName = candidate;
                }
            }

            if (bestScore == 0)
            {
                var left = phoneCol > 1 ? CellText(ws.Cell(r, phoneCol - 1)) : null;
                bestName = string.IsNullOrWhiteSpace(left) ? $"Контакт {(phone.Length >= 4 ? phone[^4..] : phone)}" : left;
            }

            if (bestName == null || !IsPlausiblePersonnelName(bestName))
                continue;

            var key = $"{phone}\t{(bestName ?? "").Trim()}";
            if (!seen.Add(key))
                continue;

            var letter = FirstLetterOfSurname(bestName!);
            result.People.Add(new DutyPerson(bestName!, phone!, letter, null, null));
        }
    }

    private static int NameHeuristic(string candidate)
    {
        if (!candidate.Any(char.IsLetter))
            return 0;
        var letters = candidate.Count(char.IsLetter);
        var cyr = HasCyrillic(candidate);
        var score = letters + (cyr ? 5 : 0);
        if (candidate.Any(char.IsDigit))
            score -= 3;
        return score;
    }

    private static int GuessDutyColumn(IXLWorksheet ws, int startRow, int dayCol, int length)
    {
        var maxCol = ws.LastColumnUsed()?.ColumnNumber() ?? dayCol;

        foreach (var col in new[] { dayCol + 1, dayCol - 1 })
        {
            if (col < 1 || col > maxCol)
                continue;

            var letterish = 0;
            var nonEmpty = 0;

            for (var d = 0; d < length; d++)
            {
                var t = CellText(ws.Cell(startRow + d, col));
                if (string.IsNullOrWhiteSpace(t))
                    continue;
                nonEmpty++;
                var trimmed = t.Trim();
                if (trimmed.Any(char.IsLetter) && trimmed.Length <= 4)
                    letterish++;
            }

            if (nonEmpty >= length / 4 && letterish >= Math.Max(3, length / 6))
                return col;
        }

        return dayCol + 1 <= maxCol ? dayCol + 1 : dayCol - 1 >= 1 ? dayCol - 1 : 0;
    }

    private static string CellText(IXLCell cell)
    {
        try
        {
            return Whitespace().Replace(cell.GetFormattedString().Trim(), " ");
        }
        catch
        {
            return string.Empty;
        }
    }

    private static bool TryGetDayNumber(IXLCell cell, out int day)
    {
        day = 0;
        if (cell.IsEmpty())
            return false;

        if (cell.DataType == XLDataType.Number && cell.TryGetValue(out double dv))
        {
            var rounded = (int)Math.Round(dv, MidpointRounding.AwayFromZero);
            if (Math.Abs(dv - rounded) < 1e-6 && rounded is >= 1 and <= 31)
            {
                day = rounded;
                return true;
            }
        }

        var text = CellText(cell);
        return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out day)
               && day is >= 1 and <= 31;
    }

    private static bool HasCyrillic(string s)
    {
        foreach (var ch in s)
        {
            if (ch is >= '\u0400' and <= '\u052F')
                return true;
        }

        return false;
    }

    private static string FirstLetterOfSurname(string name)
    {
        name = Whitespace().Replace(name.Trim(), " ");
        if (string.IsNullOrEmpty(name))
            return "?";

        var parts = name.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        foreach (var part in parts)
        {
            if (!part.Any(char.IsLetter))
                continue;
            return char.ToUpper(part.Trim()[0], RuCulture).ToString();
        }

        return char.ToUpper(name[0], RuCulture).ToString();
    }

    private static string ExtractScheduleLetterToken(string rawCell)
    {
        var t = Whitespace().Replace(rawCell.Trim(), " ");
        t = TrimTrailingDutyDecorations(t);
        if (string.IsNullOrEmpty(t))
            return string.Empty;

        foreach (var ch in t)
        {
            if (!char.IsLetter(ch))
                continue;
            return NormalizeDutyLetterToUpperToken(ch);
        }

        return string.Empty;
    }

    private static string TrimTrailingDutyDecorations(string t)
    {
        while (t.Length > 0 && (char.IsPunctuation(t[^1]) || char.IsSeparator(t[^1])))
            t = t[..^1];

        return t.TrimEnd();
    }

    private static string NormalizeDutyLetterToUpperToken(char ch)
    {
        if (TryMapLatinDutyInitial(ch) is { } mapped)
            return mapped;
        return char.ToUpper(ch, RuCulture).ToString();
    }

    /// <summary>Excel/клавиатура иногда дают латиницу вместо кириллической однобуквы (p → п).</summary>
    private static string? TryMapLatinDutyInitial(char ch)
    {
        if (ch is (< 'A' or > 'z') or (> 'Z' and < 'a'))
            return null;

        return char.ToUpperInvariant(ch) switch
        {
            'P' => "П",
            'L' => "Л",
            'K' => "К",
            'R' => "Р",
            'V' => "В",
            'B' => "В",
            'N' => "Н",
            'H' => "Н",
            'M' => "М",
            'D' => "Д",
            'G' => "Г",
            'E' => "Е",
            'I' => "И",
            'S' => "С",
            'T' => "Т",
            'Y' => "У",
            'O' => "О",
            'F' => "Ф",
            'A' => "А",
            'C' => "С",
            'X' => "Х",
            'Z' => "З",
            'W' => "Ш",
            'Q' => "Я",
            'J' => "Ж",
            'U' => "Ц",
            _ => null
        };
    }
}
