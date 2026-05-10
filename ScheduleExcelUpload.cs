using Telegram.Bot;
using Telegram.Bot.Types;

/// <summary>
/// Сохраняет Excel с расписанием, присланный как Document. Парсинг — отдельный шаг позже.
/// </summary>
internal static class ScheduleExcelUpload
{
    private static readonly string[] ExcelExtensions = [".xlsx", ".xls"];

    public static bool LooksLikeExcel(Document document)
    {
        var ext = Path.GetExtension(document.FileName ?? string.Empty);
        if (!string.IsNullOrEmpty(ext) &&
            ExcelExtensions.Contains(ext, StringComparer.OrdinalIgnoreCase))
            return true;

        return document.MimeType is "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
               or "application/vnd.ms-excel";
    }

    /// <returns>(успех, полный путь при успехе; при ошибке — текст для сообщения пользователю)</returns>
    public static async Task<(bool ok, string path, string? errorReply)> SaveAsync(
        ITelegramBotClient bot,
        Document document,
        CancellationToken cancellationToken)
    {
        if (!LooksLikeExcel(document))
        {
            return (false, string.Empty,
                "Нужен файл Excel с расширением .xlsx или .xls (или с подходящим MIME-типом).");
        }

        var dir = Path.Combine(AppContext.BaseDirectory, "data", "schedules");
        Directory.CreateDirectory(dir);

        var original = SanitizeFileName(document.FileName ?? "schedule.xlsx");
        var stamped = $"{DateTime.UtcNow:yyyyMMdd_HHmmss}_{document.FileUniqueId}_{original}";
        var fullPath = Path.Combine(dir, stamped);

        var file = await bot.GetFile(document.FileId, cancellationToken);
        if (string.IsNullOrEmpty(file.FilePath))
        {
            return (false, string.Empty,
                "Telegram не вернул путь к файлу. Попробуйте отправить файл ещё раз.");
        }

        await using (var stream = new FileStream(fullPath, FileMode.CreateNew, FileAccess.Write,
                     FileShare.None, bufferSize: 81920, useAsync: true))
        {
            await bot.DownloadFile(file.FilePath, stream, cancellationToken);
        }

        return (true, fullPath, null);
    }

    private static string SanitizeFileName(string name)
    {
        var fileName = Path.GetFileName(name.Trim());
        foreach (var c in Path.GetInvalidFileNameChars())
            fileName = fileName.Replace(c, '_');

        if (string.IsNullOrWhiteSpace(fileName))
            fileName = "schedule.xlsx";

        return fileName.Length > 120 ? fileName[..120] : fileName;
    }
}
