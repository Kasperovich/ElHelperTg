using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>Сохранённые утверждённые графики по году и месяцу (JSON в data/schedules/approved).</summary>
internal static class ApprovedScheduleStore
{
    private static readonly JsonSerializerOptions SerializerOptions =
        new()
        {
            WriteIndented = true,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            Converters = { new JsonStringEnumConverter() },
        };

    internal static string ApprovedDirectory =>
        Path.Combine(AppContext.BaseDirectory, "data", "schedules", "approved");

    internal static string FilePathForMonth(int year, int month) =>
        Path.Combine(ApprovedDirectory, $"{year}-{month:00}.json");

    internal static bool Exists(int year, int month) =>
        File.Exists(FilePathForMonth(year, month));

    internal static async Task<DutyScheduleSnapshot?> TryLoadMonthAsync(int year, int month, CancellationToken ct)
    {
        var path = FilePathForMonth(year, month);
        if (!File.Exists(path))
            return null;

        try
        {
            await using var fs = new FileStream(
                path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 81920,
                FileOptions.Asynchronous | FileOptions.SequentialScan);

            return await JsonSerializer.DeserializeAsync<DutyScheduleSnapshot>(fs, SerializerOptions, ct)
                .ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Не удалось прочитать расписание {path}: {ex.Message}");
            return null;
        }
    }

    internal static async Task SaveApprovedAsync(DutyScheduleSnapshot snapshot, CancellationToken ct)
    {
        Directory.CreateDirectory(ApprovedDirectory);
        var path = FilePathForMonth(snapshot.ScheduleYear, snapshot.ScheduleMonth);
        var toWrite =
            snapshot with { SavedAtUtc = DateTimeOffset.UtcNow };

        await using var stream = new FileStream(
            path, FileMode.Create, FileAccess.Write, FileShare.Read, bufferSize: 81920,
            FileOptions.Asynchronous | FileOptions.SequentialScan);
        await JsonSerializer.SerializeAsync(stream, toWrite, SerializerOptions, ct).ConfigureAwait(false);
    }
}
