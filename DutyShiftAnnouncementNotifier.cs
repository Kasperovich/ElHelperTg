using System.Collections.Concurrent;
using Telegram.Bot;
using Telegram.Bot.Types.ReplyMarkups;

/// <summary>При смене активной смены (день/ночь) рассылает одинаковое сообщение подписанным чатам.</summary>
internal static class DutyShiftAnnouncementNotifier
{
    private static readonly TimeSpan PollInterval = TimeSpan.FromSeconds(20);

    internal static async Task RunAsync(
        ITelegramBotClient botClient,
        ConcurrentDictionary<long, byte> subscriberChats,
        CancellationToken ct)
    {
        try
        {
            using var timer = new PeriodicTimer(PollInterval);
            DutySlashReports.ActiveDutyAnnKey? baseline = null;
            var seeded = false;

            while (await timer.WaitForNextTickAsync(ct).ConfigureAwait(false))
            {
                if (subscriberChats.IsEmpty)
                    continue;

                var announcement = await DutySlashReports.TryGetShiftStartAnnouncementAsync(DateTime.Now, ct)
                    .ConfigureAwait(false);
                var keyNow = announcement?.Key;

                if (!seeded)
                {
                    baseline = keyNow;
                    seeded = true;
                    continue;
                }

                if (DutySlashReports.SameDutyAnnouncementKey(baseline, keyNow))
                    continue;

                baseline = keyNow;

                if (announcement is not { } msg || string.IsNullOrEmpty(msg.Text))
                    continue;

                foreach (var chatId in subscriberChats.Keys.ToArray())
                {
                    try
                    {
                        await botClient.SendMessage(
                            chatId,
                            msg.Text,
                            replyMarkup: new ReplyKeyboardRemove(),
                            cancellationToken: ct).ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Уведомление о смене в чат {chatId}: {ex.Message}");
                    }
                }

                Console.WriteLine(
                    $"Рассылка «начало смены» ({msg.Key.ColumnDate:yyyy-MM-dd} {msg.Key.Shift}) — "
                    + $"{subscriberChats.Count} чат(ов).");
            }
        }
        catch (OperationCanceledException)
        {
            // штатная остановка
        }
    }
}
