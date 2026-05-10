using System.Collections.Concurrent;
using Telegram.Bot;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;

/// <summary>Подтверждение корректности Excel и сохранение графика на месяц; перезапись — со вторым согласием.</summary>
internal static class ScheduleConfirmationFlow
{
    private enum SessionStep { AwaitingFirstDecision, AwaitingOverwriteDecision }

    private sealed class PendingState
    {
        public Guid Id { get; init; }
        public required long ChatId { get; init; }
        public required long TelegramUserId { get; init; }
        public required string SourceExcelPath { get; init; }
        public required string TelegramDocumentFileName { get; init; }
        public required DutyScheduleSnapshot Snapshot { get; init; }
        public SessionStep Step { get; set; } = SessionStep.AwaitingFirstDecision;
    }

    private static readonly ConcurrentDictionary<Guid, PendingState> Sessions = new();

    internal static Guid RegisterNewUpload(long chatId, long telegramUserId, string excelPath,
        string? telegramDocName, DutyScheduleSnapshot snapshot)
    {
        var id = Guid.NewGuid();

        if (LatestSessionByChat.TryGetValue(chatId, out var oldId))
            Sessions.TryRemove(oldId, out _);

        var state = new PendingState
        {
            Id = id,
            ChatId = chatId,
            TelegramUserId = telegramUserId,
            SourceExcelPath = excelPath,
            TelegramDocumentFileName = string.IsNullOrEmpty(telegramDocName) ? "—" : telegramDocName,
            Snapshot = snapshot,
        };

        Sessions[id] = state;
        LatestSessionByChat[chatId] = id;
        return id;
    }

    internal static readonly ConcurrentDictionary<long, Guid> LatestSessionByChat = new();

    private static string SessionKey(Guid sessionId) => sessionId.ToString("N");

    private static void TryDeleteSourceExcel(string absolutePath)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(absolutePath) || !File.Exists(absolutePath))
                return;

            File.Delete(absolutePath);
            Console.WriteLine($"Удалён исходный Excel после утверждения расписания: {absolutePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Не удалось удалить исходный Excel «{absolutePath}»: {ex.Message}");
        }
    }

    internal static InlineKeyboardMarkup FirstStepMarkup(Guid sessionId)
    {
        var s = SessionKey(sessionId);
        return new InlineKeyboardMarkup(
        [
            new[]
            {
                InlineKeyboardButton.WithCallbackData("\u2705 Да, взять в работу", $"s1y:{s}"),
                InlineKeyboardButton.WithCallbackData("\u274C Отмена", $"s1n:{s}"),
            },
        ]);
    }

    private static InlineKeyboardMarkup OverwriteStepMarkup(Guid sessionId)
    {
        var s = SessionKey(sessionId);
        return new InlineKeyboardMarkup(
        [
            new[]
            {
                InlineKeyboardButton.WithCallbackData("\U0001F504 Перезаписать", $"s2y:{s}"),
                InlineKeyboardButton.WithCallbackData("\u2B95 Оставить старый график", $"s2n:{s}"),
            },
        ]);
    }

    internal static Task SendConfirmationPromptAsync(
        ITelegramBotClient bot, long chatId, Guid sessionId, CancellationToken ct) =>
        bot.SendMessage(
            chatId,
            "\u2728 Посмотрите полный текст выше.\n\n"
            + "<b>Всё верно?</b> Если да — график сохранится на этот месяц и можно будет брать в работу.\n"
            + "<i>Новый файл в чате отменит этот запрос.</i>",
            replyMarkup: FirstStepMarkup(sessionId),
            parseMode: ParseMode.Html,
            cancellationToken: ct);

    internal static async Task HandleCallbackAsync(ITelegramBotClient bot, CallbackQuery cq, CancellationToken ct)
    {
        if (cq.Data is null)
        {
            await bot.AnswerCallbackQuery(cq.Id, cancellationToken: ct).ConfigureAwait(false);
            return;
        }

        var colon = cq.Data.IndexOf(':', StringComparison.Ordinal);
        if (colon <= 0 || colon >= cq.Data.Length - 1 ||
            cq.Message is null ||
            cq.From is null)
        {
            await bot.AnswerCallbackQuery(
                cq.Id,
                "Некорректный ответ кнопки.",
                showAlert: true,
                cancellationToken: ct).ConfigureAwait(false);
            return;
        }

        var prefix = cq.Data[..colon];
        if (!Guid.TryParseExact(cq.Data.AsSpan(colon + 1), "N", out var gid) ||
            !Sessions.TryGetValue(gid, out var state))
        {
            await bot.AnswerCallbackQuery(
                cq.Id,
                "Запрос устарел — пришлите Excel снова.",
                showAlert: true,
                cancellationToken: ct).ConfigureAwait(false);
            return;
        }

        if (cq.From.Id != state.TelegramUserId || cq.Message.Chat.Id != state.ChatId)
        {
            await bot.AnswerCallbackQuery(
                cq.Id,
                "Подтвердить может только отправитель файла.",
                showAlert: true,
                cancellationToken: ct).ConfigureAwait(false);
            return;
        }

        var chatId = cq.Message.Chat.Id;
        var messageId = cq.Message.MessageId;

        try
        {
            switch (prefix)
            {
                case "s1n":
                    if (state.Step != SessionStep.AwaitingFirstDecision)
                    {
                        await bot.AnswerCallbackQuery(cq.Id, cancellationToken: ct).ConfigureAwait(false);
                        return;
                    }

                    Sessions.TryRemove(gid, out _);
                    LatestSessionByChat.TryRemove(chatId, out _);
                    await bot.EditMessageText(
                        chatId, messageId,
                        "\u274C Файл <b>не</b> принят в работу. Можете прислать другой Excel.",
                        parseMode: ParseMode.Html,
                        replyMarkup: null,
                        cancellationToken: ct).ConfigureAwait(false);
                    await bot.AnswerCallbackQuery(cq.Id, cancellationToken: ct).ConfigureAwait(false);
                    return;

                case "s1y":
                    if (state.Step != SessionStep.AwaitingFirstDecision)
                    {
                        await bot.AnswerCallbackQuery(cq.Id, cancellationToken: ct).ConfigureAwait(false);
                        return;
                    }

                    var y = state.Snapshot.ScheduleYear;
                    var m = state.Snapshot.ScheduleMonth;
                    var monthTitle = ScheduleLocaleFormatting.RussianMonthYear(y, m);

                    if (ApprovedScheduleStore.Exists(y, m))
                    {
                        state.Step = SessionStep.AwaitingOverwriteDecision;
                        await bot.EditMessageText(
                            chatId, messageId,
                            $"<b>{monthTitle}</b> — для этого месяца график <b>уже сохранён</b>.\n"
                            + "\nЗаменить его данными из текущего файла?",
                            parseMode: ParseMode.Html,
                            replyMarkup: OverwriteStepMarkup(gid),
                            cancellationToken: ct).ConfigureAwait(false);
                    }
                    else
                    {
                        await ApprovedScheduleStore.SaveApprovedAsync(state.Snapshot, ct).ConfigureAwait(false);
                        TryDeleteSourceExcel(state.SourceExcelPath);
                        Sessions.TryRemove(gid, out _);
                        LatestSessionByChat.TryRemove(chatId, out _);
                        await bot.EditMessageText(
                            chatId, messageId,
                            $"\u2705 График на <b>{monthTitle}</b> сохранён и взят в работу.",
                            parseMode: ParseMode.Html,
                            replyMarkup: null,
                            cancellationToken: ct).ConfigureAwait(false);
                        DutyNotifySubscribers.EnsureSubscribed(chatId);
                        Console.WriteLine(
                            $"Утверждённое расписание: {ApprovedScheduleStore.FilePathForMonth(y, m)}");
                    }

                    await bot.AnswerCallbackQuery(cq.Id, cancellationToken: ct).ConfigureAwait(false);
                    return;

                case "s2n":
                    if (state.Step != SessionStep.AwaitingOverwriteDecision)
                    {
                        await bot.AnswerCallbackQuery(cq.Id, cancellationToken: ct).ConfigureAwait(false);
                        return;
                    }

                    Sessions.TryRemove(gid, out _);
                    LatestSessionByChat.TryRemove(chatId, out _);
                    await bot.EditMessageText(
                        chatId, messageId,
                        "\uD83D\uDCCC Оставляем уже сохранённый график на этот месяц. Новый файл можно прислать снова.",
                        parseMode: ParseMode.Html,
                        replyMarkup: null,
                        cancellationToken: ct).ConfigureAwait(false);
                    await bot.AnswerCallbackQuery(cq.Id, cancellationToken: ct).ConfigureAwait(false);
                    return;

                case "s2y":
                    if (state.Step != SessionStep.AwaitingOverwriteDecision)
                    {
                        await bot.AnswerCallbackQuery(cq.Id, cancellationToken: ct).ConfigureAwait(false);
                        return;
                    }

                    await ApprovedScheduleStore.SaveApprovedAsync(state.Snapshot, ct).ConfigureAwait(false);
                    TryDeleteSourceExcel(state.SourceExcelPath);
                    var my = state.Snapshot.ScheduleMonth;
                    var vy = state.Snapshot.ScheduleYear;
                    var mt2 = ScheduleLocaleFormatting.RussianMonthYear(vy, my);
                    Sessions.TryRemove(gid, out _);
                    LatestSessionByChat.TryRemove(chatId, out _);
                    await bot.EditMessageText(
                        chatId, messageId,
                        $"\u2705 График на <b>{mt2}</b> <b>перезаписан</b> текущим файлом и взят в работу.",
                        parseMode: ParseMode.Html,
                        replyMarkup: null,
                        cancellationToken: ct).ConfigureAwait(false);
                    DutyNotifySubscribers.EnsureSubscribed(chatId);
                    Console.WriteLine(
                        $"Перезаписано утверждённое расписание: {ApprovedScheduleStore.FilePathForMonth(vy, my)}");
                    await bot.AnswerCallbackQuery(cq.Id, cancellationToken: ct).ConfigureAwait(false);
                    return;
            }

            await bot.AnswerCallbackQuery(cq.Id, cancellationToken: ct).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Callback расписания: {ex.Message}");
            await bot.AnswerCallbackQuery(
                cq.Id,
                "Ошибка при сохранении — попробуйте снова.",
                showAlert: true,
                cancellationToken: ct).ConfigureAwait(false);
        }
    }
}
