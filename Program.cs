using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using System.Collections.Concurrent;
using System.Reflection;
using Telegram.Bot;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;

var builder = Host.CreateApplicationBuilder(args);
builder.Configuration.AddUserSecrets(Assembly.GetExecutingAssembly());

string? ResolveBotToken()
{
    string? nz(string? s) =>
        string.IsNullOrWhiteSpace(s) ? null : s.Trim();

    // Конфиг (appsettings*, BotConfiguration__BotToken, dotnet user-secrets) и популярные имена переменных в контейнере/ПaaS.
    foreach (var s in new[]
             {
                 builder.Configuration["BotConfiguration:BotToken"],
                 Environment.GetEnvironmentVariable("BOT_TOKEN"),
                 Environment.GetEnvironmentVariable("TELEGRAM_BOT_TOKEN"),
                 Environment.GetEnvironmentVariable("TG_BOT_TOKEN"),
                 Environment.GetEnvironmentVariable("BotConfiguration__BotToken"),
             })
    {
        var t = nz(s);
        if (t != null)
            return t;
    }

    return null;
}

var token = ResolveBotToken();
if (token is null)
{
    throw new InvalidOperationException(
        "Не задан токен бота.\n\n"
        + "Docker / хостинг: задайте переменную окружения BOT_TOKEN (или TELEGRAM_BOT_TOKEN, TG_BOT_TOKEN) "
        + "со значением токена @BotFather, либо BotConfiguration__BotToken.\n"
        + "Пример Docker: docker run -e BOT_TOKEN=\"<token>\" …\n\n"
        + "Локально: dotnet user-secrets set \"BotConfiguration:BotToken\" \"<token>\"; "
        + "или $env:BOT_TOKEN=\"<token>\"; dotnet run; "
        + "или appsettings.json рядом с сборкой / проектом: {\"BotConfiguration\":{\"BotToken\":\"...\"}}.");
}

using var cts = new CancellationTokenSource();
Console.CancelKeyPress += (_, e) =>
{
    e.Cancel = true;
    cts.Cancel();
};

var subscribedChats = new ConcurrentDictionary<long, byte>();
var bot = new TelegramBotClient(token);
var receiverOptions = new ReceiverOptions
{
    AllowedUpdates = [UpdateType.Message, UpdateType.CallbackQuery],
};

bot.StartReceiving(HandleUpdateAsync, HandlePollingErrorAsync, receiverOptions, cts.Token);
_ = RunPeriodicTimeNotifierAsync(bot, subscribedChats, cts.Token);
_ = DutyShiftAnnouncementNotifier.RunAsync(bot, DutyNotifySubscribers.Chats, cts.Token);

var me = await bot.GetMe(cts.Token);
Console.WriteLine($"Бот @{me.Username} слушает обновления. Ctrl+C — выход.");
try
{
    await bot.SetMyCommands(
    [
        new BotCommand { Command = "start", Description = "О боте и командах графика" },
        new BotCommand
        {
            Command = "duty_now",
            Description = "Кто дежурит сейчас",
        },
        new BotCommand { Command = "duty_today", Description = "Дежурства на сутки (день и ночь)" },
        new BotCommand { Command = "duty_twodays", Description = "Дежурства на два дня" },
        new BotCommand { Command = "duty_week", Description = "Дежурства на 7 дней (с сегодня)" },
        new BotCommand { Command = "duty_month", Description = "Дежурства на весь месяц" },
        new BotCommand { Command = "duty_help", Description = "Справка по командам графика" },
        new BotCommand { Command = "duty_notify_on", Description = "Уведомления о начале смены в этот чат" },
        new BotCommand { Command = "duty_notify_off", Description = "Отключить уведомления о сменах" },
    ],
        cancellationToken: cts.Token);
}
catch (Exception ex)
{
    Console.WriteLine($"SetMyCommands: {ex.Message}");
}

try
{
    await Task.Delay(Timeout.Infinite, cts.Token);
}
catch (OperationCanceledException)
{
    // штатная остановка
}

async Task HandleUpdateAsync(ITelegramBotClient botClient, Update update, CancellationToken ct)
{
    if (update.CallbackQuery is { } cq)
    {
        await ScheduleConfirmationFlow.HandleCallbackAsync(botClient, cq, ct);
        return;
    }

    if (update.Message is not { } message)
        return;

    if (message.Document is { } document)
    {
        try
        {
            var (ok, path, err) = await ScheduleExcelUpload.SaveAsync(botClient, document, ct);
            if (!ok)
            {
                if (!string.IsNullOrWhiteSpace(err))
                    await botClient.SendMessage(message.Chat.Id, err, cancellationToken: ct);
                return;
            }

            Console.WriteLine($"Принят Excel для разбора: {path}");
            try
            {
                var parsed = DutyScheduleExcelParser.Parse(path);
                var uploadInfo = parsed.BuildUploadInfoFirstMessage(document.FileName);
                await botClient.SendMessage(
                    message.Chat.Id,
                    uploadInfo,
                    replyMarkup: new ReplyKeyboardRemove(),
                    cancellationToken: ct);

                var fullReport = parsed.BuildUploadDutyAndScheduleMessage();
                await TelegramChatSend.SendLongTextAsync(
                    botClient,
                    message.Chat.Id,
                    fullReport,
                    ct,
                    new ReplyKeyboardRemove());

                var snap =
                    DutyScheduleSnapshotFactory.TryFromParse(parsed, document.FileName ?? "—", path);
                if (snap is null)
                {
                    await botClient.SendMessage(
                        message.Chat.Id,
                        "📆 Месяц графика не определён — сохранить его на календарный месяц нельзя. "
                        + "Укажите месяц в имени файла или в названии видимого листа Excel и пришлите файл снова.",
                        cancellationToken: ct);
                }
                else if (message.From is null)
                {
                    await botClient.SendMessage(
                        message.Chat.Id,
                        "Не удалось определить отправителя в Telegram — подтверждение по кнопкам недоступно.",
                        cancellationToken: ct);
                }
                else
                {
                    var sid = ScheduleConfirmationFlow.RegisterNewUpload(
                        message.Chat.Id,
                        message.From.Id,
                        path,
                        document.FileName,
                        snap);

                    await ScheduleConfirmationFlow.SendConfirmationPromptAsync(botClient, message.Chat.Id, sid,
                        ct);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Разбор расписания: {ex.Message}");
                await botClient.SendMessage(
                    message.Chat.Id,
                    "Не удалось разобрать Excel — пришлите скрин структуры или уточните верстку.",
                    cancellationToken: ct);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при загрузке Excel: {ex.Message}");
            await botClient.SendMessage(message.Chat.Id,
                "Не удалось сохранить файл. Попробуйте ещё раз или проверьте размер файла.",
                cancellationToken: ct);
        }

        return;
    }

    if (message.Text is not { } text)
        return;

    var rawCommand = text.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
        .FirstOrDefault()
        ?.Split('@', 2)[0];

    var command = rawCommand?.ToLowerInvariant();

    async Task ReplyDutyReportAsync(Func<CancellationToken, Task<string>> build)
    {
        var reply = await build(ct);
        await TelegramChatSend.SendLongTextAsync(
            botClient,
            message.Chat.Id,
            reply,
            ct,
            new ReplyKeyboardRemove());
    }

    switch (command)
    {
        case "/start":
            await botClient.SendMessage(
                message.Chat.Id,
                "Привет! Команды по графику — в меню слева от поля ввода (иконка «/»). "
                + "Чтобы загрузить новый график, отправьте файл Excel в чат.\n"
                + "Справка: /duty_help",
                replyMarkup: new ReplyKeyboardRemove(),
                cancellationToken: ct);
            break;

        case "/duty_now":
            await ReplyDutyReportAsync(DutySlashReports.DutyNowAsync);
            break;
        case "/duty_today":
            await ReplyDutyReportAsync(DutySlashReports.DutyTodayAsync);
            break;
        case "/duty_twodays":
            await ReplyDutyReportAsync(DutySlashReports.DutyTwoDaysAsync);
            break;
        case "/duty_week":
            await ReplyDutyReportAsync(DutySlashReports.DutyWeekAsync);
            break;
        case "/duty_month":
            await ReplyDutyReportAsync(DutySlashReports.DutyMonthAsync);
            break;
        case "/duty_help":
            await botClient.SendMessage(
                message.Chat.Id,
                "📋 Справка по сохранённому графику (данные после подтверждения загрузки Excel):\n"
                + "\n/duty_now — (1) активная смена по времени ПК.\n"
                + "Будний столбец: день 09:00–18:00; ночь 18:00–08:00 утра столбца следующего числа. "
                + "Выходной столбец (красная заливка): день 09:00–21:00; ночь 21:00–09:00 утра следующего числа. "
                + "На буднем с 08:00 до 09:00 — промежуток без дежурства; ночью до утра и после вечернего времени столбец — текущее календарное число.\n"
                + "/duty_today — (2) суточный фрагмент: день и ночной ряд текущего числа из графика.\n"
                + "/duty_twodays — (3) сегодня и завтра.\n"
                + "/duty_week — (4) семь календарных дней начиная с сегодня.\n"
                + "/duty_month — (5) все дни календарного месяца (того же, что сегодня).\n"
                + "\nСначала пришлите Excel и утвердите его кнопкой «Да, взять в работу».",
                replyMarkup: new ReplyKeyboardRemove(),
                cancellationToken: ct);
            break;

        case "/hello":
            await botClient.SendMessage(message.Chat.Id, "HelloWorld", cancellationToken: ct);
            break;
        case "/time_on":
            subscribedChats.TryAdd(message.Chat.Id, 0);
            await botClient.SendMessage(
                message.Chat.Id,
                "Уведомления о времени включены — раз в минуту приходит локальное время ПК.\nОтключить: /time_off",
                cancellationToken: ct);
            break;
        case "/time_off":
            subscribedChats.TryRemove(message.Chat.Id, out _);
            await botClient.SendMessage(message.Chat.Id, "Уведомления о времени выключены.", cancellationToken: ct);
            break;

        case "/duty_notify_on":
            DutyNotifySubscribers.EnsureSubscribed(message.Chat.Id);
            await botClient.SendMessage(
                message.Chat.Id,
                "🔔 На начало каждой дневной и ночной смены будут приходить сообщения "
                + "(локальное время компьютера или сервера, где работает бот).\nВыключить: /duty_notify_off",
                cancellationToken: ct);
            break;

        case "/duty_notify_off":
            DutyNotifySubscribers.Chats.TryRemove(message.Chat.Id, out _);
            await botClient.SendMessage(
                message.Chat.Id,
                "Уведомления о начале смен выключены. Включить: /duty_notify_on",
                cancellationToken: ct);
            break;
    }
}

static Task HandlePollingErrorAsync(ITelegramBotClient _, Exception ex, CancellationToken __)
{
    Console.WriteLine($"Polling error: {ex.Message}");
    return Task.CompletedTask;
}

static async Task RunPeriodicTimeNotifierAsync(
    ITelegramBotClient botClient,
    ConcurrentDictionary<long, byte> chats,
    CancellationToken ct)
{
    try
    {
        using var timer = new PeriodicTimer(TimeSpan.FromMinutes(1));
        while (await timer.WaitForNextTickAsync(ct))
        {
            if (chats.IsEmpty)
                continue;

            var message = $"Текущее время: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";

            foreach (var chatId in chats.Keys.ToArray())
            {
                try
                {
                    await botClient.SendMessage(chatId, message, cancellationToken: ct);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Не удалось отправить время в чат {chatId}: {ex.Message}");
                }
            }
        }
    }
    catch (OperationCanceledException)
    {
        // штатная остановка
    }
}
