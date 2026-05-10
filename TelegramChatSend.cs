using Telegram.Bot;
using Telegram.Bot.Types;
using Telegram.Bot.Types.ReplyMarkups;

/// <summary>Отправка длинных ответов в Telegram (лимит ~4096 символов на сообщение).</summary>
internal static class TelegramChatSend
{
    public const int TelegramMaxMessageLength = 4096;

    /// <summary>Запас под префикс «(часть n/m)».</summary>
    private const int ChunkBudget = 3800;

    public static async Task SendLongTextAsync(
        ITelegramBotClient bot,
        ChatId chatId,
        string text,
        CancellationToken cancellationToken = default,
        ReplyKeyboardRemove? removeReplyKeyboardWithLastChunk = null)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            await bot.SendMessage(chatId, "—", replyMarkup: removeReplyKeyboardWithLastChunk, cancellationToken: cancellationToken);
            return;
        }

        var chunks = ChunkText(text.TrimEnd(), ChunkBudget).ToList();
        for (var i = 0; i < chunks.Count; i++)
        {
            var body = chunks.Count > 1
                ? $"({i + 1}/{chunks.Count})\n{chunks[i]}"
                : chunks[i];

            if (body.Length > TelegramMaxMessageLength)
                body = body[..TelegramMaxMessageLength];

            var last = i == chunks.Count - 1;
            await bot.SendMessage(
                chatId,
                body,
                replyMarkup: last ? removeReplyKeyboardWithLastChunk : null,
                cancellationToken: cancellationToken);
        }
    }

    internal static IEnumerable<string> ChunkText(string text, int maxChunk)
    {
        if (text.Length <= maxChunk)
        {
            yield return text;
            yield break;
        }

        var start = 0;
        while (start < text.Length)
        {
            var chunkEnd = Math.Min(start + maxChunk, text.Length);
            if (chunkEnd < text.Length)
            {
                var sliceLen = chunkEnd - start;
                var li = text.AsSpan(start, sliceLen).LastIndexOf('\n');
                if (li >= sliceLen / 4)
                    chunkEnd = start + li + 1;
            }

            var piece = text[start..chunkEnd].TrimEnd();
            if (piece.Length > 0)
                yield return piece;

            start = chunkEnd;
            while (start < text.Length && (text[start] == '\n' || text[start] == '\r'))
                start++;
        }
    }
}
