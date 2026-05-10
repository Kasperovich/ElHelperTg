using System.Collections.Concurrent;

/// <summary>Чаты, куда слать авто-уведомления о начале дежурной смены (по локальному времени процесса).</summary>
internal static class DutyNotifySubscribers
{
    internal static readonly ConcurrentDictionary<long, byte> Chats = new();

    internal static void EnsureSubscribed(long chatId) => Chats.TryAdd(chatId, 0);
}
