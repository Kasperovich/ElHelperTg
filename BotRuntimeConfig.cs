using System.Collections;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;

/// <summary>Токен бота из конфигурации и окружения (регистр имени переменной не важен).</summary>
internal static class BotRuntimeConfig
{
    internal static readonly string[] PreferredEnvVarNames =
    [
        "BOT_TOKEN",
        "TELEGRAM_BOT_TOKEN",
        "TG_BOT_TOKEN",
        "BotConfiguration__BotToken",
    ];

    internal static readonly string TokenFileEnv = "BOT_TOKEN_FILE";

    internal static string? TryGetTelegramBotToken(HostApplicationBuilder builder)
    {
        static string? Nz(string? s) =>
            string.IsNullOrWhiteSpace(s) ? null : s.Trim();

        if (Nz(builder.Configuration["BotConfiguration:BotToken"]) is { } fromCfg)
            return fromCfg;

        foreach (DictionaryEntry kv in Environment.GetEnvironmentVariables())
        {
            var key = kv.Key?.ToString();
            if (string.IsNullOrWhiteSpace(key))
                continue;

            var match = PreferredEnvVarNames.Any(w =>
                string.Equals(key, w, StringComparison.OrdinalIgnoreCase));
            if (!match)
                continue;

            if (Nz(kv.Value?.ToString()) is { } v)
                return v;
        }

        var tokenPath = Environment.GetEnvironmentVariable(TokenFileEnv);
        if (!string.IsNullOrWhiteSpace(tokenPath))
        {
            try
            {
                if (File.Exists(tokenPath) &&
                    Nz(File.ReadAllText(tokenPath).ReplaceLineEndings().Trim()) is { } fromFile)
                    return fromFile;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[ElHelper] BOT_TOKEN_FILE={tokenPath}: {ex.Message}");
            }
        }

        return null;
    }

    internal static IReadOnlyList<string> SensitiveEnvKeysForDiagnostics()
    {
        var hints = new List<string>();
        foreach (DictionaryEntry entry in Environment.GetEnvironmentVariables())
        {
            var key = entry.Key?.ToString();
            if (string.IsNullOrEmpty(key))
                continue;

            if (key.Contains("TOKEN", StringComparison.OrdinalIgnoreCase) ||
                key.Contains("TELEGRAM", StringComparison.OrdinalIgnoreCase))
                hints.Add(key);
        }

        hints.Sort(StringComparer.OrdinalIgnoreCase);
        return hints;
    }
}
