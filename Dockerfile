# docker build -t el-helper-tg .
# При запуске контейнера обязательно передайте токен Telegram, например:
#   docker run -e BOT_TOKEN="123456:ABC..." el-helper-tg
# Альтернативные имена: TELEGRAM_BOT_TOKEN, TG_BOT_TOKEN, BotConfiguration__BotToken
#
# Railway (https://railway.com): в сервисе откройте Variables → добавьте BOT_TOKEN
# со значением токена от @BotFather. Пересоберите/задеплойте после сохранения.
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

COPY ElHelper.csproj .
RUN dotnet restore ElHelper.csproj

COPY . .
RUN dotnet publish ElHelper.csproj -c Release -o /app/publish --no-restore

FROM mcr.microsoft.com/dotnet/runtime:8.0 AS final
WORKDIR /app
COPY --from=build /app/publish .

ENTRYPOINT ["dotnet", "ElHelper.dll"]
