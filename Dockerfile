# docker build -t el-helper-tg .
# Токен: BOT_TOKEN или BotConfiguration__BotToken (переменные окружения в рантайме).
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
