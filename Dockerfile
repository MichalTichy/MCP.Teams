# syntax=docker/dockerfile:1.7
FROM mcr.microsoft.com/dotnet/sdk:10.0.301-noble AS build
WORKDIR /source
COPY TeamsMcp.csproj .
RUN dotnet restore TeamsMcp.csproj
COPY . .
RUN dotnet publish TeamsMcp.csproj --configuration Release --output /app/publish --no-restore /p:UseAppHost=false

FROM mcr.microsoft.com/dotnet/aspnet:10.0.9-noble
WORKDIR /app
RUN groupadd --system teamsmcp \
    && useradd --system --gid teamsmcp --home-dir /app --shell /usr/sbin/nologin teamsmcp \
    && mkdir -p /app/data \
    && chown -R teamsmcp:teamsmcp /app
COPY --from=build --chown=teamsmcp:teamsmcp /app/publish .
ENV ASPNETCORE_URLS=http://+:8080 \
    TeamsMcp__DataPath=/app/data
VOLUME ["/app/data"]
EXPOSE 8080
USER teamsmcp
ENTRYPOINT ["dotnet", "TeamsMcp.dll"]
