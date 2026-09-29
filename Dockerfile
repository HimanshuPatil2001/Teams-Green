# Base image for production execution (Alpine-based for minimal footprint)
FROM mcr.microsoft.com/dotnet/aspnet:8.0-alpine AS base
WORKDIR /app
EXPOSE 8080
EXPOSE 8081

# Install ICU libraries (globalization support required for MS SQL Client)
RUN apk add --no-cache icu-libs
ENV DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=false

# SDK image for building the application (Alpine-based)
FROM mcr.microsoft.com/dotnet/sdk:8.0-alpine AS build
ARG BUILD_CONFIGURATION=Release
WORKDIR /src
COPY ["TeamsGreen.csproj", "."]
RUN dotnet restore "./TeamsGreen.csproj"
COPY . .
WORKDIR "/src/."
RUN dotnet build "./TeamsGreen.csproj" -c $BUILD_CONFIGURATION -o /app/build

# Publish stage to optimize compilation
FROM build AS publish
ARG BUILD_CONFIGURATION=Release
RUN dotnet publish "./TeamsGreen.csproj" -c $BUILD_CONFIGURATION -o /app/publish /p:UseAppHost=false

# Final runtime image
FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "TeamsGreen.dll"]