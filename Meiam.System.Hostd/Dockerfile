# 请参阅 https://aka.ms/customizecontainer 以了解如何自定义调试容器，以及 Visual Studio 如何使用此 Dockerfile 生成映像以更快地进行调试。

# 此阶段用于在快速模式(默认为调试配置)下从 VS 运行时
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS base
USER app
WORKDIR /app
EXPOSE 8001
EXPOSE 8002

# 安装字体
USER root
RUN apt-get update && apt-get install -y --no-install-recommends \
    fontconfig \
    fonts-dejavu-core \
    fonts-liberation && \
    fc-cache -f -v && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

# 安装 tzdata（如果基础镜像中没有）
RUN apt-get update && apt-get install -y tzdata
# 设置时区为 Asia/Shanghai
ENV TZ=Asia/Shanghai

# 此阶段用于生成服务项目
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
ARG BUILD_CONFIGURATION=Debug
WORKDIR /src
COPY ["Meiam.System.Hostd/Meiam.System.Hostd.csproj", "Meiam.System.Hostd/"]
COPY ["Meiam.System.Common/Meiam.System.Common.csproj", "Meiam.System.Common/"]
COPY ["Meiam.System.Core/Meiam.System.Core.csproj", "Meiam.System.Core/"]
COPY ["Meiam.System.Extensions/Meiam.System.Extensions.csproj", "Meiam.System.Extensions/"]
COPY ["Meiam.System.Interfaces/Meiam.System.Interfaces.csproj", "Meiam.System.Interfaces/"]
COPY ["Meiam.System.Model/Meiam.System.Model.csproj", "Meiam.System.Model/"]
COPY ["Meiam.System.Tasks/Meiam.System.Tasks.csproj", "Meiam.System.Tasks/"]
RUN dotnet restore "./Meiam.System.Hostd/Meiam.System.Hostd.csproj"
COPY . .
WORKDIR "/src/Meiam.System.Hostd"
RUN dotnet build "./Meiam.System.Hostd.csproj" -c $BUILD_CONFIGURATION -o /app/build

# 此阶段用于发布要复制到最终阶段的服务项目
FROM build AS publish
ARG BUILD_CONFIGURATION=Debug
RUN dotnet publish "./Meiam.System.Hostd.csproj" -c $BUILD_CONFIGURATION -o /app/publish /p:UseAppHost=false

# 此阶段在生产中使用，或在常规模式下从 VS 运行时使用(在不使用调试配置时为默认值)
FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Meiam.System.Hostd.dll"]