dotnet publish -r win-x64 --self-contained=true /p:PublishSingleFile=true
dotnet publish -r linux-x64 --self-contained=false /p:PublishSingleFile=true