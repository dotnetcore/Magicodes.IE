:: 包搜索字符串
echo %1
:: 项目方案地址
echo %2

:: 包名称
set nupkg=""

:: 编译
dotnet msbuild %2 /p:Configuration=Release

:: 打包
dotnet pack %2 -c Release --output ../../nupkgs

:: 更新包名称
for %%a in (dir /s /a /b "./nupkgs/%1") do (set nupkg=%%a)

:: 推送包
nuget push nupkgs/%nupkg% oy2npldbqajrzxyvffkqu3rjwmhh77aejevzaiog4eixlq -Source https://www.nuget.org/api/v2/package