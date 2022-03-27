@echo off
rem ソースファイルと同じディレクトリにこのバッチファイルを配置してください。
rem バッチファイルを実行すると同じディレクトリにあるソースファイルがコンパイルされます。
rem バージョン部分は各環境に置き換えて下さい。
cd /d C:\Windows\Microsoft.NET\Framework\v4.0.30319

rem set srcPath=%~dp0*.cs %~dp0*.resx %~dp0Properties\*.cs %~dp0Properties\*.resx
set srcPath=%~dp0CSharp_WinAPI_TextShow\*.cs %~dp0CSharp_WinAPI_TextShow\Properties\*.cs
set exePath=%~dp0Bin\Debug\%~n0.exe

set dllPaths=system.dll,system.drawing.dll,system.windows.forms.dll

csc.exe /t:winexe /optimize+ /out:%exePath% %srcPath% /r:%dllPaths%

echo %ERRORLEVEL%

if %ERRORLEVEL% == 0 (
  goto SUCCESS
)

pause

:SUCCESS