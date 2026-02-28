@echo off

set fdir=%WINDIR%\Microsoft.NET\Framework64

if not exist %fdir% (
	set fdir=%WINDIR%\Microsoft.NET\Framework
)

set msbuild4=%fdir%\v4.0.30319\msbuild.exe

%msbuild4% ..\scratchpad\HWPF\HWPF.csproj /p:Configuration=Release /t:Rebuild /p:OutputPath=..\..\Build\ScratchPad\Release\Net40

FOR /F "tokens=*" %%G IN ('DIR /B /AD /S obj') DO RMDIR /S /Q "%%G"

set msbuild2=%fdir%\v3.5\msbuild.exe

%msbuild2% ..\scratchpad\HWPF\HWPF.net2.csproj /p:Configuration=Release /t:Rebuild /p:OutputPath=..\..\Build\ScratchPad\Release\Net20

FOR /F "tokens=*" %%G IN ('DIR /B /AD /S obj') DO RMDIR /S /Q "%%G"
pause