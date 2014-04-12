@echo off

set CWD=%CD%

If "%TEMP%" == "" set TEMP=C:\TEMP
if "%FXCOPDIR%" == "" set FXCOPDIR=C:\Program Files\Microsoft Fxcop 10.0
set WorkingDir=%TEMP%\CodeAnalysis-%USERNAME%

set TRANSFORM=""
if exist "%CWD%\Config\Top10Report.xsl" set TRANSFORM=/oXSL:"%CWD%\Config\Top10Report.xsl"


rem XSL is located in Config\Top20Report.xsl
if not exist "%WorkingDir%\Output" mkdir "%WorkingDir%\Output"
rem "%FXCOPDIR%\FxCopCmd.exe" /f:"..\Lib\NPOI.dll"  /o:"%WorkingDir%\Output\NPOI.xml" /s /oXSL:"%CWD%\Config\Top20Report.xsl" /r:"%FXCOPDIR%\Rules"
"%FXCOPDIR%\FxCopCmd.exe" /f:"..\Lib\NPOI.dll"  /o:"%WorkingDir%\Output\NPOI.xml" /s %TRANSFORM% /r:"%FXCOPDIR%\Rules"

rem show result in IE
start Explorer.exe /n,/e,"%WorkingDir%"\Output\NPOI.xml
pause
