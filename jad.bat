@echo off

set TMP_DRIVE=J

:loopStart

if "%~1" == "" goto :loopEnd
echo ＿
echo ■%1

set TMP_PATH=%~dp1
subst %TMP_DRIVE%: %TMP_PATH:~0,-1%
if NOT "%ERRORLEVEL%"=="0" (
echo JAD用の一時ドライブ「%TMP_DRIVE%:」が既に使用されています。
pause
goto :end
)

rem C:\Progra~1\Java\jad.exe -ff -i -lnc -o -space -t -8 -s .java -nonlb %*
rem echo J:\%~nx1
C:\Progra~1\Java\jad.exe -ff -i -lnc -o -space -t -8 -s .java -nonlb J:\%~nx1
shift

subst /d %TMP_DRIVE%:

goto :loopStart

:loopEnd
pause

:end
@echo on
