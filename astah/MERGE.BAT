@ECHO OFF
REM 結合バッチファイル：div2.exe により生成

ECHO 分割されたファイルを復元します

IF NOT EXIST "astah.001"  GOTO ERR1
IF NOT EXIST "astah.002"  GOTO ERR1
IF NOT EXIST "astah.003"  GOTO ERR1
IF     EXIST "astah.exe"  GOTO ERR2

COPY "astah.001" "astah.exe"
IF ERRORLEVEL 1 GOTO ERR3
COPY /B "astah.exe"+"astah.002"+"astah.003"
IF ERRORLEVEL 1 GOTO ERR3

DEL "astah.001"
DEL "astah.002"
DEL "astah.003"


REN "astah.exe" "astah.exe"

ECHO このバッチファイルも消してください
GOTO END

:ERR1
ECHO ファイルが一部存在しません
ECHO 処理を中止します
GOTO END

:ERR2
ECHO 復元先のファイルが既に存在しています
ECHO 処理を中止します
GOTO END

:ERR3
ECHO 復元に失敗しました
ECHO 処理を中止します
GOTO END

:END
ECHO 終了するには、何かキーを押してください
PAUSE >NUL

