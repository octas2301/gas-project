@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo === 広告運用GAS フォルダ内のファイル（隠し含む）===
dir /a
echo.
echo === 既存のプロジェクト設定を削除します ===
if exist .clasp.json ( del .clasp.json & echo .clasp.json を削除しました。 ) else ( echo .clasp.json はありません。 )
if exist appsscript.json ( del appsscript.json & echo appsscript.json を削除しました。 )
echo.
echo === clasp clone を実行します ===
clasp clone "1VU31ni9FMuaqOV1hxCdcRd42eJMQ0ATLIXmW1pViizxz_OPC9Gj8Uf9R"
echo.
pause
