@Echo Off

echo =========================

echo Now clearing all users Java cache folder

set "docandset=%homedrive%\users"

     >> c:\Delete.log echo/ "Java\Deployment\cache\6.0\"

for /f "delims=" %%a in ('dir "%docandset%" /ad /b') do (

for %%b in (

"%docandset%\%%a\AppData\LocalLow\Sun\Java\Deployment\cache\6.0"

) do (

echo %%b >> c:\Delete.log

cd /d %%b >> c:\Delete.log  2>&1

rd /s /q %%b >> c:\Delete.log  2>&1

del /f /s /q %%b >> c:\Delete.log  2>&1

)

)

echo =========================

echo COMPLETE!!!! 

echo =========================

echo All users Java cache folder has been cleared.

PAUSE
