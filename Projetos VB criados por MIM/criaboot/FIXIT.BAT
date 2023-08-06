IF "%CDPROB%"=="1" GOTO CDGONE
if "%config%"=="HELP" GOTO HELP
if "%config%"=="NOCD" GOTO NOCDHLP
echo.
echo.
echo Your CD-ROM is drive %CDROM%
echo.
echo To view Help, type HELP and then press ENTER.
GOTO END

:NOCDHLP
echo.
echo.
echo To view Help, type HELP and then press ENTER.
GOTO END

:HELP
echo. 
echo. 
echo Press CTRL+ALT+DELETE to restart your computer.
echo. 
echo. 
echo Your CD-ROM is drive %CDROM%
echo. 
echo If you are unable to start Windows and have tried the
echo steps outlined in the Help file, please run ScanDisk and
echo then try to run Windows Setup again.  To run ScanDisk, type:
echo.      
echo     scandisk /all
echo. 
echo and then press ENTER.
echo. 
echo After ScanDisk completes, type:
echo. 
echo     %CDROM%:\setup.exe
echo. 
echo and then press ENTER. Follow the directions on your screen.
echo.
echo To view Help again, type HELP and then press ENTER.
GOTO END

:CDGONE
echo.
echo.
echo Press CTRL+ALT+DELETE to restart your computer.
echo.
echo There may be a problem with your CD-ROM. Check to make
echo sure it is properly connected and restart your computer.
echo. 
echo To view Help, or to view Help again, type HELP
echo and then press ENTER.

:END
