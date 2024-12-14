@ECHO off  

CALL :WinGetInstall Microsoft.VCRedist.2005.x86
CALL :WinGetInstall Microsoft.VCRedist.2005.x64
CALL :WinGetInstall Microsoft.VCRedist.2008.x86
CALL :WinGetInstall Microsoft.VCRedist.2008.x64
CALL :WinGetInstall Microsoft.VCRedist.2010.x86
CALL :WinGetInstall Microsoft.VCRedist.2010.x64
CALL :WinGetInstall Microsoft.VCRedist.2012.x86
CALL :WinGetInstall Microsoft.VCRedist.2012.x64
CALL :WinGetInstall Microsoft.VCRedist.2013.x86
CALL :WinGetInstall Microsoft.VCRedist.2013.x64
CALL :WinGetInstall Microsoft.VCRedist.2015+.x86
CALL :WinGetInstall Microsoft.VCRedist.2015+.x64
CALL :WinGetInstall Microsoft.DotNet.DesktopRuntime.6
CALL :WinGetInstall 7zip.7zip
CALL :WinGetInstall Adobe.Acrobat.Reader.64-bit
CALL :WinGetInstall CrystalRich.LockHunter
::CALL :WinGetInstall MathiasSvensson.MultiCommander
CALL :WinGetInstall Mozilla.Firefox
CALL :WinGetInstall Notepad++
CALL :WinGetInstall Oracle.VirtualBox
::CALL :WinGetInstall PDFsam.PDFsam
CALL :WinGetInstall Python.Python.3
::CALL :WinGetInstall QL-Win.QuickLook
CALL :WinGetInstall VideoLAN.VLC
CALL :WinGetInstall WinMerge.WinMerge
CALL :WinGetInstall SourceFoundry.HackFonts
CALL :WinGetInstall WhatsApp.WhatsApp
::CALL :WinGetInstall Telegram.TelegramDesktop

goto :eof

:WinGetInstall
REM Use WinGet to install package
winget install %1 --accept-package-agreements --accept-source-agreements
if %ERRORLEVEL% EQU 0 ECHO %1 installed successfully.  
goto :eof