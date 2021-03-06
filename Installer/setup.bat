@echo off

REM Step1: remove old version
rmdir /S /Q %APPDATA%\Microsoft\AddIns\TeX4Office_Editor

del %APPDATA%\Microsoft\AddIns\TeX4Office*
del %APPDATA%\Microsoft\Word\STARTUP\TeX4Office*
del %APPDATA%\Microsoft\Excel\XLSTART\TeX4Office*

REM Step2: copy files
xcopy /E /Y /C /R /Q AddIns\*.* %APPDATA%\Microsoft\AddIns\
xcopy /E /Y /C /R /Q Word\STARTUP\*.* %APPDATA%\Microsoft\Word\STARTUP\
xcopy /E /Y /C /R /Q Excel\XLSTART\*.* %APPDATA%\Microsoft\Excel\XLSTART\

REM Step3: open the setup manual, and remind the users to load the add-ins (PowerPoint)

echo ================================================================
echo [提示] 請參閱安裝手冊，在 PowerPoint 中載入 TeX4Office 增益集。
echo ================================================================

pause

echo ================================================================
echo [提示] 安裝完成。等待 PowerPoint 關閉後安裝程式會自動結束。
echo ================================================================

C:
cd "C:\Program Files\Microsoft Office\Office*"
IF EXIST "powerpnt.exe"  powerpnt.exe
cd "C:\Program Files (x86)\Microsoft Office\Office*"
IF EXIST "powerpnt.exe"  powerpnt.exe
	