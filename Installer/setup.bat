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
echo [����] �аѾ\�w�ˤ�U�A�b PowerPoint �����J TeX4Office �W�q���C
echo ================================================================

pause

echo ================================================================
echo [����] �w�˧����C���� PowerPoint ������w�˵{���|�۰ʵ����C
echo ================================================================

IF EXIST "C:\Program Files\Microsoft Office" (
	C:
	cd "C:\Program Files\Microsoft Office\Office*"
	powerpnt.exe
) ELSE (

	REM IF EXIST "C:\Program Files (x86)\Microsoft Office" (
		C:
		cd "C:\Program Files (x86)\Microsoft Office\Office*"
		powerpnt.exe

	REM ) ELSE (
		REM echo " "
		REM echo [���~] �b C:\Program Files\ �� C:\Program Files(x86)\ �䤣��w�w�˪� Microsoft Office�I
		REM echo " "
	REM )
	
)