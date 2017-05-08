@echo off

REM Step1: remove old version
rmdir /S /Q %APPDATA%\Microsoft\AddIns\TeX4Office_Editor

del %APPDATA%\Microsoft\AddIns\TeX4Office*
del %APPDATA%\Microsoft\Word\STARTUP\TeX4Office*
del %APPDATA%\Microsoft\Excel\XLSTART\TeX4Office*