@echo off
rem ===================================================================
rem  Student labor-fee helper - Windows launcher
rem  Finds a usable Python 3 and hands over to run.py
rem  (macOS / Linux: use ./run.sh or python3 run.py)
rem
rem  NOTE: keep this file pure ASCII. cmd.exe reads .cmd files using the
rem  OEM codepage, so UTF-8 text here would be garbled into bogus commands.
rem
rem  The interpreter is kept in TWO variables: PYEXE (already quoted) and
rem  PYARGS (e.g. -3). Joining them into one quoted string would leave a
rem  trailing space and cmd would fail with "is not recognized".
rem ===================================================================
setlocal EnableExtensions
cd /d "%~dp0"
title Student Fee Helper

set "PYEXE="
set "PYARGS="

rem Enable Python UTF-8 mode (modern default, avoids mojibake on Windows).
rem Do NOT set PYTHONIOENCODING here: on a real console Python uses the
rem console API (so Chinese shows correctly), and when output is piped to a
rem file it falls back to the system encoding (so the log stays readable).
set "PYTHONUTF8=1"

rem --- Pass 1: prefer an interpreter that already has the dependencies ---
rem "py -0p" lists every registered Python; the leftmost is the default.
for /f "tokens=1,*" %%A in ('py -0p 2^>nul') do call :try_registered "%%B"
call :try_deps python
call :try_deps python3
call :try_deps py -3

rem --- Pass 2: any working Python 3.8+; run.py installs what is missing ---
call :try_any python
call :try_any python3
call :try_any py -3
call :try_any "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
call :try_any "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
call :try_any "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
call :try_any "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
call :try_any "%LOCALAPPDATA%\Programs\Python\Python39\python.exe"
call :try_any "C:\Python313\python.exe"
call :try_any "C:\Python312\python.exe"
call :try_any "C:\Python311\python.exe"
call :try_any "D:\python\python.exe"

if not defined PYEXE goto :no_python

%PYEXE% %PYARGS% "%~dp0run.py" %*
set "CODE=%ERRORLEVEL%"

rem -1 means the console window was closed (or Ctrl+C) - that is a normal exit.
if "%CODE%"=="0" goto :done
if "%CODE%"=="-1" goto :done
if "%CODE%"=="-1073741510" goto :done
echo.
echo   The program exited with code %CODE%.
echo   Please send me a screenshot of the messages above.
echo.
pause

:done
endlocal
exit /b 0


:try_registered
rem %1 = path reported by "py -0p"; skip if we already have an interpreter.
if defined PYEXE exit /b 0
if "%~1"=="" exit /b 0
"%~1" -c "import docx, xlrd" >nul 2>nul
if errorlevel 1 exit /b 0
set "PYEXE="%~1""
exit /b 0


:try_deps
rem %1 = executable (quoted if it is a path), %2 = extra argument.
if defined PYEXE exit /b 0
set "CAND=%~1"
if not "%~2"=="" set "CAND=%~1 %~2"
%1 %2 -c "import docx, xlrd" >nul 2>nul
if errorlevel 1 exit /b 0
set "PYEXE="%~1""
set "PYARGS=%~2"
exit /b 0


:try_any
if defined PYEXE exit /b 0
%1 %2 -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 8) else 1)" >nul 2>nul
if errorlevel 1 exit /b 0
set "PYEXE="%~1""
set "PYARGS=%~2"
exit /b 0


:no_python
echo.
echo   [ERROR] No usable Python 3.8+ was found.
echo.
echo   Please install Python 3.8 or newer:
echo       https://www.python.org/downloads/
echo.
echo   During setup, tick "Add python.exe to PATH".
echo.
echo   If it is already installed, try running this and send me the output:
echo       python run.py
echo.
pause
exit /b 1
