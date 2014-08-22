@echo off

SET SOLUTION=
set TARGET=
set CONFIG=
set PLATFORM=
set CMDLINE=

:: set FRAMEWORK=v2.0.50727
:: set FRAMEWORK=v3.5
set FRAMEWORK=v4.0.30319
set MSBUILDARGS=/v:M /fl /flp:LogFile=msbuild.log;Verbosity=Normal
set MSBUILD="%WINDIR%\Microsoft.NET\Framework\%FRAMEWORK%\MsBuild.exe"
if not exist %MSBUILD% echo Unable to locate %MSBUILD%&goto end

IF /I "%1" EQU "?"     goto Usage
IF /I "%1" EQU "/?"    goto Usage
IF /I "%1" EQU "/h"    goto Usage
IF /I "%1" EQU "/help" goto Usage
IF "%1" EQU ""         goto NoSolutionFile
if not exist %1        goto FileNotFound
SET SOLUTION=%1

:target
if /I "%2" EQU "clean"   set TARGET=Clean&set CONFIG=Debug&goto Platform
if /I "%2" EQU "debug"   set TARGET=Rebuild&set CONFIG=Debug&goto Platform
if /I "%2" EQU "release" set TARGET=Rebuild&set CONFIG=Release&goto Platform
goto InvalidTarget

:Platform
IF "%3" EQU ""           set PLATFORM="Any CPU"&goto Execute
if /I "%3" EQU "x86"     set PLATFORM=%3&goto Execute
if /I "%3" EQU "x64"     set PLATFORM=%3&goto Execute
if /I "%3" EQU "Any CPU" set PLATFORM=%3&goto Execute
goto InvalidPlatform

:Execute
echo.
echo Solution:      %SOLUTION%
echo Target:        %TARGET%
echo Configuration: %CONFIG%
echo Platform:      %PLATFORM%
echo.

set CMDLINE=%MSBUILD% %SOLUTION% /t:%TARGET% /p:Configuration=%CONFIG% /p:Platform=%PLATFORM% %MSBUILDARGS%
%CMDLINE%
if errorlevel 1 goto BuildError

if "%TARGET%" NEQ "Clean" goto Executed

rem For Target=Clean, run it also with Config=Release
echo Target=%TARGET%. Ran Config=%CONFIG%, Now running Config=Release
echo.
set CMDLINE=%MSBUILD% %SOLUTION% /t:%TARGET% /p:Configuration=Release /p:Platform=%PLATFORM% %MSBUILDARGS%
%CMDLINE%
if errorlevel 1 goto BuildError

:Executed
echo.
echo Executed %CMDLINE%
echo.

goto end

:NoSolutionFile
echo.
echo Invalid solution filename.
echo.
goto Usage

:FileNotFound
echo.
echo File not found, %1
echo.
goto Usage

:InvalidPlatform
echo.
echo Invalid platform in command line arguments
echo.
goto Usage

:InvalidTarget
echo.
echo Invalid target in command line arguments
echo.
goto Usage

:Usage
echo.
echo.******************
echo    Sample usage
echo.******************
echo.
echo build.bat [SolutionFile] [target] [platform]
echo.
echo SolutionFile
echo.
echo    solution filename. eg: <path>\app.sln
echo.
echo Targets
echo.
echo    clean   - 1st: target=clean,   config=debug
echo              2nd: target=clean,   config=release
echo.
echo    debug   - 1st: target=clean,   config=debug
echo              2nd: target=rebuild, config=debug
echo.
echo    release - 1st: target=clean,   config=release
echo              2nd: target=rebuild, config=release
echo.
pause
echo Platform (not required, default: "Any CPU")
echo.
echo.   "Any CPU" - (default) compiles your assembly to run on any platform.
echo.
echo.   x86     - compiles your assembly to be run by the 32-bit, x86-compatible 
echo.             common language runtime.
echo.
echo.   x64     - compiles your assembly to be run by the 64-bit common language 
echo.             runtime on a computer that supports the AMD64 or EM64T 
echo.             instruction set.
echo. 
echo.   On a 64-bit Windows operating system:
echo. 
echo.     Assemblies compiled with /platform:x86 will Execute on the 32 bit CLR 
echo      running under WOW64.
echo. 
echo.     Executables compiled with the /platform:"Any CPU" will Execute on the 
echo.     64 bit CLR.
echo. 
echo.     DLLs compiled with the /platform:"Any CPU" will Execute on the same CLR
echo.     as the process into which it is being loaded.
echo.
goto end

:BuildError
echo.
echo.
echo *********************************************
echo    Error occurred during build, see above.
echo *********************************************
echo.
echo.

exit /b 1

goto end

:end