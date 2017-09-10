echo ON
REM For some unknown reason can't use NOT in batch files with paths with spaces :(
if EXIST "%VS150COMNTOOLS%" (
        goto :pathfound
)
set VS150COMNTOOLS=C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\Common7\Tools\
if EXIST "%VS150COMNTOOLS%" (
        goto :pathfound
)
echo VS 2017 Community NOT FOUND
goto :eof
:pathfound
call "%VS150COMNTOOLS%\VsMSBuildCmd.bat"

if EXIST "c:\Program Files (x86)\NUnit 2.6.4\bin\nunit-console.exe" (
    ECHO FOUND NUnit SKIPPING Installation.
) ELSE (
    ECHO Installing NUnit via chocolatey
    choco install NUnit
)

if EXIST "c:\Program Files (x86)\NUGet" (
    ECHO FOUND NUGet SKIPPING Installation.
) ELSE (
    ECHO Installing NUGet via chocolatey
    choco install Nuget.CommandLine
)

ECHO NUGET Package Installation.
call nuget restore

