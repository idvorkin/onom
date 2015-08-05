@echo OFF
if EXIST "%VS140COMNTOOLS%" (
    ECHO FOUND VS14  - CALLING vsvars32.bat
    call "%VS140COMNTOOLS%\vsvars32.bat"
) ELSE (
    ECHO VS14 NOT FOUND - BUILD WILL FAIL.
)

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

