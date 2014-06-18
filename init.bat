@echo OFF
if EXIST "%VS120COMNTOOLS%" (
    ECHO FOUND VS12  - CALLING vcvars32.bat
    call "%VS120COMNTOOLS%\..\..\VC\bin\vcvars32.bat"
) ELSE (
    ECHO VS NOT FOUND - BUILD WILL FAIL.
)

if EXIST "c:\Program Files (x86)\NUnit 2.6.3\bin\nunit-console.exe" (
    ECHO FOUND NUnit SKIPPING Installation.
) ELSE (
    ECHO Installing NUnit via chocolatey
    choco install NUnit
)

if EXIST "c:\Chocolatey\bin\nuget.bat" (
    ECHO FOUND NUGet SKIPPING Installation.
) ELSE (
    ECHO Installing NUGet via chocolatey
    choco install Nuget.CommandLine
)

ECHO NUGET Package Installation.
call nuget restore

