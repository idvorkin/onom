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

ECHO NUGET Package Installation.
call nuget restore

