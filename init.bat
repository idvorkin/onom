@echo OFF
if EXIST "%VS120COMNTOOLS%" (
    ECHO FOUND VS12 CALLING
    call "%VS120COMNTOOLS%\..\..\VC\bin\vcvars32.bat"
) ELSE (
    ECHO NOT FOUND
)

