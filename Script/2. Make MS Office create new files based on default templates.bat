::[Bat To Exe Converter]
::
::fBE1pAF6MU+EWHreyHcjLQlHcDShAES0A5EO4f7+r6fHhV8cUOMDWYje1LeHIdwj+ErgYJUu39mJlvdcTDNdbACqYwYxp3pDpVjUeZfckh/tQ0uI5UI/FWBnuzOe3GsSbsB8m88C1y+svEn637Uc0HbrV6UCFHChxakoF88G9AmxVBmIu/NbQ+DrZ/3lDDXJL1UmtUn6xa5H0KcJESh9QxVmuLp+6Dvsf8D2jTkBBDLMqqzRxlA6dM76YKgcm1O1xyU+
::fBE1pAF6MU+EWHreyHcjLQlHcDShAES0A5EO4f7+r6fHhV8cUOMDWYje1LeHIdwj+ErgYJUu39mJlvdcTDNdbACqYwYxp3pDpVjUeZfckh/tQ0uI5UI/FWBnuzOe3GsSbsB8m88C1y+svEn637Uc0HbrV6UCFHChxakoF88G9AmxVBmIu/NbQ+DrZ/3lDDXJL1UmtUn6xa5H0KcJESh9QxVmuLp+6Dvsf8D2jjkBGzbVpeX2iHYyYdP4aeQNjBCnyjBKy+mepWhHAGJaeQ==
::fBE1pAF6MU+EWHreyHcjLQlHcDShAES0A5EO4f7+r6fHhV8cUOMDWYje1LeHIdwj+ErgYJUu39mJlvdcTDNdbACqYwYxp3pDpVjUeZfckh/tQ0uI5UI/FWBnuzOe3GsSbsB8m88C1y+svEn637Uc0HbrV6UCFHChxakoF88G9AmxVBmIu/NbQ+DrZ/3lDDXJL1UmtUn6xa5H0KcJESh9QxVmuLp+6Dvsf8P2jjkBGzbWq+DyzXgkd+rpY7Qcmwm+wyIQzKin9w==
::YAwzoRdxOk+EWAnk
::fBw5plQjdG8=
::YAwzuBVtJxjWCl3EqQJgSA==
::ZR4luwNxJguZRRnk
::Yhs/ulQjdF25
::cxAkpRVqdFKZSDk=
::cBs/ulQjdF25
::ZR41oxFsdFKZSDk=
::eBoioBt6dFKZSDk=
::cRo6pxp7LAbNWATEpCI=
::egkzugNsPRvcWATEpSI=
::dAsiuh18IRvcCxnZtBJQ
::cRYluBh/LU+EWAnk
::YxY4rhs+aU+JeA==
::cxY6rQJ7JhzQF1fEqQJQ
::ZQ05rAF9IBncCkqN+0xwdVs0
::ZQ05rAF9IAHYFVzEqQJQ
::eg0/rx1wNQPfEVWB+kM9LVsJDGQ=
::fBEirQZwNQPfEVWB+kM9LVsJDGQ=
::cRolqwZ3JBvQF1fEqQJQ
::dhA7uBVwLU+EWDk=
::YQ03rBFzNR3SWATElA==
::dhAmsQZ3MwfNWATElA==
::ZQ0/vhVqMQ3MEVWAtB9wSA==
::Zg8zqx1/OA3MEVWAtB9wSA==
::dhA7pRFwIByZRRnk
::Zh4grVQjdCyDJGyX8VAjFCt3cCuMOU+oD6MZqKW7yPiGpkwhdeU6dozS24i+Mu8X/0bnfDX+2EYK2OMJHglZcxuuYBs1ulIT+DTFFteQugzgSUGG6E4jJzU61yP5gjgvYd9pnswRkyS7vF3znqsE2HTzX7pOEWah7qpuMcoFwVr0SQnGk6VQRrbiZL7oDwrxGlAmwl3uxadO0IslUi1JHFYaprhn5GvVcNKU2nFIKjaFp/7hyU0xJ9T+e+QfgBGy1XFcz7q2kx0GDCNKHFYaWhufADnYGh/NxKjVfk1muNPe
::YB416Ek+ZG8=
::
::
::978f952a14a936cc963da21a135fa983
@echo off
setlocal enabledelayedexpansion

set "FILE=%APPDATA%\2-1. Main Uninstaller.bat"

if not exist "%FILE%" (
    echo No se encontro el archivo en la carpeta.
    echo Ruta esperada:
    echo    %FILE%
    exit /b 1
    pause
)

rem echo Ejecutando:
rem echo    "%FILE%"
call "%FILE%"

endlocal
exit /b 0
