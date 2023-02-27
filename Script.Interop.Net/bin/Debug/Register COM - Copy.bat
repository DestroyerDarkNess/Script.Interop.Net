@echo off
Set regasm_x86="C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe"
Set regasm_x64="C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"

Set LibPath=Script.Interop.Net.dll

ECHO Registering COM "%LibPath%"

FOR /f "tokens=2 delims==" %%f IN ('wmic os get osarchitecture /value ^| find "="') DO SET "OS_ARCH=%%f"
IF "%OS_ARCH%"=="32-bit" GOTO :32bit
IF "%OS_ARCH%"=="64-bit" GOTO :64bit

ECHO OS Architecture %OS_ARCH% is not supported!
PAUSE
EXIT 1

:32bit
ECHO "32 bit Operating System"
ECHO "Registering Library"
%regasm_x86% /codebase %LibPath%
GOTO :SUCCESS

:64bit
ECHO "64 bit Operating System"
ECHO "Registering Library"
%regasm_x64% /codebase %LibPath%
GOTO :SUCCESS

:SUCCESS
PAUSE
EXIT 0


