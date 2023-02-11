# Script.Interop.Net 
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://github.com/dotnet/winforms/blob/main/LICENSE.TXT) <sub>* This is an Experimental Project, Contributions are welcome.</sub>

**Script.Interop.Net** It's a COM library that adds the ability to use the .NET Framework from older Windows scripting languages; such as **VBScript / JScript** .

# Features

The objectives of the list that have: [❌] - It means I have no idea how to solve it.

- [x] Import of .NET Framework or third-party libraries. made in .NET
- [x] Find Classes and Objects
- [x] Creation and storage of Instances [From VBS]
- [x] Object Management [From VBS]
- [ ] Event Handling  ❌ **[This is really necessary, but I can't figure out how to do it.]**
- [ ] Shared/Static Methods are not supported ❌ **[I have to figure out a way to call it from .Net and parse the parameters supplied from the script.]**
- [ ] Create a Type Converter **[Based on the type 'string' , type conversions will be performed between interop.]**
- [ ] Event Calling **[This is simpler, I can get the event pointer and launch ".Invoke" (vbs) or RaiseEvent from .NET]**
- [ ] Invent a way to make Callbacks from .NET to VBS Script. (or some socket-like system)
- [ ] If you have any other ideas about a feature to implement, please make an Issue ✳️

# How to use ?

1. Register the COM library. **[You can download from Release]**
     
```Batch
@echo off
Set regasm_x86="C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe"
Set regasm_x64="C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"

Set LibPath="Script.Interop.Net.dll"

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

```
     
2. Create a small vbs script:

```VBScript
' Call Core
Dim InteropDotNet  : Set InteropDotNet = CreateObject("Script.Interop.Net.Linker")
' Create Import/using (equivalent)
Dim AssemblyTarget : Set AssemblyTarget = InteropDotNet.GetAssembly("System.Windows.Forms")
' Get Form Class From AssemblyTarget 
Dim ClassType : Set ClassType = AssemblyTarget.GetTypeByAssembly("Form")
' Create New Form Instance.
Dim FormNewInstance : Set FormNewInstance = ClassType.CreateInstance()
' Set Form Title
FormNewInstance.Text = "New Form"
' Show Form
FormNewInstance.ShowDialog()
```




