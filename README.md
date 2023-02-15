# Script.Interop.Net 
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://github.com/dotnet/winforms/blob/main/LICENSE.TXT) <sub>* This is an Experimental Project, Contributions are welcome.</sub>

**Script.Interop.Net** It's a COM library that adds the ability to use the .NET Framework from older Windows scripting languages; such as **VBScript / JScript** .

# Features

The objectives of the list that have: [❌] - It means I have no idea how to solve it.

- [x] Import of .NET Framework or third-party libraries. made in .NET
- [x] Find Classes and Objects
- [x] Creation and storage of Instances [From script VBS/JS]
- [x] Object Management [From script VBS/JS]
- [ ] Event Handling  ❌ **[This is really necessary, but I can't figure out how to do it.]**
- [x] Static methods supported. **[Fixed]**
- [ ] Invent a way to make Callbacks from .NET to VBS/JS Script. ❌ (or some socket-like system)
- [ ] Create a Type Converter **[Based on the type 'string' , type conversions will be performed between interop.]**
- [ ] Event Calling **[This is simpler, I can get the event pointer and launch ".Invoke" (VBS/JS) or RaiseEvent from .NET]**
- [ ] If you have any other ideas about a feature to implement, please make an Issue ✳️

# How to use ?

1. Register the COM library. **[You can download from Release]** 
     
2. Example Scripts using this library:

- Create a .NET Form (Winform)

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

- Read a plain text file:

```VBScript
'File Path Example
Dim FilePath : FilePath = "File.txt"

' Call Core
Dim InteropDotNet  : Set InteropDotNet = CreateObject("Script.Interop.Net.Linker")
' Create Import/using (equivalent)
Dim AssemblyTarget : Set AssemblyTarget = InteropDotNet.GetAssembly("mscorlib")
' Get Form Class From AssemblyTarget 
Dim ClassType : Set ClassType = AssemblyTarget.GetTypeByAssembly("File")

'Get Methods By Name and Result Type
Dim Methods : Set Methods = ClassType.GetMethods("ReadAllText!String")

'Filter By Parameters and Call Function -> System.IO.File.ReadAllText(String)
Dim Result : Result = Methods.InvokeCall("-String " & FilePath)

' Show Content
Msgbox Result
```
