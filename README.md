# Script.Interop.Net 
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://github.com/dotnet/winforms/blob/main/LICENSE.TXT) <sub>* This is an Experimental Project, Contributions are welcome.</sub>

**Script.Interop.Net** It's a COM library that adds the ability to use the .NET Framework from older Windows scripting languages; such as **VBScript / JScript** .

# Features

- [x] Import of .NET Framework or third-party libraries. made in .NET
- [x] Find Classes and Objects
- [x] Creation and storage of Instances [From VBS]
- [x] Object Management [From VBS]
- [ ] Event Handling  ❌ **[This is really necessary, but I can't figure out how to do it.]**
- [ ] If you have any other ideas about a feature to implement, please make an Issue ✳️

# How to use ?

1. Register the COM library. **[You can download from Release]**
     
```
git status
git add
git commit
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




