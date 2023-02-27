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

