
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



