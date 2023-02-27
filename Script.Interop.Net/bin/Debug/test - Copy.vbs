Function GetRefTest()
   Dim Splash
   Splash = "GetRefTest Version 1.0"   & vbCrLf
   Splash = Splash & Chr(169) & " YourCompany 1999 "
   MsgBox Splash
End Function

' Call Core
Dim InteropDotNet  : Set InteropDotNet = CreateObject("Script.Interop.Net.Linker")

'------------------------------------------------------------------------------------------------------
'-    Main UI
'------------------------------------------------------------------------------------------------------

Dim ClassType : Set ClassType = InteropDotNet.GetByRoot("System.Windows.Forms!Form")

Dim FormNewInstance : Set FormNewInstance = ClassType.CreateInstance()

' Set Form Title
FormNewInstance.Text = "New Form"


'------------------------------------------------------------------------------------------------------
'-    Controls
'------------------------------------------------------------------------------------------------------

' Get Form Class From AssemblyTarget 
Dim ButtonType : Set ButtonType =  InteropDotNet.GetByRoot("System.Windows.Forms!Button")
Dim ButtonInstance : Set ButtonInstance = ButtonType.CreateInstance()
ButtonInstance.Text = "Clickme"
ButtonInstance.Height = 80


Dim LocationProperty : Set LocationProperty = ButtonInstance.GetPropertyValue("Location")
ButtonInstance.Location = PointClass

'------------------------------------------------------------------------------------------------------
'-    Make UI
'------------------------------------------------------------------------------------------------------

Dim PropertyEx : Set PropertyEx = ClassType.GetPropertyValue("Controls")
Dim MethodToCall : Set MethodToCall = PropertyEx.GetMethod("Add")
MethodToCall.InvokeCall(ButtonInstance)

'------------------------------------------------------------------------------------------------------
'-    Show GUI
'------------------------------------------------------------------------------------------------------

' Show Form
FormNewInstance.ShowDialog()


Dim bCloseClick : bCloseClick = True

 Do
     WScript.Sleep 200
 Loop Until bCloseClick

