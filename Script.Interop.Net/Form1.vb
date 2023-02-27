Imports System.Reflection

Public Class Form1
    '  Dim EventInfoEx As MethodInfo
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim test As EventTracer = New EventTracer
        Dim method As MethodInfo = GetType(EventTracer).GetMethod("WriteTrace")
        Dim eventInfo As EventInfo = Button1.[GetType]().GetEvent("Click")
        Dim handler As [Delegate] = [Delegate].CreateDelegate(eventInfo.EventHandlerType, test, method)
        eventInfo.AddEventHandler(Button1, handler)
        ' EventInfoEx = method

    End Sub
    '

End Class

