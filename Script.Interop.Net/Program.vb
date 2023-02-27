Imports System.Reflection

Public Class Program

    Public Shared Sub Main()
        Dim ButonNew As New Windows.Forms.Button With {.Name = "Button1"}
        Dim LabelNew As New Windows.Forms.Label With {.Name = "Label1"}

        Dim InteropDotNet As New Linker
        Dim ClassType = InteropDotNet.GetByRoot("mscorlib!Console")
        Dim Methods = ClassType.GetMethods("WriteLine")
        Dim Result = Methods.InvokeCall("-String " & "Finalization")

        Console.ReadLine()
        Dim ConsolePost As Integer = Console.CursorTop
        For i As Integer = 0 To 100
            ' 
            Dim Data As String = $"Progress:{i}%"
            ' Console.Write(Data)

            Dim MethodsWrite = ClassType.GetMethods("Write")
            Dim ResultWrite = MethodsWrite.InvokeCall("-String " & Data)
            System.Threading.Thread.Sleep(25)

            Dim MethodsSetCursorPosition = ClassType.GetMethods("SetCursorPosition")
            Dim ResultSetCursorPosition = MethodsSetCursorPosition.InvokeCall("-Int32 0 -Int32 " & ConsolePost)


            ' Console.SetCursorPosition(0, ConsolePost)
        Next

        Console.Write(vbCr & "Done!          ")



        '   Dim LabelEx As Windows.Forms.Button = New Windows.Forms.Button
        '    LabelEx.Location.X
        ' Dim asdasd = New Form1
        '  asdasd.Height = 80
        'Dim titles As Object = asdasd.[GetType]().GetProperty("Controls").GetValue(asdasd)

        ''For Each exad In titles.GetMethods
        ''    Console.WriteLine(exad.Name)
        ''Next
        ''Console.ReadLine()

        'Dim addMethod As MethodInfo = titles.[GetType]().GetMethods().Where(Function(m) m.Name = "Add" AndAlso m.GetParameters().Count() = 1).FirstOrDefault()
        'addMethod.Invoke(titles, New Object() {LabelEx})

        'asdasd.ShowDialog()

        'Dim NewClassEx As New EventExo
        'NewClassEx.MakeDialog()




    End Sub


End Class


Public Class EventExo

    Public Sub MakeDialog()

        Dim ASD As New Windows.Forms.Form
        ASD.Show()
    End Sub

    Public Sub RegisterEvents(ByVal FormTarget As Object)
        Dim tExForm As Type = FormTarget.GetType

        Dim evClick As EventInfo = FormTarget.GetType.GetEvent("Click")

        Dim tDelegate As Type = evClick.EventHandlerType

        Dim miHandler As MethodInfo = tExForm.GetMethod("evnt", BindingFlags.NonPublic Or BindingFlags.Instance)
        Dim d As [Delegate] = [Delegate].CreateDelegate(tDelegate, Me, miHandler)

        Dim miAddHandler As MethodInfo = evClick.GetAddMethod()
        Dim addHandlerArgs() As Object = {d}

        miAddHandler.Invoke(FormTarget, addHandlerArgs)
    End Sub



End Class