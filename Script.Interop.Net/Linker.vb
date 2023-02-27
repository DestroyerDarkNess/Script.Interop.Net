
<ComClass(Linker.ClassId, Linker.InterfaceId, Linker.EventsId)>
Public Class Linker

#Region " COM Defines "

    Public Const ClassId As String = "69E7D203-B755-41F3-90A4-11B336CCF06E"
    Public Const InterfaceId As String = "9665100C-5717-463C-B7B5-76779ADF36D3"
    Public Const EventsId As String = "A937B081-25F0-41B1-9AC1-3B5CC7E0642B"

#End Region

#Region " Constructor "

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region " Declare "

    Public Property FrameworkPath As String = DotNetFrameworkLocator.GetInstallationLocation
    Private Property InstancesLoaded As New Dictionary(Of String, Library)

#End Region

#Region " Methods "

    Public Function GetlibsLoaded() As String
        Dim KeysLibs As List(Of String) = InstancesLoaded.Keys.ToList
        Return String.Join(",", KeysLibs.ToArray)
    End Function

    Public Function GetlibsLoadedByName(ByVal NameEx As String) As String
        Dim KeysLibs As List(Of String) = InstancesLoaded.Keys.ToList
        Return KeysLibs.Find(Function(p) p = NameEx).ToString
    End Function

    Public Function GetAssembly(ByVal Name As String) As Library
        Try
            For Each InsOld As KeyValuePair(Of String, Library) In InstancesLoaded
                If IO.Path.GetFileNameWithoutExtension(InsOld.Key) = IO.Path.GetFileNameWithoutExtension(Name) Then
                    Return InsOld.Value
                End If
            Next

            If Name.EndsWith(".dll") = False Then
                Name += ".dll"
            End If

            Dim TargetPath As String = String.Empty

            If IO.File.Exists(Name) Then
                TargetPath = Name
            ElseIf IO.File.Exists(IO.Path.Combine(FrameworkPath, Name)) Then
                TargetPath = IO.Path.Combine(FrameworkPath, Name)
            End If

            If TargetPath = String.Empty Then
                Return Nothing
            Else
                Dim LibraryPath As Byte() = IO.File.ReadAllBytes(TargetPath)
                Dim assembly As Reflection.Assembly = System.Reflection.Assembly.Load(LibraryPath)
                Dim NewLib As New Library With {.Asm = assembly}
                InstancesLoaded.Add(IO.Path.GetFileNameWithoutExtension(Name), NewLib)

                Return DirectCast(NewLib, Object)
            End If

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function GetByRoot(ByVal Name As String) As TypeWrapper
        Dim RootNamespace As String = Name
        Dim TypeName As String = String.Empty

        If Name.Contains("!") = True Then
            Dim NamType() As String = Name.Split("!")
            RootNamespace = NamType(0)
            TypeName = NamType(1)
            Dim ObjectLib As Library = GetAssembly(RootNamespace)
            Dim ObjectClass As TypeWrapper = ObjectLib.GetTypeByAssembly(TypeName)
            Return ObjectClass
        End If

        Return Nothing
    End Function

    Public Function CreateByRoot(ByVal Name As String) As Object
        Dim RootNamespace As String = Name
        Dim TypeName As String = String.Empty

        If Name.Contains("!") = True Then
            Dim NamType() As String = Name.Split("!")
            RootNamespace = NamType(0)
            TypeName = NamType(1)
            Dim ObjectLib As Library = GetAssembly(RootNamespace)
            Dim ObjectClass As TypeWrapper = ObjectLib.GetTypeByAssembly(TypeName)
            Return ObjectClass.CreateInstance

        End If

        Return Nothing
    End Function

    Public Function GetExam(ByRef Name As Object) As String
        MsgBox("net " & Name)
        Name += " holaperra"
        Return Name
    End Function


#End Region

End Class
