Public Class Library

    Public Property Asm As Reflection.Assembly = Nothing

    Public Function GetTypeByAssembly(ByVal TypeName As String) As TypeWrapper
        Try
            Dim AssembliesList As List(Of Type) = Asm.GetTypes.ToList
            Dim ClassTarget As Type = AssembliesList.Find(Function(p) p.Name = TypeName)
            If ClassTarget Is Nothing Then
                Return Nothing
            Else
                Dim NewTypeWrapper As New TypeWrapper With {.Name = TypeName, .TypeDef = ClassTarget}
                Return NewTypeWrapper
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

End Class

Public Class TypeWrapper
    Public Property Instance As Object = Nothing
    Public Property TypeDef As Type = Nothing
    Public Property Name As String = String.Empty
    Public Property Data() As Object = Nothing

    Public Function CreateInstance() As Object
        Try
            If Data Is Nothing Then
                Instance = DirectCast(Activator.CreateInstance(TypeDef), Object)
                Return Instance
            Else
                Instance = DirectCast(Activator.CreateInstance(TypeDef, Data), Object)
                Return Instance
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
End Class
