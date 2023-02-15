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

    Public Property _Instance As Object = Nothing
    Public ReadOnly Property Instance As Object
        Get
            Return _Instance
        End Get
    End Property

    Public Property _TypeDef As Object = Nothing
    Public Property TypeDef As Object
        Set(value As Object)
            _TypeDef = value
        End Set
        Get
            Return _TypeDef
        End Get
    End Property

    Public Property _Name As String = String.Empty
    Public Property Name As String
        Set(value As String)
            _Name = value
        End Set
        Get
            Return _Name
        End Get
    End Property

    Private Property _Data() As Object = Nothing
    Public Property Data() As Object
        Set(value As Object)
            _Data = value
        End Set
        Get
            Return _Data
        End Get
    End Property

    Public Function CreateInstance() As Object
        Try
            If _Data Is Nothing Then
                _Instance = DirectCast(Activator.CreateInstance(TypeDef), Object)
                Return Instance
            Else
                _Instance = DirectCast(Activator.CreateInstance(TypeDef, _Data), Object)
                Return Instance
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function GetMethods(ByVal NameAndType As String) As Object
        Dim MethodName As String = NameAndType
        Dim TypeName As String = String.Empty

        If NameAndType.Contains("!") = True Then
            Dim NamType() As String = NameAndType.Split("!")
            MethodName = NamType(0)
            TypeName = NamType(1)
        End If
        '   Console.WriteLine("Method : " & MethodName & "   Type : " & TypeName)

        Dim MethodList As New List(Of Reflection.MethodInfo)
        For Each Method As Reflection.MethodInfo In DirectCast(_TypeDef, Type).GetMethods
            If String.Equals(Method.Name, MethodName, StringComparison.OrdinalIgnoreCase) Then

                If TypeName = String.Empty Then

                    If Method.ReturnType.Name = "Void" Then
                        MethodList.Add(Method)
                    End If

                ElseIf String.Equals(Method.ReturnType.Name, TypeName, StringComparison.OrdinalIgnoreCase) Then

                    MethodList.Add(Method)

                End If

            End If
        Next
        ' Console.WriteLine("Method Count: " & MethodList.Count)
        If MethodList.Count = 0 Then
            Return Nothing
        Else
            Dim MethodWrapperEx As New MethodWrapper
            MethodWrapperEx.MethodDef = MethodList.ToList
            Return MethodWrapperEx
        End If
    End Function


End Class

Public Class MethodWrapper

    Public Property MethodDef As List(Of Reflection.MethodInfo) = Nothing

    Public Function InvokeCall(Optional ByVal Parameters As String = "") As Object

        Dim FastArgumentParser As Core.FastArgumentParser = New Core.FastArgumentParser()
        Dim ResultEx As Object = Nothing

        If Parameters = "" Then

            For Each Method As Reflection.MethodInfo In _MethodDef
                If Method.GetParameters().Count = 0 Then
                    ResultEx = Method.Invoke(Nothing, Nothing)
                    Exit For
                End If
            Next

        Else

            Dim CommandLineArgs As String() = Parameters.Split(" ")

            FastArgumentParser.Parse(CommandLineArgs)

            Dim ArgumentArray As New Dictionary(Of Integer, List(Of Object))

            For MethodEx As Integer = 0 To (_MethodDef.Count - 1)

                ' Console.WriteLine("Processing : " & _MethodDef(MethodEx).Name)

                Dim Method As Reflection.MethodInfo = _MethodDef(MethodEx)
                Dim IsFuntionValid As Boolean = False
                Dim Parameter As List(Of Reflection.ParameterInfo) = Method.GetParameters().ToList
                '  Console.WriteLine("Parameters : " & Parameter.Count & "   Arguments : " & FastArgumentParser.UnknownArguments.Count)
                If Parameter.Count = FastArgumentParser.UnknownArguments.Count Then
                    IsFuntionValid = True
                Else
                    IsFuntionValid = HaveOptional(Parameter)
                End If

                If IsFuntionValid = False Then
                    Exit For
                End If

                ' Console.WriteLine("IsValid : " & IsFuntionValid & vbNewLine & vbNewLine)

                Dim ParamEx As New List(Of Object)

                For i As Integer = 0 To (Parameter.Count - 1)

                    Try

                        Dim Param As Reflection.ParameterInfo = Parameter(i)
                        Dim ArgParsed As Core.IArgument = FastArgumentParser.UnknownArguments(i)

                        '  Console.WriteLine("Param : " & Param.ParameterType.Name & "    Arg   " & ArgParsed.Name)
                        ' Console.WriteLine(String.Equals(Param.ParameterType.Name, ArgParsed.Name.Substring(1, ArgParsed.Name.Length - 1), StringComparison.OrdinalIgnoreCase))

                        If String.Equals(Param.ParameterType.Name, ArgParsed.Name.Substring(1, ArgParsed.Name.Length - 1), StringComparison.OrdinalIgnoreCase) Then

                            If ArgParsed.Value Is Nothing Then
                                ' Console.WriteLine("Value : Is Nothing")
                                If Param.IsOptional Then
                                    ParamEx.Add(Param.DefaultValue)
                                Else
                                    ParamEx.Clear()
                                    Exit For
                                End If
                            Else
                                '  Console.WriteLine("Adding : ")
                                Dim CtypeEx As Object = ArgParsed.ConvertType
                                ParamEx.Add(CtypeEx)
                            End If

                        Else
                            ParamEx.Clear()
                            Exit For
                        End If

                    Catch ex As Exception
                        ParamEx.Clear()
                        Exit For
                    End Try

                Next

                If Not ParamEx.Count = 0 Then
                    ArgumentArray.Add(MethodEx, ParamEx)
                End If

            Next
            '  Console.WriteLine("ArgumentArray Count : " & ArgumentArray.Count)
            If Not ArgumentArray.Count = 0 Then
                Dim Index As Integer = ArgumentArray.FirstOrDefault.Key
                ResultEx = _MethodDef(Index).Invoke(Nothing, ArgumentArray.FirstOrDefault.Value.ToArray)
                '   Console.WriteLine("Result  : " & ResultEx)
            End If

        End If

        Return ResultEx

    End Function

    Public Function GetMethod() As Object()
        Return MethodDef.ToArray
    End Function

    Private Function HaveOptional(ByVal Param As List(Of Reflection.ParameterInfo)) As Boolean
        For Each Parameter As Reflection.ParameterInfo In Param
            If Parameter.IsOptional = True Then
                Return True
            End If
        Next
        Return False
    End Function


End Class


Public Class TypeConverter




End Class