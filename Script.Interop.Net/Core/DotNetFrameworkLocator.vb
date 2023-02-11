Imports Microsoft.Win32

Friend Module DotNetFrameworkLocator
    Function GetInstallationLocation() As String
        Try
            Const subkey As String = "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\"

            Using ndpKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey(subkey, False)
                If ndpKey Is Nothing Then Return String.Empty
                Dim value = TryCast(ndpKey.GetValue("InstallPath"), String)

                If value IsNot Nothing Then
                    Return value
                Else
                    Return String.Empty
                End If
            End Using
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function
End Module
