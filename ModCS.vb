Imports System.IO
Imports System.Configuration

Module ModCS
    Dim st As String
    Public Function ReadCS() As String
        'Using sr As StreamReader = New StreamReader(Application.StartupPath & "\SQLSettings.dat")
        'st = sr.ReadLine()

        Dim settings As ConnectionStringSettingsCollection = ConfigurationManager.ConnectionStrings

        For Each cs As ConnectionStringSettings In settings
            st = cs.ConnectionString
        Next

        'End Using
        Return st
    End Function
    Public ReadOnly cs As String = ReadCS()
End Module
