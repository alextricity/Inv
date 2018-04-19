Imports System.Data.Odbc
Module Classes
    Public con As OdbcConnection = Nothing
    Public cmd, cmd1 As OdbcCommand
    Public rdr As OdbcDataReader = Nothing
    Public ds As DataSet
    Public adp As OdbcDataAdapter
    Public dtable As DataTable
    Public TempFileNames2 As String
End Module
