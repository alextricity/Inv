Imports System.IO

Module ModCS
    Dim st As String
    Public Function ReadCS() As String
        Using sr As StreamReader = New StreamReader(Application.StartupPath & "\SQLSettings.dat")
            st = "DSN=sims;dbq=C:\Users\aroberts\AppData\Roaming\Inv\SIMSmdb.mdb;defaultdir=C:\Users\aroberts\AppData\Roaming\Inv;driverid=25;fil=MS Access;maxbuffersize=2048;pagetimeout=5;uid=admin;Integrated Security=True" 'sr.ReadLine()
        End Using
        Return st
    End Function
    Public ReadOnly cs As String = ReadCS()
End Module
