Imports System.Data.Odbc
Imports System.Security.Cryptography
Imports System.IO
Imports System.Text
Imports System.Collections.Generic
Imports System.Net
Imports System.Web
Imports Excel = Microsoft.Office.Interop.Excel
Module ModFunc
    Public Function CheckForInternetConnection() As Boolean
        Try
            Using client = New WebClient()
                Using stream = client.OpenRead("http://www.google.com")
                    Return True
                End Using
            End Using
        Catch
            Return False
        End Try
    End Function
    Sub SMS(ByVal st1 As String)
        con = New OdbcConnection(cs)
        con.Open()
        Dim cb As String = "insert into SMS(Message,Date) VALUES ('" & st1 & "','" & System.DateTime.Now & ")"
        cmd = New OdbcCommand(cb)
        cmd.Connection = con
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
    Sub LogFunc(ByVal st1 As String, ByVal st2 As String)
        con = New OdbcConnection(cs)
        con.Open()
        Dim cb As String = "insert into Logs (UserID, [Date], Operation) VALUES ('" & st1 & "', '" & System.DateTime.Now & "','" & st2 & "')"
        cmd = New OdbcCommand(cb)
        cmd.Connection = con
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
    Sub SMSFunc(ByVal st1 As String, ByVal st2 As String, ByVal st3 As String)
        st3 = st3.Replace("@MobileNo", st1).Replace("@Message", st2)
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse = Nothing
        Dim myUri As New Uri(st3)
        request = DirectCast(WebRequest.Create(myUri), HttpWebRequest)
        response = DirectCast(request.GetResponse(), HttpWebResponse)
    End Sub
    Public Function Encrypt(password As String) As String
        Dim strmsg As String = String.Empty
        Dim encode As Byte() = New Byte(password.Length - 1) {}
        encode = Encoding.UTF8.GetBytes(password)
        strmsg = Convert.ToBase64String(encode)
        Return strmsg
    End Function

    Public Function Decrypt(encryptpwd As String) As String
        Dim decryptpwd As String = String.Empty
        Dim encodepwd As New UTF8Encoding()
        Dim Decode As Decoder = encodepwd.GetDecoder()
        Dim todecode_byte As Byte() = Convert.FromBase64String(encryptpwd)
        Dim charCount As Integer = Decode.GetCharCount(todecode_byte, 0, todecode_byte.Length)
        Dim decoded_char As Char() = New Char(charCount - 1) {}
        Decode.GetChars(todecode_byte, 0, todecode_byte.Length, decoded_char, 0)
        decryptpwd = New [String](decoded_char)
        Return decryptpwd
    End Function
    Sub RefreshRecords()
        'Dim obj As frmMainMenu = DirectCast(Application.OpenForms("frmMainMenu"), frmMainMenu)
        'obj.Getdata()
        'obj.DataGridView1.Refresh()
        'obj.DataGridView1.Update()
    End Sub
    Sub ExportExcel(ByVal obj As Object)
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = obj.RowCount
            colsTotal = obj.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = obj.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = obj.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub
    Sub LedgerSave(ByVal a As DateTime, ByVal b As String, ByVal c As String, ByVal d As String, ByVal e As Decimal, ByVal f As Decimal, ByVal g As String)
        con = New OdbcConnection(cs)
        con.Open()
        Dim cb As String = "insert into LedgerBook(Date, Name, LedgerNo, Label,Debit,Credit,PartyID) Values (@d1,@d2,@d3,@d4,@d5,@d6,@d7)"
        cmd = New OdbcCommand(cb)
        cmd.Parameters.AddWithValue("@d1", a)
        cmd.Parameters.AddWithValue("@d2", b)
        cmd.Parameters.AddWithValue("@d3", c)
        cmd.Parameters.AddWithValue("@d4", d)
        cmd.Parameters.AddWithValue("@d5", e)
        cmd.Parameters.AddWithValue("@d6", f)
        cmd.Parameters.AddWithValue("@d7", g)
        cmd.Connection = con
        cmd.ExecuteReader()
        con.Close()
    End Sub
    Sub LedgerDelete(ByVal a As String)
        con = New OdbcConnection(cs)
        con.Open()
        Dim cq As String = "delete from LedgerBook where LedgerNo=@d1"
        cmd = New OdbcCommand(cq)
        cmd.Parameters.AddWithValue("@d1", a)
        cmd.Connection = con
        cmd.ExecuteReader()
        con.Close()
    End Sub
    Sub LedgerUpdate(ByVal a As DateTime, ByVal b As String, ByVal e As Decimal, ByVal f As Decimal, ByVal g As String, ByVal h As String)
        con = New OdbcConnection(cs)
        con.Open()
        Dim cb As String = "Update LedgerBook set Date=@d1, Name=@d2,Debit=@d3,Credit=@d4 where LedgerNo=@d5 and Label=@d6"
        cmd = New OdbcCommand(cb)
        cmd.Parameters.AddWithValue("@d1", a)
        cmd.Parameters.AddWithValue("@d2", b)
        cmd.Parameters.AddWithValue("@d3", e)
        cmd.Parameters.AddWithValue("@d4", f)
        cmd.Parameters.AddWithValue("@d5", g)
        cmd.Parameters.AddWithValue("@d6", h)
        cmd.Connection = con
        cmd.ExecuteReader()
        con.Close()
    End Sub
    Sub SupplierLedgerSave(ByVal a As DateTime, ByVal b As String, ByVal c As String, ByVal d As String, ByVal e As Decimal, ByVal f As Decimal, ByVal g As String)
        con = New OdbcConnection(cs)
        con.Open()
        Dim cb As String = "insert into SupplierLedgerBook(Date, Name, LedgerNo, Label,Debit,Credit,PartyID) Values (@d1,@d2,@d3,@d4,@d5,@d6,@d7)"
        cmd = New OdbcCommand(cb)
        cmd.Parameters.AddWithValue("@d1", a)
        cmd.Parameters.AddWithValue("@d2", b)
        cmd.Parameters.AddWithValue("@d3", c)
        cmd.Parameters.AddWithValue("@d4", d)
        cmd.Parameters.AddWithValue("@d5", e)
        cmd.Parameters.AddWithValue("@d6", f)
        cmd.Parameters.AddWithValue("@d7", g)
        cmd.Connection = con
        cmd.ExecuteReader()
        con.Close()
    End Sub
    Sub SupplierLedgerDelete(ByVal a As String)
        con = New OdbcConnection(cs)
        con.Open()
        Dim cq As String = "delete from SupplierLedgerBook where LedgerNo=@d1"
        cmd = New OdbcCommand(cq)
        cmd.Parameters.AddWithValue("@d1", a)
        cmd.Connection = con
        cmd.ExecuteReader()
        con.Close()
    End Sub
    Sub SupplierLedgerUpdate(ByVal a As DateTime, ByVal b As String, ByVal e As Decimal, ByVal f As Decimal, ByVal g As String, ByVal h As String)
        con = New OdbcConnection(cs)
        con.Open()
        Dim cb As String = "Update SupplierLedgerBook set Date=@d1, Name=@d2,Debit=@d3,Credit=@d4 where LedgerNo=@d5 and Label=@d6"
        cmd = New OdbcCommand(cb)
        cmd.Parameters.AddWithValue("@d1", a)
        cmd.Parameters.AddWithValue("@d2", b)
        cmd.Parameters.AddWithValue("@d3", e)
        cmd.Parameters.AddWithValue("@d4", f)
        cmd.Parameters.AddWithValue("@d5", g)
        cmd.Parameters.AddWithValue("@d6", h)
        cmd.Connection = con
        cmd.ExecuteReader()
        con.Close()
    End Sub
End Module
