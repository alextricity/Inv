Imports System.Data.Odbc
Imports System.IO

Public Class frmSupplier
    Dim s As String
    Dim Photoname As String = ""
    Dim IsImageChanged As Boolean = False
    Sub Reset()
        txtSupplierName.Text = ""
        txtAddress.Text = ""
        txtRemarks.Text = ""
        txtSupplierName.Text = ""
        txtSupplierID.Text = ""
        txtContactNo.Text = ""
        txtEmailID.Text = ""
        cmbState.Text = ""
        txtZipCode.Text = ""
        txtCity.Text = ""
        txtSupplierName.Focus()
        btnSave.Enabled = True
        btnUpdate.Enabled = False
        btnDelete.Enabled = False
        auto()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Function GenerateID() As String
        con = New OdbcConnection(cs)
        Dim value As String = "0000"
        Try
            ' Fetch the latest ID from the database
            con.Open()
            cmd = New OdbcCommand("SELECT TOP 1 ID FROM Supplier ORDER BY ID DESC", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If rdr.HasRows Then
                rdr.Read()
                value = rdr.Item("ID")
            End If
            rdr.Close()
            ' Increase the ID by 1
            value += 1
            ' Because incrementing a string with an integer removes 0's
            ' we need to replace them. If necessary.
            If value <= 9 Then 'Value is between 0 and 10
                value = "000" & value
            ElseIf value <= 99 Then 'Value is between 9 and 100
                value = "00" & value
            ElseIf value <= 999 Then 'Value is between 999 and 1000
                value = "0" & value
            End If
        Catch ex As Exception
            ' If an error occurs, check the connection state and close it if necessary.
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            value = "0000"
        End Try
        Return value
    End Function
    Sub auto()
        Try
            txtID.Text = GenerateID()
            txtSupplierID.Text = "S-" + GenerateID()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try

            If Len(Trim(txtSupplierName.Text)) = 0 Then
                MessageBox.Show("Please enter supplier name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtSupplierName.Focus()
                Exit Sub
            End If
            If Len(Trim(txtAddress.Text)) = 0 Then
                MessageBox.Show("Please Enter Address", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtAddress.Focus()
                Exit Sub
            End If
            If Len(Trim(txtCity.Text)) = 0 Then
                MessageBox.Show("Please Enter City", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtCity.Focus()
                Exit Sub
            End If
            If Len(Trim(txtContactNo.Text)) = 0 Then
                MessageBox.Show("Please Enter Contact No.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtContactNo.Focus()
                Exit Sub
            End If
            con = New OdbcConnection(cs)
            con.Open()
            Dim cb1 As String = "insert into Suppliers(T_ID, TransactionID, Date, PaymentMode, SupplierID, Amount,Remarks) VALUES (@d1,@d2,@d3,@d4,@d5,@d6,@d7)"
            cmd = New OdbcCommand(cb1)
            cmd.Parameters.AddWithValue("@d1", Val(txtT_ID.Text))
            cmd.Parameters.AddWithValue("@d2", "")
            cmd.Parameters.AddWithValue("@d3", System.DateTime.Now.Date)
            cmd.Parameters.AddWithValue("@d4", "")
            cmd.Parameters.AddWithValue("@d5", Val(txtID.Text))
            cmd.Parameters.AddWithValue("@d6", 0)
            cmd.Parameters.AddWithValue("@d7", "")
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            LogFunc(lblUser.Text, "added the new supplier having supplier id '" & txtSupplierID.Text & "'")
            MessageBox.Show("Successfully saved", "Supplier Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            fillState()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Function GenerateID1() As String
        con = New OdbcConnection(cs)
        Dim value As String = "0000"
        Try
            ' Fetch the latest ID from the database
            con.Open()
            cmd = New OdbcCommand("SELECT TOP 1 T_ID FROM Payment ORDER BY T_ID DESC", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If rdr.HasRows Then
                rdr.Read()
                value = rdr.Item("T_ID")
            End If
            rdr.Close()
            ' Increase the ID by 1
            value += 1
            ' Because incrementing a string with an integer removes 0's
            ' we need to replace them. If necessary.
            If value <= 9 Then 'Value is between 0 and 10
                value = "000" & value
            ElseIf value <= 99 Then 'Value is between 9 and 100
                value = "00" & value
            ElseIf value <= 999 Then 'Value is between 999 and 1000
                value = "0" & value
            End If
        Catch ex As Exception
            ' If an error occurs, check the connection state and close it if necessary.
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            value = "0000"
        End Try
        Return value
    End Function
    Sub auto1()
        Try
            txtT_ID.Text = GenerateID1()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            If MessageBox.Show("Do you really want to delete this record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                DeleteRecord()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub DeleteRecord()
        Try
            Dim RowsAffected As Integer = 0
            con = New OdbcConnection(cs)
            con.Open()
            Dim cl As String = "SELECT Supplier.ID FROM Supplier INNER JOIN Stock ON Supplier.ID = Stock.SupplierID where Supplier.ID=@d1"
            cmd = New OdbcCommand(cl)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", Val(txtID.Text))
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Purchase Entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con.Close()
            con = New OdbcConnection(cs)
            con.Open()
            Dim cl1 As String = "SELECT Supplier.ID FROM Supplier INNER JOIN Payment ON Supplier.ID = Payment.SupplierID where Supplier.ID=@d1 and Amount >0"
            cmd = New OdbcCommand(cl1)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", Val(txtID.Text))
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Payment Entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con.Close()
            con = New OdbcConnection(cs)
            con.Open()
            Dim cq1 As String = "delete from Payment where SupplierID =" & txtID.Text & " and Amount=0"
            cmd = New OdbcCommand(cq1)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            con = New OdbcConnection(cs)
            con.Open()
            Dim cq As String = "delete from Supplier where ID =" & txtID.Text & ""
            cmd = New OdbcCommand(cq)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                LogFunc(lblUser.Text, "deleted the supplier record having supplier id '" & txtSupplierID.Text & "'")
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
                fillState()
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Len(Trim(txtSupplierName.Text)) = 0 Then
            MessageBox.Show("Please enter supplier name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtSupplierName.Focus()
            Exit Sub
        End If
        If Len(Trim(txtAddress.Text)) = 0 Then
            MessageBox.Show("Please Enter Address", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAddress.Focus()
            Exit Sub
        End If
        If Len(Trim(txtCity.Text)) = 0 Then
            MessageBox.Show("Please Enter City", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtCity.Focus()
            Exit Sub
        End If
        If cmbState.Text = "" Then
            MessageBox.Show("Please enter state", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbState.Focus()
            Return
        End If
        If Len(Trim(txtContactNo.Text)) = 0 Then
            MessageBox.Show("Please Enter Contact No.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtContactNo.Focus()
            Exit Sub
        End If

        Try
            con = New OdbcConnection(cs)
            con.Open()
            Dim cb1 As String = "update LedgerBook set [Name]=@d3 where PartyID=@d1 and Name=@d2"
            cmd = New OdbcCommand(cb1)
            cmd.Parameters.AddWithValue("@d1", txtSupplierID.Text)
            cmd.Parameters.AddWithValue("@d2", txtSupName.Text)
            cmd.Parameters.AddWithValue("@d3", txtSupplierName.Text)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            con = New OdbcConnection(cs)
            con.Open()
            Dim cb3 As String = "update SupplierLedgerBook set [Name]=@d3 where PartyID=@d1 and Name=@d2"
            cmd = New OdbcCommand(cb3)
            cmd.Parameters.AddWithValue("@d1", txtSupplierID.Text)
            cmd.Parameters.AddWithValue("@d2", txtSupName.Text)
            cmd.Parameters.AddWithValue("@d3", txtSupplierName.Text)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            con = New OdbcConnection(cs)
            con.Open()
            Dim cb2 As String = "update LedgerBook set Label=@d1 where Label=@d2"
            cmd = New OdbcCommand(cb2)
            cmd.Parameters.AddWithValue("@d1", "Payment to " & txtSupplierName.Text & "")
            cmd.Parameters.AddWithValue("@d2", "Payment to " & txtSupName.Text & "")
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            con = New OdbcConnection(cs)
            con.Open()
            Dim cb As String = "update supplier set SupplierID=@d2,[Name]=@d3, Address=@d5,City=@d6, ContactNo=@d7, EmailID=@d8,Remarks=@d9,State=@d10,ZipCode=@d11 where ID=@d1"
            cmd = New OdbcCommand(cb)
            cmd.Parameters.AddWithValue("@d2", txtSupplierID.Text)
            cmd.Parameters.AddWithValue("@d3", txtSupplierName.Text)
            cmd.Parameters.AddWithValue("@d5", txtAddress.Text)
            cmd.Parameters.AddWithValue("@d6", txtCity.Text)
            cmd.Parameters.AddWithValue("@d7", txtContactNo.Text)
            cmd.Parameters.AddWithValue("@d8", txtEmailID.Text)
            cmd.Parameters.AddWithValue("@d9", txtRemarks.Text)
            cmd.Parameters.AddWithValue("@d10", cmbState.Text)
            cmd.Parameters.AddWithValue("@d11", txtZipCode.Text)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", Val(txtID.Text))
            cmd.ExecuteNonQuery()
            LogFunc(lblUser.Text, "updated the supplier having supplier id '" & txtSupplierID.Text & "'")
            MessageBox.Show("Successfully updated", "Supplier Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Enabled = False
            fillState()
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Reset()
    End Sub


    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click

    End Sub

    Private Sub frmSupplier_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fillState()
    End Sub
    Sub fillState()
        Try
            con = New OdbcConnection(cs)
            con.Open()
            adp = New OdbcDataAdapter()
            adp.SelectCommand = New OdbcCommand("SELECT distinct RTRIM(State) FROM Supplier order by 1", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbState.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbState.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub cmbState_Format(sender As System.Object, e As System.Windows.Forms.ListControlConvertEventArgs) Handles cmbState.Format
        If (e.DesiredType Is GetType(String)) Then
            e.Value = e.Value.ToString.Trim
        End If
    End Sub
End Class
