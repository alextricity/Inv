Imports System.Data.Odbc

Public Class frmVouchersEntry


    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Sub Reset()
        txtVoucherID.Text = ""
        txtName.Text = ""
        txtDetails.Text = ""
        txtParticulars.Text = ""
        txtNotes.Text = ""
        txtVoucherNo.Text = ""
        txtAmount.Text = ""
        txtGrandTotal.Text = ""
        dtpDate.Text = Today
        DataGridView1.Rows.Clear()
        btnPrint.Enabled = False
        btnSave.Enabled = True
        btnDelete.Enabled = False
        btnUpdate.Enabled = False
        btnAdd.Enabled = True
        btnRemove.Enabled = False
        Clear()
        auto()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Try
            If txtParticulars.Text = "" Then
                MessageBox.Show("Please enter particulars", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtParticulars.Focus()
                Exit Sub
            End If
            If txtAmount.Text = "" Then
                MessageBox.Show("Please enter amount", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtAmount.Focus()
                Exit Sub
            End If
            If DataGridView1.Rows.Count = 0 Then
                Dim k As Double = 0
                k = GrandTotal()
                k = Math.Round(k, 2)
                txtGrandTotal.Text = k
                Clear()
                Exit Sub
            End If
            DataGridView1.Rows.Add(txtParticulars.Text, Val(txtAmount.Text), txtNotes.Text)
            Dim j As Double = 0
            j = GrandTotal()
            j = Math.Round(j, 2)
            txtGrandTotal.Text = j
            Clear()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub Clear()
        txtParticulars.Text = ""
        txtAmount.Text = ""
        txtNotes.Text = ""
        btnAdd.Enabled = True
        btnRemove.Enabled = False
    End Sub

    Public Function GrandTotal() As Double
        Dim sum As Double = 0
        Try
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                sum = sum + r.Cells(1).Value
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return sum
    End Function


    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Reset()
        Reset()
    End Sub
    Private Function GenerateID() As String
        con = New OdbcConnection(cs)
        Dim value As String = "0000"
        Try
            ' Fetch the latest ID from the database
            con.Open()
            cmd = New OdbcCommand("SELECT TOP 1 ID FROM Voucher ORDER BY ID DESC", con)
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
            txtVoucherID.Text = GenerateID()
            txtVoucherNo.Text = "V-" + GenerateID()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Sub Print()
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            'Dim rpt As New rptVoucher 'The report you created.
            Dim myConnection As OdbcConnection
            Dim MyCommand, MyCommand1 As New OdbcCommand()
            Dim myDA, myDA1 As New OdbcDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New OdbcConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand1.Connection = myConnection
            MyCommand.CommandText = "SELECT Voucher.ID, Voucher.VoucherNo, Voucher.Date, Voucher.Name, Voucher.Details, Voucher.GrandTotal, Voucher_OtherDetails.VD_ID, Voucher_OtherDetails.VoucherID,Voucher_OtherDetails.Particulars, Voucher_OtherDetails.Amount, Voucher_OtherDetails.Note FROM Voucher INNER JOIN Voucher_OtherDetails ON Voucher.ID = Voucher_OtherDetails.VoucherID  where VoucherNo='" & txtVoucherNo.Text & "'"
            MyCommand1.CommandText = "SELECT * from Company"
            myDA.Fill(myDS, "Voucher")
            myDA.Fill(myDS, "Voucher_OtherDetails")
            myDA1.Fill(myDS, "Company")

            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If txtName.Text = "" Then
                MessageBox.Show("Please enter voucher name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtName.Focus()
                Exit Sub
            End If
            If DataGridView1.Rows.Count = 0 Then
                MessageBox.Show("sorry no data added to grid", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            con = New OdbcConnection(cs)
            con.Open()
            Dim cb As String = "insert into Voucher(Id, VoucherNo, Date,Name,Details,GrandTotal) Values (@d1,@d2,@d3,@d4,@d5,@d7)"
            cmd = New OdbcCommand(cb)
            cmd.Parameters.AddWithValue("@d1", Val(txtVoucherID.Text))
            cmd.Parameters.AddWithValue("@d2", txtVoucherNo.Text)
            cmd.Parameters.AddWithValue("@d3", dtpDate.Value.Date)
            con.Close()
            con = New OdbcConnection(cs)
            con.Open()
            Dim cb1 As String = "insert into Voucher_OtherDetails(VoucherID,Particulars,Amount,Note) VALUES (" & txtVoucherID.Text & ",@d1,@d2,@d3)"
            Dim st As String = "added the new voucher having voucher no.'" & txtVoucherNo.Text & "'"
            LogFunc(lblUser.Text, st)
            LedgerSave(dtpDate.Value.Date, txtName.Text, txtVoucherNo.Text, "Expenses", Val(txtGrandTotal.Text), 0, "")
            btnSave.Enabled = False
            MessageBox.Show("Successfully saved", "Voucher", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Print()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub DeleteRecord()
        Try
            Dim RowsAffected As Integer = 0
            con = New OdbcConnection(cs)
            con.Open()
            Dim ct As String = "delete from Voucher where ID=" & txtVoucherID.Text & ""
            cmd = New OdbcCommand(ct)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            If RowsAffected > 0 Then
                LedgerDelete(txtVoucherNo.Text)
                Dim st As String = "deleted the voucher having voucher no.'" & txtVoucherNo.Text & "'"
                LogFunc(lblUser.Text, st)
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
                Reset()
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
            End If
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Try
            If txtName.Text = "" Then
                MessageBox.Show("Please enter voucher name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtName.Focus()
                Exit Sub
            End If
            If DataGridView1.Rows.Count = 0 Then
                MessageBox.Show("sorry no data added to grid", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            con.Open()
            Dim ct As String = "delete from Voucher_OtherDetails where VoucherID=" & txtVoucherID.Text & ""
            cmd = New OdbcCommand(ct)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            con = New OdbcConnection(cs)
            con.Open()
            Dim cb1 As String = "insert into Voucher_OtherDetails(VoucherID,Particulars,Amount,Note) VALUES (" & txtVoucherID.Text & ",@d1,@d2,@d3)"
            cmd = New OdbcCommand(cb1)
            cmd.Connection = con
            ' Prepare command for repeated execution
            cmd.Prepare()
            ' Data to be inserted
             con.Close()
            Dim st As String = "updated the voucher having voucher no.'" & txtVoucherNo.Text & "'"
            LogFunc(lblUser.Text, st)
            LedgerUpdate(dtpDate.Value.Date, txtName.Text, Val(txtGrandTotal.Text), 0, txtVoucherNo.Text, "Expenses")
            btnUpdate.Enabled = False
            MessageBox.Show("Successfully updated", "Voucher", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub DataGridView1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseClick
        btnRemove.Enabled = True
    End Sub


    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        Try
            For Each row As DataGridViewRow In DataGridView1.SelectedRows
                DataGridView1.Rows.Remove(row)
            Next
            Dim k As Double = 0
            k = GrandTotal()
            k = Math.Round(k, 2)
            txtGrandTotal.Text = k
            btnRemove.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Print()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub txtAmount_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtAmount.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtAmount.Text
            Dim selectionStart = Me.txtAmount.SelectionStart
            Dim selectionLength = Me.txtAmount.SelectionLength

            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)

            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an integer that is longer than 16 digits.
                e.Handled = True
            ElseIf Double.TryParse(text, New Double) AndAlso text.IndexOf("."c) < text.Length - 3 Then
                'Reject a real number with two many decimal places.
                e.Handled = False
            End If
        Else
            'Reject all other characters.
            e.Handled = True
        End If
    End Sub

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
    
    End Sub

    Private Sub frmVoucher_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
