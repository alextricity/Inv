Imports System.Data.SqlClient


Public Class frmSalesmanLedger

    Dim a, b, c As String
    Sub fillSalesman()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT RTRIM(Name) FROM Salesman order by 1", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbSalesman.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbSalesman.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub Reset()
        dtpDateFrom.Text = Today
        dtpDateTo.Text = Today
        cmbSalesman.Text = ""
        txtSalesmanID.Text = ""
    End Sub
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub


    Private Sub btnClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub


    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
            If Len(Trim(cmbSalesman.Text)) = 0 Then
                MessageBox.Show("Please Select Salesman", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSalesman.Focus()
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "Select * FROM InvoiceInfo INNER JOIN SalesMan ON InvoiceInfo.SalesmanID = SalesMan.SM_ID INNER JOIN Salesman_Commission ON InvoiceInfo.Inv_ID = Salesman_Commission.InvoiceID where InvoiceDate between @d2 and @d3 and Salesman_ID=@d1"
            cmd = New SqlCommand(ct)
            cmd.Parameters.AddWithValue("@d1", txtSalesmanID.Text)
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d3", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            cmd.Connection = con
            rdr = cmd.ExecuteReader()

            If Not rdr.Read() Then
                MessageBox.Show("Sorry...No record found", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If (rdr IsNot Nothing) Then
                    rdr.Close()
                End If
                Return
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            '  Dim rpt As New rptSalesmanLedger 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand As New SqlCommand()
            Dim myDA As New SqlDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New SqlConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT InvoiceInfo.Inv_ID, InvoiceInfo.InvoiceNo, InvoiceInfo.InvoiceDate, InvoiceInfo.CustomerID, InvoiceInfo.SalesmanID, InvoiceInfo.GrandTotal, InvoiceInfo.TotalPaid, InvoiceInfo.Balance, InvoiceInfo.Remarks,SalesMan.SM_ID, SalesMan.SalesMan_ID, SalesMan.Name, SalesMan.Address, SalesMan.City, SalesMan.State, SalesMan.ZipCode, SalesMan.ContactNo, SalesMan.EmailID, SalesMan.Remarks AS Expr1,SalesMan.Photo, SalesMan.CommissionPer, Salesman_Commission.ID, Salesman_Commission.InvoiceID, Salesman_Commission.CommissionPer AS Expr2, Salesman_Commission.Commission FROM InvoiceInfo INNER JOIN SalesMan ON InvoiceInfo.SalesmanID = SalesMan.SM_ID INNER JOIN Salesman_Commission ON InvoiceInfo.Inv_ID = Salesman_Commission.InvoiceID where InvoiceDate between @d2 and @d3 and Salesman_ID=@d1 order by Inv_ID"
            MyCommand.Parameters.AddWithValue("@d1", txtSalesmanID.Text)
            MyCommand.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            MyCommand.Parameters.Add("@d3", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "InvoiceInfo")
            myDA.Fill(myDS, "Salesman")
            myDA.Fill(myDS, "Salesman_Commission")
          
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmSalesReport_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fillSalesman()
    End Sub

    Private Sub cmbSupplierName_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbSalesman.SelectedIndexChanged
        Try
            a = ""
            b = ""
            c = ""
            txtSalesmanID.Text = ""
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT RTRIM(Salesman_ID),RTRIM(Address),RTRIM(City),RTRIM(ContactNo) FROM Salesman WHERE Name=@d1"
            cmd.Parameters.AddWithValue("@d1", cmbSalesman.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtSalesmanID.Text = rdr.GetValue(0)
                a = rdr.GetValue(1)
                b = rdr.GetValue(2)
                c = rdr.GetValue(3)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
End Class
