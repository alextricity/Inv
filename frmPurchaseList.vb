﻿Imports System.Data.Odbc

Imports System.IO

Public Class frmPurchaseList

    Public Sub Getdata()
        Try
            con = New OdbcConnection(cs)
            con.Open()
            cmd = New OdbcCommand("SELECT ST_ID, RTRIM(InvoiceNo), Date,RTRIM(PurchaseType),Supplier.ID, RTRIM(Supplier.SupplierID),RTRIM(Name), SubTotal, DiscountPer, Discount, FreightCharges, OtherCharges,PreviousDue, Total, RoundOff, GrandTotal, TotalPayment, PaymentDue, RTRIM(Stock.Remarks) from Supplier,Stock where Supplier.ID=Stock.SupplierID order by [Date]", con)
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6), rdr(7), rdr(8), rdr(9), rdr(10), rdr(11), rdr(12), rdr(13), rdr(14), rdr(15), rdr(16), rdr(17), rdr(18))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub frmLogs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Getdata()
    End Sub

    Private Sub dgw_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                If lblSet.Text = "Purchase" Then
                    Dim dr As DataGridViewRow = dgw.SelectedRows(0)


                    con = New OdbcConnection(cs)
                    con.Open()
                    Dim sql As String = "SELECT PID,RTRIM(Product.ProductCode),RTRIM(Productname),Price,Qty,TotalAmount from Stock,Stock_Product,product where product.PID=Stock_product.ProductID and Stock.ST_ID=Stock_Product.StockID and ST_ID=" & dr.Cells(0).Value & ""
                    cmd = New OdbcCommand(sql, con)
                    rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                    While (rdr.Read() = True)
                        frmStockStatus.DataGridView1.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5))
                    End While
                    con.Close()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dgw_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles dgw.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If dgw.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            dgw.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlLightLight
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub
    Sub Reset()
        txtSupplierName.Text = ""
        dtpDateFrom.Text = Today
        dtpDateTo.Text = Today
        Getdata()
    End Sub
    Private Sub btnReset_Click(sender As System.Object, e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub txtSupplierName_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtSupplierName.TextChanged
        Try
            con = New OdbcConnection(cs)
            con.Open()
            cmd = New OdbcCommand("SELECT ST_ID, RTRIM(InvoiceNo), Date,RTRIM(PurchaseType),Supplier.ID, RTRIM(Supplier.SupplierID),RTRIM(Name), SubTotal, DiscountPer, Discount, FreightCharges, OtherCharges,PreviousDue, Total, RoundOff, GrandTotal, TotalPayment, PaymentDue, RTRIM(Stock.Remarks) from Supplier,Stock where Supplier.ID=Stock.SupplierID  and [Name] like '%" & txtSupplierName.Text & "%' order by [Date]", con)
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6), rdr(7), rdr(8), rdr(9), rdr(10), rdr(11), rdr(12), rdr(13), rdr(14), rdr(15), rdr(16), rdr(17), rdr(18))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub btnExportExcel_Click(sender As System.Object, e As System.EventArgs) Handles btnExportExcel.Click
        ExportExcel(dgw)
    End Sub

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        Try
            con = New OdbcConnection(cs)
            con.Open()
            cmd = New OdbcCommand("SELECT ST_ID, RTRIM(InvoiceNo), Date,RTRIM(PurchaseType),Supplier.ID, RTRIM(Supplier.SupplierID),RTRIM(Name), SubTotal, DiscountPer, Discount, FreightCharges, OtherCharges,PreviousDue, Total, RoundOff, GrandTotal, TotalPayment, PaymentDue, RTRIM(Stock.Remarks) from Supplier,Stock where Supplier.ID=Stock.SupplierID  and [Date] between @d1 and @d2 order by [Date]", con)
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6), rdr(7), rdr(8), rdr(9), rdr(10), rdr(11), rdr(12), rdr(13), rdr(14), rdr(15), rdr(16), rdr(17), rdr(18))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
