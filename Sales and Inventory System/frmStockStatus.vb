Imports System.Data.Odbc
Public Class frmStockStatus

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub
    Sub Reset()
        txtProductName.Text = ""
        Getdata()
    End Sub
    Private Sub dgw_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)

    End Sub

    Private Sub dgw_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs)

    End Sub
    Public Sub Getdata()
        Try
            con = New OdbcConnection(cs)
            con.Open()
            cmd = New OdbcCommand("SELECT RTRIM(Product.ProductCode),RTRIM(ProductName),CostPrice,SellingPrice,Discount,VAT,Qty from Temp_Stock,Product where Product.PID=Temp_Stock.ProductID and Qty > 0 order by ProductName", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            DataGridView1.Rows.Clear()
            While (rdr.Read() = True)
                DataGridView1.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6))
            End While
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                con = New OdbcConnection(cs)
                con.Open()
                Dim ct As String = "select ReorderPoint from Product where ProductCode=@d1"
                cmd = New OdbcCommand(ct)
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@d1", r.Cells(0).Value.ToString())
                rdr = cmd.ExecuteReader()
                If (rdr.Read()) Then
                    Dim i As Integer
                    i = rdr(0)
                    If r.Cells(6).Value < i Then
                        r.DefaultCellStyle.BackColor = Color.Red
                    End If
                End If
            Next
            con.Close()
            DataGridView1.ClearSelection()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub frmcategory_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Reset()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub btnReset_Click(sender As System.Object, e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub

    Private Sub btnExportExcel_Click(sender As System.Object, e As System.EventArgs) Handles btnExportExcel.Click
        ExportExcel(DataGridView1)
    End Sub

    Private Sub txtProductName_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtProductName.TextChanged
        Try
            con = New OdbcConnection(cs)
            con.Open()
            cmd = New OdbcCommand("SELECT RTRIM(Product.ProductCode),RTRIM(ProductName),CostPrice,SellingPrice,Discount,VAT,Qty from Temp_Stock,Product where Product.PID=Temp_Stock.ProductID and qty > 0 and ProductName like '%" & txtProductName.Text & "%' order by ProductName", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            DataGridView1.Rows.Clear()
            While (rdr.Read() = True)
                DataGridView1.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6))
            End While
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                con = New OdbcConnection(cs)
                con.Open()
                Dim ct As String = "select ReorderPoint from Product where ProductCode=@d1"
                cmd = New OdbcCommand(ct)
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@d1", r.Cells(0).Value.ToString())
                rdr = cmd.ExecuteReader()
                If (rdr.Read()) Then
                    Dim i As Integer
                    i = rdr(0)
                    If r.Cells(6).Value < i Then
                        r.DefaultCellStyle.BackColor = Color.Red
                    End If
                End If
            Next
            con.Close()
            DataGridView1.ClearSelection()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView1_RowPostPaint(sender As System.Object, e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView1.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView1.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ButtonHighlight
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub
End Class
