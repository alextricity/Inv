Imports System.Data.Odbc
Imports System.IO

Imports Microsoft.SqlServer.Management.Smo
Imports System.Globalization

Public Class frmMainMenu
    Dim Filename As String

    Public Sub Getdata()

    End Sub
    Private Sub AboutToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        frmAbout.ShowDialog()
    End Sub
    Sub Backup()
        Try
            Dim dt As DateTime = Today
            Dim destdir As String = "SIS_DB " & System.DateTime.Now.ToString("dd-MM-yyyy_h-mm-ss") & ".bak"
            Dim objdlg As New SaveFileDialog
            objdlg.FileName = destdir
            objdlg.ShowDialog()
            Filename = objdlg.FileName
            Cursor = Cursors.WaitCursor
            Timer2.Enabled = True
            con = New OdbcConnection(cs)
            con.Open()
            Dim cb As String = "backup database SIS_DB to disk='" & Filename & "'with init,stats=10"
            cmd = New OdbcCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub BackupToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub Timer2_Tick(sender As System.Object, e As System.EventArgs) Handles Timer2.Tick
        Cursor = Cursors.Default
        Timer2.Enabled = False
    End Sub

    Private Sub RestoreToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub RegistrationToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles RegistrationToolStripMenuItem.Click
        frmRegistration.lblUser.Text = lblUser.Text
        frmRegistration.Reset()
        frmRegistration.ShowDialog()
    End Sub

    Private Sub LogsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles LogsToolStripMenuItem.Click
        frmLogs.Reset()
        frmLogs.lblUser.Text = lblUser.Text
        frmLogs.ShowDialog()
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Dim dt As DateTime = Today
        lblDateTime.Text = dt.ToString("dd/MM/yyyy")
        lblTime.Text = TimeOfDay.ToString("h:mm:ss tt")
    End Sub

    Private Sub CalculatorToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CalculatorToolStripMenuItem.Click
        Try
            System.Diagnostics.Process.Start("Calc.exe")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub NotepadToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles NotepadToolStripMenuItem.Click
        Try
            System.Diagnostics.Process.Start("Notepad.exe")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub WordpadToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles WordpadToolStripMenuItem.Click
        Try
            System.Diagnostics.Process.Start("wordpad.exe")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub MSWordToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles MSWordToolStripMenuItem.Click
        Try
            System.Diagnostics.Process.Start("winword.exe")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TaskManagerToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles TaskManagerToolStripMenuItem.Click
        Try
            System.Diagnostics.Process.Start("TaskMgr.exe")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SystemInfoToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SystemInfoToolStripMenuItem.Click

    End Sub
    Sub LogOut()

        frmProductEntry.Hide()
        Dim st As String = "Successfully logged out"
        LogFunc(lblUser.Text, st)
        Me.Hide()
        frmLogin.Show()
        frmLogin.UserID.Text = ""
        frmLogin.Password.Text = ""
        frmLogin.UserID.Focus()
    End Sub
    Private Sub LogoutToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles LogoutToolStripMenuItem.Click
        Try
            If MessageBox.Show("Do you really want to logout from application?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                If MessageBox.Show("Do you want backup database before logout?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                    Backup()
                    LogOut()
                Else
                    LogOut()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmMainMenu_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        e.Cancel = True
    End Sub

    Private Sub CompanyInfoToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CompanyInfoToolStripMenuItem.Click
        frmCompany.lblUser.Text = lblUser.Text
        frmCompany.Reset()
        frmCompany.ShowDialog()
    End Sub

    Private Sub CustomerToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CustomerToolStripMenuItem.Click
        frmCustomer.lblUser.Text = lblUser.Text
        frmCustomer.Reset()
        frmCustomer.ShowDialog()
    End Sub

    Private Sub CategoryToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CategoryToolStripMenuItem.Click

    End Sub

    Private Sub SubCategoryToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SubCategoryToolStripMenuItem.Click

    End Sub

    Private Sub SupplierToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SupplierToolStripMenuItem.Click
        frmSupplier.lblUser.Text = lblUser.Text
        frmSupplier.Reset()
        frmSupplier.ShowDialog()
    End Sub

    Private Sub CustomerToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub SupplierToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs)

    End Sub


    Private Sub StockToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles StockToolStripMenuItem.Click

    End Sub


    Sub Reset()

    End Sub
    Private Function HandleRegistry() As Boolean
        Dim firstRunDate As Date
        Dim st As Date = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\InventorySoft2", "Set", Nothing)
        firstRunDate = st
        If firstRunDate = Nothing Then
            firstRunDate = System.DateTime.Today.Date
            My.Computer.Registry.SetValue("HKEY_LOCAL_MACHINE\SOFTWARE\InventorySoft2", "Set", firstRunDate)
        ElseIf (Now - firstRunDate).Days > 7 Then
            Return False
        End If
        Return True
    End Function
    Private Sub frmMainMenu_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'Dim result As Boolean = HandleRegistry()
        'If result = False Then 'something went wrong
        'MessageBox.Show("Trial expired" & vbCrLf & "for purchasing the full version of software call us at +919827858191", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
        'End
        'End If
        If lblUserType.Text = "Admin" Then
            MasterEntryToolStripMenuItem.Enabled = True
            RegistrationToolStripMenuItem.Enabled = True
            LogsToolStripMenuItem.Enabled = True
            DatabaseToolStripMenuItem.Enabled = True
            CustomerToolStripMenuItem.Enabled = True
            SupplierToolStripMenuItem.Enabled = True
            ProductToolStripMenuItem.Enabled = True
            StockToolStripMenuItem.Enabled = True
            ServiceToolStripMenuItem.Enabled = True
            StockInToolStripMenuItem.Enabled = True
            BillingToolStripMenuItem.Enabled = True
            QuotationToolStripMenuItem.Enabled = True

            VoucherToolStripMenuItem.Enabled = True
            SalesmanToolStripMenuItem3.Enabled = True
            SendSMSToolStripMenuItem.Enabled = True
        End If
        If lblUserType.Text = "Sales Person" Then
            MasterEntryToolStripMenuItem.Enabled = False
            RegistrationToolStripMenuItem.Enabled = False
            LogsToolStripMenuItem.Enabled = False
            DatabaseToolStripMenuItem.Enabled = False
            CustomerToolStripMenuItem.Enabled = True
            SupplierToolStripMenuItem.Enabled = False
            ProductToolStripMenuItem.Enabled = False
            StockToolStripMenuItem.Enabled = False
            ServiceToolStripMenuItem.Enabled = True
            StockInToolStripMenuItem.Enabled = True
            BillingToolStripMenuItem.Enabled = True
            QuotationToolStripMenuItem.Enabled = True
            'RecordToolStripMenuItem.Enabled = False
            'ReportsToolStripMenuItem.Enabled = False
            VoucherToolStripMenuItem.Enabled = False
            SalesmanToolStripMenuItem3.Enabled = False
            SendSMSToolStripMenuItem.Enabled = False
        End If
        If lblUserType.Text = "Inventory Manager" Then
            MasterEntryToolStripMenuItem.Enabled = False
            RegistrationToolStripMenuItem.Enabled = False
            LogsToolStripMenuItem.Enabled = False
            DatabaseToolStripMenuItem.Enabled = False
            CustomerToolStripMenuItem.Enabled = False
            SupplierToolStripMenuItem.Enabled = False
            ProductToolStripMenuItem.Enabled = True
            StockToolStripMenuItem.Enabled = True
            ServiceToolStripMenuItem.Enabled = False
            StockInToolStripMenuItem.Enabled = True
            BillingToolStripMenuItem.Enabled = False
            QuotationToolStripMenuItem.Enabled = False
            'RecordToolStripMenuItem.Enabled = False
            'ReportsToolStripMenuItem.Enabled = False
            VoucherToolStripMenuItem.Enabled = False
            SalesmanToolStripMenuItem3.Enabled = False
            SendSMSToolStripMenuItem.Enabled = False
        End If
        ' Getdata()
        'DataGridView1.ClearSelection()
        'DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        'DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        'DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        'DataGridView1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        'DataGridView1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
    End Sub

    Private Sub StockToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub StockInToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles StockInToolStripMenuItem.Click
        frmServices.lblUser.Text = lblUser.Text
        frmServices.lblUserType.Text = lblUserType.Text
        frmServices.Reset()
        frmServices.ShowDialog()
    End Sub

    Private Sub btnExportExcel_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub txtProductName_TextChanged(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub btnReset_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub ContactsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ContactsToolStripMenuItem.Click

    End Sub

    Private Sub IndividualToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub ProductToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ProductToolStripMenuItem.Click
        frmProductEntry.lblUser.Text = lblUser.Text
        frmProductEntry.lblUserType.Text = lblUserType.Text
        frmProductEntry.Reset()
        frmProductEntry.ShowDialog()
    End Sub

    Private Sub ProductToolStripMenuItem2_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub ServiceToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ServiceToolStripMenuItem.Click
        frmStockStatus.Reset()
        frmStockStatus.ShowDialog()
    End Sub

    Private Sub ServiceToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs)

    End Sub


    Private Sub QuotationToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub ProductsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ProductsToolStripMenuItem.Click
        frmSales.lblUser.Text = lblUser.Text
        frmSales.lblUserType.Text = lblUserType.Text
        frmSales.Reset()
        frmSales.ShowDialog()
    End Sub

    Private Sub ProductsRepairToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ProductsRepairToolStripMenuItem.Click

    End Sub

    Private Sub BillingProductsServiceToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub SMSSettingToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub SalesToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub ServiceBillingToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub StockInAndStockOutToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub PurchaseToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub ProfitAndLossToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub VoucherToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles VoucherToolStripMenuItem.Click
        frmVouchersEntry.Reset()
        frmVouchersEntry.lblUser.Text = lblUser.Text
        frmVouchersEntry.ShowDialog()
    End Sub

    Private Sub ExpenditureToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub CreditorsAndDebtorsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub OverallToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub SQLServerSettingToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub PurchaseDaybookToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)
        frmPurchaseDaybook.Reset()
        frmPurchaseDaybook.ShowDialog()
    End Sub

    Private Sub GeneralLedgerToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub GeneralDaybookToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub PaymentToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PaymentToolStripMenuItem.Click

    End Sub

    Private Sub PaymentsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub TrialBalanceToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub SupplierLedgerToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub CustomerLedgerToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub SMSToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub


    Private Sub SalesmanToolStripMenuItem2_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub SalesmanToolStripMenuItem3_Click(sender As System.Object, e As System.EventArgs) Handles SalesmanToolStripMenuItem3.Click

    End Sub

    Private Sub SalesmanLedgerToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub SalesmanCommissionToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub TaxToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub SendSMSToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SendSMSToolStripMenuItem.Click

    End Sub

    Private Sub CreditTermsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub CreditTermsStatementsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Try
            With OpenFileDialog1
                .Filter = ("DB Backup File|*.bak;")
                .FilterIndex = 4
            End With
            'Clear the file name
            OpenFileDialog1.FileName = ""

            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                Cursor = Cursors.WaitCursor
                Timer2.Enabled = True
                OdbcConnection.ReleaseObjectPool()
                con = New OdbcConnection(cs)
                con.Open()
                Dim cb As String = "USE Master ALTER DATABASE SIS_DB SET Single_User WITH Rollback Immediate Restore database SIS_DB FROM disk='" & OpenFileDialog1.FileName & "' WITH REPLACE ALTER DATABASE SIS_DB SET Multi_User "
                cmd = New OdbcCommand(cb)
                cmd.Connection = con
                cmd.ExecuteReader()
                con.Close()
                Dim st As String = "Sucessfully performed the restore"
                LogFunc(lblUser.Text, st)
                MessageBox.Show("Successfully performed", "Database Restore", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DatabaseToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DatabaseToolStripMenuItem.Click
        Backup()
    End Sub

    Private Sub SMSToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles SMSToolStripMenuItem1.Click
        frmSMSSetting.Reset()
        frmSMSSetting.ShowDialog()
    End Sub

    Private Sub SQLServerToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SQLServerToolStripMenuItem.Click
        
    End Sub

    Private Sub CustomersToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CustomersToolStripMenuItem.Click

    End Sub

    Private Sub SalesmanToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub SuppliersToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SuppliersToolStripMenuItem.Click
     
    End Sub

    Private Sub SalesmanToolStripMenuItem4_Click(sender As System.Object, e As System.EventArgs) Handles SalesmanToolStripMenuItem4.Click
    
    End Sub

   

    Private Sub PurchaseReportToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PurchaseReportToolStripMenuItem.Click
        frmPurchaseReport.Reset()
        frmPurchaseReport.ShowDialog()
    End Sub

    Private Sub ExpenditureReportToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExpenditureReportToolStripMenuItem.Click
        frmVoucherReport.Reset()
        frmVoucherReport.ShowDialog()
    End Sub

  

    Private Sub ProfitAndLossReportsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ProfitAndLossReportsToolStripMenuItem.Click
        frmProfitAndLossReport.Reset()
        frmProfitAndLossReport.ShowDialog()
    End Sub

    Private Sub DayBookToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

   

    Private Sub DayBookToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles DayBookToolStripMenuItem1.Click

    End Sub

    Private Sub PurchaseDaybookToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles PurchaseDaybookToolStripMenuItem1.Click
        frmPurchaseDaybook.Reset()
        frmPurchaseDaybook.ShowDialog()
    End Sub

    Private Sub GeneralDaybookToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles GeneralDaybookToolStripMenuItem1.Click
      
    End Sub

    Private Sub GeneralLedgerToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles GeneralLedgerToolStripMenuItem1.Click
   
    End Sub

    Private Sub SalesmanLedgerToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles SalesmanLedgerToolStripMenuItem1.Click
        frmSalesmanLedger.Reset()
        frmSalesmanLedger.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem16_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem16.Click
    End Sub

    Private Sub ToolStripMenuItem16_Click_1(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem16.Click
        frmTrialBalanceAccounting.Reset()
        frmTrialBalanceAccounting.ShowDialog()
    End Sub

    Private Sub SalesmanCommissionToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles SalesmanCommissionToolStripMenuItem1.Click
    
    End Sub

  

    Private Sub CustomerLedgerToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles CustomerLedgerToolStripMenuItem1.Click
        frmCustomerLedger.Reset()
        frmCustomerLedger.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem9_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem9.Click
        frmServices.lblUser.Text = lblUser.Text
        frmServices.lblUserType.Text = lblUserType.Text
        frmServices.Reset()
        frmServices.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem2.Click
        frmRegistration.lblUser.Text = lblUser.Text
        frmRegistration.Reset()
        frmRegistration.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem3.Click
        frmCustomer.lblUser.Text = lblUser.Text
        frmCustomer.Reset()
        frmCustomer.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem5.Click

    End Sub

    Private Sub ToolStripMenuItem6_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem6.Click
        frmProductEntry.lblUser.Text = lblUser.Text
        frmProductEntry.lblUserType.Text = lblUserType.Text
        frmProductEntry.Reset()
        frmProductEntry.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem7_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem7.Click

    End Sub

    Private Sub ToolStripMenuItem8_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem8.Click
        frmStockStatus.Reset()
        frmStockStatus.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem10_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem10.Click

    End Sub

    Private Sub BillingToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles BillingToolStripMenuItem.Click

    End Sub

    Private Sub ProductsToolStripMenuItem2_Click(sender As System.Object, e As System.EventArgs) Handles ProductsToolStripMenuItem2.Click
        frmSales.lblUser.Text = lblUser.Text
        frmSales.lblUserType.Text = lblUserType.Text
        frmSales.Reset()
        frmSales.ShowDialog()
    End Sub

    Private Sub ServiceToolStripMenuItem1_Click_1(sender As System.Object, e As System.EventArgs) Handles ServiceToolStripMenuItem1.Click

    End Sub

    Private Sub ToolStripMenuItem11_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem11.Click
        
    End Sub

    Private Sub ToolStripMenuItem17_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem17.Click

    End Sub

    Private Sub SqlServerToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles SqlServerToolStripMenuItem1.Click

    End Sub

    Private Sub SMSToolStripMenuItem_Click_1(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub ToolStripMenuItem18_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem18.Click
        frmContactMe.ShowDialog()
    End Sub

    Private Sub LogoutToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles LogoutToolStripMenuItem1.Click
        Try
            If MessageBox.Show("Do you really want to logout from application?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                If MessageBox.Show("Do you want backup database before logout?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                    Backup()
                    LogOut()
                Else
                    LogOut()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem4.Click
        frmSupplier.lblUser.Text = lblUser.Text
        frmSupplier.Reset()
        frmSupplier.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem14_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem14.Click
        frmLogs.Reset()
        frmLogs.lblUser.Text = lblUser.Text
        frmLogs.ShowDialog()
    End Sub
End Class