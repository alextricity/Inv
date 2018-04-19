Imports System.Data.Odbc
Public Class frmSplash

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Try
            If System.IO.File.Exists(Application.StartupPath & "\SQLSettings.dat") Then
                If txtActivationID.Text = TextBox1.Text Then
                    ProgressBar1.Visible = True
                    ProgressBar1.Value = ProgressBar1.Value + 2
                    If (ProgressBar1.Value = 10) Then
                        lblSet.Text = "Reading modules.."
                    ElseIf (ProgressBar1.Value = 20) Then
                        lblSet.Text = "Turning on modules."
                    ElseIf (ProgressBar1.Value = 40) Then
                        lblSet.Text = "Starting modules.."
                    ElseIf (ProgressBar1.Value = 60) Then
                        lblSet.Text = "Loading modules.."
                    ElseIf (ProgressBar1.Value = 80) Then
                        lblSet.Text = "Done Loading modules.."
                    ElseIf (ProgressBar1.Value = 100) Then
                        frmLogin.Show()
                        Timer1.Enabled = False
                        Me.Hide()
                    End If
                End If
            Else
                ProgressBar1.Visible = True
                ProgressBar1.Value = ProgressBar1.Value + 2
                If (ProgressBar1.Value = 10) Then
                    lblSet.Text = "Reading modules.."
                ElseIf (ProgressBar1.Value = 20) Then
                    lblSet.Text = "Turning on modules."
                ElseIf (ProgressBar1.Value = 40) Then
                    lblSet.Text = "Starting modules.."
                ElseIf (ProgressBar1.Value = 60) Then
                    lblSet.Text = "Loading modules.."
                ElseIf (ProgressBar1.Value = 80) Then
                    lblSet.Text = "Done Loading modules.."
                ElseIf (ProgressBar1.Value = 100) Then

                    Timer1.Enabled = False
                    Me.Hide()
                End If
            End If
            If System.IO.File.Exists(Application.StartupPath & "\SQLSettings.dat") Then
                If txtActivationID.Text <> TextBox1.Text Then
                    ProgressBar1.Visible = True
                    ProgressBar1.Value = ProgressBar1.Value + 2
                    If (ProgressBar1.Value = 10) Then
                        lblSet.Text = "Reading modules.."
                    ElseIf (ProgressBar1.Value = 20) Then
                        lblSet.Text = "Turning on modules."
                    ElseIf (ProgressBar1.Value = 40) Then
                        lblSet.Text = "Starting modules.."
                    ElseIf (ProgressBar1.Value = 60) Then
                        lblSet.Text = "Loading modules.."
                    ElseIf (ProgressBar1.Value = 80) Then
                        lblSet.Text = "Done Loading modules.."
                    ElseIf (ProgressBar1.Value = 100) Then
                        '
                        Timer1.Enabled = False
                        Me.Hide()
                    End If
                End If
            Else
                ProgressBar1.Visible = True
                ProgressBar1.Value = ProgressBar1.Value + 2
                If (ProgressBar1.Value = 10) Then
                    lblSet.Text = "Reading modules.."
                ElseIf (ProgressBar1.Value = 20) Then
                    lblSet.Text = "Turning on modules."
                ElseIf (ProgressBar1.Value = 40) Then
                    lblSet.Text = "Starting modules.."
                ElseIf (ProgressBar1.Value = 60) Then
                    lblSet.Text = "Loading modules.."
                ElseIf (ProgressBar1.Value = 80) Then
                    lblSet.Text = "Done Loading modules.."
                ElseIf (ProgressBar1.Value = 100) Then

                    Timer1.Enabled = False
                    Me.Hide()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error!Unable to load the sqlsettings.dat file")
            End
        End Try

    End Sub

    Private Sub frmSplash1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Try
            'If System.IO.File.Exists(Application.StartupPath & "\SQLSettings.dat") Then
            ' Dim i As System.Management.ManagementObject
            ' Dim searchInfo_Processor As New System.Management.ManagementObjectSearcher("Select * from Win32_Processor")
            ' For Each i In searchInfo_Processor.Get()
            ' txtHardwareID.Text = i("ProcessorID").ToString
            ' Next
            ' Dim searchInfo_Board As New System.Management.ManagementObjectSearcher("Select * from Win32_BaseBoard")
            ' For Each i In searchInfo_Board.Get()
            ' txtSerialNo.Text = i("SerialNumber").ToString
            ' Next
            ' Dim st As String = (txtHardwareID.Text) + (txtSerialNo.Text)
            ' '   TextBox1.Text = Encryption.MakePassword(st, 659)
            ' con = New OdbcConnection(cs)
            ' con.Open()
            ' Dim ct As String = "select RTRIM(ActivationID) from Activation where HardwareID=@d1 and SerialNo=@d2"
            ' cmd = New OdbcCommand(ct)
            ' cmd.Connection = con
            ' cmd.Parameters.AddWithValue("@d1", Encrypt(txtHardwareID.Text.Trim))
            ' cmd.Parameters.AddWithValue("@d2", Encrypt(txtSerialNo.Text.Trim))
            ' rdr = cmd.ExecuteReader()
            ' If rdr.Read() Then
            ' txtActivationID.Text = Decrypt(rdr.GetValue(0))
            ' End If
            ' End If
        Catch ex As Exception
            '  MsgBox(ex.Message, MsgBoxStyle.Critical, "Error! Unable to load the sqlsettings.dat file")
            '  End
        End Try
    End Sub

End Class