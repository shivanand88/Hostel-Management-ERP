Imports System.Net.Mail
Imports System.Data.OleDB
Public Class frmPasswordRecovery

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim ds As New DataSet()
            Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;")
            con.Open()
            Dim cmd As New OleDbCommand("SELECT [Password] FROM Users Where username = 'admin'", con)
            Dim da As New OleDbDataAdapter(cmd)
            da.Fill(ds)
            con.Close()
            If ds.Tables(0).Rows.Count > 0 Then
                Dim Msg As New MailMessage()
                ' Sender e-mail address.
                Msg.From = New MailAddress("abcd@gmail.com")
                ' Recipient e-mail address.
                Msg.[To].Add("abcd@gmail.com")
                Msg.Subject = "Your Password Details"
                Msg.Body = "Your Password: " & Convert.ToString(ds.Tables(0).Rows(0)("Password")) & ""
                Msg.IsBodyHtml = True
                ' your remote SMTP server IP.
                Dim smtp As New SmtpClient()
                smtp.Host = "smtp.gmail.com"
                smtp.Port = 587
                smtp.Credentials = New System.Net.NetworkCredential("abcd@gmail.com", "abcd")
                smtp.EnableSsl = True
                smtp.Send(Msg)
                MessageBox.Show("Password Successfully sent " & vbCrLf & "Please check your mail", "Thank you", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Hide()
                frmlogin.Show()
                frmlogin.UserName.Text = ""
                frmlogin.Password.Text = ""
                frmlogin.UserName.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub





    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub frmPasswordRecovery_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmLogin.Show()
        frmLogin.UserName.Text = ""
        frmLogin.Password.Text = ""
        frmLogin.UserName.Focus()
    End Sub

    Private Sub frmPasswordRecovery_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
