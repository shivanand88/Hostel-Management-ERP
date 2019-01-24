Imports System.Data.OleDB
Public Class frmLogin
    Sub reset()
        UserName.Text = ""
        Password.Text = ""

    End Sub
    Private Sub frmLogin_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        End
    End Sub

    Private Sub frmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        reset()
        UserName.Focus()
    End Sub

    Private Sub linkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkLabel1.LinkClicked
        Me.Hide()
        frmChangePassword.Show()
        frmChangePassword.UserName.Text = ""
        frmChangePassword.OldPassword.Text = ""
        frmChangePassword.NewPassword.Text = ""
        frmChangePassword.ConfirmPassword.Text = ""
        frmChangePassword.UserName.Focus()
    End Sub

    Private Sub linkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkLabel3.LinkClicked
        Me.Hide()
        frmPasswordRecovery.Show()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        End
    End Sub

    Private Sub btnOK_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If Len(Trim(UserName.Text)) = 0 Then
            MessageBox.Show("Please enter user name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            UserName.Focus()
            Exit Sub
        End If
        If Len(Trim(Password.Text)) = 0 Then
            MessageBox.Show("Please enter password", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Password.Focus()
            Exit Sub
        End If
        Try
            Dim myConnection As OleDbConnection
            myConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;")

            Dim myCommand As OleDbCommand

            myCommand = New OleDbCommand("SELECT Username,[password] FROM Users WHERE username = @username AND password = @UserPassword", myConnection)

            Dim uName As New OleDbParameter("@username", SqlDbType.VarChar)

            Dim uPassword As New OleDbParameter("@UserPassword", SqlDbType.VarChar)

            uName.Value = UserName.Text

            uPassword.Value = Password.Text

            myCommand.Parameters.Add(uName)

            myCommand.Parameters.Add(uPassword)

            myCommand.Connection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)

            Dim Login As Object = 0

            If myReader.HasRows Then

                myReader.Read()

                Login = myReader(Login)

            End If

            If Login = Nothing Then

                MsgBox("Login is Failed...Try again !", MsgBoxStyle.Critical, "Login Denied")
                UserName.Clear()
                Password.Clear()
                UserName.Focus()

            Else

                ProgressBar1.Visible = True
                ProgressBar1.Maximum = 25000
                ProgressBar1.Minimum = 0
                ProgressBar1.Value = 4
                ProgressBar1.Step = 1

                For i = 0 To 25000
                    ProgressBar1.PerformStep()
                Next
                Me.Hide()
                frmMain.lblUser.Text = UserName.Text
                frmMain.Show()
            End If
            myCommand.Dispose()
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
