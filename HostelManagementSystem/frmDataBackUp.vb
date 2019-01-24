Imports System.IO

Public Class frmDataBackUp


  

   


    Private Sub frmDataBackUp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Set the two button as disabled
        btnbackup.Enabled = False
        btnrestore.Enabled = False
    End Sub
    Private Sub frmDataBackUp_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
        
    End Sub





    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnbackup.Click

        Try
            'call the SavefileDialog box
            SFD.ShowDialog()
            'Set the title
            SFD.Title = "Save File"
            'Set a specific filter
            SFD.Filter = "(*.mdb)|*.accdb"
            If SFD.ShowDialog = Windows.Forms.DialogResult.OK Then
                'set the destination of a file
                txtDestination.Text = SFD.FileName
                Dim portfolioPath As String = My.Application.Info.DirectoryPath
                'create a backup by using Filecopy Command to copy the file from  location to destination
                FileCopy(txtLocation.Text, txtDestination.Text)
                MsgBox("Database Backup Created Successfully!")
                'Reload the form
                Call frmDataBackUp_Load(sender, e)
                txtLocation.Text = Nothing
                txtDestination.Text = "Destination..."
            End If


        Catch ex As Exception
            'catch some error that may occur
            MsgBox(ex.Message)

        End Try

    End Sub

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        Try
            'set the Title of a Openfolder Dialog Box 
            OFD.Title = "Please Select MS Access Database File"
            'Set a specific Filter  and Index
            OFD.Filter = "MS Access Database Files (*.mdb)|*.accdb"
            OFD.FilterIndex = 1
            ' when the Browse Image button is click it will open the OpenfileDialog
            'this line of will check if the dialogresult selected is cancel then 
            If OFD.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
                'Set the location
                txtLocation.Text = OFD.FileName

            End If
        Catch ex As Exception
            'catch some error that may occur
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub txtlocation_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLocation.TextChanged
        ' it Disable the two button when the location is empty else do the opposite
        If txtLocation.Text = "" Then
            btnbackup.Enabled = False
            btnrestore.Enabled = False

        Else
            btnbackup.Enabled = True
            btnrestore.Enabled = True

        End If

    End Sub

  


    Private Sub btnrestore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnrestore.Click
        Try
            'call the SavefileDialog box
            SFD.ShowDialog()
            'Set the title
            SFD.Title = "Save File"
            'Set a specific filter
            SFD.Filter = "(*.mdb)|*.accdb"
            If SFD.ShowDialog = Windows.Forms.DialogResult.OK Then
                'set the destination of a file
                txtDestination.Text = SFD.FileName
                Dim portfolioPath As String = My.Application.Info.DirectoryPath
                If MessageBox.Show("Restoring the database will erase any changes you have made since you last backup. Are you sure you want to do this?", "Confirm Delete", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.OK Then
                    'Restore the database from a backup copy. 
                    FileCopy(txtLocation.Text, txtDestination.Text)
                    MsgBox("Database Restoration Successful")


                End If

                'Reload the form
                Call frmDataBackUp_Load(sender, e)
                txtLocation.Text = Nothing
                txtDestination.Text = "Destination..."
                MsgBox("Relogin to Access New DataBase ")
                End
            End If


        Catch ex As Exception
            'catch some error that may occur
            MsgBox(ex.Message)

        End Try
    End Sub
End Class
