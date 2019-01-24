Imports System.Data.OleDb
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text


Public Class frmConcession
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"

    Sub Reset()
        txtHosteler.Text = ""
        txtUSN.Text = ""
        txtHostelerID.Text = ""
        txtHostelerName.Text = ""
        txtDueAmount.Text = ""
        txtConAmount.Text = ""
        cmbAuthBy.Text = ""
        txtRemarks.Text = ""
        dtpCondate.Text = Today
        btnSave.Enabled = True
        DataGridView2.Refresh()
    End Sub

    Private Sub frmConcession_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub

    Private Sub frmConcession_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtHosteler.Focus()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Hostelers.HostelerId as [Hosteler ID],Hostelers.HostelerName as [Hosteler Name],Hostelers.USN as [USN],DueAmount.TotalDueAmount as[Due Amount], DueAmount.AcadYear as [Academic Year] FROM Hostelers INNER JOIN DueAmount ON Hostelers.[HostelerID] = DueAmount.[HostelerID]  where TotalDueAmount > '0'  order by HostelerName", con) ' Hostelers,FeePayment where Hostelers.HostelerID=FeePayment.HostelerID group by Hostelers.HostelerID,Hostelername,Hostelname,RoomNo having sum(DuePayment >0) 
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "DueAmount")
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView2.DataSource = myDataSet.Tables("DueAmount").DefaultView
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtHosteler_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHosteler.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Hostelers.HostelerId as [Hosteler ID],Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN],DueAmount.TotalDueAmount as[Due Amount], DueAmount.AcadYear as [Academic Year] FROM Hostelers INNER JOIN DueAmount ON Hostelers.[HostelerID] = DueAmount.[HostelerID]  where TotalDueAmount > '0' and HostelerName like '%" & txtHosteler.Text & "%' order by HostelerName", con) ' Hostelers,FeePayment where Hostelers.HostelerID=FeePayment.HostelerID group by Hostelers.HostelerID,Hostelername,Hostelname,RoomNo having sum(DuePayment >0) 
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "DueAmount")
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView2.DataSource = myDataSet.Tables("DueAmount").DefaultView
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub txtUSN_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUSN.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Hostelers.HostelerId as [Hosteler ID],Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN],DueAmount.TotalDueAmount as[Due Amount], DueAmount.AcadYear as [Academic Year] FROM Hostelers INNER JOIN DueAmount ON Hostelers.[HostelerID] = DueAmount.[HostelerID]  where TotalDueAmount > '0' and Hostelers.USN like '%" & txtUSN.Text & "%' order by HostelerName", con) ' Hostelers,FeePayment where Hostelers.HostelerID=FeePayment.HostelerID group by Hostelers.HostelerID,Hostelername,Hostelname,RoomNo having sum(DuePayment >0) 
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "DueAmount")
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView2.DataSource = myDataSet.Tables("DueAmount").DefaultView
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub DataGridView2_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.RowHeaderMouseClick


        Dim dr As DataGridViewRow = DataGridView2.SelectedRows(0)
        txtHostelerID.Text = dr.Cells(0).Value.ToString
        txtHostelerName.Text = dr.Cells(1).Value.ToString
        txtDueAmount.Text = dr.Cells(3).Value.ToString
        txtAcadYear.Text = dr.Cells(4).Value.ToString
        txtConAmount.Focus()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtHostelerID.Text)) = 0 Then
            MessageBox.Show("Please Enter hosteler id", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtHostelerID.Focus()
            Exit Sub
        End If
        If Len(Trim(txtAcadYear.Text)) = 0 Then
            MessageBox.Show("Please Select Academic Year", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtAcadYear.Focus()
            Exit Sub
        End If
        If Len(Trim(txtConAmount.Text)) = 0 Then
            MessageBox.Show("Please enter Concession amount", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtConAmount.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbAuthBy.Text)) = 0 Then
            MessageBox.Show("Please Select Aothurized By ", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbAuthBy.Focus()
            Exit Sub
        End If

        Try

            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "insert into Concession(HostelerID,DueAmt,Con_Amount,DueAmtRemaining,AcadYear,Authorized_By,Con_Date,Remarks) VALUES ('" & txtHostelerID.Text & "','" & txtDueAmount.Text & "','" & txtConAmount.Text & "','" & txtDueleft.Text & "','" & txtAcadYear.Text & "','" & cmbAuthBy.Text & "',#" & dtpCondate.Text & "#, '" & txtRemarks.Text & "')"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            con.Open()
            Dim cb1 As String = "update DueAmount set TotalDueAmount='" & txtDueleft.Text & "' where HostelerID='" & txtHostelerID.Text & "' and AcadYear='" & txtAcadYear.Text & "'"
            'Dim cb1 As String = "update DueAmount set TotalDueAmount =(select serviceCharges - sum(Totalpaid) from Feepayment where HostelerID='" & cmbHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "' group by HostelerID,ServiceCharges,ExtraCharge,AcadYear,PayableCharges,Fine)  where HostelerID='" & cmbHostelerID.Text & "'and AcadYear='" & cmbAcadYear.Text & "'"
            cmd = New OleDbCommand(cb1)
            cmd.Connection = con
            cmd.ExecuteReader()
            MessageBox.Show("Successfully Saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Reset()
            con.Close()
            'MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnNewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewRecord.Click
        Reset()
    End Sub

    Private Sub btnGetdata_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetdata.Click
        frmConcessionRecord.Show()
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Len(Trim(txtHostelerID.Text)) = 0 Then
            MessageBox.Show("Please Enter hosteler id", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtHostelerID.Focus()
            Exit Sub
        End If
        If Len(Trim(txtAcadYear.Text)) = 0 Then
            MessageBox.Show("Please Select Academic Year", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtAcadYear.Focus()
            Exit Sub
        End If
        If Len(Trim(txtConAmount.Text)) = 0 Then
            MessageBox.Show("Please enter Concession amount", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtConAmount.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbAuthBy.Text)) = 0 Then
            MessageBox.Show("Please Select Aothurized By ", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbAuthBy.Focus()
            Exit Sub
        End If
        Try

            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "update Concession set HostelerID ='" & txtHostelerID.Text & "',DueAmt = '" & txtDueAmount.Text & "',Con_Amount=" & txtConAmount.Text & ",DueAmtRemaining = '" & txtDueleft.Text & "', AcadYear='" & txtAcadYear.Text & "',Con_date=#" & dtpCondate.Text & "#, Remarks = '" & txtRemarks.Text & "' where  HostelerID='" & txtHostelerID.Text & "'and AcadYear='" & txtAcadYear.Text & "'"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            con.Open()
            Dim cb1 As String = "update DueAmount set TotalDueAmount='" & txtDueleft.Text & "' where HostelerID='" & txtHostelerID.Text & "'and AcadYear='" & txtAcadYear.Text & "'"
            'Dim cb1 As String = "update DueAmount set TotalDueAmount =(select serviceCharges - sum(Totalpaid) from Feepayment where HostelerID='" & cmbHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "' group by HostelerID,ServiceCharges,ExtraCharge,AcadYear,PayableCharges,Fine)  where HostelerID='" & cmbHostelerID.Text & "'and AcadYear='" & cmbAcadYear.Text & "'"
            cmd = New OleDbCommand(cb1)
            cmd.Connection = con
            cmd.ExecuteReader()
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Reset()
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtConAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtConAmount.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

   


    Private Sub txtConAmount_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtConAmount.TextChanged
        txtDueleft.Text = Val(txtDueAmount.Text) - Val(txtConAmount.Text)
        
    End Sub


    Private Sub txtConAmount_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtConAmount.Validating
        If Val(txtConAmount.Text) > Val(txtDueAmount.Text) Then
            MessageBox.Show("Amount Entered is more than Due Amount", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            txtConAmount.Text = ""
            txtDueleft.Text = ""
            txtConAmount.Focus()
        End If
    End Sub
End Class
