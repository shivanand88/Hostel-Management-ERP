Imports System.Data.OleDb
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Public Class frmGuestCheckOut
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"
    Sub Reset()

        txtHostelerID.Text = ""
        txtHostelerName.Text = ""
        txtHostelName.Text = ""
        txtRoomNo.Text = ""
        Label2.Visible = False
        txtTotalDueAmount.Visible = False
        txtTotalDueAmount.Text = ""
        txtAcadYear.Text = ""
        txtHosteler.Text = ""
        txtUSN.Text = ""
        dtpPaymentDate.Text = Today
        btnSave.Enabled = True



    End Sub
    Private Sub btnNewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewRecord.Click
        Reset()
    End Sub

    Private Sub frmGuestCheckOut_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub
    Private Sub txtHosteler_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHosteler.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name],USN as [USN] from Hostelers where HostelerName like '" & txtHosteler.Text & "%' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub txtUSN_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUSN.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name],USN as [USN] from Hostelers where USN like '" & txtUSN.Text & "%' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub DataGridView2_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.RowHeaderMouseClick
        Label2.Visible = True
        txtTotalDueAmount.Visible = True
        Try

            Dim dr As DataGridViewRow = DataGridView2.SelectedRows(0)
            txtHostelerID.Text = dr.Cells(0).Value.ToString
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT HostelerName,RoomNo,hostelname,AcadYear FROM Hostelers WHERE HostelerID= '" & txtHostelerID.Text & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtHostelerName.Text = rdr.GetString(0)
                txtRoomNo.Text = rdr.GetString(1)
                txtHostelName.Text = rdr.GetString(2)
                txtAcadYear.Text = rdr.GetString(3)
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

    Private Sub frmGuestCheckOut_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name],USN as [USN] from Hostelers order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtHostelerID.Text)) = 0 Then
            MessageBox.Show("Please Select Hosteler", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtHostelerName.Focus()
            Exit Sub
        End If


        con = New OleDbConnection(cs)
        con.Open()
        Dim ct3 As String = "select TotalDueAmount,AcadYear from DueAmount where HostelerID = '" & txtHostelerID.Text & "' and AcadYear= '" & txtAcadYear.Text & "'"
        cmd = New OleDbCommand(ct3)
        cmd.Connection = con
        rdr = cmd.ExecuteReader()
        Try
            rdr.Read()
            Label2.Visible = True
            txtTotalDueAmount.Visible = True
            txtTotalDueAmount.Text = rdr.GetValue(0)
            rdr.Close()

        Catch ex As Exception
            If (rdr IsNot Nothing) Then
                MsgBox("Cant Check out Register first", MessageBoxButtons.OK, MessageBoxIcon.Error)
                rdr.Close()
            End If
            Exit Sub
        End Try
        con.Close()


        Try

            
            If Len(Trim(txtHostelerID.Text)) = 0 Then
                MessageBox.Show("Please retrieve hosteler id", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtHostelerID.Focus()
                Exit Sub
            End If
            If Len(Trim(txtHostelerName.Text)) = 0 Then
                MessageBox.Show("Please retrieve hosteler name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtHostelerName.Focus()
                Exit Sub
            End If
            If Len(Trim(txtHostelerName.Text)) = 0 Then
                MessageBox.Show("Please select hostel name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtHostelerName.Focus()
                Exit Sub
            End If
            If Len(Trim(txtRoomNo.Text)) = 0 Then
                MessageBox.Show("Please select room no.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtRoomNo.Focus()
                Exit Sub
            End If


            If MessageBox.Show(" " & txtTotalDueAmount.Text & " Rs Due Pending! Are you sure want to check out?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then

                ' selection of hosteler ID from Checkout table if it exists if not ID already exists

                
                con = New OleDbConnection(cs)
                con.Open()
                cmd = con.CreateCommand()
                cmd.CommandText = "SELECT NoOfBeds,bedsavailable FROM Room WHERE RoomNo= '" & txtRoomNo.Text & "' and HostelName='" & txtHostelName.Text & "'"
                rdr = cmd.ExecuteReader()
                If rdr.Read() Then
                    txtNoOfBeds.Text = rdr.GetInt32(0)
                    txtBedAvailable.Text = rdr.GetInt32(1)
                End If
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
                If (Val(txtBedAvailable.Text) > Val(txtNoOfBeds.Text)) Then
                    MessageBox.Show("Unable to allocate room", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
                con = New OleDbConnection(cs)
                con.Open()
                Dim ct As String = "update room set BedsAvailable =  BedsAvailable + 1 where HostelName= '" & txtHostelName.Text & "' and RoomNo='" & txtRoomNo.Text & "'"
                cmd = New OleDbCommand(ct)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                con.Close()


                con = New OleDbConnection(cs)
                con.Open()
                Dim cb As String = "insert into CheckOut(HostelerID,HostelerName,HostelName,RoomNo,CheckOutDate,AcadYear) VALUES('" & txtHostelerID.Text & "','" & txtHostelerName.Text & "','" & txtHostelName.Text & "','" & txtRoomNo.Text & "',#" & DateTime.Now.ToShortDateString() & "# ,'" & txtAcadYear.Text & "')"
                cmd = New OleDbCommand(cb)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                con.Close()
                con = New OleDbConnection(cs)
                con.Open()
                Dim cb1 As String = "insert into CheckOut_History(HostelerID,HostelerName,HostelName,RoomNo,CheckOutDate,AcadYear) VALUES('" & txtHostelerID.Text & "','" & txtHostelerName.Text & "','" & txtHostelName.Text & "','" & txtRoomNo.Text & "',#" & DateTime.Now.ToShortDateString() & "#,'" & txtAcadYear.Text & "')"
                cmd = New OleDbCommand(cb1)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                con.Close()
                MessageBox.Show("Successfully Check out", "Record updated", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'End If
                Reset()

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
End Class