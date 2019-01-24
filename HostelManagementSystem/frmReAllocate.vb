Imports System.Data.OleDb
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

Public Class frmReAllocate
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"
    Sub GetHosteler()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name],USN as [USN] ,AcadYear as [Academic Year] from Hostelers where Status = 'DEACTIVE' and AcadYear = '" & cmbAcadYear1.Text & "' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub Reset()
        txtHosteler.Text = ""
        txtUSN.Text = ""
        txtHostelerID.Text = ""
        txtHostelerName.Text = ""
        txtHostelName.Text = ""
        txtRoomNo.Text = ""
        txtAcadYearOld.Text = ""
        cmbHostelName.Text = ""
        cmbRoomNo.Text = ""
        cmbAcadYear.Text = ""
        txtRoomType.Text = ""
        dtpDateOfJoining.Text = Today
        btnReAllot.Enabled = True
    End Sub
    Private Sub frmReAllocate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillHostelName()
        '  GetHosteler()
    End Sub
    Private Sub cmbAcadYear1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAcadYear1.SelectedIndexChanged
        GetHosteler()
    End Sub
    Sub fillHostelName()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct RTRIM(HostelName) FROM Room", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbHostelName.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbHostelName.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub frmReAllocate_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub


    Private Sub txtHosteler_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHosteler.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name],USN as [USN] from Hostelers where Status = 'DEACTIVE' and AcadYear = '" & cmbAcadYear1.Text & "' and  HostelerName like '%" & txtHosteler.Text & "%' order by HostelerName", con)
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
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name],USN as [USN] from Hostelers where Status = 'DEACTIVE' and AcadYear = '" & cmbAcadYear1.Text & "' and USN like '%" & txtUSN.Text & "%' order by HostelerName", con)
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
        Try

            Dim dr As DataGridViewRow = DataGridView2.SelectedRows(0)
            txtHostelerID.Text = dr.Cells(0).Value.ToString
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT HostelerName,HostelName,RoomNo,AcadYear FROM Hostelers WHERE HostelerID= '" & txtHostelerID.Text & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtHostelerName.Text = rdr.GetString(0)
                txtHostelName.Text = rdr.GetString(1)
                txtRoomNo.Text = rdr.GetString(2)
                txtAcadYearOld.Text = rdr.GetString(3)
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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If Len(Trim(cmbHostelName.Text)) = 0 Then
                MessageBox.Show("Please select hostel name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbHostelName.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbRoomNo.Text)) = 0 Then
                MessageBox.Show("Please select room no.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbRoomNo.Focus()
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT NoOfBeds FROM Room WHERE RoomNo= '" & cmbRoomNo.Text & "' and HostelName='" & cmbHostelName.Text & "' and BedsAvailable > 0"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                MessageBox.Show("Bed Available in selected room no.", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Bed not Available in selected room no.", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
    Private Sub cmbHostelName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbHostelName.SelectedIndexChanged
        cmbRoomNo.Items.Clear()
        cmbRoomNo.Text = ""
        cmbRoomNo.Enabled = True
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select distinct RTRIM(RoomNo), BedsAvailable from Room where HostelName= '" & cmbHostelName.Text & "' and BedsAvailable > 0  "
            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()
            While rdr.Read()
                cmbRoomNo.Items.Add(rdr(0))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub cmbRoomNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbRoomNo.SelectedIndexChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT RoomType FROM Room WHERE RoomNo= '" & cmbRoomNo.Text & "' and HostelName='" & cmbHostelName.Text & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtRoomType.Text = rdr.GetString(0)
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

   
    Private Sub btnReAllot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReAllot.Click
        Try

            If Len(Trim(txtHostelerName.Text)) = 0 Then
                MessageBox.Show("Please enter hosteler name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtHostelerName.Focus()
                Exit Sub
            End If
            If Len(Trim(txtHostelerName.Text)) = 0 Then
                MessageBox.Show("Please select hostel name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtHostelerName.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbRoomNo.Text)) = 0 Then
                MessageBox.Show("Please select room no.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbRoomNo.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbAcadYear.Text)) = 0 Then
                MessageBox.Show("Please select room no.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbAcadYear.Focus()
                Exit Sub
            End If
           
            If MessageBox.Show("Are you sure want to re-allocate?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                con = New OleDbConnection(cs)
                con.Open()
                cmd = con.CreateCommand()
                cmd.CommandText = "SELECT hostelerID,AcadYear FROM CheckIn where HostelerID='" & txtHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "'"
                rdr = cmd.ExecuteReader()
                If rdr.Read() Then
                    MessageBox.Show("Hosteler is already check In from " & cmbAcadYear.Text & "Check Out First Or Select Next Academic Year", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    cmbAcadYear.Text = ""
                    cmbAcadYear.Focus()
                    rdr.Close()
                    Exit Sub
                End If

                con = New OleDbConnection(cs)
                con.Open()
                cmd = con.CreateCommand()
                cmd.CommandText = "SELECT hostelerID,AcadYear FROM Checkout where HostelerID='" & txtHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "'"
                rdr = cmd.ExecuteReader()
                If rdr.Read() Then
                    MessageBox.Show("Hosteler is already check out from " & cmbAcadYear.Text & "Please select Next academic year", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    cmbAcadYear.Focus()
                    Exit Sub
                    If (rdr IsNot Nothing) Then
                        MessageBox.Show("No records Found", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        rdr.Close()
                    End If
                    Exit Sub
                End If

                con = New OleDbConnection(cs)
                con.Open()
                cmd = con.CreateCommand()
                cmd.CommandText = "SELECT BedsAvailable FROM Room WHERE RoomNo= '" & cmbRoomNo.Text & "' and HostelName='" & cmbHostelName.Text & "' and BedsAvailable <= 0"
                rdr = cmd.ExecuteReader()
                If rdr.Read() Then
                    MessageBox.Show("Bed not Available in selected room no.", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    cmbRoomNo.Focus()
                    If (rdr IsNot Nothing) Then
                        rdr.Close()
                    End If
                    Exit Sub
                End If
                con = New OleDbConnection(cs)
                con.Open()
                Dim cb As String = "update Hostelers set RoomNo='" & cmbRoomNo.Text & "',HostelName='" & cmbHostelName.Text & "',AcadYear='" & cmbAcadYear.Text & "' where HostelerID='" & txtHostelerID.Text & "'"
                cmd = New OleDbCommand(cb)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                con.Close()
                con = New OleDbConnection(cs)
                con.Open()
                Dim ct1 As String = "insert into CheckIn (HostelerID,HostelerName,HostelName,RoomNo,AcadYear,CheckInDate) values('" & txtHostelerID.Text & "','" & txtHostelerName.Text & "','" & cmbHostelName.Text & "','" & cmbRoomNo.Text & "','" & cmbAcadYear.Text & "',#" & dtpDateOfJoining.Text & "#)"
                cmd = New OleDbCommand(ct1)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                con.Close()

                con = New OleDbConnection(cs)
                con.Open()
                Dim ct As String = "update room set BedsAvailable = BedsAvailable - 1 where HostelName= '" & cmbHostelName.Text & "' and RoomNo='" & cmbRoomNo.Text & "'"
                cmd = New OleDbCommand(ct)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                MessageBox.Show("Successfully re-allocated", "Hosteler", MessageBoxButtons.OK, MessageBoxIcon.Information)
                con.Close()
               
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Reset()
        GetHosteler()
    End Sub

    Private Sub btnNewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewRecord.Click
        Reset()
    End Sub

   
  
End Class