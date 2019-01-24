Imports System.Data.OleDb
Public Class frmRoom
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"

    Sub Reset()
        txtRoomNo.Text = ""
        cmbHostelName.SelectedIndex = -1
        cmbRoomType.SelectedIndex = -1
        txtNoOfBeds.Text = ""
        btnDelete.Enabled = False
        btnUpdate_record.Enabled = False
        btnSave.Enabled = True
        cmbHostelName.Focus()
    End Sub
    Private Sub btnNewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewRecord.Click
        Reset()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
      If Len(Trim(cmbHostelName.Text)) = 0 Then
            MessageBox.Show("Please select hostel name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbHostelName.Focus()
            Exit Sub
        End If
        If Len(Trim(txtRoomNo.Text)) = 0 Then
            MessageBox.Show("Please enter room no.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtRoomNo.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbRoomType.Text)) = 0 Then
            MessageBox.Show("Please select room type", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbRoomType.Focus()
            Exit Sub
        End If
        If Len(Trim(txtNoOfBeds.Text)) = 0 Then
            MessageBox.Show("Please enter no. of beds", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtNoOfBeds.Focus()
            Exit Sub
        End If
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select RoomNo,HostelName from Room where RoomNo='" & txtRoomNo.Text & "' and HostelName='" & cmbHostelName.Text & "'"
            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Record Already Exists", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Reset()
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "insert into Room(RoomNo,HostelName,RoomType,NoOfBeds,BedsAvailable) VALUES ('" & txtRoomNo.Text & "','" & cmbHostelName.Text & "','" & cmbRoomType.Text & "'," & txtNoOfBeds.Text & "," & txtNoOfBeds.Text & ")"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Autocomplete()
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnUpdate_record_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate_record.Click
        Try
            Dim diff, add As Integer
            diff = Val(txtNoOfBeds.Text) - Val(TextBox1.Text)
            add = Val(TextBox2.Text) + diff
            If (add < 0) Then
                add = 0
            End If
            If (Val(txtNoOfBeds.Text) = 0) Then
                add = 0
            End If

            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "update CheckIn  set  RoomNo= '" & txtRoomNo.Text & "'  where   Hostelname='" & cmbHostelName.Text & "' and RoomNo ='" & txtRoomNoDummy.Text & "'"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()


            con = New OleDbConnection(cs)
            con.Open()
            Dim cb1 As String = "update CheckOut  set  RoomNo= '" & txtRoomNo.Text & "'  where   Hostelname='" & cmbHostelName.Text & "' and RoomNo ='" & txtRoomNoDummy.Text & "'"
            cmd = New OleDbCommand(cb1)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()

            con = New OleDbConnection(cs)
            con.Open()
            Dim cb2 As String = "update CheckOut_History  set  RoomNo= '" & txtRoomNo.Text & "'  where   Hostelname='" & cmbHostelName.Text & "' and RoomNo ='" & txtRoomNoDummy.Text & "'"
            cmd = New OleDbCommand(cb2)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()


            con = New OleDbConnection(cs)
            con.Open()
            Dim cb3 As String = "update Hostelers  set  RoomNo= '" & txtRoomNo.Text & "'  where   Hostelname='" & cmbHostelName.Text & "' and RoomNo ='" & txtRoomNoDummy.Text & "'"
            cmd = New OleDbCommand(cb3)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()

            con = New OleDbConnection(cs)
            con.Open()
            Dim cb4 As String = "update RoomAllotment  set  RoomNo= '" & txtRoomNo.Text & "'  where   Hostelname='" & cmbHostelName.Text & "' and RoomNo ='" & txtRoomNoDummy.Text & "'"
            cmd = New OleDbCommand(cb4)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()

            con = New OleDbConnection(cs)
            con.Open()
            Dim cb5 As String = "update Room set RoomType='" & cmbRoomType.Text & "',NoOfBeds = " & txtNoOfBeds.Text & ",BedsAvailable= " & add & " where   Hostelname='" & cmbHostelName.Text & "' and RoomNo= '" & txtRoomNo.Text & "'"
            cmd = New OleDbCommand(cb5)
            cmd.Connection = con
            cmd.ExecuteReader()
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate_record.Enabled = False
            Autocomplete()
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub DeleteRecord()
        Try
            Dim RowsAffected As Integer = 0
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select RoomNo,HostelName from Hostelers where RoomNo = '" & txtRoomNo.Text & "' and HostelName='" & cmbHostelName.Text & "'"
            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Reset()
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim cq As String = "delete from Room where RoomNo='" & txtRoomNo.Text & "' and HostelName = '" & cmbHostelName.Text & "'"
            cmd = New OleDbCommand(cq)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Autocomplete()
                Reset()
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            If MessageBox.Show("Do you really want to delete this record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                DeleteRecord()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub Autocomplete()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim cmd As New OleDbCommand("SELECT distinct RoomNo FROM Room", con)
            Dim ds As New DataSet()
            Dim da As New OleDbDataAdapter(cmd)
            da.Fill(ds, "Room")
            Dim col As New AutoCompleteStringCollection()
            Dim i As Integer = 0
            For i = 0 To ds.Tables(0).Rows.Count - 1
                col.Add(ds.Tables(0).Rows(i)("RoomNo").ToString())
            Next
            txtRoomNo.AutoCompleteSource = AutoCompleteSource.CustomSource
            txtRoomNo.AutoCompleteCustomSource = col
            txtRoomNo.AutoCompleteMode = AutoCompleteMode.Suggest
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        frmRoomRecord.cmbHostelName.Text = ""
        frmRoomRecord.fillHostelName()
        frmRoomRecord.DataGridView1.DataSource = Nothing
        frmRoomRecord.Show()
    End Sub

    Private Sub frmRoom_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
           Me.Hide()
        frmMain.Show()
    End Sub



    Private Sub frmRoom_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillHostelName()
        Autocomplete()
    End Sub
    Sub fillHostelName()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct RTRIM(HostelName) FROM hostel", CN)
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

    Private Sub txtNoOfBeds_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNoOfBeds.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub
End Class