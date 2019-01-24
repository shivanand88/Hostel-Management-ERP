Imports System.Data.OleDb
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Public Class frmRoomChange
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"
    Sub Reset()
        cmbAcadYear.Text = ""
        cmbHostelerID.Text = ""
        txtHostelerName.Text = ""
        cmbHostelName.Text = ""
        cmbRoomNo.Text = ""
        btnSave.Enabled = True
        txtHosteler.Text = ""
        txtUSN.Text = ""
        txtRoomno.Text = ""
        txtHostelname.Text = ""
        txtBedsAvailable.Text = ""


    End Sub

   
    Private Sub txtHosteler_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHosteler.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name],USN as [USN] from Hostelers where HostelerName like '%" & txtHosteler.Text & "%' order by HostelerName", con)
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
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name],USN as [USN] from Hostelers where USN like '%" & txtUSN.Text & "%' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub cmbHostelerID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbHostelerID.SelectedIndexChanged
        Try

            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT HostelerName,RoomNo,hostelname,AcadYear FROM Hostelers WHERE HostelerID= '" & cmbHostelerID.Text & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then

                txtHostelerName.Text = rdr.GetString(0)
                txtRoomno.Text = rdr.GetString(1)
                txtHostelname.Text = rdr.GetString(2)
                cmbAcadYear.Text = rdr.GetString(3)
                'passing value to dummy textboxes for ccomparison
               
                cmbAcadYear.Enabled = False
                cmbHostelerID.Enabled = False
                txtHostelerName.Enabled = False
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
    Private Sub DataGridView2_RowHeaderMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.RowHeaderMouseClick

        Try

            Dim dr As DataGridViewRow = DataGridView2.SelectedRows(0)
            cmbHostelerID.Text = dr.Cells(0).Value.ToString
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT HostelerName,RoomNo,hostelname,AcadYear FROM Hostelers WHERE HostelerID= '" & cmbHostelerID.Text & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtHostelerName.Text = rdr.GetString(0)
                txtRoomno.Text = rdr.GetString(1)
                txtHostelname.Text = rdr.GetString(2)
                cmbAcadYear.Text = rdr.GetString(3)
                'passing value to dummy textboxes for ccomparison
                cmbAcadYear.Enabled = False
                cmbHostelerID.Enabled = False
                txtHostelerName.Enabled = False

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
    Private Sub DataGridView2_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView2.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView2.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView2.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub
    Sub fillHostelerID()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct RTRIM(HostelerID) FROM Hostelers", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbHostelerID.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbHostelerID.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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
    Private Sub frmRoomChange_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillHostelerID()
        fillHostelName()
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
    Private Sub frmRoomChange_closing(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Hide()
        frmMain.Show()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If Len(Trim(Val(txtBedsAvailable.Text))) = 0 Then
                MessageBox.Show("No Beds Available in selected Room", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbRoomNo.Focus()
                Exit Sub
            End If
            If Len(Trim(txtHostelerName.Text)) = 0 Then
                MessageBox.Show("Please enter hosteler name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtHostelerName.Focus()
                Exit Sub
            End If
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
            If Len(Trim(cmbAcadYear.Text)) = 0 Then
                MessageBox.Show("Please select room no.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbRoomNo.Focus()
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
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con = New OleDbConnection(cs)
            con.Open()
            If (cmbHostelName.Text <> txtHostelerName.Text And cmbRoomNo.Text = txtRoomno.Text) Then
                Dim ct As String = "update room set BedsAvailable =  BedsAvailable - 1 where HostelName= '" & cmbHostelName.Text & "' and RoomNo='" & cmbRoomNo.Text & "'"
                Dim ct1 As String = "update room set BedsAvailable =  BedsAvailable + 1 where HostelName= '" & txtHostelname.Text & "' and RoomNo='" & txtRoomno.Text & "'"
                cmd = New OleDbCommand(ct)
                Dim cmd1 As OleDbCommand
                cmd1 = New OleDbCommand(ct1)
                cmd.Connection = con
                cmd1.Connection = con
                cmd.ExecuteNonQuery()
                cmd1.ExecuteNonQuery()
                con.Close()
            End If
            con = New OleDbConnection(cs)
            con.Open()
            If (cmbHostelName.Text = txtHostelerName.Text And cmbRoomNo.Text <> txtRoomno.Text) Then
                Dim ct As String = "update room set  BedsAvailable =  BedsAvailable - 1 where HostelName= '" & cmbHostelName.Text & "' and RoomNo='" & cmbRoomNo.Text & "'"
                Dim ct1 As String = "update room set  BedsAvailable =  BedsAvailable + 1 where HostelName= '" & txtHostelname.Text & "' and RoomNo='" & txtRoomno.Text & "'"
                cmd = New OleDbCommand(ct)
                Dim cmd1 As OleDbCommand
                cmd1 = New OleDbCommand(ct1)
                cmd.Connection = con
                cmd1.Connection = con
                cmd.ExecuteNonQuery()
                cmd1.ExecuteNonQuery()
                con.Close()
            End If
            con = New OleDbConnection(cs)
            con.Open()
            If (cmbHostelName.Text <> txtHostelerName.Text And cmbRoomNo.Text <> txtRoomno.Text) Then
                Dim ct As String = "update room set BedsAvailable =  BedsAvailable - 1 where HostelName= '" & cmbHostelName.Text & "' and RoomNo='" & cmbRoomNo.Text & "'"
                Dim ct1 As String = "update room set  BedsAvailable =  BedsAvailable + 1 where HostelName= '" & txtHostelname.Text & "' and RoomNo='" & txtRoomno.Text & "'"
                cmd = New OleDbCommand(ct)
                Dim cmd1 As OleDbCommand
                cmd1 = New OleDbCommand(ct1)
                cmd.Connection = con
                cmd1.Connection = con
                cmd.ExecuteNonQuery()
                cmd1.ExecuteNonQuery()
                con.Close()
            End If

            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "update CheckIn set RoomNo='" & cmbRoomNo.Text & "',HostelName='" & cmbHostelName.Text & "' where HostelerID='" & cmbHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "'"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb1 As String = "update RoomAllotment set RoomNo='" & cmbRoomNo.Text & "',HostelName='" & cmbHostelName.Text & "' where HostelerID='" & cmbHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "'"
            cmd = New OleDbCommand(cb1)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb2 As String = "update Hostelers set HostelerName='" & txtHostelerName.Text & "',RoomNo='" & cmbRoomNo.Text & "',HostelName='" & cmbHostelName.Text & "' where HostelerID='" & cmbHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "'"
            cmd = New OleDbCommand(cb2)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            MessageBox.Show("Successfully Changed Room", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            con.Close()
            Reset()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
       
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
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
        Label2.Visible = False
        txtBedsAvailable.Visible = False
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select distinct RTRIM(RoomNo),BedsAvailable from Room where HostelName= '" + cmbHostelName.Text & "' and BedsAvailable > 0 "
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
            cmd.CommandText = "SELECT BedsAvailable FROM Room WHERE RoomNo= '" & cmbRoomNo.Text & "' and HostelName='" & cmbHostelName.Text & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtBedsAvailable.Text = rdr.GetInt32(0)
                Label2.Visible = True
                txtBedsAvailable.Visible = True

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

   
    Private Sub btnNewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewRecord.Click
        Reset()
    End Sub
End Class