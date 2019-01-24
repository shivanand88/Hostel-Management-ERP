Imports System.Data.OleDb
Public Class frmHostel
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"

    Sub Reset()
        txtAddress.Text = ""
        txtContactNo.Text = ""
        txtHostelName.Text = ""
        txtManagedBy.Text = ""
        txtPhoneNo.Text = ""
        cmbHostelType.Text = ""
        btnSave.Enabled = True
        btnUpdate_record.Enabled = False
        btnDelete.Enabled = False
        txtHostelName.Focus()
    End Sub

    Private Sub btnNewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewRecord.Click
        Reset()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtHostelName.Text)) = 0 Then
            MessageBox.Show("Please enter hostel name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtHostelName.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbHostelType.Text)) = 0 Then
            MessageBox.Show("Please select Hostel Type.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbHostelType.Focus()
            Exit Sub
        End If
        If Len(Trim(txtAddress.Text)) = 0 Then
            MessageBox.Show("Please enter address", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtAddress.Focus()
            Exit Sub
        End If
        If Len(Trim(txtPhoneNo.Text)) = 0 Then
            MessageBox.Show("Please enter PhoneNo no.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtPhoneNo.Focus()
            Exit Sub
        End If
        If Len(Trim(txtManagedBy.Text)) = 0 Then
            MessageBox.Show("Please enter managed by", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtManagedBy.Focus()
            Exit Sub
        End If
        If Len(Trim(txtContactNo.Text)) = 0 Then
            MessageBox.Show("Please enter contact no.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtContactNo.Focus()
            Exit Sub
        End If

        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select HostelName from Hostel where hostelname='" & txtHostelName.Text & "'"
            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("hostel name already exists", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtHostelName.Text = ""
                txtHostelName.Focus()
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "insert into Hostel(Hostelname,Hostel_address,Hostel_Phone,ManagedBy,Hostel_contactno,HostelType) VALUES ('" & txtHostelName.Text & "','" & txtAddress.Text & "','" & txtPhoneNo.Text & "','" & txtManagedBy.Text & "','" & txtContactNo.Text & "','" & cmbHostelType.Text & "')"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            If MessageBox.Show("Do you really want to delete the record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.Yes Then
                delete_records()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub delete_records()
        Try
            Dim RowsAffected As Integer = 0
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select Hostelname from Room where Hostelname = '" & txtHostelName.Text & "'"
            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Reset()
                If Not rdr Is Nothing Then
                End If
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct1 As String = "select Hostelname from hostelers where Hostelname = '" & txtHostelName.Text & "'"
            cmd = New OleDbCommand(ct1)
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
            Dim ct2 As String = "select Hostelname from quotation where Hostelname = '" & txtHostelName.Text & "'"
            cmd = New OleDbCommand(ct2)
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
            Dim cq As String = "delete from hostel where hostelname= '" & txtHostelName.Text & "'"
            cmd = New OleDbCommand(cq)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then

                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
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

    Private Sub btnUpdate_record_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate_record.Click
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "update Hostel set HostelName='" & txtHostelName.Text & "', Hostel_address='" & txtAddress.Text & "',Hostel_Phone='" & txtPhoneNo.Text & "',ManagedBy='" & txtManagedBy.Text & "',Hostel_contactno='" & txtContactNo.Text & "', HostelType='" & cmbHostelType.Text & "' where hostelname= '" & TextBox1.Text & "'"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate_record.Enabled = False
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Reset()
        frmHostelRecord.GetData()
        frmHostelRecord.Show()
    End Sub

    Private Sub txtPhoneNo_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPhoneNo.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtContactNo_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtContactNo.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

   
End Class