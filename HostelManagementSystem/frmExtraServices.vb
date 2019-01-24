Imports System.Data.OleDb
Public Class frmExtraServices
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim scAutoComplete As New AutoCompleteStringCollection
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"

    Sub Reset()
        cmbHostelerID.Text = ""
        txtHostelerName.Text = ""
        txtBranch.Text = ""
        txtRoomNo.Text = ""
        txtHosteler.Text = ""
        btnSave.Enabled = True
        btnDelete.Enabled = False
        dtpServiceDate.Text = Today
        txtTotalCharges.Text = ""
        DataGridView1.Rows.Clear()
        DataGridView2.Refresh()
    End Sub

   
    Private Sub DataGridView1_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        AddHandler e.Control.TextChanged, AddressOf CellTextChanged



        If DataGridView1.CurrentCell.ColumnIndex = 0 AndAlso TypeOf e.Control Is TextBox Then
            With DirectCast(e.Control, TextBox)
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .AutoCompleteSource = AutoCompleteSource.CustomSource
                .AutoCompleteCustomSource = scAutoComplete
            End With
        Else
            With DirectCast(e.Control, TextBox)
                .AutoCompleteMode = Nothing
                .AutoCompleteSource = AutoCompleteSource.CustomSource
                .AutoCompleteCustomSource = Nothing
            End With
        End If
        con.Close()

    End Sub

    Private Sub CellTextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim txt As DataGridViewTextBoxEditingControl
        Dim col2 As Integer
        Dim col3 As Integer

        Try
            If DataGridView1.CurrentCell.ColumnIndex = 1 Then
                txt = DirectCast(sender, DataGridViewTextBoxEditingControl)
                If txt.Text.Length > 0 Then
                    col2 = Integer.Parse(txt.Text)
                End If
                If DataGridView1.CurrentRow.Cells("Column3").Value IsNot Nothing Then
                    col3 = Integer.Parse(DataGridView1.CurrentRow.Cells("Column3").Value)
                End If
            ElseIf DataGridView1.CurrentCell.ColumnIndex = 2 Then
                txt = DirectCast(sender, DataGridViewTextBoxEditingControl)
                If DataGridView1.CurrentRow.Cells("Column2").Value IsNot Nothing Then
                    col2 = Integer.Parse(DataGridView1.CurrentRow.Cells("Column2").Value)
                End If
                If txt.Text.Length > 0 Then
                    col3 = Integer.Parse(txt.Text)
                End If
            End If
            DataGridView1.CurrentRow.Cells("Column4").Value = col2 * col3
            Dim sum As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                sum = sum + r.Cells(3).Value
            Next
            txtTotalCharges.Text = sum
        Catch ex As Exception

        End Try
    End Sub
    Private Sub DataGridView1_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView1.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView1.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If Len(Trim(cmbHostelerID.Text)) = 0 Then
                MessageBox.Show("Please select hosteler id", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbHostelerID.Focus()
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "insert into ExtraServices(HostelerID,ServiceDate,Item, Quantity,UnitPrice, TotalPrice) VALUES (@d5,@d6,@d1,@d2,@d3,@d4)"
            cmd = New OleDbCommand(cb) '
            cmd.Connection = con
            cmd.Parameters.Add(New OleDbParameter("@d5", OleDbType.VarChar, 20))
            cmd.Parameters.Add(New OleDbParameter("@d6", OleDbType.VarChar, 30))
            cmd.Parameters.Add(New OleDbParameter("@d1", OleDbType.VarChar, 150))
            cmd.Parameters.Add(New OleDbParameter("@d2", OleDbType.Integer, 6))
            cmd.Parameters.Add(New OleDbParameter("@d3", OleDbType.Integer, 6))
            cmd.Parameters.Add(New OleDbParameter("@d4", OleDbType.Integer, 6))
            ' Prepare command for repeated execution
            cmd.Prepare()
            ' Data to be inserted
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    cmd.Parameters("@d5").Value = cmbHostelerID.Text
                    cmd.Parameters("@d6").Value = dtpServiceDate.Text
                    cmd.Parameters("@d1").Value = row.Cells(0).Value
                    cmd.Parameters("@d2").Value = row.Cells(1).Value
                    cmd.Parameters("@d3").Value = row.Cells(2).Value
                    cmd.Parameters("@d4").Value = row.Cells(3).Value
                    cmd.ExecuteNonQuery()
                End If
            Next
            MessageBox.Show("Successfully saved", "Entry", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            fillItem()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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

    Private Sub frmExtraServices_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub
   
    Private Sub frmExtraServices_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillHostelerID()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name] from Hostelers order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        fillItem()
    End Sub
    Sub fillItem()
        Try

            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            Dim cmd As New OleDbCommand("SELECT distinct item FROM extraservices", con)

            Dim ds As New DataSet
            Dim da As New OleDbDataAdapter(cmd)
            da.Fill(ds, "ExtraServices")

            Dim i As Integer
            For i = 0 To ds.Tables(0).Rows.Count - 1
                scAutoComplete.Add(ds.Tables(0).Rows(i)("item").ToString())
            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView2_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.RowHeaderMouseClick

        Try
            btnDelete.Enabled = True
            Dim dr As DataGridViewRow = DataGridView2.SelectedRows(0)
            cmbHostelerID.Text = dr.Cells(0).Value.ToString
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT HostelerName,RoomNo,hostelname FROM Hostelers WHERE HostelerID= '" & cmbHostelerID.Text & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtHostelerName.Text = rdr.GetString(0)
                txtRoomNo.Text = rdr.GetString(1)
                txtBranch.Text = rdr.GetString(2)
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

    Private Sub txtHosteler_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHosteler.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name] from Hostelers where HostelerName like '" & txtHosteler.Text & "%' order by HostelerName", con)
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
            btnDelete.Enabled = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT HostelerName,RoomNo,hostelname FROM Hostelers WHERE HostelerID= '" & cmbHostelerID.Text & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtHostelerName.Text = rdr.GetString(0)
                txtRoomNo.Text = rdr.GetString(1)
                txtBranch.Text = rdr.GetString(2)
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

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            If MessageBox.Show("Do you really want to delete the records?" & vbCrLf & "all extra services records related to" & vbCrLf & "selected hosteler will be deleted permanently" & vbCrLf & "you can not restore the records", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                DeleteRecord()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub DeleteRecord()
        Try
            Dim RowsAffected As Integer = 0
            con = New OleDbConnection(cs)
            con.Open()
            Dim cq As String = "delete from ExtraServices where HostelerID = '" & cmbHostelerID.Text & "'"
            cmd = New OleDbCommand(cq)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                fillItem()
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

  
    Private Sub DataGridView2_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView2.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView2.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView2.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub
End Class