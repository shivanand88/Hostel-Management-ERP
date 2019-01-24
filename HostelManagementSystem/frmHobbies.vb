Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class frmHobbies
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"
    Sub Reset()
        txtUSN.Text = ""
        cmbHostelerID.Text = ""
        txtHostelerName.Text = ""
        txtHobby1.Text = ""
        txtHobby2.Text = ""
        txtHobby3.Text = ""
        txtHobby4.Text = ""
        txtHobby5.Text = ""
        txtHobby6.Text = ""
        txtHobby7.Text = ""
        txtHobby8.Text = ""
        txtHosteler.Text = ""
    End Sub

   
    Private Sub cmbHostelerID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbHostelerID.SelectedIndexChanged
        Try

            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT HostelerName FROM Hostelers WHERE HostelerID= '" & cmbHostelerID.Text & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtHostelerName.Text = rdr.GetString(0)
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
    Private Sub DataGridView2_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.RowHeaderMouseClick
        Try

            Dim dr As DataGridViewRow = DataGridView2.SelectedRows(0)
            cmbHostelerID.Text = dr.Cells(0).Value.ToString
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT HostelerName FROM Hostelers WHERE HostelerID= '" & cmbHostelerID.Text & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtHostelerName.Text = rdr.GetString(0)

            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            txtHobby1.Focus()
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

    Private Sub frmHobbies_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub

    Private Sub frmHobbies_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillHostelerID()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name],USN as [USN]  from Hostelers where Status='ACTIVE' order by HostelerName", con)
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
        If Len(Trim(cmbHostelerID.Text)) = 0 Then
            MessageBox.Show("Please select hosteler id", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbHostelerID.Focus()
            Exit Sub
        End If
        If Len(Trim(txtHostelerName.Text)) = 0 Then
            MessageBox.Show("Please select Hosteler Name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtHostelerName.Focus()
            Exit Sub
        End If
        If Len(Trim(txtHobby1.Text)) = 0 Then
            MessageBox.Show("Atleast enter one hobby", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtHobby1.Focus()
            Exit Sub
        End If

        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select HostelerId from Hobbies where HostelerId='" & cmbHostelerID.Text & "'"
            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Hosteler Name already exists", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Reset()
                cmbHostelerID.Focus()
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "insert into Hobbies(HostelerId,Hobby1,Hobby2,Hobby3,Hobby4,Hobby5,Hobby6,Hobby7,Hobby8,HobbyList) VALUES ('" & cmbHostelerID.Text & "','" & txtHobby1.Text & "','" & txtHobby2.Text & "','" & txtHobby3.Text & "','" & txtHobby4.Text & "','" & txtHobby5.Text & "','" & txtHobby6.Text & "','" & txtHobby7.Text & "','" & txtHobby8.Text & "','" & txtHobbyList.Text & "')"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            MessageBox.Show("Hobbies Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)

            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Reset()
    End Sub

    Private Sub btnGetHobby_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetHobby.Click
        frmHobbiesRecord.Show()

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
            Dim cq As String = "delete from Hobbies where HostelerID= '" & cmbHostelerID.Text & "' "
            cmd = New OleDbCommand(cq)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()

            MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Reset()
            con.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Len(Trim(cmbHostelerID.Text)) = 0 Then
            MessageBox.Show("Please select hosteler id", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbHostelerID.Focus()
            Exit Sub
        End If
        If Len(Trim(txtHostelerName.Text)) = 0 Then
            MessageBox.Show("Please select Hosteler Name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtHostelerName.Focus()
            Exit Sub
        End If
        If Len(Trim(txtHobby1.Text)) = 0 Then
            MessageBox.Show("Atleast Enter one hobby", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtHobby1.Focus()
            Exit Sub
        End If
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "update Hobbies  set Hobby1='" & txtHobby1.Text & "',Hobby2='" & txtHobby2.Text & "',Hobby3='" & txtHobby3.Text & "',Hobby4='" & txtHobby4.Text & "',Hobby5='" & txtHobby5.Text & "',Hobby6='" & txtHobby6.Text & "',Hobby7='" & txtHobby7.Text & "',Hobby8='" & txtHobby8.Text & "' ,HobbyList='" & txtHobbyList.Text & "' where  HostelerID='" & cmbHostelerID.Text & "'"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)

            con.Close()
            btnUpdate.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub txtUSN_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUSN.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name],USN as [USN], Status as [Status] from Hostelers where USN like '%" & txtUSN.Text & "%' and Status = 'ACTIVE' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub txtHosteler_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHosteler.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name],USN as [USN], Status as [Status] from Hostelers where HostelerName like '%" & txtHosteler.Text & "%' and Status = 'ACTIVE' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

   
    
    Private Sub txtHobby1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHobby1.TextChanged
        txtHobbyList.Text = txtHobby1.Text & "     " & txtHobby2.Text & "     " & txtHobby3.Text & "     " & txtHobby4.Text & "     " & txtHobby5.Text & "     " & txtHobby6.Text & "     " & txtHobby7.Text & "     " & txtHobby8.Text
    End Sub
    Private Sub txtHobby2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHobby2.TextChanged
        txtHobbyList.Text = txtHobby1.Text & "     " & txtHobby2.Text & "     " & txtHobby3.Text & "     " & txtHobby4.Text & "     " & txtHobby5.Text & "     " & txtHobby6.Text & "     " & txtHobby7.Text & "     " & txtHobby8.Text
    End Sub
    Private Sub txtHobby3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHobby3.TextChanged
        txtHobbyList.Text = txtHobby1.Text & "     " & txtHobby2.Text & "     " & txtHobby3.Text & "     " & txtHobby4.Text & "     " & txtHobby5.Text & "     " & txtHobby6.Text & "     " & txtHobby7.Text & "     " & txtHobby8.Text
    End Sub
    Private Sub txtHobby4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHobby4.TextChanged
        txtHobbyList.Text = txtHobby1.Text & "     " & txtHobby2.Text & "     " & txtHobby3.Text & "     " & txtHobby4.Text & "     " & txtHobby5.Text & "     " & txtHobby6.Text & "     " & txtHobby7.Text & "     " & txtHobby8.Text
    End Sub
    Private Sub txtHobby5_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHobby5.TextChanged
        txtHobbyList.Text = txtHobby1.Text & "     " & txtHobby2.Text & "     " & txtHobby3.Text & "     " & txtHobby4.Text & "     " & txtHobby5.Text & "     " & txtHobby6.Text & "     " & txtHobby7.Text & "     " & txtHobby8.Text
    End Sub
    Private Sub txtHobby6_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHobby6.TextChanged
        txtHobbyList.Text = txtHobby1.Text & "     " & txtHobby2.Text & "     " & txtHobby3.Text & "     " & txtHobby4.Text & "     " & txtHobby5.Text & "     " & txtHobby6.Text & "     " & txtHobby7.Text & "     " & txtHobby8.Text
    End Sub
    Private Sub txtHobby7_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHobby7.TextChanged

        txtHobbyList.Text = txtHobby1.Text & "     " & txtHobby2.Text & "     " & txtHobby3.Text & "     " & txtHobby4.Text & "     " & txtHobby5.Text & "     " & txtHobby6.Text & "     " & txtHobby7.Text & "     " & txtHobby8.Text
    End Sub
    Private Sub txtHobby8_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHobby8.TextChanged
        txtHobbyList.Text = txtHobby1.Text & "     " & txtHobby2.Text & "     " & txtHobby3.Text & "     " & txtHobby4.Text & "     " & txtHobby5.Text & "     " & txtHobby6.Text & "     " & txtHobby7.Text & "     " & txtHobby8.Text
    End Sub
End Class