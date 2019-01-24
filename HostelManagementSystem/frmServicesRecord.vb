Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmServicesRecord
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"

    Sub fillHostelerName()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct RTRIM(HostelerName) FROM Hostelers,Services where Hostelers.HostelerID=Services.HostelerID ", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbHostelerName.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbHostelerName.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub frmServicesRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillHostelerName()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If DataGridView1.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = DataGridView1.RowCount - 1
            colsTotal = DataGridView1.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView1.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView1.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        cmbHostelerName.Text = ""
        txtHostelerName.Text = ""
        DataGridView1.DataSource = Nothing
        GroupBox2.Visible = False
    End Sub

    Private Sub cmbHostelerName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbHostelerName.SelectedIndexChanged
        Try
            GroupBox2.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT ID as [Service Id] , Hostelers.HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],HostelName as [Hostel Name],RoomNo as [Room No],ServiceName as [Service],ServiceCharges as [Service Charges] from Services,Hostelers where Hostelers.HostelerID=Services.HostelerID and HostelerName='" & cmbHostelerName.Text & "' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Services")
            DataGridView1.DataSource = myDataSet.Tables("Services").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            Dim sum As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                sum = sum + r.Cells(6).Value
            Next
            TextBox1.Text = sum
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtHostelerName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHostelerName.TextChanged
        Try
            GroupBox2.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT ID as [Service Id] , Hostelers.HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],HostelName as [Hostel Name],RoomNo as [Room No],ServiceName as [Service],ServiceCharges as [Service Charges] from Services,Hostelers where Hostelers.HostelerID=Services.HostelerID and HostelerName like'" & txtHostelerName.Text & "%' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Services")
            DataGridView1.DataSource = myDataSet.Tables("Services").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            Dim sum As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                sum = sum + r.Cells(6).Value
            Next
            TextBox1.Text = sum
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            Me.Hide()
            frmServices.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmServices.txtServiceID.Text = dr.Cells(0).Value.ToString()
            frmServices.cmbHostelerID.Text = dr.Cells(1).Value.ToString()
            frmServices.txtHostelerName.Text = dr.Cells(2).Value.ToString()
            frmServices.txtHostelName.Text = dr.Cells(3).Value.ToString()
            frmServices.txtRoomNo.Text = dr.Cells(4).Value.ToString()
            frmServices.cmbService.Text = dr.Cells(5).Value.ToString()
            frmServices.txtServiceCharges.Text = dr.Cells(6).Value.ToString()
            frmServices.btnUpdate_record.Enabled = True
            frmServices.btnDelete.Enabled = True
            frmServices.btnSave.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
End Class