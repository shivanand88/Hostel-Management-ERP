Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class frmConcessionRecord
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"


    Private Sub frmConcessionRecord_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmConcession.Show()
    End Sub

    Private Sub btnExport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport1.Click
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

    Private Sub btnExport2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport2.Click
        If DataGridView2.RowCount = Nothing Then
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

            rowsTotal = DataGridView2.RowCount - 1
            colsTotal = DataGridView2.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView2.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView2.Rows(I).Cells(j).Value
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

    Private Sub frmConcessionRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Concession.HostelerID as [Hosteler Id], Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], Concession.DueAmt as [Due Amount],Concession.Con_Amount as [Concession Amount],Concession.DueAmtRemaining as [Due Left], Concession.AcadYear as [Acad Year], Concession.Authorized_By as [Authorized By], Concession.Con_Date as [Date], Remarks as [Remarks]  FROM Hostelers INNER JOIN Concession ON Hostelers.[HostelerID] = Concession.[HostelerID] ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Concession")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("Concession").DefaultView
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView2.DataSource = myDataSet.Tables("Concession").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub txtHosteler_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHosteler.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Concession.HostelerID as [Hosteler Id], Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], Concession.DueAmt as [Due Amount],Concession.Con_Amount as [Concession Amount],Concession.DueAmtRemaining as [Due Left], Concession.AcadYear as [Acad Year], Concession.Authorized_By as [Authorized By], Concession.Con_Date as [Date], Remarks as [Remarks] FROM Hostelers INNER JOIN Concession ON Hostelers.[HostelerID] = Concession.[HostelerID] where Hostelers.HostelerName like '%" & txtHosteler.Text & "%' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Concession")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("Concession").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub txtUSN_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUSN.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Concession.HostelerID as [Hosteler Id], Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], Concession.DueAmt as [Due Amount],Concession.Con_Amount as [Concession Amount],Concession.DueAmtRemaining as [Due Left], Concession.AcadYear as [Acad Year], Concession.Authorized_By as [Authorized By], Concession.Con_Date as [Date], Remarks as [Remarks]  FROM Hostelers INNER JOIN Concession ON Hostelers.[HostelerID] = Concession.[HostelerID] where Hostelers.USN like '%" & txtUSN.Text & "%' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Concession")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView2.DataSource = myDataSet.Tables("Concession").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnReset1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset1.Click
        txtHosteler.Text = ""
        txtHosteler.Focus()
    End Sub

    Private Sub btnReset2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset2.Click
        txtUSN.Text = ""
        txtUSN.Focus()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Concession.HostelerID as [Hosteler Id], Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], Concession.DueAmt as [Due Amount],Concession.Con_Amount as [Concession Amount],Concession.DueAmtRemaining as [Due Left], Concession.AcadYear as [Acad Year], Concession.Authorized_By as [Authorized By], Concession.Con_Date as [Date], Remarks as [Remarks]  FROM Hostelers INNER JOIN Concession ON Hostelers.[HostelerID] = Concession.[HostelerID] ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Concession")
            DataGridView3.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView3.DataSource = myDataSet.Tables("Concession").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

   
   
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If DataGridView3.RowCount = Nothing Then
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

            rowsTotal = DataGridView3.RowCount - 1
            colsTotal = DataGridView3.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView3.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView3.Rows(I).Cells(j).Value
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

   
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        DataGridView3.DataSource = Nothing
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        DataGridView4.DataSource = Nothing
        cmbAuthBy.Text = ""

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Concession.HostelerID as [Hosteler Id], Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], Concession.DueAmt as [Due Amount],Concession.Con_Amount as [Concession Amount],Concession.DueAmtRemaining as [Due Left], Concession.AcadYear as [Acad Year], Concession.Authorized_By as [Authorized By], Concession.Con_Date as [Date], Remarks as [Remarks]  FROM Hostelers INNER JOIN Concession ON Hostelers.[HostelerID] = Concession.[HostelerID] where Concession.Authorized_By = '" & cmbAuthBy.Text & "'  ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Concession")
            DataGridView4.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView4.DataSource = myDataSet.Tables("Concession").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If DataGridView4.RowCount = Nothing Then
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

            rowsTotal = DataGridView4.RowCount - 1
            colsTotal = DataGridView4.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView4.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView4.Rows(I).Cells(j).Value
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
    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            Me.Hide()
            frmConcession.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmConcession.txtHostelerID.Text = dr.Cells(0).Value.ToString()
            frmConcession.txtHostelerName.Text = dr.Cells(1).Value.ToString()

            frmConcession.txtDueAmount.Text = dr.Cells(3).Value.ToString()
            frmConcession.txtConAmount.Text = dr.Cells(4).Value.ToString()
            frmConcession.txtDueleft.Text = dr.Cells(5).Value.ToString()
            frmConcession.txtAcadYear.Text = dr.Cells(6).Value.ToString()
            frmConcession.cmbAuthBy.Text = dr.Cells(7).Value.ToString()
            frmConcession.dtpCondate.Text = dr.Cells(8).Value.ToString()
            frmConcession.txtRemarks.Text = dr.Cells(9).Value.ToString()
            frmConcession.txtHostelerID.Enabled = False
            frmConcession.txtHostelerName.Enabled = False
            frmConcession.txtAcadYear.Enabled = False
            frmConcession.btnUpdate.Enabled = True
            frmConcession.btnSave.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView2_RowHeaderMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView2.SelectedRows(0)
            Me.Hide()
            frmConcession.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmConcession.txtHostelerID.Text = dr.Cells(0).Value.ToString()
            frmConcession.txtHostelerName.Text = dr.Cells(1).Value.ToString()

            frmConcession.txtDueAmount.Text = dr.Cells(3).Value.ToString()
            frmConcession.txtConAmount.Text = dr.Cells(4).Value.ToString()
            frmConcession.txtDueleft.Text = dr.Cells(5).Value.ToString()
            frmConcession.txtAcadYear.Text = dr.Cells(6).Value.ToString()
            frmConcession.cmbAuthBy.Text = dr.Cells(7).Value.ToString()
            frmConcession.dtpCondate.Text = dr.Cells(8).Value.ToString()
            frmConcession.txtRemarks.Text = dr.Cells(9).Value.ToString()
            frmConcession.txtHostelerID.Enabled = False
            frmConcession.txtHostelerName.Enabled = False
            frmConcession.txtAcadYear.Enabled = False
            frmConcession.btnUpdate.Enabled = True
            frmConcession.btnSave.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView3_RowHeaderMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView3.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView4.SelectedRows(0)
            Me.Hide()
            frmConcession.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmConcession.txtHostelerID.Text = dr.Cells(0).Value.ToString()
            frmConcession.txtHostelerName.Text = dr.Cells(1).Value.ToString()

            frmConcession.txtDueAmount.Text = dr.Cells(3).Value.ToString()
            frmConcession.txtConAmount.Text = dr.Cells(4).Value.ToString()
            frmConcession.txtDueleft.Text = dr.Cells(5).Value.ToString()
            frmConcession.txtAcadYear.Text = dr.Cells(6).Value.ToString()
            frmConcession.cmbAuthBy.Text = dr.Cells(7).Value.ToString()
            frmConcession.dtpCondate.Text = dr.Cells(8).Value.ToString()
            frmConcession.txtRemarks.Text = dr.Cells(9).Value.ToString()
            frmConcession.txtHostelerID.Enabled = False
            frmConcession.txtHostelerName.Enabled = False
            frmConcession.txtAcadYear.Enabled = False
            frmConcession.btnUpdate.Enabled = True
            frmConcession.btnSave.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView4_RowHeaderMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView4.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView4.SelectedRows(0)
            Me.Hide()
            frmConcession.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmConcession.txtHostelerID.Text = dr.Cells(0).Value.ToString()
            frmConcession.txtHostelerName.Text = dr.Cells(1).Value.ToString()

            frmConcession.txtDueAmount.Text = dr.Cells(3).Value.ToString()
            frmConcession.txtConAmount.Text = dr.Cells(4).Value.ToString()
            frmConcession.txtDueleft.Text = dr.Cells(5).Value.ToString()
            frmConcession.txtAcadYear.Text = dr.Cells(6).Value.ToString()
            frmConcession.cmbAuthBy.Text = dr.Cells(7).Value.ToString()
            frmConcession.dtpCondate.Text = dr.Cells(8).Value.ToString()
            frmConcession.txtRemarks.Text = dr.Cells(9).Value.ToString()
            frmConcession.txtHostelerID.Enabled = False
            frmConcession.txtHostelerName.Enabled = False
            frmConcession.txtAcadYear.Enabled = False
            frmConcession.btnUpdate.Enabled = True
            frmConcession.btnSave.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class