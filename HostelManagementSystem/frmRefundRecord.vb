Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class frmRefundRecord
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"

    Private Sub frmRefundRecord_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        txtHosteler.Text = ""
        txtUSN.Text = ""
        txtHosteler.Focus()
    End Sub

   

    Private Sub txtHosteler_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHosteler.TextChanged
        Try
            GroupBox8.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Refund.HostelerID as  [Hosteler ID], Hostelers.HostelerName as  [Hosteler Name],Hostelers.USN as  [USN],Refund.RefundableAmt as [Refundable Amount],Refund.AcadYear as [Acad Year],Refund.RefundAmt as [Refund Amount],Refund.ChequeNo as [Checque No],Refund.RefundDate as [Refund Date], Refund.Remarks as [Remarks] from Hostelers,Refund where Refund.HostelerID=Hostelers.HostelerID and Hostelers.HostelerName like '%" & txtHosteler.Text & "%' order by Hostelers.HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Refund")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("Refund").DefaultView
            Dim sum As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                sum = sum + r.Cells(5).Value
            Next
            TextBox3.Text = sum
            Label5.Text = frmDueCharges.GetInWords(TextBox3.Text.Trim)
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub txtUSN_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUSN.TextChanged
        Try
            GroupBox8.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Refund.HostelerID as  [Hosteler ID], Hostelers.HostelerName as  [Hosteler Name],Hostelers.USN as  [USN],Refund.RefundableAmt as [Refundable Amount],Refund.AcadYear as [Acad Year],Refund.RefundAmt as [Refund Amount],Refund.ChequeNo as [Checque No],Refund.RefundDate as [Refund Date], Refund.Remarks as [Remarks] from Hostelers,Refund where Refund.HostelerID=Hostelers.HostelerID and Hostelers.USN like '%" & txtUSN.Text & "%' order by Hostelers.HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Refund")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("Refund").DefaultView
            Dim sum As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                sum = sum + r.Cells(5).Value
            Next
            TextBox3.Text = sum
            Label5.Text = frmDueCharges.GetInWords(TextBox3.Text.Trim)
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmRefundRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Refresh()
        DataGridView1.Refresh()
        DataGridView2.Refresh()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Refund.HostelerID as  [Hosteler ID], Hostelers.HostelerName as  [Hosteler Name],Hostelers.USN as  [USN],Refund.RefundableAmt as [Refundable Amount],Refund.AcadYear as [Acad Year],Refund.RefundAmt as [Refund Amount],Refund.ChequeNo as [Checque No],Refund.RefundDate as [Refund Date], Refund.Remarks as [Remarks] from Hostelers,Refund where Refund.HostelerID=Hostelers.HostelerID and Refund.HostelerID", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Refund")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("Refund").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    
    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
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

    Private Sub btngetData2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btngetData2.Click
        Try
            DataGridView2.Visible = True
            GroupBox7.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Refund.HostelerID as  [Hosteler ID], Hostelers.HostelerName as  [Hosteler Name],Hostelers.USN as  [USN],Refund.RefundableAmt as [Refundable Amount],Refund.AcadYear as [Acad Year],Refund.RefundAmt as [Refund Amount],Refund.ChequeNo as [Checque No],Refund.RefundDate as [Refund Date], Refund.Remarks as [Remarks] from Hostelers,Refund where Refund.HostelerID=Hostelers.HostelerID and Refund.HostelerID", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Refund")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView2.DataSource = myDataSet.Tables("Refund").DefaultView
            Dim sum As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView2.Rows
                sum = sum + r.Cells(5).Value
            Next
            TextBox2.Text = sum
            Label4.Text = frmDueCharges.GetInWords(TextBox2.Text.Trim)
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnNew2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew2.Click
        DataGridView2.DataSource = Nothing
    End Sub
    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            Me.Hide()
            frmRefund.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();

            frmRefund.cmbHostelerID.Text = dr.Cells(0).Value.ToString()
            frmRefund.txthostelerName.Text = dr.Cells(1).Value.ToString()
            frmRefund.txtRefundableAmount.Text = dr.Cells(3).Value.ToString()
            frmRefund.cmbAcadYear.Text = dr.Cells(4).Value.ToString()
            frmRefund.txtRefundAmt.Text = dr.Cells(5).Value.ToString()
            frmRefund.txtChequeNo.Text = dr.Cells(6).Value.ToString()
            frmRefund.dtpRefundDate.Text = dr.Cells(7).Value.ToString()
            frmRefund.txtRemarks.Text = dr.Cells(8).Value.ToString()
            frmRefund.cmbHostelerID.Enabled = False
            frmRefund.txthostelerName.Enabled = False
            frmRefund.cmbAcadYear.Enabled = False
            frmRefund.btnUpdate_record.Enabled = True
            frmRefund.btnDelete.Enabled = True
            frmRefund.btnSave.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub DataGridView3_RowHeaderMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView3.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView3.SelectedRows(0)
            Me.Hide()
            frmRefund.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();

            frmRefund.cmbHostelerID.Text = dr.Cells(0).Value.ToString()
            frmRefund.txthostelerName.Text = dr.Cells(1).Value.ToString()
            frmRefund.txtRefundableAmount.Text = dr.Cells(3).Value.ToString()
            frmRefund.cmbAcadYear.Text = dr.Cells(4).Value.ToString()
            frmRefund.txtRefundAmt.Text = dr.Cells(5).Value.ToString()
            frmRefund.txtChequeNo.Text = dr.Cells(6).Value.ToString()
            frmRefund.dtpRefundDate.Text = dr.Cells(7).Value.ToString()
            frmRefund.txtRemarks.Text = dr.Cells(8).Value.ToString()
            frmRefund.cmbHostelerID.Enabled = False
            frmRefund.txthostelerName.Enabled = False
            frmRefund.cmbAcadYear.Enabled = False
            frmRefund.btnUpdate_record.Enabled = True
            frmRefund.btnDelete.Enabled = True
            frmRefund.btnSave.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Len(Trim(cmbAcadYear.Text)) = 0 Then
            MessageBox.Show("Please select Academic Year", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbAcadYear.Focus()
            Exit Sub
        End If
        Try
            DataGridView3.Visible = True
            GroupBox6.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Refund.HostelerID as  [Hosteler ID], Hostelers.HostelerName as  [Hosteler Name],Hostelers.USN as  [USN],Refund.RefundableAmt as [Refundable Amount],Refund.AcadYear as [Acad Year],Refund.RefundAmt as [Refund Amount],Refund.ChequeNo as [Checque No],Refund.RefundDate as [Refund Date],Refund.Remarks as [Remarks] from Hostelers,Refund where Refund.AcadYear='" & cmbAcadYear.Text & "' and Refund.HostelerID=Hostelers.HostelerID and Refund.HostelerID", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Refund")
            DataGridView3.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView3.DataSource = myDataSet.Tables("Refund").DefaultView
            Dim sum As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView3.Rows
                sum = sum + r.Cells(5).Value
            Next
            TextBox1.Text = sum
            Label3.Text = frmDueCharges.GetInWords(TextBox1.Text.Trim)
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        cmbAcadYear.Text = ""
        DataGridView3.DataSource = Nothing
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
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
End Class