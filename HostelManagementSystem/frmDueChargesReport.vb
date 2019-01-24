Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmDueChargesReport
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"
    
    Sub GetData()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand(" SELECT Hostelers.HostelerName as [Hosteler Name],Hostelers.USN as [USN], RoomAllotment.Hostelname as [Hostel Name], RoomAllotment.RoomNo as [Room No], FeePayment.AcadYear as [Academic year],  FeePayment.ServiceCharges as [Total], FeePayment.PayableCharges as [Prev Due ], FeePayment.TotalPaid as [Total Paid],  FeePayment.DuePayment as [Present Due],  FeePayment.PaymentDate as [Payment date],  FeePayment.BankChallanNumber as [Bank Challan No] FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where  FeePayment.AcadYear=RoomAllotment.AcadYear order by Hostelers.HostelerName ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "RoomAllotment")
            myDA.Fill(myDataSet, "FeePayment")
            DataGridView3.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView3.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
            DataGridView3.DataSource = myDataSet.Tables("FeePayment").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnGetData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetData.Click

        Try
            GroupBox2.Visible = True
            Label1.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("SELECT Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], Hostelers.Purpose as [Department], DueAmount.TotalCharges as[Total Charges] , DueAmount.TotalDueAmount as[Due Amount], DueAmount.AcadYear as [Academic Year] FROM Hostelers INNER JOIN DueAmount ON Hostelers.[HostelerID] = DueAmount.[HostelerID]  where TotalDueAmount > '0' order by Hostelername", con) ' Hostelers,FeePayment where Hostelers.HostelerID=FeePayment.HostelerID group by Hostelers.HostelerID,Hostelername,Hostelname,RoomNo having sum(DuePayment >0) 
            cmd = New OleDbCommand("  SELECT Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], Hostelers.AcadYear as [Academic Year], Hostelers.HostelName as [Hostel Name], DueAmount.TotalCharges as [Total Charges], (DueAmount.TotalCharges - DueAmount.TotalDueAmount ) as [Total Paid], DueAmount.TotalDueAmount as [Remaining Due]  FROM (Hostel INNER JOIN Hostelers ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN DueAmount ON Hostelers.[HostelerID] = DueAmount.[HostelerID] where DueAmount.TotalDueAmount> '0' order by Hostelers.Hostelername ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Hostel")
            myDA.Fill(myDataSet, "DueAmount")
            'myDA.Fill(myDataSet, "Feepayment")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("Hostel").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("DueAmount").DefaultView
            'DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            Dim sum As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                sum = sum + r.Cells(6).Value
            Next
            TextBox1.Text = sum
            Label1.Text = frmDueCharges.GetInWords(TextBox1.Text.Trim)
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub





    Private Sub btnGetDataHostel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetDataHostel.Click
        Try
            GroupBox1.Visible = True
            Label2.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("SELECT Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], Hostelers.Purpose as [Department], DueAmount.TotalCharges as[Total Charges] , DueAmount.TotalDueAmount as[Due Amount], DueAmount.AcadYear as [Academic Year] FROM Hostelers INNER JOIN DueAmount ON Hostelers.[HostelerID] = DueAmount.[HostelerID]  where TotalDueAmount > '0' order by Hostelername", con) ' Hostelers,FeePayment where Hostelers.HostelerID=FeePayment.HostelerID group by Hostelers.HostelerID,Hostelername,Hostelname,RoomNo having sum(DuePayment >0) 
            cmd = New OleDbCommand("  SELECT Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], Hostelers.AcadYear as [Academic Year], Hostelers.HostelName as [Hostel Name], DueAmount.TotalCharges as [Total Fees], (DueAmount.TotalCharges - DueAmount.TotalDueAmount ) as [Total Paid], DueAmount.TotalDueAmount as [Remaining Due]  FROM (Hostel INNER JOIN Hostelers ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN DueAmount ON Hostelers.[HostelerID] = DueAmount.[HostelerID] where DueAmount.TotalDueAmount> '0' and Hostel.HostelType= '" & cmbHostel.Text & "' order by Hostelers.Hostelername ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Hostel")
            myDA.Fill(myDataSet, "DueAmount")
            'myDA.Fill(myDataSet, "Feepayment")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView2.DataSource = myDataSet.Tables("Hostel").DefaultView
            DataGridView2.DataSource = myDataSet.Tables("DueAmount").DefaultView
            'DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            Dim sum As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView2.Rows
                sum = sum + r.Cells(6).Value
            Next
            TextBox2.Text = sum
            Label2.Text = frmDueCharges.GetInWords(TextBox2.Text.Trim)
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        TextBox1.Text = ""
        Label1.Text = ""
        Label1.Visible = False
        DataGridView1.DataSource = Nothing
    End Sub

    Private Sub btnExportExcelAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportExcelAll.Click
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

    Private Sub frmDueChargesReport_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub

    Private Sub frmDueChargesReport_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        GetData()
    End Sub

    
    Private Sub btnResetHostel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResetHostel.Click
        cmbHostel.Text = ""
        TextBox2.Text = ""
        Label2.Text = ""
        Label2.Visible = False
        DataGridView2.DataSource = Nothing
    End Sub

    Private Sub btnExportExcelHostel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportExcelHostel.Click
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

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        txtUSN.Text = ""
        txtHostelerName.Text = ""
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

    Private Sub txtHostelerName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHostelerName.TextChanged
        txtUSN.Text = ""
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand(" SELECT Hostelers.HostelerName as [Hosteler Name],Hostelers.USN as [USN], RoomAllotment.Hostelname as [Hostel Name], RoomAllotment.RoomNo as [Room No], FeePayment.AcadYear as [Academic year],  FeePayment.ServiceCharges as [Total], FeePayment.PayableCharges as [Prev Due ], FeePayment.TotalPaid as [Total Paid],  FeePayment.DuePayment as [Present Due],  FeePayment.PaymentDate as [Payment date],  FeePayment.BankChallanNumber as [Bank Challan No] FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where Hostelers.HostelerName like '%" & txtHostelerName.Text & "%' and FeePayment.AcadYear=RoomAllotment.AcadYear order by Hostelers.HostelerName ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "RoomAllotment")
            myDA.Fill(myDataSet, "FeePayment")
            DataGridView3.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView3.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
            DataGridView3.DataSource = myDataSet.Tables("FeePayment").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtUSN_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUSN.TextChanged
        txtHostelerName.Text = ""
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand(" SELECT Hostelers.HostelerName as [Hosteler Name],Hostelers.USN as [USN], RoomAllotment.Hostelname as [Hostel Name], RoomAllotment.RoomNo as [Room No], FeePayment.AcadYear as [Academic year],  FeePayment.ServiceCharges as [Total], FeePayment.PayableCharges as [Prev Due ], FeePayment.TotalPaid as [Total Paid],  FeePayment.DuePayment as [Present Due],  FeePayment.PaymentDate as [Payment date],  FeePayment.BankChallanNumber as [Bank Challan No] FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where Hostelers.USN like '%" & txtUSN.Text & "%' and FeePayment.AcadYear=RoomAllotment.AcadYear order by Hostelers.HostelerName ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "RoomAllotment")
            myDA.Fill(myDataSet, "FeePayment")
            DataGridView3.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView3.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
            DataGridView3.DataSource = myDataSet.Tables("FeePayment").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

   
End Class