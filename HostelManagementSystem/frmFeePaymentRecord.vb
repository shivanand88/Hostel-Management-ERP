Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class frmFeePaymentRecord
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"
   


    Private Sub frmFeePaymentRecord_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub
    Private Sub frmFeePaymentRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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
        txtUSN.Text = ""
        txtHostelerName.Text = ""
        DataGridView1.DataSource = Nothing
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DateFrom.Text = Today
        DateTo.Text = Today
        DataGridView2.DataSource = Nothing
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
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
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT FeePaymentID as [Fee Payment ID ],Hostelers.HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],HostelName as [HostelName],RoomNo as [Room No],AcadYear as [Acasemic year],ServiceCharges as [Service Charges],PayableCharges as [Payable Charges],TotalPaid as [Total Paid],DuePayment as [Due Charges],PaymentDate as [Payment Date],BankChallanNumber as [Bank Challan No], ExtraCharges as [Extra Charges],Fine from Hostelers,FeePayment where Hostelers.HostelerID=FeePayment.HostelerID and PaymentDate between #" & DateFrom.Text & " # and #" & DateTo.Text & "# order by HostelerName,PaymentDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "FeePayment")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView2.DataSource = myDataSet.Tables("FeePayment").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT FeePaymentID as [Fee Payment ID ],Hostelers.HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],HostelName as [HostelName],RoomNo as [Room No],AcadYear as [Acasemic year],ServiceCharges as [Service Charges],PayableCharges as [Payable Charges],TotalPaid as [Total Paid],DuePayment as [Due Charges],PaymentDate as [Payment Date],BankChallanNumber as [Bank Challan No], ExtraCharges as [Extra Charges],Fine from Hostelers,FeePayment where Hostelers.HostelerID=FeePayment.HostelerID and PaymentDate between #" & DateTimePicker2.Text & " # and #" & DateTimePicker1.Text & "# and DuePayment > 0 order by HostelerName,PaymentDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "FeePayment")
            DataGridView3.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView3.DataSource = myDataSet.Tables("FeePayment").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        DateTimePicker1.Text = Today
        DateTimePicker2.Text = Today
        DataGridView3.DataSource = Nothing
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
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
    Private Sub txtUSN_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUSN.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT FeePaymentID as [Fee Payment ID ],Hostelers.HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],HostelName as [HostelName],RoomNo as [Room No],AcadYear as [Acasemic year],ServiceCharges as [Service Charges],PayableCharges as [Payable Charges],TotalPaid as [Total Paid],DuePayment as [Due Charges],PaymentDate as [Payment Date],BankChallanNumber as [Bank Challan No], ExtraCharges as [Extra Charges],Fine from Hostelers,FeePayment where Hostelers.HostelerID=FeePayment.HostelerID and USN like '%" & txtUSN.Text & "%' order by HostelerName,PaymentDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "FeePayment")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("FeePayment").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub txtHostelerName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHostelerName.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT FeePaymentID as [Fee Payment ID ],Hostelers.HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],HostelName as [HostelName],RoomNo as [Room No],AcadYear as [Acasemic year],ServiceCharges as [Service Charges],PayableCharges as [Payable Charges],TotalPaid as [Total Paid],DuePayment as [Due Charges],PaymentDate as [Payment Date],BankChallanNumber as [Bank Challan No], ExtraCharges as [Extra Charges],Fine from Hostelers,FeePayment where Hostelers.HostelerID=FeePayment.HostelerID and HostelerName like '%" & txtHostelerName.Text & "%' order by HostelerName,PaymentDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "FeePayment")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("FeePayment").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        cmbHostelerName1.Text = ""
        DateTimePicker4.Text = Today
        DateTimePicker3.Text = Today
        DataGridView4.DataSource = Nothing
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
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

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Try
            If Len(Trim(cmbHostelerName1.Text)) = 0 Then
                MessageBox.Show("Please select hosteler name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbHostelerName1.Focus()
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT FeePaymentID as [Fee Payment ID ],Hostelers.HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],HostelName as [HostelName],RoomNo as [Room No],AcadYear as [Acasemic year],ServiceCharges as [Service Charges],PayableCharges as [Payable Charges],TotalPaid as [Total Paid],DuePayment as [Due Charges],PaymentDate as [Payment Date],BankChallanNumber as [Bank Challan No], ExtraCharges as [Extra Charges],Fine from Hostelers,FeePayment where Hostelers.HostelerID=FeePayment.HostelerID and PaymentDate between #" & DateTimePicker4.Text & " # and #" & DateTimePicker3.Text & "# and HostelerName='" & cmbHostelerName1.Text & "' order by HostelerName,PaymentDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "FeePayment")
            DataGridView4.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView4.DataSource = myDataSet.Tables("FeePayment").DefaultView
            con.Close()
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

    Private Sub DataGridView2_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView2.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView2.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView2.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub DataGridView3_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView3.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView3.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView3.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub DataGridView4_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView4.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView4.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView4.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub TabControl1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.Click
        txtUSN.Text = ""
        txtHostelerName.Text = ""
        DataGridView1.DataSource = Nothing
        DateFrom.Text = Today
        DateTo.Text = Today
        DataGridView2.DataSource = Nothing
        DateTimePicker1.Text = Today
        DateTimePicker2.Text = Today
        DataGridView3.DataSource = Nothing
        cmbHostelerName1.Text = ""
        DateTimePicker4.Text = Today
        DateTimePicker3.Text = Today
        DataGridView4.DataSource = Nothing
    End Sub

   

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Try

            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptFeePayment 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New DataSet  'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT FeePayment.FeePaymentID, FeePayment.HostelerID, FeePayment.ServiceCharges, FeePayment.ExtraCharges, FeePayment.AcadYear,FeePayment.PaymentDate, FeePayment.TotalPaid, FeePayment.Fine, FeePayment.DuePayment,FeePayment.BankChallanNumber, Hostelers.HostelerID AS Expr1, Hostelers.HostelerName,Hostelers.DOB, Hostelers.Gender, Hostelers.RoomNo, Hostelers.HostelName, Hostelers.DateOfJoining, Hostelers.Purpose, Hostelers.FatherName, Hostelers.MobNo1, Hostelers.Phone1, Hostelers.MotherName, Hostelers.MobNo2, Hostelers.City, Hostelers.Address, Hostelers.Email, Hostelers.ContactNo,Hostelers.InstOfcDetails, Hostelers.Phone2, Hostelers.Agreement, Hostelers.GuardianName, Hostelers.GuardianAddress, Hostelers.MobNo3,Hostelers.Phone3, Hostelers.Photo, Hostelers.DocsPic, Hostelers.CompletionDate FROM FeePayment,Hostelers where FeePayment.HostelerID = Hostelers.HostelerID and PaymentDate between #" & DateFrom.Text & " # and #" & DateTo.Text & "# order by HostelerName,PaymentDate"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Hostelers")
            myDA.Fill(myDS, "FeePayment")
            rpt.SetDataSource(myDS)
            frmFeePaymentReport.CrystalReportViewer1.ReportSource = rpt
            frmFeePaymentReport.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Try

            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptFeePayment 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New DataSet  'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT FeePayment.FeePaymentID, FeePayment.HostelerID, FeePayment.ServiceCharges, FeePayment.ExtraCharges, FeePayment.AcadYear,FeePayment.PaymentDate, FeePayment.TotalPaid, FeePayment.Fine, FeePayment.DuePayment,FeePayment.BankChallanNumber, Hostelers.HostelerID AS Expr1, Hostelers.HostelerName,Hostelers.DOB, Hostelers.Gender, Hostelers.RoomNo, Hostelers.HostelName, Hostelers.DateOfJoining, Hostelers.Purpose, Hostelers.FatherName, Hostelers.MobNo1, Hostelers.Phone1, Hostelers.MotherName, Hostelers.MobNo2, Hostelers.City, Hostelers.Address, Hostelers.Email, Hostelers.ContactNo,Hostelers.InstOfcDetails, Hostelers.Phone2, Hostelers.Agreement, Hostelers.GuardianName, Hostelers.GuardianAddress, Hostelers.MobNo3,Hostelers.Phone3, Hostelers.Photo, Hostelers.DocsPic, Hostelers.CompletionDate FROM FeePayment,Hostelers where FeePayment.HostelerID = Hostelers.HostelerID and PaymentDate between #" & DateTimePicker2.Text & "# and #" & DateTimePicker1.Text & "# and DuePayment > 0 order by HostelerName,PaymentDate"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Hostelers")
            myDA.Fill(myDS, "FeePayment")
            rpt.SetDataSource(myDS)
            frmFeePaymentReport.CrystalReportViewer1.ReportSource = rpt
            frmFeePaymentReport.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptFeePayment 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New DataSet  'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT FeePayment.FeePaymentID, FeePayment.HostelerID, FeePayment.ServiceCharges, FeePayment.ExtraCharges, FeePayment.AcadYear,FeePayment.PaymentDate, FeePayment.TotalPaid, FeePayment.Fine, FeePayment.DuePayment,FeePayment.BankChallanNumber, Hostelers.HostelerID AS Expr1, Hostelers.HostelerName,Hostelers.DOB, Hostelers.Gender, Hostelers.RoomNo, Hostelers.HostelName, Hostelers.DateOfJoining, Hostelers.Purpose, Hostelers.FatherName, Hostelers.MobNo1, Hostelers.Phone1, Hostelers.MotherName, Hostelers.MobNo2, Hostelers.City, Hostelers.Address, Hostelers.Email, Hostelers.ContactNo,Hostelers.InstOfcDetails, Hostelers.Phone2, Hostelers.Agreement, Hostelers.GuardianName, Hostelers.GuardianAddress, Hostelers.MobNo3,Hostelers.Phone3, Hostelers.Photo, Hostelers.DocsPic, Hostelers.CompletionDate FROM FeePayment,Hostelers where FeePayment.HostelerID = Hostelers.HostelerID and PaymentDate between #" & DateTimePicker4.Text & "# and #" & DateTimePicker3.Text & "# and Hostelername='" & cmbHostelerName1.Text & "' order by HostelerName,PaymentDate"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Hostelers")
            myDA.Fill(myDS, "FeePayment")
            rpt.SetDataSource(myDS)
            frmFeePaymentReport.CrystalReportViewer1.ReportSource = rpt
            frmFeePaymentReport.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    
End Class