Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class frmPurchasedInventoryRecord
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"

    Sub fillProductName()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct RTRIM(ProductName) FROM PurchasedInventory", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbProductName.Items.Clear()
            cmbProductName1.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbProductName.Items.Add(drow(0).ToString())
                cmbProductName1.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub fillCategory()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct RTRIM(Category) FROM PurchasedInventory", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbCategory.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbCategory.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub frmPurchasedInventoryRecord_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub
    Private Sub frmPurchasedInventoryRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillProductName()
        fillCategory()
        fillPartyname()
    End Sub
    Sub fillPartyname()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct RTRIM(PartyName) FROM PurchasedInventory", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbPartyName.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbPartyName.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub cmbProductName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbProductName.SelectedIndexChanged
        Try
            GroupBox2.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT ID as [Purchase ID], ProductName as [Product Name],Category,TransactionType as [Transaction Type],PartyName as [Party Name],PurchaseDate as [Purchase Date],Quantity as [Quantity],Unit as [Unit],Price as [Unit Price],TotalPrice as [Total Price] from PurchasedInventory where ProductName='" & cmbProductName.Text & "' order by ProductName,PurchaseDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "PurchasedInventory")
            DataGridView1.DataSource = myDataSet.Tables("PurchasedInventory").DefaultView
            Dim sum As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                sum = sum + r.Cells(9).Value
            Next
            TextBox1.Text = sum
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtProductName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtProductName.TextChanged
        Try
            GroupBox2.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT ID as [Purchase ID], ProductName as [Product Name],Category,TransactionType as [Transaction Type],PartyName as [Party Name],PurchaseDate as [Purchase Date],Quantity as [Quantity],Unit as [Unit],Price as [Unit Price],TotalPrice as [Total Price] from PurchasedInventory where ProductName like '" & txtProductName.Text & "%' order by ProductName,PurchaseDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "PurchasedInventory")
            DataGridView1.DataSource = myDataSet.Tables("PurchasedInventory").DefaultView
            Dim sum As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                sum = sum + r.Cells(9).Value
            Next
            TextBox1.Text = sum
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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
        cmbProductName.Text = ""
        txtProductName.Text = ""
        DataGridView1.DataSource = Nothing
        GroupBox2.Visible = False
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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DataGridView2.DataSource = Nothing
        DateFrom.Text = Today
        DateTo.Text = Today
        GroupBox7.Visible = False
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            GroupBox7.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT ID as [Purchase ID], ProductName as [Product Name],Category,TransactionType as [Transaction Type],PartyName as [Party Name],PurchaseDate as [Purchase Date],Quantity as [Quantity],Unit as [Unit],Price as [Unit Price],TotalPrice as [Total Price] from PurchasedInventory where Purchasedate between #" & DateFrom.Text & " # and #" & DateTo.Text & "# order by ProductName,PurchaseDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "PurchasedInventory")
            DataGridView2.DataSource = myDataSet.Tables("PurchasedInventory").DefaultView
            Dim sum As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView2.Rows
                sum = sum + r.Cells(9).Value
            Next
            TextBox2.Text = sum
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView1_RowPostPaint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView1.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView1.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub DataGridView2_RowPostPaint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView2.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView2.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView2.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
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

    Private Sub DataGridView3_RowPostPaint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView3.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView3.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView3.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Try
            If Len(Trim(cmbProductName1.Text)) = 0 Then
                MessageBox.Show("Please select product name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbProductName1.Focus()
                Exit Sub
            End If
            GroupBox11.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT ID as [Purchase ID], ProductName as [Product Name],Category,TransactionType as [Transaction Type],PartyName as [Party Name],PurchaseDate as [Purchase Date],Quantity as [Quantity],Unit as [Unit],Price as [Unit Price],TotalPrice as [Total Price] from PurchasedInventory where ProductName='" & cmbProductName1.Text & "' and Purchasedate between #" & DateTimePicker2.Text & " # and #" & DateTimePicker1.Text & "# order by ProductName,PurchaseDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "PurchasedInventory")
            DataGridView3.DataSource = myDataSet.Tables("PurchasedInventory").DefaultView
            Dim sum As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView3.Rows
                sum = sum + r.Cells(9).Value
            Next
            TextBox3.Text = sum
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        cmbProductName1.Text = ""
        DateTimePicker1.Text = Today
        DateTimePicker2.Text = Today
        DataGridView3.DataSource = Nothing
        GroupBox11.Visible = False
    End Sub

    Private Sub TabControl1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.Click
        cmbProductName1.Text = ""
        DateTimePicker1.Text = Today
        DateTimePicker2.Text = Today
        DataGridView3.DataSource = Nothing
        DataGridView2.DataSource = Nothing
        DateFrom.Text = Today
        DateTo.Text = Today
        GroupBox7.Visible = False
        cmbProductName.Text = ""
        txtProductName.Text = ""
        DataGridView1.DataSource = Nothing
        GroupBox2.Visible = False
        GroupBox11.Visible = False
        cmbCategory.Text = ""
        txtCategory.Text = ""
        DataGridView4.DataSource = Nothing
        GroupBox15.Visible = False
        cmbTransType.SelectedIndex = -1
        DataGridView5.DataSource = Nothing
        GroupBox18.Visible = False
        cmbPartyName.Text = ""
        txtPartyName.Text = ""
        DataGridView6.DataSource = Nothing
        GroupBox19.Visible = False
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try
            If Len(Trim(cmbProductName.Text)) = 0 Then
                MessageBox.Show("Please select product name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbProductName.Focus()
                Exit Sub
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptPurchasedInventory 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New HMS_DBDataSet 'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT * from PurchasedInventory where ProductName='" & cmbProductName.Text & "' order by ProductName,PurchaseDate"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "PurchasedInventory")
            rpt.SetDataSource(myDS)
            frmPurchasedInventoryReport.CrystalReportViewer1.ReportSource = rpt
            frmPurchasedInventoryReport.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptPurchasedInventory 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New HMS_DBDataSet 'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT * from PurchasedInventory where Purchasedate between #" & DateFrom.Text & " # and #" & DateTo.Text & "# order by ProductName,PurchaseDate"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "PurchasedInventory")
            rpt.SetDataSource(myDS)
            frmPurchasedInventoryReport.CrystalReportViewer1.ReportSource = rpt
            frmPurchasedInventoryReport.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Try
            If Len(Trim(cmbProductName1.Text)) = 0 Then
                MessageBox.Show("Please select product name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbProductName1.Focus()
                Exit Sub
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptPurchasedInventory 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New HMS_DBDataSet 'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT * from PurchasedInventory where ProductName='" & cmbProductName1.Text & "' and Purchasedate between #" & DateTimePicker2.Text & " # and #" & DateTimePicker1.Text & "# order by ProductName,PurchaseDate"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Hostelers")
            myDA.Fill(myDS, "FeePayment")
            rpt.SetDataSource(myDS)
            frmAccountRecordReport.CrystalReportViewer1.ReportSource = rpt
            frmPurchasedInventoryReport.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbCategory_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCategory.SelectedIndexChanged
        Try
            GroupBox15.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT ID as [Purchase ID], ProductName as [Product Name],Category,TransactionType as [Transaction Type],PartyName as [Party Name],PurchaseDate as [Purchase Date],Quantity as [Quantity],Unit as [Unit],Price as [Unit Price],TotalPrice as [Total Price] from PurchasedInventory where Category='" & cmbCategory.Text & "' order by ProductName,PurchaseDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "PurchasedInventory")
            DataGridView4.DataSource = myDataSet.Tables("PurchasedInventory").DefaultView
            Dim sum As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView4.Rows
                sum = sum + r.Cells(9).Value
            Next
            TextBox4.Text = sum
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtCategory_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCategory.TextChanged
        Try
            GroupBox15.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT ID as [Purchase ID], ProductName as [Product Name],Category,TransactionType as [Transaction Type],PartyName as [Party Name],PurchaseDate as [Purchase Date],Quantity as [Quantity],Unit as [Unit],Price as [Unit Price],TotalPrice as [Total Price] from PurchasedInventory where Category like '" & txtCategory.Text & "%' order by ProductName,PurchaseDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "PurchasedInventory")
            DataGridView4.DataSource = myDataSet.Tables("PurchasedInventory").DefaultView
            Dim sum As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView4.Rows
                sum = sum + r.Cells(9).Value
            Next
            TextBox4.Text = sum
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        cmbCategory.Text = ""
        txtCategory.Text = ""
        DataGridView4.DataSource = Nothing
        GroupBox15.Visible = False
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
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

    Private Sub DataGridView4_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView4.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView4.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView4.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Try
            If Len(Trim(cmbCategory.Text)) = 0 Then
                MessageBox.Show("Please select category name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbCategory.Focus()
                Exit Sub
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptPurchasedInventory 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New HMS_DBDataSet 'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT * from PurchasedInventory where Category='" & cmbCategory.Text & "' order by ProductName,PurchaseDate"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "PurchasedInventory")
            rpt.SetDataSource(myDS)
            frmPurchasedInventoryReport.CrystalReportViewer1.ReportSource = rpt
            frmPurchasedInventoryReport.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Try
            If Len(Trim(cmbTransType.Text)) = 0 Then
                MessageBox.Show("Please select transaction type", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbTransType.Focus()
                Exit Sub
            End If
            GroupBox18.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT ID as [Purchase ID], ProductName as [Product Name],Category,TransactionType as [Transaction Type],PartyName as [Party Name],PurchaseDate as [Purchase Date],Quantity as [Quantity],Unit as [Unit],Price as [Unit Price],TotalPrice as [Total Price] from PurchasedInventory where TransactionType='" & cmbTransType.Text & "' order by ProductName,PurchaseDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "PurchasedInventory")
            DataGridView5.DataSource = myDataSet.Tables("PurchasedInventory").DefaultView
            Dim sum As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView5.Rows
                sum = sum + r.Cells(9).Value
            Next
            TextBox5.Text = sum
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        cmbTransType.SelectedIndex = -1
        DataGridView5.DataSource = Nothing
        GroupBox18.Visible = False
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        If DataGridView5.RowCount = Nothing Then
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

            rowsTotal = DataGridView5.RowCount - 1
            colsTotal = DataGridView5.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView5.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView5.Rows(I).Cells(j).Value
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

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        If DataGridView6.RowCount = Nothing Then
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

            rowsTotal = DataGridView6.RowCount - 1
            colsTotal = DataGridView6.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView6.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView6.Rows(I).Cells(j).Value
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

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        cmbPartyName.Text = ""
        txtPartyName.Text = ""
        DataGridView6.DataSource = Nothing
        GroupBox19.Visible = False
    End Sub

    Private Sub cmbPartyName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbPartyName.SelectedIndexChanged
        Try
            GroupBox19.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT ID as [Purchase ID], ProductName as [Product Name],Category,TransactionType as [Transaction Type],PartyName as [Party Name],PurchaseDate as [Purchase Date],Quantity as [Quantity],Unit as [Unit],Price as [Unit Price],TotalPrice as [Total Price] from PurchasedInventory where Partyname='" & cmbPartyName.Text & "' order by ProductName,PurchaseDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "PurchasedInventory")
            DataGridView6.DataSource = myDataSet.Tables("PurchasedInventory").DefaultView
            Dim sum As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView6.Rows
                sum = sum + r.Cells(9).Value
            Next
            TextBox6.Text = sum
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtPartyName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPartyName.TextChanged
        Try
            GroupBox19.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT ID as [Purchase ID], ProductName as [Product Name],Category,TransactionType as [Transaction Type],PartyName as [Party Name],PurchaseDate as [Purchase Date],Quantity as [Quantity],Unit as [Unit],Price as [Unit Price],TotalPrice as [Total Price] from PurchasedInventory where Partyname like '" & txtPartyName.Text & "%' order by ProductName,PurchaseDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "PurchasedInventory")
            DataGridView6.DataSource = myDataSet.Tables("PurchasedInventory").DefaultView
            Dim sum As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView6.Rows
                sum = sum + r.Cells(9).Value
            Next
            TextBox6.Text = sum
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        Try
            If Len(Trim(cmbPartyName.Text)) = 0 Then
                MessageBox.Show("Please select party name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbPartyName.Focus()
                Exit Sub
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptPurchasedInventory 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New HMS_DBDataSet 'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT * from PurchasedInventory where PartyName='" & cmbPartyName.Text & "' order by ProductName,PurchaseDate"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "PurchasedInventory")
            rpt.SetDataSource(myDS)
            frmPurchasedInventoryReport.CrystalReportViewer1.ReportSource = rpt
            frmPurchasedInventoryReport.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        Try
            If Len(Trim(cmbTransType.Text)) = 0 Then
                MessageBox.Show("Please select transaction type", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbTransType.Focus()
                Exit Sub
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptPurchasedInventory 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New HMS_DBDataSet 'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT * from PurchasedInventory where TransactionType='" & cmbTransType.Text & "' order by ProductName,PurchaseDate"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "PurchasedInventory")
            rpt.SetDataSource(myDS)
            frmPurchasedInventoryReport.CrystalReportViewer1.ReportSource = rpt
            frmPurchasedInventoryReport.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class