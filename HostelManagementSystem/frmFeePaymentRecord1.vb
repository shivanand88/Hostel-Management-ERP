Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmFeePaymentRecord1
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
            adp.SelectCommand = New OleDbCommand("SELECT distinct RTRIM(HostelerName) FROM FeePayment,Hostelers where Hostelers.HostelerID=FeePayment.HostelerID", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbHostelerName.Items.Clear()
            cmbHostelerName1.Items.Clear()
            ComboBox1.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbHostelerName.Items.Add(drow(0).ToString())
                cmbHostelerName1.Items.Add(drow(0).ToString())
                ComboBox1.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmFeePaymentRecord1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmFeePayment.Show()
    End Sub
    Private Sub frmFeePaymentRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GetData()
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
    Sub GetData()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT FeePayment.FeePaymentID as [Fee Payment ID ], FeePayment.HostelerID as [Hosteler ID], Hostelers.HostelerName as [Hosteler Name], RoomAllotment.Hostelname as [Hostel Name], RoomAllotment.RoomNo as [Room No], FeePayment.AcadYear as [Academic year],  FeePayment.ServiceCharges as [Total], FeePayment.PayableCharges as [Prev Due ], FeePayment.TotalPaid as [Total Paid],  FeePayment.DuePayment as [Present Due],  FeePayment.PaymentDate as [Payment date],  FeePayment.BankChallanNumber as [Bank Challan No] FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where  FeePayment.AcadYear=RoomAllotment.AcadYear order by HostelerName,PaymentDate desc ", con)
            'cmd = New OleDbCommand("SELECT FeePaymentID as [Fee Payment ID ],Hostelers.HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],HostelName as [HostelName],RoomNo as [Room No],AcadYear as [Acasemic year],ServiceCharges as [Service Charges],PayableCharges as [Payable Charges],TotalPaid as [Total Paid],DuePayment as [Due Charges],PaymentDate as [Payment Date],BankChallanNumber as [Bank Challan No], ExtraCharges as [Extra Charges],Fine from Hostelers,FeePayment where Hostelers.HostelerID=FeePayment.HostelerID and HostelerName='" & cmbHostelerName.Text & "' order by HostelerName,PaymentDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "RoomAllotment")
            myDA.Fill(myDataSet, "FeePayment")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("FeePayment").DefaultView
           
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbHostelerName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbHostelerName.SelectedIndexChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT FeePayment.FeePaymentID as [Fee Payment ID ], FeePayment.HostelerID as [Hosteler ID], Hostelers.HostelerName as [Hosteler Name], RoomAllotment.Hostelname as [Hostel Name], RoomAllotment.RoomNo as [Room No], FeePayment.AcadYear as [Academic year],  FeePayment.ServiceCharges as [Total], FeePayment.PayableCharges as [Prev Due ], FeePayment.TotalPaid as [Total Paid],  FeePayment.DuePayment as [Present Due],  FeePayment.PaymentDate as [Payment date],  FeePayment.BankChallanNumber as [Bank Challan No] FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where  Hostelers.HostelerName='" & cmbHostelerName.Text & "' and FeePayment.AcadYear=RoomAllotment.AcadYear order by HostelerName,PaymentDate desc ", con)
            'cmd = New OleDbCommand("SELECT FeePaymentID as [Fee Payment ID ],Hostelers.HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],HostelName as [HostelName],RoomNo as [Room No],AcadYear as [Acasemic year],ServiceCharges as [Service Charges],PayableCharges as [Payable Charges],TotalPaid as [Total Paid],DuePayment as [Due Charges],PaymentDate as [Payment Date],BankChallanNumber as [Bank Challan No], ExtraCharges as [Extra Charges],Fine from Hostelers,FeePayment where Hostelers.HostelerID=FeePayment.HostelerID and HostelerName='" & cmbHostelerName.Text & "' order by HostelerName,PaymentDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "RoomAllotment")
            myDA.Fill(myDataSet, "FeePayment")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("FeePayment").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        cmbHostelerName.Text = ""
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
            cmd = New OleDbCommand(" SELECT FeePayment.FeePaymentID as [Fee Payment ID ], FeePayment.HostelerID as [Hosteler ID], Hostelers.HostelerName as [Hosteler Name], RoomAllotment.Hostelname as [Hostel Name], RoomAllotment.RoomNo as [Room No], FeePayment.AcadYear as [Academic year],  FeePayment.ServiceCharges as [Total], FeePayment.PayableCharges as [Prev Due ], FeePayment.TotalPaid as [Total Paid],  FeePayment.DuePayment as [Present Due],  FeePayment.PaymentDate as [Payment date],  FeePayment.BankChallanNumber as [Bank Challan No] FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where FeePayment.PaymentDate between #" & DateFrom.Text & " # and #" & DateTo.Text & "# and FeePayment.AcadYear=RoomAllotment.AcadYear order by PaymentDate desc", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "RoomAllotment")
            myDA.Fill(myDataSet, "FeePayment")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView2.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
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
            cmd = New OleDbCommand(" SELECT FeePayment.FeePaymentID as [Fee Payment ID ], FeePayment.HostelerID as [Hosteler ID], Hostelers.HostelerName as [Hosteler Name], RoomAllotment.Hostelname as [Hostel Name], RoomAllotment.RoomNo as [Room No], FeePayment.AcadYear as [Academic year],  FeePayment.ServiceCharges as [Total], FeePayment.PayableCharges as [Prev Due], FeePayment.TotalPaid as [Total Paid],  FeePayment.DuePayment as [Present Due],  FeePayment.PaymentDate as [Payment date],  FeePayment.BankChallanNumber as [Bank Challan No] FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where  FeePayment.PaymentDate between #" & DateTimePicker2.Text & " # and #" & DateTimePicker1.Text & "# and FeePayment.AcadYear=RoomAllotment.AcadYear and DuePayment > 0 order by HostelerName,PaymentDate desc", con)
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
            txtHostelerName.Text = ""
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand(" SELECT FeePayment.FeePaymentID as [Fee Payment ID ], FeePayment.HostelerID as [Hosteler ID], Hostelers.HostelerName as [Hosteler Name], RoomAllotment.Hostelname as [Hostel Name], RoomAllotment.RoomNo as [Room No], FeePayment.AcadYear as [Academic year],  FeePayment.ServiceCharges as [Total], FeePayment.PayableCharges as [Prev Due], FeePayment.TotalPaid as [Total Paid],  FeePayment.DuePayment as [Present Due],  FeePayment.PaymentDate as [Payment date],  FeePayment.BankChallanNumber as [Bank Challan No] FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where  Hostelers.USN like '%" & txtUSN.Text & "%' and FeePayment.AcadYear=RoomAllotment.AcadYear order by HostelerName,PaymentDate desc", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "RoomAllotment")
            myDA.Fill(myDataSet, "FeePayment")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("FeePayment").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub txtHostelerName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHostelerName.TextChanged
        Try
            txtUSN.Text = ""
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand(" SELECT FeePayment.FeePaymentID as [Fee Payment ID ], FeePayment.HostelerID as [Hosteler ID], Hostelers.HostelerName as [Hosteler Name], RoomAllotment.Hostelname as [Hostel Name], RoomAllotment.RoomNo as [Room No], FeePayment.AcadYear as [Academic year],  FeePayment.ServiceCharges as [Total], FeePayment.PayableCharges as [Prev Due], FeePayment.TotalPaid as [Total Paid],  FeePayment.DuePayment as [Present Due],  FeePayment.PaymentDate as [Payment date],  FeePayment.BankChallanNumber as [Bank Challan No] FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where  Hostelers.HostelerName like '%" & txtHostelerName.Text & "%' and FeePayment.AcadYear=RoomAllotment.AcadYear order by HostelerName,PaymentDate desc", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "RoomAllotment")
            myDA.Fill(myDataSet, "FeePayment")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
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
            cmd = New OleDbCommand(" SELECT FeePayment.FeePaymentID as [Fee Payment ID ], FeePayment.HostelerID as [Hosteler ID], Hostelers.HostelerName as [Hosteler Name], RoomAllotment.Hostelname as [Hostel Name], RoomAllotment.RoomNo as [Room No], FeePayment.AcadYear as [Academic year],  FeePayment.ServiceCharges as [Total], FeePayment.PayableCharges as [Prev Due], FeePayment.TotalPaid as [Total Paid],  FeePayment.DuePayment as [Present Due],  FeePayment.PaymentDate as [Payment date],  FeePayment.BankChallanNumber as [Bank Challan No] FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where  FeePayment.PaymentDate between #" & DateTimePicker4.Text & " # and #" & DateTimePicker3.Text & "# and Hostelers.HostelerName='" & cmbHostelerName1.Text & "' and FeePayment.AcadYear=RoomAllotment.AcadYear order by HostelerName,PaymentDate desc", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "RoomAllotment")
            myDA.Fill(myDataSet, "FeePayment")
            DataGridView4.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView4.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
            DataGridView4.DataSource = myDataSet.Tables("FeePayment").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            Me.Hide()
            frmFeePayment.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmFeePayment.txtFeePaymentID.Text = dr.Cells(0).Value.ToString()
            frmFeePayment.cmbHostelerID.Text = dr.Cells(1).Value.ToString()
            frmFeePayment.txtHostelerName.Text = dr.Cells(2).Value.ToString()
            frmFeePayment.txtBranch.Text = dr.Cells(3).Value.ToString()
            frmFeePayment.txtRoomNo.Text = dr.Cells(4).Value.ToString()
            frmFeePayment.cmbAcadYear.Text = dr.Cells(5).Value.ToString()
            frmFeePayment.txtServiceCharges.Text = dr.Cells(6).Value.ToString()
            frmFeePayment.txtTotalCharges.Text = dr.Cells(7).Value.ToString()
            frmFeePayment.txtTotalPaid.Text = dr.Cells(8).Value.ToString()
            frmFeePayment.txtDueCharges.Text = dr.Cells(9).Value.ToString()
            frmFeePayment.dtpPaymentDate.Text = dr.Cells(10).Value.ToString()
            frmFeePayment.txtBankChallanNumber.Text = dr.Cells(11).Value.ToString()
            ' frmFeePayment.txtExtraCharges.Text = dr.Cells(12).Value.ToString()
            ' frmFeePayment.txtFine.Text = dr.Cells(13).Value.ToString()
            frmFeePayment.Update_record.Enabled = True
            frmFeePayment.Delete.Enabled = True
            frmFeePayment.btnSave.Enabled = False
            frmFeePayment.Print.Enabled = True
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

    Private Sub DataGridView2_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView2.SelectedRows(0)
            Me.Hide()
            frmFeePayment.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmFeePayment.txtFeePaymentID.Text = dr.Cells(0).Value.ToString()
            frmFeePayment.cmbHostelerID.Text = dr.Cells(1).Value.ToString()
            frmFeePayment.txtHostelerName.Text = dr.Cells(2).Value.ToString()
            frmFeePayment.txtBranch.Text = dr.Cells(3).Value.ToString()
            frmFeePayment.txtRoomNo.Text = dr.Cells(4).Value.ToString()
            frmFeePayment.cmbAcadYear.Text = dr.Cells(5).Value.ToString()
            frmFeePayment.txtServiceCharges.Text = dr.Cells(6).Value.ToString()
            frmFeePayment.txtTotalCharges.Text = dr.Cells(7).Value.ToString()
            frmFeePayment.txtTotalPaid.Text = dr.Cells(8).Value.ToString()
            frmFeePayment.txtDueCharges.Text = dr.Cells(9).Value.ToString()
            frmFeePayment.dtpPaymentDate.Text = dr.Cells(10).Value.ToString()
            frmFeePayment.txtBankChallanNumber.Text = dr.Cells(11).Value.ToString()
            ' frmFeePayment.txtExtraCharges.Text = dr.Cells(12).Value.ToString()
            ' frmFeePayment.txtFine.Text = dr.Cells(13).Value.ToString()
            frmFeePayment.Update_record.Enabled = True
            frmFeePayment.Delete.Enabled = True
            frmFeePayment.btnSave.Enabled = False
            frmFeePayment.Print.Enabled = True
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

    Private Sub DataGridView3_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView3.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView3.SelectedRows(0)
            Me.Hide()
            frmFeePayment.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmFeePayment.txtFeePaymentID.Text = dr.Cells(0).Value.ToString()
            frmFeePayment.cmbHostelerID.Text = dr.Cells(1).Value.ToString()
            frmFeePayment.txtHostelerName.Text = dr.Cells(2).Value.ToString()
            frmFeePayment.txtBranch.Text = dr.Cells(3).Value.ToString()
            frmFeePayment.txtRoomNo.Text = dr.Cells(4).Value.ToString()
            frmFeePayment.cmbAcadYear.Text = dr.Cells(5).Value.ToString()
            frmFeePayment.txtServiceCharges.Text = dr.Cells(6).Value.ToString()
            frmFeePayment.txtTotalCharges.Text = dr.Cells(7).Value.ToString()
            frmFeePayment.txtTotalPaid.Text = dr.Cells(8).Value.ToString()
            frmFeePayment.txtDueCharges.Text = dr.Cells(9).Value.ToString()
            frmFeePayment.dtpPaymentDate.Text = dr.Cells(10).Value.ToString()
            frmFeePayment.txtBankChallanNumber.Text = dr.Cells(11).Value.ToString()
            ' frmFeePayment.txtExtraCharges.Text = dr.Cells(12).Value.ToString()
            ' frmFeePayment.txtFine.Text = dr.Cells(13).Value.ToString()
            frmFeePayment.Update_record.Enabled = True
            frmFeePayment.Delete.Enabled = True
            frmFeePayment.btnSave.Enabled = False
            frmFeePayment.Print.Enabled = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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

    Private Sub DataGridView4_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView4.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView4.SelectedRows(0)
            Me.Hide()
            frmFeePayment.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmFeePayment.txtFeePaymentID.Text = dr.Cells(0).Value.ToString()
            frmFeePayment.cmbHostelerID.Text = dr.Cells(1).Value.ToString()
            frmFeePayment.txtHostelerName.Text = dr.Cells(2).Value.ToString()
            frmFeePayment.txtBranch.Text = dr.Cells(3).Value.ToString()
            frmFeePayment.txtRoomNo.Text = dr.Cells(4).Value.ToString()
            frmFeePayment.cmbAcadYear.Text = dr.Cells(5).Value.ToString()
            frmFeePayment.txtServiceCharges.Text = dr.Cells(6).Value.ToString()
            frmFeePayment.txtTotalCharges.Text = dr.Cells(7).Value.ToString()
            frmFeePayment.txtTotalPaid.Text = dr.Cells(8).Value.ToString()
            frmFeePayment.txtDueCharges.Text = dr.Cells(9).Value.ToString()
            frmFeePayment.dtpPaymentDate.Text = dr.Cells(10).Value.ToString()
            frmFeePayment.txtBankChallanNumber.Text = dr.Cells(11).Value.ToString()
            ' frmFeePayment.txtExtraCharges.Text = dr.Cells(12).Value.ToString()
            ' frmFeePayment.txtFine.Text = dr.Cells(13).Value.ToString()
            frmFeePayment.Update_record.Enabled = True
            frmFeePayment.Delete.Enabled = True
            frmFeePayment.btnSave.Enabled = False
            frmFeePayment.Print.Enabled = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

    Private Sub TabControl1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.Click
        cmbHostelerName.Text = ""
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
        ComboBox1.Text = ""
        DataGridView5.DataSource = Nothing
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        ComboBox1.Text = ""
        DataGridView5.DataSource = Nothing
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Try
            If Len(Trim(ComboBox1.Text)) = 0 Then
                MessageBox.Show("Please select hosteler name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbHostelerName1.Focus()
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand(" SELECT FeePayment.FeePaymentID as [Fee Payment ID ], FeePayment.HostelerID as [Hosteler ID], Hostelers.HostelerName as [Hosteler Name], RoomAllotment.Hostelname as [Hostel Name], RoomAllotment.RoomNo as [Room No], FeePayment.AcadYear as [Academic year],  FeePayment.ServiceCharges as [Total], FeePayment.PayableCharges as [Prev Due], FeePayment.TotalPaid as [Total Paid],  FeePayment.DuePayment as [Present Due],  FeePayment.PaymentDate as [Payment date],  FeePayment.BankChallanNumber as [Bank Challan No],  FeePayment.ExtraCharges as [Extra Charges],  FeePayment.Fine  as [Fine] FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where  HostelerName='" & ComboBox1.Text & "' and FeePayment.DuePayment > 0 order by HostelerName,PaymentDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "RoomAllotment")
            myDA.Fill(myDataSet, "FeePayment")
            DataGridView5.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView5.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
            DataGridView5.DataSource = myDataSet.Tables("FeePayment").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
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

    Private Sub DataGridView5_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView5.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView5.SelectedRows(0)
            Me.Hide()
            frmFeePayment.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmFeePayment.txtFeePaymentID.Text = dr.Cells(0).Value.ToString()
            frmFeePayment.cmbHostelerID.Text = dr.Cells(1).Value.ToString()
            frmFeePayment.txtHostelerName.Text = dr.Cells(2).Value.ToString()
            frmFeePayment.txtBranch.Text = dr.Cells(3).Value.ToString()
            frmFeePayment.txtRoomNo.Text = dr.Cells(4).Value.ToString()
            frmFeePayment.cmbAcadYear.Text = dr.Cells(5).Value.ToString()
            frmFeePayment.txtServiceCharges.Text = dr.Cells(6).Value.ToString()
            frmFeePayment.txtTotalCharges.Text = dr.Cells(7).Value.ToString()
            frmFeePayment.txtTotalPaid.Text = dr.Cells(8).Value.ToString()
            frmFeePayment.txtDueCharges.Text = dr.Cells(9).Value.ToString()
            frmFeePayment.dtpPaymentDate.Text = dr.Cells(10).Value.ToString()
            frmFeePayment.txtBankChallanNumber.Text = dr.Cells(11).Value.ToString()
            ' frmFeePayment.txtExtraCharges.Text = dr.Cells(12).Value.ToString()
            ' frmFeePayment.txtFine.Text = dr.Cells(13).Value.ToString()
            frmFeePayment.Update_record.Enabled = True
            frmFeePayment.Delete.Enabled = True
            frmFeePayment.btnSave.Enabled = False
            frmFeePayment.Print.Enabled = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView5_RowPostPaint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView5.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView5.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView5.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub btnViewReport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnViewReport.Click
        Try
            If Len(Trim(cmbHostelerName.Text)) = 0 Then
                MessageBox.Show("Please select hosteler name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbHostelerName.Focus()
                Exit Sub
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptFeePayment 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            'MyCommand.CommandText = "SELECT FeePayment.FeePaymentID, FeePayment.HostelerID, FeePayment.ServiceCharges,FeePayment.PayableCharges, FeePayment.AcadYear,FeePayment.PaymentDate, FeePayment.TotalPaid, FeePayment.DuePayment,FeePayment.BankChallanNumber, Hostelers.HostelerID AS Expr1, Hostelers.HostelerName,Hostelers.RoomNo, Hostelers.HostelName, Hostelers.DateOfJoining, Hostelers.FatherName, Hostelers.MobNo1, Hostelers.Phone1, Hostelers.MotherName, Hostelers.MobNo2, Hostelers.City, Hostelers.Address, Hostelers.Email, Hostelers.ContactNo,Hostelers.InstOfcDetails, Hostelers.Phone2, Hostelers.Agreement, Hostelers.GuardianName, Hostelers.GuardianAddress, Hostelers.MobNo3,Hostelers.Phone3, Hostelers.Photo, Hostelers.DocsPic, Hostelers.CompletionDate FROM FeePayment,Hostelers where FeePayment.HostelerID = Hostelers.HostelerID and HostelerName='" & cmbHostelerName.Text & "' order by HostelerName,PaymentDate"
            MyCommand.CommandText = "SELECT distinct FeePayment.FeePaymentID, FeePayment.HostelerID, Hostelers.HostelerName, Hostelers.Hostelname, RoomAllotment.RoomNo, FeePayment.AcadYear ,  FeePayment.ServiceCharges , FeePayment.PayableCharges, FeePayment.TotalPaid,  FeePayment.DuePayment,  FeePayment.PaymentDate,  FeePayment.BankChallanNumber,  FeePayment.ExtraCharges,  FeePayment.Fine  FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where  Hostelers.HostelerName='" & cmbHostelerName.Text & "' and FeePayment.AcadYear=RoomAllotment.AcadYear  order by HostelerName,PaymentDate asc"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Hostelers")
            myDA.Fill(myDS, "FeePayment")
            myDA.Fill(myDS, "RoomAllotment")
            rpt.SetDataSource(myDS)
            frmFeePaymentReport.CrystalReportViewer1.ReportSource = rpt
            frmFeePaymentReport.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub



End Class