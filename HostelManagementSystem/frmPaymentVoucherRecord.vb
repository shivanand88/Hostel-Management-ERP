Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmPaymentVoucherRecord
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"

    Private Sub cmbTransactionType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbTransactionType.SelectedIndexChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("Select VoucherNo as [Voucher No],ReceiverName as [Receiver's Name],TransactionType as [Transaction Type],ACNo as [A/C No],Amount as [Transaction Amount],PaymentDate as [Transaction Date] from Voucher where TransactionType='" & cmbTransactionType.Text & "' order by PaymentDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Voucher")
            DataGridView1.DataSource = myDataSet.Tables("Voucher").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("Select VoucherNo as [Voucher No],ReceiverName as [Receiver's Name],TransactionType as [Transaction Type],ACNo as [A/C No],Amount as [Transaction Amount],PaymentDate as [Transaction Date] from Voucher where PaymentDate between #" & DateTimePicker2.Text & "# and #" & DateTimePicker1.Text & "#  order by PaymentDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Voucher")
            DataGridView1.DataSource = myDataSet.Tables("Voucher").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
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

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        cmbTransactionType.SelectedIndex = -1
        DateTimePicker1.Text = Today
        DateTimePicker2.Text = Today
        DataGridView1.DataSource = Nothing
    End Sub

  
    Private Sub frmPaymentVoucherRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            Me.Hide()
            frmPaymentVoucher.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmPaymentVoucher.txtVoucherNo.Text = dr.Cells(0).Value.ToString()
            frmPaymentVoucher.txtReceiverName.Text = dr.Cells(1).Value.ToString()
            frmPaymentVoucher.cmbTransactionType.Text = dr.Cells(2).Value.ToString()
            frmPaymentVoucher.txtACNo.Text = dr.Cells(3).Value.ToString()
            frmPaymentVoucher.txtAmount.Text = dr.Cells(4).Value.ToString()
            frmPaymentVoucher.dtpPaymentDate.Text = dr.Cells(5).Value.ToString()
            frmPaymentVoucher.btnUpdate_record.Enabled = True
            frmPaymentVoucher.btnDelete.Enabled = True
            frmPaymentVoucher.btnSave.Enabled = False
            frmPaymentVoucher.btnPrint.Enabled = True
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If (cmbTransactionType.Text = "") Then
                MessageBox.Show("Please select transaction type", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptVouchers 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New Voucher_DBDataSet 'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT  VoucherNo, ReceiverName, TransactionType, ACNo, Amount, PaymentDate FROM Voucher where TransactionType='" & cmbTransactionType.Text & "' order by PaymentDate"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Voucher")
            rpt.SetDataSource(myDS)
            frmPaymentVouchersReport.CrystalReportViewer1.ReportSource = rpt
            frmPaymentVouchersReport.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try

            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptVouchers 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New Voucher_DBDataSet 'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT  VoucherNo, ReceiverName, TransactionType, ACNo, Amount, PaymentDate FROM Voucher where Paymentdate between #" & DateTimePicker2.Text & "# and #" & DateTimePicker1.Text & "# order by PaymentDate"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Voucher")
            rpt.SetDataSource(myDS)
            frmPaymentVouchersReport.CrystalReportViewer1.ReportSource = rpt
            frmPaymentVouchersReport.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class