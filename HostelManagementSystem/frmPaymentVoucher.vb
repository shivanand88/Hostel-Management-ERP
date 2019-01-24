Imports System.Data.OleDb
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Public Class frmPaymentVoucher


    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"
    Sub Reset()
        txtAmount.Text = ""
        btnDelete.Enabled = False
        btnUpdate_record.Enabled = False
        btnSave.Enabled = True
        cmbTransactionType.SelectedIndex = -1
        txtReceiverName.Text = ""
        txtVoucherNo.Text = ""
        txtACNo.Text = ""
        dtpPaymentDate.Text = Today
        btnPrint.Enabled = False
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtReceiverName.Text)) = 0 Then
            MessageBox.Show("Please enter receiver's name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtReceiverName.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbTransactionType.Text)) = 0 Then
            MessageBox.Show("Please select transaction type", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbTransactionType.Focus()
            Exit Sub
        End If

        If Len(Trim(txtAmount.Text)) = 0 Then
            MessageBox.Show("Please enter transaction amount", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtAmount.Focus()
            Exit Sub
        End If
        Try

            txtVoucherNo.Text = "V-" & GetUniqueKey(5)
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "insert into Voucher(VoucherNo,ReceiverName,TransactionType,ACNo,Amount,PaymentDate) VALUES ('" & txtVoucherNo.Text & "','" & txtReceiverName.Text & "','" & cmbTransactionType.Text & "','" & txtACNo.Text & "'," & txtAmount.Text & ",#" & dtpPaymentDate.Text & "#)"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            con.Close()
            btnSave.Enabled = False
            btnPrint.Enabled = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Shared Function GetUniqueKey(ByVal maxSize As Integer) As String
        Dim chars As Char() = New Char(61) {}
        chars = "123456789".ToCharArray()
        Dim data As Byte() = New Byte(0) {}
        Dim crypto As New RNGCryptoServiceProvider()
        crypto.GetNonZeroBytes(data)
        data = New Byte(maxSize - 1) {}
        crypto.GetNonZeroBytes(data)
        Dim result As New StringBuilder(maxSize)
        For Each b As Byte In data
            result.Append(chars(b Mod (chars.Length)))
        Next
        Return result.ToString()
    End Function

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
            Dim cq As String = "delete from Voucher where VoucherNo= '" & txtVoucherNo.Text & "'"
            cmd = New OleDbCommand(cq)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then

                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
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

    Private Sub btnUpdate_record_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate_record.Click
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "update Voucher set ReceiverName='" & txtReceiverName.Text & "',TransactionType='" & cmbTransactionType.Text & "',ACNo='" & txtACNo.Text & "',Amount=" & txtAmount.Text & ",PaymentDate=#" & dtpPaymentDate.Text & "# where VoucherNo='" & txtVoucherNo.Text & "'"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            con.Close()
            btnUpdate_record.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnNewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewRecord.Click
        Reset()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        frmPaymentVoucherRecord.Show()
        Me.Reset()
        frmPaymentVoucherRecord.DateTimePicker1.Text = Today
        frmPaymentVoucherRecord.DateTimePicker2.Text = Today
        frmPaymentVoucherRecord.cmbTransactionType.SelectedIndex = -1
        frmPaymentVoucherRecord.DataGridView1.DataSource = Nothing
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try

            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptPaymentVoucher 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT  VoucherNo, ReceiverName, TransactionType, ACNo, Amount, PaymentDate FROM Voucher where VoucherNo='" & txtVoucherNo.Text & "'"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Voucher")
            rpt.SetDataSource(myDS)
            frmPaymentVoucherReport.CrystalReportViewer1.ReportSource = rpt
            frmPaymentVoucherReport.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub txtAmount_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAmount.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtACNo_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtACNo.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub frmPaymentVoucher_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub

    Private Sub frmPaymentVoucher_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FillCheckListBox()

    End Sub
    Private Sub FillCheckListBox()

        Dim CN As New OleDbConnection(cs)
        CN.Open()
        adp = New OleDbDataAdapter()
        adp.SelectCommand = New OleDbCommand("SELECT Distinct HostelName FROM Hostel", con)
        ds = New DataSet("ds")
        adp.Fill(ds)
        dtable = ds.Tables(0)
        lstHostel.Items.Clear()

        For Each drow As DataRow In dtable.Rows
            lstHostel.Items.Add(drow(0).ToString())
        Next




    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click


        If lstHostel.Items.Count > 0 Then
            For i = 0 To lstHostel.CheckedItems.Count - 1
                MsgBox(lstHostel.CheckedItems(i))
            Next
        End If

    End Sub

End Class