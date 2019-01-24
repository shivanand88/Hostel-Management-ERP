Imports System.Data.OleDb
Public Class frmTransaction
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"

    Sub Reset()
        txtEmployeeName.Text = ""
        txtTransactionAmount.Text = ""
        dtpTransactionDate.Text = Today
        cmbTransactionType.SelectedIndex = -1
        cmbTransDetails.Text = ""
        cmbMonth.Text = ""
        txtAmountReceived.Text = ""
        txtDueAmount.Text = ""
        txtEmployeeName.Focus()
        btnDelete.Enabled = False
        btnUpdate_record.Enabled = False
        btnSave.Enabled = True
        txtAmountReceived.ReadOnly = True
    End Sub
    Private Sub txtAdvance_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTransactionAmount.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnNewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewRecord.Click
        Reset()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtEmployeeName.Text)) = 0 Then
            MessageBox.Show("Please enter employee/party name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtEmployeeName.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbTransactionType.Text)) = 0 Then
            MessageBox.Show("Please select transaction type", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbTransactionType.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbMonth.Text)) = 0 Then
            MessageBox.Show("Please select month", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbMonth.Focus()
            Exit Sub
        End If
        If Len(Trim(txtTransactionAmount.Text)) = 0 Then
            MessageBox.Show("Please enter transaction amount", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtTransactionAmount.Focus()
            Exit Sub
        End If
        Try

            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "insert into Trans(Employee_Party_Name,TransactionType,TransactionMonth,TransactionDate,TransactionAmount,AmountReceived,DueAmount,TransactionDetails) VALUES ('" & txtEmployeeName.Text & "','" & cmbTransactionType.Text & "','" & cmbMonth.Text & "',#" & dtpTransactionDate.Text & "#," & txtTransactionAmount.Text & ",0," & Val(txtTransactionAmount.Text) - Val(txtAmountReceived.Text) & ",'" & cmbTransDetails.Text & "')"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            Autocomplete()
            fillTransDetails()
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub Autocomplete()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim cmd As New OleDbCommand("SELECT distinct Employee_Party_Name FROM Trans", con)
            Dim ds As New DataSet()
            Dim da As New OleDbDataAdapter(cmd)
            da.Fill(ds, "Transaction")
            Dim col As New AutoCompleteStringCollection()
            Dim i As Integer = 0
            For i = 0 To ds.Tables(0).Rows.Count - 1
                col.Add(ds.Tables(0).Rows(i)("Employee_Party_Name").ToString())
            Next
            txtEmployeeName.AutoCompleteSource = AutoCompleteSource.CustomSource
            txtEmployeeName.AutoCompleteCustomSource = col
            txtEmployeeName.AutoCompleteMode = AutoCompleteMode.Suggest
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub frmTransaction_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub

    Private Sub frmTransaction_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Autocomplete()
        fillTransDetails()
    End Sub

    Private Sub btnUpdate_record_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate_record.Click
        Try
            If (txtAmountReceived.Text = "") Then
                txtAmountReceived.Text = 0
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "update Trans set Employee_Party_Name='" & txtEmployeeName.Text & "',TransactionType='" & cmbTransactionType.Text & "',TransactionMonth='" & cmbMonth.Text & "',TransactionDate=#" & dtpTransactionDate.Text & "#,TransactionAmount=" & txtTransactionAmount.Text & ",AmountReceived= " & txtAmountReceived.Text & ",TransactionDetails='" & cmbTransDetails.Text & "',DueAmount=" & txtDueAmount.Text & " where ID= " & txtTransactionID.Text & ""
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate_record.Enabled = False
            Autocomplete()
            fillTransDetails()
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            If MessageBox.Show("Do you really want to delete this record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                DeleteRecord()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub DeleteRecord()
        Try
            Dim RowsAffected As Integer = 0
            con = New OleDbConnection(cs)
            con.Open()
            Dim cq As String = "delete from Trans where ID = " & txtTransactionID.Text & ""
            cmd = New OleDbCommand(cq)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Autocomplete()
                fillTransDetails()
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Reset()
        frmTransactionRecord1.fillEmployeeName()
        frmTransactionRecord1.DataGridView3.DataSource = Nothing
        frmTransactionRecord1.DateFrom.Text = Today
        frmTransactionRecord1.DateTo.Text = Today
        frmTransactionRecord1.DataGridView2.DataSource = Nothing
        frmTransactionRecord1.cmbEmployeeName.Text = ""
        frmTransactionRecord1.txtEmployeeName.Text = ""
        frmTransactionRecord1.DataGridView1.DataSource = Nothing
        frmTransactionRecord1.Show()
    End Sub
    Sub fillTransDetails()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct RTRIM(TransactionDetails) FROM trans", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbTransDetails.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbTransDetails.Items.Add(drow(0).ToString())
            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtAmountReceived_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmountReceived.TextChanged
        txtDueAmount.Text = Val(txtTransactionAmount.Text) - Val(txtAmountReceived.Text)
    End Sub

    Private Sub txtAmountReceived_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtAmountReceived.Validating
        If Val(txtAmountReceived.Text) > Val(txtTransactionAmount.Text) Then
            MessageBox.Show("Received/Returned amount is more than Transaction amount", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtAmountReceived.Text = ""
            txtAmountReceived.Focus()
        End If
    End Sub
End Class