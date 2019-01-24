Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class frmTransactionRecord1
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"
    Sub fillEmployeeName()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct RTRIM(Employee_Party_Name) FROM Trans", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbEmployeeName.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbEmployeeName.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub cmbEmployeeName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbEmployeeName.SelectedIndexChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT ID as [Transaction ID], Employee_Party_name as [Employee/Party Name],TransactionType as [Transaction Type],TransactionDetails as [Transaction Details],TransactionMonth as [Month],TransactionDate as [Transaction Date],TransactionAmount as [Transaction Amount],AmountReceived as [Amount Received/Returned],DueAmount as [Due Amount] from Trans where Employee_Party_Name='" & cmbEmployeeName.Text & "' order by Employee_Party_Name,TransactionDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Trans")
            DataGridView1.DataSource = myDataSet.Tables("Trans").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmAdvanceEntryRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillEmployeeName()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        cmbEmployeeName.Text = ""
        txtEmployeeName.Text = ""
        DataGridView1.DataSource = Nothing
    End Sub

    Private Sub txtEmployeeName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEmployeeName.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT ID as [Transaction ID], Employee_Party_name as [Employee/Party Name],TransactionType as [Transaction Type],TransactionDetails as [Transaction Details],TransactionMonth as [Month],TransactionDate as [Transaction Date],TransactionAmount as [Transaction Amount],AmountReceived as [Amount Received/Returned],DueAmount as [Due Amount] from Trans where Employee_Party_Name like '" & txtEmployeeName.Text & "%' order by Employee_Party_Name,TransactionDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Trans")
            DataGridView1.DataSource = myDataSet.Tables("Trans").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DateFrom.Text = Today
        DateTo.Text = Today
        DataGridView2.DataSource = Nothing
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT ID as [Transaction ID], Employee_Party_name as [Employee/Party Name],TransactionType as [Transaction Type],TransactionDetails as [Transaction Details],TransactionMonth as [Month],TransactionDate as [Transaction Date],TransactionAmount as [Transaction Amount],AmountReceived as [Amount Received/Returned],DueAmount as [Due Amount] from Trans where TransactionDate between #" & DateFrom.Text & " # and #" & DateTo.Text & "# order by Employee_Party_Name,TransactionDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Trans")
            DataGridView2.DataSource = myDataSet.Tables("Trans").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT ID as [Transaction ID], Employee_Party_name as [Employee/Party Name],TransactionType as [Transaction Type],TransactionDetails as [Transaction Details],TransactionMonth as [Month],TransactionDate as [Transaction Date],TransactionAmount as [Transaction Amount],AmountReceived as [Amount Received/Returned],DueAmount as [Due Amount] from Trans order by Employee_Party_Name,TransactionDate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Trans")
            DataGridView3.DataSource = myDataSet.Tables("Trans").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        DataGridView3.DataSource = Nothing
    End Sub

    Private Sub TabControl1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.Click
        DataGridView3.DataSource = Nothing
        DateFrom.Text = Today
        DateTo.Text = Today
        DataGridView2.DataSource = Nothing
        cmbEmployeeName.Text = ""
        txtEmployeeName.Text = ""
        DataGridView1.DataSource = Nothing
    End Sub

    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            Me.Hide()
            frmTransaction.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmTransaction.txtTransactionID.Text = dr.Cells(0).Value.ToString()
            frmTransaction.txtEmployeeName.Text = dr.Cells(1).Value.ToString()
            frmTransaction.cmbTransactionType.Text = dr.Cells(2).Value.ToString()
            frmTransaction.cmbTransDetails.Text = dr.Cells(3).Value.ToString()
            frmTransaction.cmbMonth.Text = dr.Cells(4).Value.ToString()
            frmTransaction.dtpTransactionDate.Text = dr.Cells(5).Value.ToString()
            frmTransaction.txtTransactionAmount.Text = dr.Cells(6).Value.ToString()
            frmTransaction.txtAmountReceived.Text = dr.Cells(7).Value.ToString()
            frmTransaction.txtDueAmount.Text = dr.Cells(8).Value.ToString()
            frmTransaction.btnUpdate_record.Enabled = True
            frmTransaction.btnDelete.Enabled = True
            frmTransaction.btnSave.Enabled = False
            frmTransaction.txtAmountReceived.ReadOnly = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView2_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView2.SelectedRows(0)
            Me.Hide()
            frmTransaction.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmTransaction.txtTransactionID.Text = dr.Cells(0).Value.ToString()
            frmTransaction.txtEmployeeName.Text = dr.Cells(1).Value.ToString()
            frmTransaction.cmbTransactionType.Text = dr.Cells(2).Value.ToString()
            frmTransaction.cmbTransDetails.Text = dr.Cells(3).Value.ToString()
            frmTransaction.cmbMonth.Text = dr.Cells(4).Value.ToString()
            frmTransaction.dtpTransactionDate.Text = dr.Cells(5).Value.ToString()
            frmTransaction.txtTransactionAmount.Text = dr.Cells(6).Value.ToString()
            frmTransaction.txtAmountReceived.Text = dr.Cells(7).Value.ToString()
            frmTransaction.txtDueAmount.Text = dr.Cells(8).Value.ToString()
            frmTransaction.btnUpdate_record.Enabled = True
            frmTransaction.btnDelete.Enabled = True
            frmTransaction.btnSave.Enabled = False
            frmTransaction.txtAmountReceived.ReadOnly = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView3_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView3.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView3.SelectedRows(0)
            Me.Hide()
            frmTransaction.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmTransaction.txtTransactionID.Text = dr.Cells(0).Value.ToString()
            frmTransaction.txtEmployeeName.Text = dr.Cells(1).Value.ToString()
            frmTransaction.cmbTransactionType.Text = dr.Cells(2).Value.ToString()
            frmTransaction.cmbTransDetails.Text = dr.Cells(3).Value.ToString()
            frmTransaction.cmbMonth.Text = dr.Cells(4).Value.ToString()
            frmTransaction.dtpTransactionDate.Text = dr.Cells(5).Value.ToString()
            frmTransaction.txtTransactionAmount.Text = dr.Cells(6).Value.ToString()
            frmTransaction.txtAmountReceived.Text = dr.Cells(7).Value.ToString()
            frmTransaction.txtDueAmount.Text = dr.Cells(8).Value.ToString()
            frmTransaction.btnUpdate_record.Enabled = True
            frmTransaction.btnDelete.Enabled = True
            frmTransaction.btnSave.Enabled = False
            frmTransaction.txtAmountReceived.ReadOnly = False
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
End Class