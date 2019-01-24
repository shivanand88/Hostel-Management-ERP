Imports System.Data.OleDb
Public Class frmPurchaseInventory
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"

    Sub Reset()
        cmbProductName.Text = ""
        txtQuantity.Text = ""
        cmbUnit.SelectedIndex = -1
        txtUnitPrice.Text = ""
        txtTotalPrice.Text = ""
        cmbCategory.Text = ""
        cmbPartyName.Text = ""
        cmbTransType.SelectedIndex = -1
        dtpPurchaseDate.Text = Today
        btnDelete.Enabled = False
        btnDelete.Enabled = False
        btnSave.Enabled = True
        cmbProductName.Focus()
    End Sub
    Public Sub DeleteRecord()
        Try
            Dim RowsAffected As Integer = 0
            con = New OleDbConnection(cs)
            con.Open()
            Dim cq As String = "delete from PurchasedInventory where ID = " & txtPurchaseID.Text & ""
            cmd = New OleDbCommand(cq)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                fillProductName()
                fillCategory()
                fillPartyname()
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

    Private Sub frmPurchaseInventory_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillProductName()
        fillCategory()
        fillPartyname()
    End Sub

    Private Sub btnNewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewRecord.Click
        Reset()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(cmbProductName.Text)) = 0 Then
            MessageBox.Show("Please select product name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbProductName.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbCategory.Text)) = 0 Then
            MessageBox.Show("Please enter Category", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbCategory.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbTransType.Text)) = 0 Then
            MessageBox.Show("Please select transaction type", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbTransType.Focus()
            Exit Sub
        End If
        If Len(Trim(txtQuantity.Text)) = 0 Then
            MessageBox.Show("Please select quantity", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtQuantity.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbUnit.Text)) = 0 Then
            MessageBox.Show("Please select unit", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbUnit.Focus()
            Exit Sub
        End If
        If Len(Trim(txtUnitPrice.Text)) = 0 Then
            MessageBox.Show("Please enter unit price", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtUnitPrice.Focus()
            Exit Sub
        End If
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "insert into PurchasedInventory(ProductName,Quantity,Unit,Price,PurchaseDate,TotalPrice,Category,TransactionType,PartyName) VALUES ('" & cmbProductName.Text & "'," & txtQuantity.Text & ",'" & cmbUnit.Text & "'," & txtUnitPrice.Text & ",#" & dtpPurchaseDate.Text & "#," & txtTotalPrice.Text & ",'" & cmbCategory.Text & "','" & cmbTransType.Text & "','" & cmbPartyName.Text & "')"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            fillProductName()
            fillCategory()
            fillPartyname()
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnUpdate_record_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate_record.Click
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "update PurchasedInventory set ProductName='" & cmbProductName.Text & "',Quantity=" & txtQuantity.Text & ", Unit='" & cmbUnit.Text & "',Price=" & txtUnitPrice.Text & ",PurchaseDate=#" & dtpPurchaseDate.Text & "#,TotalPrice=" & txtTotalPrice.Text & ",Category='" & cmbCategory.Text & "',TransactionType='" & cmbTransType.Text & "',PartyName='" & cmbPartyName.Text & "' where ID= " & txtPurchaseID.Text & ""
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate_record.Enabled = False
            fillProductName()
            fillCategory()
            fillPartyname()
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
            For Each drow As DataRow In dtable.Rows
                cmbProductName.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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
    Private Sub txtprice_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtUnitPrice.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Reset()
        frmPurchasedInventoryRecord1.fillCategory()
        frmPurchasedInventoryRecord1.fillProductName()
        frmPurchasedInventoryRecord1.fillPartyname()
        frmPurchasedInventoryRecord1.cmbProductName1.Text = ""
        frmPurchasedInventoryRecord1.DateTimePicker1.Text = Today
        frmPurchasedInventoryRecord1.DateTimePicker2.Text = Today
        frmPurchasedInventoryRecord1.DataGridView3.DataSource = Nothing
        frmPurchasedInventoryRecord1.DataGridView2.DataSource = Nothing
        frmPurchasedInventoryRecord1.DateFrom.Text = Today
        frmPurchasedInventoryRecord1.DateTo.Text = Today
        frmPurchasedInventoryRecord1.GroupBox7.Visible = False
        frmPurchasedInventoryRecord1.cmbProductName.Text = ""
        frmPurchasedInventoryRecord1.txtProductName.Text = ""
        frmPurchasedInventoryRecord1.DataGridView1.DataSource = Nothing
        frmPurchasedInventoryRecord1.GroupBox2.Visible = False
        frmPurchasedInventoryRecord1.GroupBox11.Visible = False
        frmPurchasedInventoryRecord1.cmbCategory.Text = ""
        frmPurchasedInventoryRecord1.txtCategory.Text = ""
        frmPurchasedInventoryRecord1.DataGridView4.DataSource = Nothing
        frmPurchasedInventoryRecord1.GroupBox15.Visible = False
        frmPurchasedInventoryRecord1.cmbTransType.SelectedIndex = -1
        frmPurchasedInventoryRecord1.DataGridView5.DataSource = Nothing
        frmPurchasedInventoryRecord1.GroupBox18.Visible = False
        frmPurchasedInventoryRecord1.cmbPartyName.Text = ""
        frmPurchasedInventoryRecord1.txtPartyName.Text = ""
        frmPurchasedInventoryRecord1.DataGridView6.DataSource = Nothing
        frmPurchasedInventoryRecord1.GroupBox19.Visible = False
        frmPurchasedInventoryRecord1.Show()
    End Sub

    Private Sub txtQuantity_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtQuantity.KeyPress
        Dim keyChar = e.KeyChar
        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtQuantity.Text
            Dim selectionStart = Me.txtQuantity.SelectionStart
            Dim selectionLength = Me.txtQuantity.SelectionLength
            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)
            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an integer that is longer than 16 digits.
                e.Handled = True
            ElseIf Double.TryParse(text, New Double) AndAlso text.IndexOf("."c) < text.Length - 3 Then
                'Reject a real number with two many decimal places.
                e.Handled = False
            End If
        Else
            'Reject all other characters.
            e.Handled = True
        End If
    End Sub

    Private Sub txtprice_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUnitPrice.TextChanged
        txtTotalPrice.Text = CInt(Val(txtQuantity.Text) * Val(txtUnitPrice.Text))
    End Sub

    Private Sub txtQuantity_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtQuantity.TextChanged
        txtTotalPrice.Text = CInt(Val(txtQuantity.Text) * Val(txtUnitPrice.Text))
    End Sub
End Class