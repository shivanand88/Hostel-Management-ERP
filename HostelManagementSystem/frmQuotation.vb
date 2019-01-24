Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Security.Cryptography
Imports System.Text
Public Class frmQuotation
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"
    Sub reset()
        cmbHostelName.SelectedIndex = -1
        txtStudentName.Text = ""
        txtContactNo.Text = ""
        txtChargesPerMonth1.Text = ""
        txtMonth1.Text = ""
        txtTotalCharges1.Text = ""
        txtChargesPerMonth2.Text = ""
        txtMonth2.Text = ""
        txtTotalCharges2.Text = ""
        txtGrandTotal.Text = ""
        txtQuotationID.Text = ""
        dtpQuotationDate.Text = Today
        btnSave.Enabled = True
        btnDelete.Enabled = False
        btnUpdate_record.Enabled = False
        btnPrint.Enabled = False
        cmbHostelName.Focus()
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
    Private Sub txtChargesPerMonth1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtChargesPerMonth1.TextChanged
        txtTotalCharges1.Text = Val(txtChargesPerMonth1.Text) * Val(txtMonth1.Text)
        txtGrandTotal.Text = Val(txtTotalCharges1.Text) + Val(txtTotalCharges2.Text)
    End Sub

    Private Sub txtMonth1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMonth1.TextChanged
        txtTotalCharges1.Text = Val(txtChargesPerMonth1.Text) * Val(txtMonth1.Text)
        txtGrandTotal.Text = Val(txtTotalCharges1.Text) + Val(txtTotalCharges2.Text)
    End Sub

    Private Sub txtChargesPerMonth2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtChargesPerMonth2.TextChanged
        txtTotalCharges2.Text = Val(txtChargesPerMonth2.Text) * Val(txtMonth2.Text)
        txtGrandTotal.Text = Val(txtTotalCharges1.Text) + Val(txtTotalCharges2.Text)
    End Sub

    Private Sub txtMonth2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMonth2.TextChanged
        txtTotalCharges2.Text = Val(txtChargesPerMonth2.Text) * Val(txtMonth2.Text)
        txtGrandTotal.Text = Val(txtTotalCharges1.Text) + Val(txtTotalCharges2.Text)
    End Sub
    Sub fillHostelName()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct RTRIM(HostelName) FROM hostel", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbHostelName.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbHostelName.Items.Add(drow(0).ToString())
            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmQuotation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillHostelName()
    End Sub

    Private Sub btnNewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewRecord.Click
        reset()
    End Sub

    Private Sub txtContactNo_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtContactNo.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtChargesPerMonth1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtChargesPerMonth1.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtMonth1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMonth1.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtChargesPerMonth2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtChargesPerMonth2.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtMonth2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMonth2.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(cmbHostelName.Text)) = 0 Then
            MessageBox.Show("Please select hostel name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbHostelName.Focus()
            Exit Sub
        End If
        If Len(Trim(txtStudentName.Text)) = 0 Then
            MessageBox.Show("Please enter student name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtStudentName.Focus()
            Exit Sub
        End If
        If Len(Trim(txtContactNo.Text)) = 0 Then
            MessageBox.Show("Please enter contact no.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtContactNo.Focus()
            Exit Sub
        End If
        If Len(Trim(txtChargesPerMonth1.Text)) = 0 Then
            MessageBox.Show("Please enter charges per month", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtChargesPerMonth1.Focus()
            Exit Sub
        End If
        If Len(Trim(txtMonth1.Text)) = 0 Then
            MessageBox.Show("Please enter no. of months", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtMonth1.Focus()
            Exit Sub
        End If
        If Len(Trim(txtChargesPerMonth2.Text)) = 0 Then
            MessageBox.Show("Please enter charges per month", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtChargesPerMonth2.Focus()
            Exit Sub
        End If
        If Len(Trim(txtMonth2.Text)) = 0 Then
            MessageBox.Show("Please enter no. of months", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtMonth2.Focus()
            Exit Sub
        End If
        Try
            txtQuotationID.Text = "Q-" & GetUniqueKey(6)
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "insert into Quotation(QuotationID, Hostelname, StudentName, ContactNo, ChargesPerMonth1, NoOfMonth1, TotalCharges1, ChargesPerMonth2, NoOfMonth2,TotalCharges2, GrandTotal, QuotationDate) VALUES ('" & txtQuotationID.Text & "','" & cmbHostelName.Text & "','" & txtStudentName.Text & "','" & txtContactNo.Text & "'," & txtChargesPerMonth1.Text & "," & txtMonth1.Text & "," & txtTotalCharges1.Text & "," & txtChargesPerMonth2.Text & "," & txtMonth2.Text & "," & txtTotalCharges2.Text & "," & txtGrandTotal.Text & ",#" & dtpQuotationDate.Text & "#)"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            btnPrint.Enabled = True
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
            Dim cq As String = "delete from Quotation where QuotationID = '" & txtQuotationID.Text & "'"
            cmd = New OleDbCommand(cq)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                reset()
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
            Dim cb As String = "update Quotation set Hostelname='" & cmbHostelName.Text & "', StudentName='" & txtStudentName.Text & "', ContactNo='" & txtContactNo.Text & "', ChargesPerMonth1=" & txtChargesPerMonth1.Text & ", NoOfMonth1=" & txtMonth1.Text & ", TotalCharges1=" & txtTotalCharges1.Text & ", ChargesPerMonth2=" & txtChargesPerMonth2.Text & ", NoOfMonth2=" & txtMonth2.Text & ",TotalCharges2=" & txtTotalCharges2.Text & ", GrandTotal=" & txtGrandTotal.Text & ", QuotationDate=#" & dtpQuotationDate.Text & "# where QuotationID='" & txtQuotationID.Text & "'"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate_record.Enabled = False
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.reset()
        frmQuotationRecord.GetData()
        frmQuotationRecord.Show()
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
           
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptQuotation 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New Quotation_DBDataSet 'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT Hostel.HostelName, Hostel.Hostel_Address, Hostel.Hostel_Phone, Hostel.ManagedBy, Hostel.Hostel_ContactNo, Quotation.QuotationID,Quotation.Hostelname AS Expr1, Quotation.StudentName, Quotation.ContactNo, Quotation.ChargesPerMonth1, Quotation.NoOfMonth1, Quotation.TotalCharges1, Quotation.ChargesPerMonth2, Quotation.NoOfMonth2, Quotation.TotalCharges2, Quotation.GrandTotal, Quotation.QuotationDate FROM (Hostel INNER JOIN Quotation ON Hostel.HostelName = Quotation.Hostelname) where Quotation.QuotationID='" & txtQuotationID.Text & "'"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Hostel")
            myDA.Fill(myDS, "Quotation")
            rpt.SetDataSource(myDS)
            frmQuotationReport.CrystalReportViewer1.ReportSource = rpt
            frmQuotationReport.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub
End Class