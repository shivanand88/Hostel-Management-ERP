Imports System.Data.OleDb
Imports System.Security.Cryptography
Imports System.Text
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class frmDueFeePayment
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"

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

    Private Sub frmDueFeePayment_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub
    Private Sub frmDueFeePayment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            'GroupBox2.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID],HostelerName as [HostelerName],HostelName as [Hostel Name],RoomNo as [Room No],TotalDueAmount as [Due Amount] from DueAmount where TotalDueAmount > '0' order by Hostelername", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            'myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "DueAmount")
            'DataGridView1.DataSource = myDataSet.Tables("FeePayment").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("DueAmount").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Sub Reset()
        txtBranch.Text = ""
        txtFeePaymentID.Text = ""
        txtHostelerName.Text = ""
        txtServiceCharges.Text = ""
        txtRoomNo.Text = ""
        txtTotalCharges.Text = ""
        txtTotalPaid.Text = ""
        txtBankChallanNumber.Text = ""
        txtHostelerID.Text = ""
        cmbMonth.Text = ""
        txtFine.Text = 0
        txtDueCharges.Text = ""
        txtBankChallanNumber.Text = ""
        dtpPaymentDate.Text = Today
        btnSave.Enabled = True
        Delete.Enabled = False
        Update_record.Enabled = False
        Print.Enabled = False
        
    End Sub

    Private Sub cmbHostelerID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtTotalCharges.Text = ""
        txtFine.Text = 0
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT HostelerID,HostelerName,Hostelname,RoomNo,TotalDueAmount FROM DueAmount WHERE HostelerID= '" & txtHostelerID.Text & "'  "
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtHostelerID.Text = rdr.GetString(0)
                txtHostelerName.Text = rdr.GetString(1)
                txtBranch.Text = rdr.GetString(2)
                txtRoomNo.Text = rdr.GetString(3)
                txtServiceCharges.Text = rdr.GetValue(4)

            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            cmbMonth.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub dataGridView1_RowHeaderMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dataGridView1.RowHeaderMouseClick
        txtTotalCharges.Text = ""
        txtFine.Text = 0
        Try
            Dim dr As DataGridViewRow = dataGridView1.SelectedRows(0)
            txtHostelerID.Text = dr.Cells(0).Value.ToString
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT HostelerID,HostelerName,Hostelname,RoomNo,TotalDueAmount FROM DueAmount WHERE HostelerID= '" & txtHostelerID.Text & "' "
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtHostelerID.Text = rdr.GetString(0)
                txtHostelerName.Text = rdr.GetString(1)
                txtBranch.Text = rdr.GetString(2)
                txtRoomNo.Text = rdr.GetString(3)
                txtServiceCharges.Text = rdr.GetValue(4)

            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            cmbMonth.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub NewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewRecord.Click
        Reset()
    End Sub
    Sub fillHostelerID()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct RTRIM(HostelerID) FROM Hostelers", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            txtHostelerID.Clear()
            For Each drow As DataRow In dtable.Rows
                'txtHostelerID.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtTotalPaid_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTotalPaid.TextChanged
        txtDueCharges.Text = CInt((Val(txtTotalCharges.Text) + Val(txtFine.Text)) - Val(txtTotalPaid.Text))

    End Sub


    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtHostelerID.Text)) = 0 Then
            MessageBox.Show("Please select hosteler id", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtHostelerID.Focus()
            Exit Sub
        End If
        'Month selection is temporarily in activated
        'If Len(Trim(cmbMonth.Text)) = 0 Then
        'MessageBox.Show("Please select month", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'cmbMonth.Focus()
        'Exit Sub
        'End If
        If (txtTotalPaid.Text) = "" Then
            txtTotalPaid.Text = 0
        End If

        If (txtFine.Text) = "" Then
            txtFine.Text = 0
        End If
        Try
            txtFeePaymentID.Text = "F-" & GetUniqueKey(7)
            con = New OleDbConnection(cs)
            con.Open()
            '''''''''''''''''  "insert into FeePayment(FeePaymentID,HostelerID,ServiceCharges,ExtraCharges,FeeMonth,PaymentDate,TotalPaid,DuePayment,Fine,BankChallanNumber) VALUES ('" & txtFeePaymentID.Text & "','" & cmbHostelerID.Text & "'," & txtServiceCharges.Text & "," & txtExtraCharges.Text & ",'" & cmbMonth.Text & "',#" & dtpPaymentDate.Text & "#," & txtTotalPaid.Text & "," & txtDueCharges.Text & "," & txtFine.Text & "," & txtBankChallanNumber.Text & ")"
            Dim cb As String = "insert into FeePayment(FeePaymentID,HostelerID,ServiceCharges,ExtraCharges,FeeMonth,PaymentDate,TotalPaid,DuePayment,Fine,BankChallanNumber) VALUES ('" & txtFeePaymentID.Text & "','" & txtHostelerID.Text & "'," & txtServiceCharges.Text & "," & txtFine.Text & ",'" & cmbMonth.Text & "',#" & dtpPaymentDate.Text & "#," & txtTotalPaid.Text & "," & txtDueCharges.Text & "," & txtFine.Text & "," & txtBankChallanNumber.Text & ")"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            con.Close()
            con.Open()
            Dim cb1 As String = "update DueAmount set TotalDueAmount= '" & txtDueCharges.Text & "' where HostelerID='" & txtHostelerID.Text & "' "
            'Dim cb1 As String = "update FeePayment set HostelerID='" & txtHostelerID.Text & "',ServiceCharges=" & txtServiceCharges.Text & ",FeeMonth='" & cmbMonth.Text & "',PaymentDate=#" & dtpPaymentDate.Text & "#,TotalPaid=" & txtTotalPaid.Text & ",DuePayment=" & txtDueCharges.Text & ",Fine=" & txtFine.Text & " , BankChallanNumber=" & txtBankChallanNumber.Text & " where FeePaymentID='" & txtFeePaymentID.Text & "'"
            cmd = New OleDbCommand(cb1)
            cmd.Connection = con
            cmd.ExecuteReader()
            MessageBox.Show("Successfully Debited from Due", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            Print.Enabled = True
            con.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHosteler.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID],HostelerName as [HostelerName],HostelName as [Hostel Name],RoomNo as [Room No],TotalDueAmount as [Due Amount] from DueAmount where TotalDueAmount >'0' and HostelerName like '" & txtHosteler.Text & "%' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "DueAmount")
            dataGridView1.DataSource = myDataSet.Tables("DueAmount").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Update_record_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Update_record.Click
        Try

            If (txtFine.Text = "") Then
                txtFine.Text = 0
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "update FeePayment set HostelerID='" & txtHostelerID.Text & "',ServiceCharges=" & txtServiceCharges.Text & ",FeeMonth='" & cmbMonth.Text & "',PaymentDate=#" & dtpPaymentDate.Text & "#,TotalPaid=" & txtTotalPaid.Text & ",DuePayment=" & txtDueCharges.Text & ",Fine=" & txtFine.Text & " , BankChallanNumber=" & txtBankChallanNumber.Text & " where FeePaymentID='" & txtFeePaymentID.Text & "'"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Update_record.Enabled = False
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Delete.Click
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
            Dim cq As String = "delete from FeePayment where FeePaymentID = '" & txtFeePaymentID.Text & "'"
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

    Private Sub button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button2.Click
        Me.Reset()
        frmFeePaymentRecord1.fillHostelerName()
        frmFeePaymentRecord1.cmbHostelerName.Text = ""
        frmFeePaymentRecord1.txtHostelerName.Text = ""
        frmFeePaymentRecord1.DataGridView1.DataSource = Nothing
        frmFeePaymentRecord1.DateFrom.Text = Today
        frmFeePaymentRecord1.DateTo.Text = Today
        frmFeePaymentRecord1.DataGridView2.DataSource = Nothing
        frmFeePaymentRecord1.DateTimePicker1.Text = Today
        frmFeePaymentRecord1.DateTimePicker2.Text = Today
        frmFeePaymentRecord1.DataGridView3.DataSource = Nothing
        frmFeePaymentRecord1.cmbHostelerName1.Text = ""
        frmFeePaymentRecord1.DateTimePicker4.Text = Today
        frmFeePaymentRecord1.DateTimePicker3.Text = Today
        frmFeePaymentRecord1.DataGridView4.DataSource = Nothing
        frmFeePaymentRecord1.ComboBox1.Text = ""
        frmFeePaymentRecord1.DataGridView5.DataSource = Nothing
        frmFeePaymentRecord1.Show()
    End Sub



    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        Try

            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptDueFeePaymentReceipt 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT FeePaymentID,Hostelers.HostelerID,Hostelername,HostelName,FeeMonth,PaymentDate,RoomNo,ServiceCharges,TotalPaid,TotalDueAmount,BankChallanNumber from Hostelers,FeePayment,DueAmount where Hostelers.HostelerID=FeePayment.HostelerID and Hostelers.HostelerID=DueAmount.HostelerID and FeePaymentID='" & txtFeePaymentID.Text & "'"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Hostelers")
            myDA.Fill(myDS, "FeePayment")
            myDA.Fill(myDS, "DuePayment")
            rpt.SetDataSource(myDS)
            frmDueFeePaymentReceipt.CrystalReportViewer1.ReportSource = rpt
            frmDueFeePaymentReceipt.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub txtTotalPaid_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtTotalPaid.Validating
        If Val(txtTotalPaid.Text) > (Val(txtTotalCharges.Text) + Val(txtFine.Text)) Then
            MessageBox.Show("Total paid is more than total charges and fine", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtDueCharges.Text = ""
            txtTotalPaid.Text = ""
            txtTotalPaid.Focus()
        End If
    End Sub

    Private Sub txtServiceCharges_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtServiceCharges.TextChanged
        txtTotalCharges.Text = Val(txtServiceCharges.Text) + Val(txtFine.Text)
    End Sub

    
End Class