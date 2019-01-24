Imports System.Data.OleDb
Imports System.Security.Cryptography
Imports System.Text
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class frmFeePayment
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

    Private Sub frmFeePayment_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub
    

   
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Me.frmFeePayment_Load(New Object, New System.EventArgs)
        'btnRefreshForm.PerformClick()
    End Sub
    Private Sub btnRefreshForm_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRefreshForm.Click
        'fillHostelerID()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Hostelers.HostelerID, Hostelers.HostelerName, RoomAllotment.Hostelname, RoomAllotment.RoomNo,DueAmount.totalCharges, DueAmount.AcadYear, DueAmount.TotalDueAmount FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN DueAmount ON Hostelers.[HostelerID] = DueAmount.[HostelerID] where   DueAmount.Acadyear=RoomAllotment.AcadYear and DueAmount.TotalDueAmount > '0'  ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()

            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "RoomAllotment")
            myDA.Fill(myDataSet, "DueAmount")
            dataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            dataGridView1.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
            dataGridView1.DataSource = myDataSet.Tables("DueAmount").DefaultView

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
        cmbHostelerID.Text = ""
        cmbAcadYear.Text = ""
        txtDueCharges.Text = ""
        txtBankChallanNumber.Text = ""
        dtpPaymentDate.Text = Today


        ' fine and extra charges texboxes are extra in database  make them zero .
        txtFine.Text = 0
        txtExtraCharges.Text = 0
        dataGridView1.Update()
        btnSave.Enabled = True
        Delete.Enabled = False
        Update_record.Enabled = False
        Print.Enabled = False
    End Sub

    Private Sub cmbHostelerID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbHostelerID.SelectedIndexChanged
        txtExtraCharges.Text = 0
        txtFine.Text = 0
        txtServiceCharges.Enabled = False
        txtTotalCharges.Enabled = False
        cmbAcadYear.Enabled = False
        txtTotalPaid.Text = ""
        txtBankChallanNumber.Text = ""
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT Hostelers.HostelerName, RoomAllotment.Hostelname, RoomAllotment.RoomNo,DueAmount.TotalCharges, DueAmount.AcadYear, DueAmount.TotalDueAmount FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN DueAmount ON Hostelers.[HostelerID] = DueAmount.[HostelerID] where Hostelers.HostelerID = '" & cmbHostelerID.Text & "' and TotalDueAmount >'0' and  DueAmount.Acadyear=RoomAllotment.AcadYear"

            'cmd.CommandText = "SELECT Hostelers.HostelerName, Hostelers.HostelName, Hostelers.RoomNo, DueAmount.TotalCharges, Hostelers.AcadYear, DueAmount.TotalDueAmount FROM Hostelers INNER JOIN DueAmount ON Hostelers.[HostelerID] = DueAmount.[HostelerID]  where Hostelers.HostelerID = '" & cmbHostelerID.Text & "' and TotalDueAmount>'0'" 'group by hostelername,roomno,hostelname"
            'cmd.CommandText = "SELECT distinct HostelerName,RoomNo,hostelname,TotalCharges,TotalDueAmount FROM Hostelers,RegCharges,DueAmount WHERE  Hostelers.HostelerID= '" & cmbHostelerID.Text & "' and Hostelers.HostelerID=RegCharges.HostelerID and RegCharges.HostelerID=DueAmount.HostelerID " 'group by hostelername,roomno,hostelname"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtHostelerName.Text = rdr.GetString(0)
                txtBranch.Text = rdr.GetString(1)
                txtRoomNo.Text = rdr.GetString(2)
                txtServiceCharges.Text = rdr.GetValue(3)
                cmbAcadYear.Text = rdr.GetString(4)
                txtTotalCharges.Text = rdr.GetValue(5)
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            txtTotalPaid.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub dataGridView1_RowHeaderMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dataGridView1.RowHeaderMouseClick
        txtTotalCharges.Text = ""
        txtExtraCharges.Text = 0
        txtFine.Text = 0
        txtServiceCharges.Enabled = False
        txtTotalCharges.Enabled = False
        cmbAcadYear.Enabled = False
        txtTotalPaid.Text = ""
        txtBankChallanNumber.Text = ""
        Try
            Dim dr As DataGridViewRow = dataGridView1.SelectedRows(0)
            cmbHostelerID.Text = dr.Cells(0).Value.ToString
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT distinct Hostelers.HostelerName, RoomAllotment.Hostelname, RoomAllotment.RoomNo,DueAmount.TotalCharges, DueAmount.AcadYear, DueAmount.TotalDueAmount FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN DueAmount ON Hostelers.[HostelerID] = DueAmount.[HostelerID] where Hostelers.HostelerID = '" & cmbHostelerID.Text & "' and TotalDueAmount >'0' and  DueAmount.Acadyear=RoomAllotment.AcadYear"

            ' cmd.CommandText = "SELECT Hostelers.HostelerName, Hostelers.HostelName, Hostelers.RoomNo, DueAmount.TotalCharges, DueAmount.AcadYear, DueAmount.TotalDueAmount FROM Hostelers INNER JOIN DueAmount ON Hostelers.[HostelerID] = DueAmount.[HostelerID]  where Hostelers.HostelerID = '" & cmbHostelerID.Text & "' " 'group by hostelername,roomno,hostelname"
            'cmd.CommandText = "SELECT distinct HostelerName,RoomNo,hostelname,TotalCharges,TotalDueAmount FROM Hostelers,RegCharges,DueAmount WHERE  Hostelers.HostelerID= '" & cmbHostelerID.Text & "' and Hostelers.HostelerID=RegCharges.HostelerID and RegCharges.HostelerID=DueAmount.HostelerID " 'group by hostelername,roomno,hostelname"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtHostelerName.Text = rdr.GetString(0)
                txtBranch.Text = rdr.GetString(1)
                txtRoomNo.Text = rdr.GetString(2)
                txtServiceCharges.Text = rdr.GetValue(3)
                cmbAcadYear.Text = rdr.GetString(4)
                txtTotalCharges.Text = rdr.GetValue(5)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            txtTotalPaid.Focus()
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
            cmbHostelerID.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbHostelerID.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtTotalPaid_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTotalPaid.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub



    Private Sub txtTotalPaid_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTotalPaid.TextChanged
        txtDueCharges.Text = CInt((Val(txtTotalCharges.Text) + Val(txtFine.Text)) - Val(txtTotalPaid.Text))

    End Sub

    Private Sub txtTotalPaid_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtTotalPaid.Validating
        If Val(txtTotalPaid.Text) > (Val(txtTotalCharges.Text) + Val(txtFine.Text)) Then
            MessageBox.Show("Total paid is more than total charges", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtDueCharges.Text = ""
            txtTotalPaid.Text = ""
            txtTotalPaid.Focus()

        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click


        If Len(Trim(cmbHostelerID.Text)) = 0 Then
            MessageBox.Show("Please select hosteler id", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbHostelerID.Focus()
            Exit Sub
        End If
        If txtTotalPaid.Text = "" Then
            MessageBox.Show("Please Enter Amount", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtTotalPaid.Focus()
            Exit Sub
        End If
        'Month selection is temporarily in activated
        'If Len(Trim(cmbAcadYear.Text)) = 0 Then
        'MessageBox.Show("Please select month", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'cmbAcadYear.Focus()
        'Exit Sub
        'End If
        
        If (txtTotalPaid.Text) = "" Then
            txtTotalPaid.Text = ""
        End If
        If (txtExtraCharges.Text) = "" Then
            txtExtraCharges.Text = 0
        End If
        If (txtFine.Text) = "" Then
            txtFine.Text = 0
        End If
        If (txtBankChallanNumber.Text) = "" Then
            txtBankChallanNumber.Text = 0
        End If
        Try
            txtFeePaymentID.Text = "F-" & GetUniqueKey(7)
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "insert into FeePayment(FeePaymentID,HostelerID,ServiceCharges,ExtraCharges,AcadYear,PaymentDate,PayableCharges,TotalPaid,DuePayment,Fine,BankChallanNumber) VALUES ('" & txtFeePaymentID.Text & "','" & cmbHostelerID.Text & "'," & txtServiceCharges.Text & "," & txtExtraCharges.Text & ",'" & cmbAcadYear.Text & "',#" & dtpPaymentDate.Text & "#,'" & txtTotalCharges.Text & "','" & txtTotalPaid.Text & "','" & txtDueCharges.Text & "','" & txtFine.Text & "','" & txtBankChallanNumber.Text & "')"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            Dim cb1 As String = "update DueAmount set TotalDueAmount='" & txtDueCharges.Text & "'where HostelerID='" & cmbHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "'"
            cmd = New OleDbCommand(cb1)
            cmd.Connection = con
            cmd.ExecuteReader()
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            Print.Enabled = True
            btnRefreshForm.PerformClick()
            con.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtHosteler_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHosteler.GotFocus
        txtUSN.Text = ""
        Timer2.Enabled = False
    End Sub
    Private Sub txtUSN_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUSN.GotFocus
        Timer2.Enabled = False
        txtHosteler.Text = ""
    End Sub

    Private Sub txtHosteler_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHosteler.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Hostelers.HostelerID as [Hosteler Id], Hostelers.HostelerName as [Hosteler Name],Hostelers.USN as [USN], Hostelers.HostelName as [Hostel Name], Hostelers.RoomNo as [Room No], DueAmount.AcadYear as [Acad Year], DueAmount.TotalDueAmount as [Due Amount] FROM Hostelers INNER JOIN DueAmount ON Hostelers.[HostelerID] = DueAmount.[HostelerID] where TotalDueAmount > '0' and HostelerName like '%" & txtHosteler.Text & "%' order by HostelerName ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()

            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "DueAmount")
            dataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            dataGridView1.DataSource = myDataSet.Tables("DueAmount").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub txtUSN_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUSN.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Hostelers.HostelerID as [Hosteler Id], Hostelers.HostelerName as [Hosteler Name],Hostelers.USN as [USN], Hostelers.HostelName as [Hostel Name], Hostelers.RoomNo as [Room No], DueAmount.AcadYear as [Acad Year], DueAmount.TotalDueAmount as [Due Amount] FROM Hostelers INNER JOIN DueAmount ON Hostelers.[HostelerID] = DueAmount.[HostelerID] where TotalDueAmount > '0' and USN like '%" & txtUSN.Text & "%' order by HostelerName ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()

            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "DueAmount")
            dataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            dataGridView1.DataSource = myDataSet.Tables("DueAmount").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub frmFeePayment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        fillHostelerID()

        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Hostelers.HostelerID as [Hosteler Id], Hostelers.HostelerName as [Hosteler Name],Hostelers.USN as [USN], Hostelers.HostelName as [Hostel Name], Hostelers.RoomNo as [Room No], DueAmount.AcadYear as [Acad Year], DueAmount.TotalDueAmount as [Due Amount] FROM Hostelers INNER JOIN DueAmount ON Hostelers.[HostelerID] = DueAmount.[HostelerID] where TotalDueAmount > '0' and USN like '" & txtUSN.Text & "%' order by HostelerName ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()

            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "DueAmount")
            dataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            dataGridView1.DataSource = myDataSet.Tables("DueAmount").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

       

    End Sub

    Private Sub Update_record_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Update_record.Click
        Try
            If (txtExtraCharges.Text) = "" Then
                txtExtraCharges.Text = 0
            End If
            If (txtFine.Text = "") Then
                txtFine.Text = 0
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "update FeePayment set HostelerID='" & cmbHostelerID.Text & "',ServiceCharges='" & txtServiceCharges.Text & "',ExtraCharges='" & txtExtraCharges.Text & "',AcadYear='" & cmbAcadYear.Text & "',PaymentDate=#" & dtpPaymentDate.Text & "#,PayableCharges=" & txtTotalCharges.Text & ",TotalPaid=" & txtTotalPaid.Text & ",DuePayment='" & txtDueCharges.Text & "',Fine='" & txtFine.Text & "' , BankChallanNumber='" & txtBankChallanNumber.Text & "' where FeePaymentID='" & txtFeePaymentID.Text & "'"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            ' MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Update_record.Enabled = False
            con.Close()
            con.Open()
            Dim cb1 As String = "update DueAmount set TotalDueAmount='" & txtDueCharges.Text & "' where HostelerID='" & cmbHostelerID.Text & "'and AcadYear='" & cmbAcadYear.Text & "'"
            'Dim cb1 As String = "update DueAmount set TotalDueAmount =(select serviceCharges - sum(Totalpaid) from Feepayment where HostelerID='" & cmbHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "' group by HostelerID,ServiceCharges,ExtraCharge,AcadYear,PayableCharges,Fine)  where HostelerID='" & cmbHostelerID.Text & "'and AcadYear='" & cmbAcadYear.Text & "'"
            cmd = New OleDbCommand(cb1)
            cmd.Connection = con
            cmd.ExecuteReader()
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Update_record.Enabled = False
            con.Close()
            btnRefreshForm.PerformClick()
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
            Dim cq As String = "delete from FeePayment where FeePaymentID = '" & txtFeePaymentID.Text & "' and AcadYear='" & cmbAcadYear.Text & "'"
            cmd = New OleDbCommand(cq)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then

                Dim cb2 As String = "update DueAmount set TotalDueAmount= TotalDueAmount + val('" & txtTotalPaid.Text & "') where HostelerID='" & cmbHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "'"
                cmd = New OleDbCommand(cb2)
                cmd.Connection = con
                cmd.ExecuteReader()
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
        '   frmFeePaymentRecord1.DataGridView1.DataSource = Nothing
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
            Dim rpt As New rptFeePaymentReceipt 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT FeePaymentID,Hostelers.HostelerID,Hostelername,HostelName,FeePayment.AcadYear,PaymentDate,RoomNo,ServiceCharges,PayableCharges,TotalPaid,DuePayment,BankChallanNumber from Hostelers,Feepayment where Hostelers.HostelerID=FeePayment.HostelerID and FeePayment.AcadYear=Hostelers.AcadYear and FeePaymentID='" & txtFeePaymentID.Text & "'"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Hostelers")
            myDA.Fill(myDS, "FeePayment")
            rpt.SetDataSource(myDS)
            frmFeePaymentReceipt.CrystalReportViewer1.ReportSource = rpt
            frmFeePaymentReceipt.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

  
    Private Sub txtExtraCharges_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtExtraCharges.TextChanged
        txtTotalCharges.Text = Val(txtExtraCharges.Text) + Val(txtServiceCharges.Text)
    End Sub

    Private Sub txtServiceCharges_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtServiceCharges.TextChanged
        txtTotalCharges.Text = Val(txtServiceCharges.Text) + Val(txtFine.Text)
    End Sub
   
   
    Private Sub frmRegCharges_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub


   
End Class