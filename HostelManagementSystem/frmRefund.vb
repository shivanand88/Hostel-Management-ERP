Imports System.Data.OleDb
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Public Class frmRefund
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"


    Sub Reset()
        cmbHostelerID.Text = ""
        txthostelerName.Text = ""
        txtHosteler.Text = ""
        txtUSN.Text = ""
        cmbAcadYear.Text = ""
        txtRefundAmt.Text = ""
        txtRefundableAmount.Text = ""
        txtChequeNo.Text = ""
        txtRemarks.Text = ""
        dtpRefundDate.Text = Today
        btnSave.Enabled = True
    End Sub
    Private Sub cmbHostelerID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbHostelerID.SelectedIndexChanged
        Try
            btnDelete.Enabled = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT HostelerName,Acadyear FROM Hostelers WHERE HostelerID= '" & cmbHostelerID.Text & "' and Status = 'ACTIVE' "
            'cmd.CommandText = "SELECT  Hostelers.HostelerName,RegCharges.AcadYear, RegCharges.CautionMoney  FROM Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID] where RegCharges.HostelerID='" & cmbHostelerID.Text & "' and   RegCharges.CautionMoney>0  "
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtHostelerName.Text = rdr.GetString(0)
                cmbAcadYear.Text = rdr.GetString(1)
                '    txtRefundableAmount.Text = rdr.GetString(2)
                txtRefundAmt.Focus()
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub frmRefund_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub
    Private Sub frmRefund_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
       
        Reset()
        fillHostelerID()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name],USN as [USN] from Hostelers order by HostelerName", con)
            cmd = New OleDbCommand("SELECT RegCharges.HostelerID as [Hosteler ID], Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], RegCharges.CautionMoney as [Deposit] ,RegCharges.AcadYear FROM Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID] where CautionMoney>0 and Hostelers.Status = 'ACTIVE' ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Regcharges")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("Regcharges").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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
    Private Sub DataGridView1_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs)
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView1.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView1.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub
    Private Sub btnNewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewRecord.Click
        Reset()
    End Sub
    Private Sub txtHosteler_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHosteler.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name],USN as [USN] from Hostelers order by HostelerName", con)
            cmd = New OleDbCommand("SELECT RegCharges.HostelerID as [Hosteler ID], Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], RegCharges.CautionMoney as [Deposit] ,RegCharges.AcadYear as [Acad Year] FROM Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID] where RegCharges.CautionMoney>0 and Hostelers.Status = 'ACTIVE' and HostelerName like '%" & txtHosteler.Text & "%' order by HostelerName ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Regcharges")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("Regcharges").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub txtUSN_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUSN.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name],USN as [USN] from Hostelers order by HostelerName", con)
            cmd = New OleDbCommand("SELECT RegCharges.HostelerID as [Hosteler ID], Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], RegCharges.CautionMoney as [Deposit] ,RegCharges.AcadYear FROM Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID] where RegCharges.CautionMoney>0 and Hostelers.Status = 'ACTIVE' and USN like '%" & txtUSN.Text & "%' order by HostelerName ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Regcharges")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("Regcharges").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        txtRefundableAmount.Text = ""
        txtRefundAmt.Text = ""
        txtChequeNo.Text = ""
        dtpRefundDate.Text = Today
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            cmbHostelerID.Text = dr.Cells(0).Value.ToString
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            'cmd.CommandText = "SELECT HostelerName,AcadYear FROM Hostelers WHERE HostelerID= '" & cmbHostelerID.Text & "'"
            cmd.CommandText = "SELECT  Regcharges.HostelerId as [Hosteler Id],Hostelers.HostelerName as [Hosteler Name],  RegCharges.CautionMoney as [Deposit] ,RegCharges.AcadYear as [Acad Year] FROM Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID] where Regcharges.HostelerID= '" & cmbHostelerID.Text & "' "
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                ' cmbHostelerID.Text = rdr.GetString(0)
                txthostelerName.Text = rdr.GetString(1)
                txtRefundableAmount.Text = rdr.GetInt32(2).ToString
                cmbAcadYear.Text = rdr.GetString(3)

                txtRefundAmt.Focus()
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(cmbHostelerID.Text)) = 0 Then
            MessageBox.Show("Please Select Hosteler", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbHostelerID.Focus()
            Exit Sub
        End If
        If Len(Trim(txtRefundAmt.Text)) = 0 Then
            MessageBox.Show("Please Enter Amount", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtRefundAmt.Focus()
            Exit Sub
        End If
        If Len(Trim(txtChequeNo.Text)) = 0 Then
            MessageBox.Show("Please Enter Cheque No", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtChequeNo.Focus()
            Exit Sub
        End If

        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT hostelerID,AcadYear FROM Checkout where HostelerID='" & cmbHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                Dim cb1 As String = "insert into Refund(HostelerID,RefundableAmt,RefundAmt,ChequeNo,RefundDate,AcadYear,Remarks) values ('" & cmbHostelerID.Text & "','" & txtRefundableAmount.Text & "','" & txtRefundAmt.Text & "','" & txtChequeNo.Text & "',#" & dtpRefundDate.Text & "#,'" & cmbAcadYear.Text & "','" & txtRemarks.Text & "')"
                cmd = New OleDbCommand(cb1)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                con.Close()

                con = New OleDbConnection(cs)
                con.Open()
                Dim cb2 As String = "update Hostelers set Status='DEACTIVE' where HostelerID='" & cmbHostelerID.Text & "' "
                cmd = New OleDbCommand(cb2)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                con.Close()
                MessageBox.Show(" " & txthostelerName.Text & " ,Refunded SuccessFully", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Exit Sub
            End If
            If (rdr IsNot Nothing) Then
                MessageBox.Show("Pleae Checkout hosteler From " & cmbAcadYear.Text & " First", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Reset()
               
                Exit Sub
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    
    Private Sub btnGetData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetData.Click
        Me.Hide()
        frmRefundRecord.Refresh()
        frmRefundRecord.Show()
    End Sub

    Private Sub btnUpdate_record_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate_record.Click
        If Len(Trim(txtRefundAmt.Text)) = 0 Then
            MessageBox.Show("Please Enter Amount", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtRefundAmt.Focus()
            Exit Sub
        End If
        If Len(Trim(txtChequeNo.Text)) = 0 Then
            MessageBox.Show("Please Enter Cheque No", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtChequeNo.Focus()
            Exit Sub
        End If
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "update Refund  set RefundableAmt='" & txtRefundableAmount.Text & "',RefundAmt='" & txtRefundAmt.Text & "',ChequeNo='" & txtChequeNo.Text & "', RefundDate=#" & dtpRefundDate.Text & "#, Remarks='" & txtRemarks.Text & "' where  HostelerID='" & cmbHostelerID.Text & "' and acadyear='" & cmbAcadYear.Text & "'"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Reset()
            con.Close()
            btnUpdate_record.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
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
            Dim cq As String = "delete from Refund where HostelerID= '" & cmbHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "'"
            cmd = New OleDbCommand(cq)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()

            MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Reset()
            con.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtRefundAmt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRefundAmt.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

   
    Private Sub txtRefundAmt_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtRefundAmt.Validating
        If Val(txtRefundAmt.Text) > Val(txtRefundableAmount.Text) Then
            MessageBox.Show("Entered Amount is more than refund amount", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            txtRefundAmt.Text = ""
            txtRefundAmt.Focus()
        End If
    End Sub
End Class