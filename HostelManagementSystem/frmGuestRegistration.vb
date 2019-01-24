Imports System.Data.OleDb
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Public Class frmGuestRegistration
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"
    Sub Reset()
        txtCautionMoney.Text = ""
        txtFormFee.Text = ""
        cmbHostelerID.Text = ""
        cmbAcadYear.Text = ""
        txtHostelerName.Text = ""
        txtHostelName.Text = ""
        txtRentalCharges.Text = ""
        txtTotalCharges.Text = ""
        txtRoomNo.Text = ""
        txtHosteler.Text = ""
        txtUSN.Text = ""
        txtReceiptNo.Text = ""
        txtPrevDue.Text = ""
        dtpPaymentDate.Text = Today
        btnSave.Enabled = True
        btnDelete.Enabled = False
        btnUpdate_record.Enabled = False
        btnPrint.Enabled = False
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
    Private Sub frmGuestRegistration_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub
    Private Sub frmGuestRegistration_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Reset()
        fillHostelerID()
        Try
            con = New OleDbConnection(cs)
            con.Open()

            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name],USN as [USN]  from Hostelers order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub cmbHostelerID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbHostelerID.SelectedIndexChanged
        Try

            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT HostelerName,RoomNo,hostelname,AcadYear FROM Hostelers WHERE HostelerID= '" & cmbHostelerID.Text & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtHostelerName.Text = rdr.GetString(0)
                txtRoomNo.Text = rdr.GetString(1)
                txtHostelName.Text = rdr.GetString(2)
                cmbAcadYear.Text = rdr.GetString(3)
                cmbAcadYear.Enabled = False
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
    Private Sub txtHosteler_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHosteler.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name] from Hostelers where HostelerName like '" & txtHosteler.Text & "%' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub txtUSN_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUSN.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name],USN as [USN] from Hostelers where USN like '" & txtUSN.Text & "%' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        frmGuestRegistrationRecord.Show()
        frmGuestRegistrationRecord.GetData()
    End Sub
    Private Sub txtCautionMoney_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCautionMoney.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtRentalCharges_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRentalCharges.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtFormFee_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFormFee.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtPreDue_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPrevDue.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtFormFee_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFormFee.TextChanged
        txtTotalCharges.Text = Val(txtCautionMoney.Text) + Val(txtRentalCharges.Text) + Val(txtFormFee.Text) + Val(txtPrevDue.Text)

    End Sub

    Private Sub txtCautionMoney_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCautionMoney.TextChanged
        txtTotalCharges.Text = Val(txtCautionMoney.Text) + Val(txtRentalCharges.Text) + Val(txtFormFee.Text) + Val(txtPrevDue.Text)
    End Sub

    Private Sub txtPrevDue_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRentalCharges.TextChanged
        txtTotalCharges.Text = Val(txtCautionMoney.Text) + Val(txtRentalCharges.Text) + Val(txtFormFee.Text) + Val(txtPrevDue.Text)

    End Sub
    Private Sub txtRentalCharges_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrevDue.TextChanged
        txtTotalCharges.Text = Val(txtCautionMoney.Text) + Val(txtRentalCharges.Text) + Val(txtFormFee.Text) + Val(txtPrevDue.Text)

    End Sub
    Private Sub DataGridView2_RowHeaderMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.RowHeaderMouseClick
        txtCautionMoney.Focus()
        Try

            Dim dr As DataGridViewRow = DataGridView2.SelectedRows(0)
            cmbHostelerID.Text = dr.Cells(0).Value.ToString
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT HostelerName,RoomNo,hostelname,AcadYear FROM Hostelers WHERE HostelerID= '" & cmbHostelerID.Text & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtHostelerName.Text = rdr.GetString(0)
                txtRoomNo.Text = rdr.GetString(1)
                txtHostelName.Text = rdr.GetString(2)
                cmbAcadYear.Text = rdr.GetString(3)
                cmbAcadYear.Enabled = False
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
        con = New OleDbConnection(cs)
        con.Open()
        Dim ct3 As String = "select HostelerId,sum(TotalDueAmount) from DueAmount where HostelerID = '" & cmbHostelerID.Text & "' group by HostelerID "
        cmd = New OleDbCommand(ct3)
        cmd.Connection = con
        rdr = cmd.ExecuteReader()
        Try
            rdr.Read()

            txtPrevDue.Visible = True
            txtPrevDue.Text = rdr.GetValue(1)
            rdr.Close()
        Catch ex As Exception
            txtPrevDue.Text = "0"
            MsgBox(" No Previous Balance", MessageBoxButtons.OK, MessageBoxIcon.Information)
            rdr.Close()
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
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(cmbHostelerID.Text)) = 0 Then
            MessageBox.Show("Please select hosteler id", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbHostelerID.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbAcadYear.Text)) = 0 Then
            MessageBox.Show("Please Select Academic Year", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbAcadYear.Focus()
            Exit Sub
        End If
        If Len(Trim(txtCautionMoney.Text)) = 0 Then
            MessageBox.Show("Please enter caution money", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtCautionMoney.Focus()
            Exit Sub
        End If
        If Len(Trim(txtFormFee.Text)) = 0 Then
            MessageBox.Show("Please enter Form Fee ", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtCautionMoney.Focus()
            Exit Sub
        End If

        If Len(Trim(txtRentalCharges.Text)) = 0 Then
            MessageBox.Show("Please enter rental Charges", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtRentalCharges.Focus()
            Exit Sub
        End If
        Try
           

            txtReceiptNo.Text = "R-" & GetUniqueKey(4)
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "insert into RegCharges(HostelerID,CautionMoney,RentalCharges,FormFee,PrevDue,TotalCharges,PaymentDate,ReceiptNumber,AcadYear) VALUES ('" & cmbHostelerID.Text & "'," & txtCautionMoney.Text & "," & txtRentalCharges.Text & "," & txtFormFee.Text & "," & txtPrevDue.Text & "," & txtTotalCharges.Text & ",#" & dtpPaymentDate.Text & "#,'" & txtReceiptNo.Text & "','" & cmbAcadYear.Text & "')"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()

            con = New OleDbConnection(cs)
            con.Open()
            Dim cb1 As String = "insert into DueAmount(HostelerID,TotalCharges,TotalDueAmount,ReceiptNumber,AcadYear) VALUES ('" & cmbHostelerID.Text & "','" & txtTotalCharges.Text & "','" & txtTotalCharges.Text & "','" & txtReceiptNo.Text & "','" & cmbAcadYear.Text & "')"
            cmd = New OleDbCommand(cb1)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            ' update table due amount if any previous due amouut exists make it to zero 
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb6 As String = "update DueAmount set TotalDueAmount='0' where HostelerID='" & cmbHostelerID.Text & "' and TotalDueAmount='" & txtPrevDue.Text & "'"
            cmd = New OleDbCommand(cb6)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            'MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            con.Close()


            con = New OleDbConnection(cs)
            con.Open()
            Dim cb3 As String = "delete from RoomAllotment where HostelerID= '" & cmbHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "'"
            cmd = New OleDbCommand(cb3)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb5 As String = "insert into RoomAllotment(HostelerID,Hostelname,RoomNo,AcadYear) values ('" & cmbHostelerID.Text & "','" & txtHostelName.Text & "','" & txtRoomNo.Text & "','" & cmbAcadYear.Text & "')"
            cmd = New OleDbCommand(cb5)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            MessageBox.Show(" " & txtHostelerName.Text & " ,Registered Proceed  to Fee Payment", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            con.Close()
            Reset()
            btnSave.Enabled = False
            btnPrint.Enabled = True
            rdr.Close()
            con.Close()

            Exit Sub



        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub btnNewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewRecord.Click
        Reset()
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
            Dim cq As String = "delete from RegCharges where RecieptNumber= '" & txtReceiptNo.Text & "' "
            cmd = New OleDbCommand(cq)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            con.Close()
            If RowsAffected > 0 Then
                con = New OleDbConnection(cs)
                con.Open()
                Dim cb2 As String = "delete from DueAmount where HostelerID= RecieptNumber= '" & txtReceiptNo.Text & "'"
                cmd = New OleDbCommand(cb2)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                con.Close()
                con = New OleDbConnection(cs)
                con.Open()
                Dim cb3 As String = "delete from RoomAllotment where HostelerID= '" & cmbHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "'"
                cmd = New OleDbCommand(cb3)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                con.Close()
                con = New OleDbConnection(cs)
                con.Open()
                Dim cb4 As String = "delete from FeePayment where HostelerID= '" & cmbHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "'"
                cmd = New OleDbCommand(cb4)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                con = New OleDbConnection(cs)
                con.Open()
                Dim cb5 As String = "delete from Concession where HostelerID= '" & cmbHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "'"
                cmd = New OleDbCommand(cb5)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
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
            MessageBox.Show("No Record Found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
    End Sub
    Private Sub btnUpdate_record_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate_record.Click
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT hostelerID,AcadYear FROM FeePayment where HostelerID='" & cmbHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                MessageBox.Show("Fee payment Done! please Delete All records from Fee Payment  for " & cmbAcadYear.Text & "", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Reset()
                Me.Hide()
                frmFeePaymentRecord1.Show()

                Exit Sub

            End If

            If (rdr IsNot Nothing) Then
                con = New OleDbConnection(cs)
                con.Open()
                Dim cb As String = "update RegCharges  set CautionMoney=" & txtCautionMoney.Text & ",RentalCharges=" & txtRentalCharges.Text & ",FormFee=" & txtFormFee.Text & ",PrevDue=" & txtPrevDue.Text & ",TotalCharges=" & txtTotalCharges.Text & ",PaymentDate= #" & dtpPaymentDate.Text & "#, AcadYear='" & cmbAcadYear.Text & "' where  HostelerID='" & cmbHostelerID.Text & "' and acadyear='" & cmbAcadYear.Text & "'"
                cmd = New OleDbCommand(cb)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                ' MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                con.Close()
                con.Open()
                Dim cb1 As String = "update DueAmount set TotalCharges='" & txtTotalCharges.Text & "',TotalDueAmount='" & txtTotalCharges.Text & "',AcadYear='" & cmbAcadYear.Text & "' where HostelerID='" & cmbHostelerID.Text & "' and acadyear='" & cmbAcadYear.Text & "'"
                cmd = New OleDbCommand(cb1)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                con.Close()
                btnUpdate_record.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    
End Class