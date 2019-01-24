Imports System.Data.OleDb
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Public Class frmRegCharges
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"
    Sub GetHosteler()
        Try
            con = New OleDbConnection(cs)
            con.Open()

            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name],USN as [USN]  from Hostelers where Status ='ACTIVE' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

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
        txtOtherFee.Text = ""
        txtRemarks.Text = ""
        txtFeePaid.Text = ""
        txtPrevDue.Text = ""
        dtpPaymentDate.Text = Today
        btnSave.Enabled = True
        btnDelete.Enabled = False
        btnUpdate_record.Enabled = False
        btnPrint.Enabled = False
    End Sub
    Private Sub cmbHostelerID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbHostelerID.SelectedIndexChanged
        Try
            btnDelete.Enabled = True
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
            If txtPrevDue.Text = "0" Then
                MsgBox(" No Previous Balance", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            rdr.Close()
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

    Private Sub frmRegCharges_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub
    Private Sub frmRegCharges_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Reset()
        fillHostelerID()
        GetHosteler()
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
            MessageBox.Show("Please enter Deposit", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtCautionMoney.Focus()
            Exit Sub
        End If
        If Len(Trim(txtRentalCharges.Text)) = 0 Then
            MessageBox.Show("Please enter rental Fee", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtRentalCharges.Focus()
            Exit Sub
        End If
        If Len(Trim(txtFormFee.Text)) = 0 Then
            MessageBox.Show("Please enter Form Fee ", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtFormFee.Focus()
            Exit Sub
        End If
        If Len(Trim(txtOtherFee.Text)) = 0 Then
            MessageBox.Show("Please enter Other Fees If any", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtOtherFee.Focus()
            Exit Sub
        End If
        
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT hostelerID,AcadYear FROM RegCharges where HostelerID='" & cmbHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                MessageBox.Show("Already registered for " & cmbAcadYear.Text & "! please Check out and Re Allocate", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Reset()
                Me.Hide()
                frmCheckOut.Show()
                Exit Sub
                
            End If
        
            If (rdr IsNot Nothing) Then
                txtReceiptNo.Text = "R-" & GetUniqueKey(6)
                con = New OleDbConnection(cs)
                con.Open()
                Dim cb As String = "insert into RegCharges(HostelerID,CautionMoney,RentalCharges,FormFee,OtherFee,PrevDue,TotalCharges,PaymentDate,ReceiptNumber,AcadYear, Remarks) VALUES ('" & cmbHostelerID.Text & "'," & txtCautionMoney.Text & "," & txtRentalCharges.Text & "," & txtFormFee.Text & "," & txtOtherFee.Text & "," & txtPrevDue.Text & "," & txtTotalCharges.Text & ",#" & dtpPaymentDate.Text & "#,'" & txtReceiptNo.Text & "','" & cmbAcadYear.Text & "' ,'" & txtRemarks.Text & "')"
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
            End If


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
            Dim cq As String = "delete from RegCharges where HostelerID= '" & cmbHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "'"
            cmd = New OleDbCommand(cq)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            con.Close()
            If RowsAffected > 0 Then
                con = New OleDbConnection(cs)
                con.Open()
                Dim cb2 As String = "delete from DueAmount where HostelerID= '" & cmbHostelerID.Text & "' and AcadYear='" & cmbAcadYear.Text & "'"
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
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnUpdate_record_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate_record.Click
        Try
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
                MessageBox.Show("Please enter Deposit", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtCautionMoney.Focus()
                Exit Sub
            End If
            If Len(Trim(txtRentalCharges.Text)) = 0 Then
                MessageBox.Show("Please enter rental Fee", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtRentalCharges.Focus()
                Exit Sub
            End If
            If Len(Trim(txtFormFee.Text)) = 0 Then
                MessageBox.Show("Please enter Form Fee ", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtCautionMoney.Focus()
                Exit Sub
            End If
            If Len(Trim(txtOtherFee.Text)) = 0 Then
                MessageBox.Show("Please enter Other Fees If any", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtOtherFee.Focus()
                Exit Sub
            End If


            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "update RegCharges  set CautionMoney=" & txtCautionMoney.Text & ",RentalCharges=" & txtRentalCharges.Text & ",FormFee=" & txtFormFee.Text & ",OtherFee=" & txtOtherFee.Text & ",PrevDue=" & txtPrevDue.Text & ",TotalCharges=" & txtTotalCharges.Text & ",PaymentDate= #" & dtpPaymentDate.Text & "#, AcadYear='" & cmbAcadYear.Text & "', Remarks='" & txtRemarks.Text & "' where  HostelerID='" & cmbHostelerID.Text & "' and acadyear='" & cmbAcadYear.Text & "'"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            ' MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            con.Close()
            con.Open()
            Dim cb1 As String = "update DueAmount set TotalCharges='" & txtTotalCharges.Text & "',TotalDueAmount='" & txtTotalCharges.Text - txtFeePaid.Text & "',AcadYear='" & cmbAcadYear.Text & "' where HostelerID='" & cmbHostelerID.Text & "' and acadyear='" & cmbAcadYear.Text & "'"
            cmd = New OleDbCommand(cb1)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            con.Close()
            btnUpdate_record.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        frmRegRecord.Show()
        frmRegRecord.GetData()
    End Sub

    Private Sub txtHosteler_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHosteler.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name] from Hostelers where HostelerName like '%" & txtHosteler.Text & "%' and Status ='ACTIVE'order by HostelerName", con)
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
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID], HostelerName as [Hosteler Name],USN as [USN] from Hostelers where USN like '%" & txtUSN.Text & "%' and Status ='ACTIVE' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try

            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptRegReceipt 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT distinct RegCharges.ReceiptNumber , RegCharges.HostelerID, Hostelers.HostelerName, RoomAllotment.Hostelname, RoomAllotment.RoomNo, RegCharges.AcadYear, RegCharges.CautionMoney, RegCharges.RentalCharges, RegCharges.FormFee, RegCharges.PrevDue, RegCharges.TotalCharges, RegCharges.PaymentDate FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID] where ReceiptNumber='" & txtReceiptNo.Text & "' and RegCharges.AcadYear=RoomAllotment.AcadYear"
            'MyCommand.CommandText = "SELECT RegCharges.ReceiptNumber, RegCharges.HostelerID, RegCharges.CautionMoney, RegCharges.RentalCharges, RegCharges.FormFee,RegCharges.PrevDue,RegCharges.TotalCharges,RegCharges.AcadYear, RegCharges.PaymentDate,Hostelers.HostelerID AS Expr1, Hostelers.HostelerName, Hostelers.DOB, Hostelers.Gender, Hostelers.RoomNo, Hostelers.HostelName,Hostelers.DateOfJoining, Hostelers.Purpose, Hostelers.FatherName, Hostelers.MobNo1, Hostelers.Phone1, Hostelers.MotherName, Hostelers.MobNo2,Hostelers.City, Hostelers.Address, Hostelers.Email, Hostelers.ContactNo, Hostelers.InstOfcDetails, Hostelers.Phone2, Hostelers.Agreement,Hostelers.GuardianName, Hostelers.GuardianAddress, Hostelers.MobNo3, Hostelers.Phone3, Hostelers.Photo, Hostelers.DocsPic,Hostelers.CompletionDate FROM(RegCharges INNER JOIN Hostelers ON RegCharges.HostelerID = Hostelers.HostelerID) where ReceiptNumber='" & txtReceiptNo.Text & "' and RegCharges.AcadYear=Hostelers.AcadYear"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Hostelers")
            myDA.Fill(myDS, "RegCharges")
            myDA.Fill(myDS, "RoomAllotment")
            rpt.SetDataSource(myDS)
            frmRegistrationReport.CrystalReportViewer1.ReportSource = rpt
            frmRegistrationReport.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
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
    Private Sub txtOtherFee_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOtherFee.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtFormFee_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFormFee.TextChanged
        txtTotalCharges.Text = Val(txtCautionMoney.Text) + Val(txtRentalCharges.Text) + Val(txtFormFee.Text) + Val(txtOtherFee.Text) + Val(txtPrevDue.Text)

    End Sub

    Private Sub txtCautionMoney_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCautionMoney.TextChanged
        txtTotalCharges.Text = Val(txtCautionMoney.Text) + Val(txtRentalCharges.Text) + Val(txtFormFee.Text) + Val(txtOtherFee.Text) + Val(txtPrevDue.Text)
    End Sub

    Private Sub txtPrevDue_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRentalCharges.TextChanged
        txtTotalCharges.Text = Val(txtCautionMoney.Text) + Val(txtRentalCharges.Text) + Val(txtFormFee.Text) + Val(txtOtherFee.Text) + Val(txtPrevDue.Text)

    End Sub
    Private Sub txtRentalCharges_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrevDue.TextChanged
        txtTotalCharges.Text = Val(txtCautionMoney.Text) + Val(txtRentalCharges.Text) + Val(txtFormFee.Text) + Val(txtOtherFee.Text) + Val(txtPrevDue.Text)

    End Sub

   
   
    Private Sub txtOtherFee_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOtherFee.TextChanged
        txtTotalCharges.Text = Val(txtCautionMoney.Text) + Val(txtRentalCharges.Text) + Val(txtFormFee.Text) + Val(txtOtherFee.Text) + Val(txtPrevDue.Text)
    End Sub



    Private Sub txtReceiptNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReceiptNo.TextChanged
        Try

            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT FeePayment.HostelerID, FeePayment.AcadYear,Sum(FeePayment.TotalPaid)  FROM FeePayment where FeePayment.HostelerID = '" & cmbHostelerID.Text & "' and FeePayment.AcadYear = '" & cmbAcadYear.Text & "' GROUP BY FeePayment.HostelerID, FeePayment.AcadYear "

            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtFeePaid.Text = rdr.GetValue(2)

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
End Class