Imports System.Data.OleDb
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

Public Class frmHostelers
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"
    Sub Reset()
        txtUsn.Text = ""
        txtAddress.Text = ""
        txtAgreement.Text = ""
        txtContactNo.Text = ""
        dtpDOB.Text = Today
        dtpDateOfJoining.Text = Today
        txtEmail.Text = ""
        txtFatherName.Text = ""
        txtGuardianAddress.Text = ""
        txtGuardianContactNo.Text = ""
        txtGuardianName.Text = ""
        txtGuardianPhoneNo.Text = ""
        txtPhone.Text = ""
        txtMotherName.Text = ""
        txtMobNo2.Text = ""
        txtMobNo1.Text = ""
        txtInstPhoneNo.Text = ""
        txtInOfcDeatils.Text = ""
        txtHostelerName.Text = ""
        txtHostelerID.Text = ""
        txtCaste.Text = ""
        txtSubCaste.Text = ""
        cmbHostelName.SelectedIndex = -1
        cmbGender.SelectedIndex = -1
        cmbPurpose.Text = ""
        cmbRoomNo.Text = ""
        cmbRelegion.Text = ""
        cmbCategory.Text = ""
        cmbAcadYear.Text = ""
        txtCity.Text = ""
        txtRoomType.Text = ""
        txtHostelname.Text = ""
        txtRoomno.Text = ""
        cmbStatus.SelectedIndex = -1
        cmbRoomNo.Enabled = False
        btnPrint.Enabled = False
        Picture.Image = My.Resources.photo
        PictureBox2.Image = My.Resources.images1
        btnSave.Enabled = True
        btnUpdate_record.Enabled = False
        btnDelete.Enabled = False
        btnReallocation.Enabled = False
        dtpCompletionDate.Text = Today
        txtHostelerName.Focus()
        txtAgreement.ReadOnly = False
    End Sub

    Private Sub frmHostelers_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub
    Sub Autocomplete()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim cmd As New OleDbCommand("SELECT distinct city FROM Hostelers", con)
            Dim ds As New DataSet()
            Dim da As New OleDbDataAdapter(cmd)
            da.Fill(ds, "Hostelers")
            Dim col As New AutoCompleteStringCollection()
            Dim i As Integer = 0
            For i = 0 To ds.Tables(0).Rows.Count - 1
                col.Add(ds.Tables(0).Rows(i)("City").ToString())
            Next
            txtCity.AutoCompleteSource = AutoCompleteSource.CustomSource
            txtCity.AutoCompleteCustomSource = col
            txtCity.AutoCompleteMode = AutoCompleteMode.Suggest
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim cmd As New OleDbCommand("SELECT distinct Caste FROM Hostelers", con)
            Dim ds As New DataSet()
            Dim da As New OleDbDataAdapter(cmd)
            da.Fill(ds, "Hostelers")
            Dim col As New AutoCompleteStringCollection()
            Dim i As Integer = 0
            For i = 0 To ds.Tables(0).Rows.Count - 1
                col.Add(ds.Tables(0).Rows(i)("Caste").ToString())
            Next
            txtCaste.AutoCompleteSource = AutoCompleteSource.CustomSource
            txtCaste.AutoCompleteCustomSource = col
            txtCaste.AutoCompleteMode = AutoCompleteMode.Suggest
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub frmHostelers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Autocomplete()
        fillHostelName()
    End Sub
    Sub fillHostelName()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct RTRIM(HostelName) FROM Room", CN)
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

    

    Private Sub txtInstPhoneNo_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtInstPhoneNo.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtFixedDeposit_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtEmail_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtEmail.Validating
        Dim rEMail As New System.Text.RegularExpressions.Regex("^[a-zA-Z][\w\.-]{2,28}[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$")
        If txtEmail.Text.Length > 0 Then
            If Not rEMail.IsMatch(txtEmail.Text) Then
                MessageBox.Show("invalid email address", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                txtEmail.SelectAll()
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub txtPhone_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPhone.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

   
    Public Sub DeleteRecord()
        Try
            Dim RowsAffected As Integer = 0
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select HostelerID from FeePayment where HostelerID = '" & txtHostelerID.Text & "' and AcadYear= '" & cmbAcadYear.Text & "'"
            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use Delete records from Fee payment first", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Reset()
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            
            
            
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct4 As String = "select HostelerID from RegCharges where HostelerID = '" & txtHostelerID.Text & "' and AcadYear= '" & cmbAcadYear.Text & "'"
            cmd = New OleDbCommand(ct4)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use Delete record From Registration", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Reset()
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            

            con = New OleDbConnection(cs)
            con.Open()
            Dim cq11 As String = "delete from CheckIn where HostelerID = '" & txtHostelerID.Text & "' "
            cmd = New OleDbCommand(cq11)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()

            con = New OleDbConnection(cs)
            con.Open()
            Dim ct5 As String = "update room set BedsAvailable =  BedsAvailable + 1 where HostelName= '" & cmbHostelName.Text & "' and RoomNo='" & cmbRoomNo.Text & "'"
            cmd = New OleDbCommand(ct5)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()

            con = New OleDbConnection(cs)
            con.Open()
            Dim cq1 As String = "delete from Hostelers where HostelerID = '" & txtHostelerID.Text & "'"
            cmd = New OleDbCommand(cq1)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Autocomplete()
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

        frmHostelersRecord1.fillRoomNo()
        frmHostelersRecord1.fillCity()

        frmHostelersRecord1.cmbBranch.Text = ""
        frmHostelersRecord1.DataGridView4.DataSource = Nothing
        frmHostelersRecord1.cmbRoomNo.Text = ""
        frmHostelersRecord1.DataGridView3.DataSource = Nothing
        frmHostelersRecord1.DateFrom.Text = Today
        frmHostelersRecord1.DateTo.Text = Today
        frmHostelersRecord1.DataGridView2.DataSource = Nothing

        frmHostelersRecord1.txtHostelerName.Text = ""
        frmHostelersRecord1.DataGridView1.DataSource = Nothing
        frmHostelersRecord1.DataGridView6.DataSource = Nothing
        frmHostelersRecord1.DateTimePicker1.Text = Today
        frmHostelersRecord1.DataGridView7.DataSource = Nothing
        frmHostelersRecord1.cmbCity.Text = ""
       
        frmHostelersRecord1.Show()
    End Sub

   

    Private Sub cmbHostelName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbHostelName.SelectedIndexChanged
        cmbRoomNo.Items.Clear()
        cmbRoomNo.Text = ""
        cmbRoomNo.Enabled = True
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select distinct RTRIM(RoomNo),BedsAvailable from Room where HostelName= '" & cmbHostelName.Text & "' and BedsAvailable > 0"
            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()
            While rdr.Read()
                cmbRoomNo.Items.Add(rdr(0))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub cmbRoomNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbRoomNo.SelectedIndexChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT RoomType FROM Room WHERE RoomNo= '" & cmbRoomNo.Text & "' and HostelName='" & cmbHostelName.Text & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtRoomType.Text = rdr.GetString(0)

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

   

    Private Sub txtContactNo_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtContactNo.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtMobNo1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMobNo1.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtMobNo2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMobNo2.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

   

    Private Sub btnSave_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If Len(Trim(txtHostelerName.Text)) = 0 Then
                MessageBox.Show("Please enter hosteler name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtHostelerName.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbGender.Text)) = 0 Then
                MessageBox.Show("Please select gender", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbGender.Focus()
                Exit Sub
            End If
            
            If Len(Trim(cmbPurpose.Text)) = 0 Then
                MessageBox.Show("Please select Department", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbPurpose.Focus()
                Exit Sub
            End If

            If Len(Trim(dtpDateOfJoining.Text)) = 0 Then
                MessageBox.Show("Please select Date Of joining", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                dtpDateOfJoining.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbAcadYear.Text)) = 0 Then
                MessageBox.Show("Please Select Academic Year.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbAcadYear.Focus()
                Exit Sub
            End If

            If Len(Trim(txtContactNo.Text)) = 0 Then
                MessageBox.Show("Please enter contact no.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtContactNo.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbHostelName.Text)) = 0 Then
                MessageBox.Show("Please select hostel name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbHostelName.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbRoomNo.Text)) = 0 Then
                MessageBox.Show("Please select room no.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbRoomNo.Focus()
                Exit Sub
            End If
            If Len(Trim(txtFatherName.Text)) = 0 Then
                MessageBox.Show("Please enter father's name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtFatherName.Focus()
                Exit Sub
            End If
            If Len(Trim(txtAddress.Text)) = 0 Then
                MessageBox.Show("Please enter address", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtAddress.Focus()
                Exit Sub
            End If

            
            If Len(Trim(cmbRelegion.Text)) = 0 Then
                MessageBox.Show("Please Select Relegion.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbRelegion.Focus()
                Exit Sub
            End If
           
            If Len(Trim(txtCaste.Text)) = 0 Then
                MessageBox.Show("Please enter Caste name.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtCaste.Focus()
                Exit Sub
            End If

            If Len(Trim(cmbCategory.Text)) = 0 Then
                MessageBox.Show("Please Select Categoory.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbCategory.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbStatus.Text)) = 0 Then
                MessageBox.Show("Please Select Status.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbStatus.Focus()
                Exit Sub
            End If
           
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT BedsAvailable FROM Room WHERE RoomNo= '" & cmbRoomNo.Text & "' and HostelName='" & cmbHostelName.Text & "' and BedsAvailable <= 0"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                MessageBox.Show("Bed not Available in selected room no.", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                cmbRoomNo.Focus()
                If (rdr IsNot Nothing) Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            txtHostelerID.Text = "H-" & GetUniqueKey(6)
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "insert into Hostelers(HostelerID,HostelerName,USN,AcadYear,DOB,RoomNo,DateOfJoining,Purpose,HostelName,FatherName,MobNo1,Phone1,MotherName,MobNo2,Address,Email,ContactNo,InstOfcDetails,Phone2,Agreement,GuardianName,GuardianAddress,MobNo3,Phone3,City,Photo,DocsPic,CompletionDate,Gender,Relegion,Caste,Subcaste,Category,Status) VALUES('" & txtHostelerID.Text & "','" & txtHostelerName.Text & "','" & txtUsn.Text & "','" & cmbAcadYear.Text & "',#" & dtpDOB.Text & "#,'" & cmbRoomNo.Text & "',#" & dtpDateOfJoining.Text & "#,'" & cmbPurpose.Text & "','" & cmbHostelName.Text & "','" & txtFatherName.Text & "','" & txtMobNo1.Text & "','" & txtPhone.Text & "','" & txtMotherName.Text & "','" & txtMobNo2.Text & "','" & txtAddress.Text & "','" & txtEmail.Text & "','" & txtContactNo.Text & "','" & txtInOfcDeatils.Text & "','" & txtInstPhoneNo.Text & "','" & txtAgreement.Text & "','" & txtGuardianName.Text & "','" & txtGuardianAddress.Text & "','" & txtGuardianContactNo.Text & "','" & txtGuardianPhoneNo.Text & "','" & txtCity.Text & "',@image,@docspic,#" & dtpCompletionDate.Text & "#,'" & cmbGender.Text & "','" & cmbRelegion.Text & "','" & txtCaste.Text & "','" & txtSubCaste.Text & "','" & cmbCategory.Text & "','" & cmbStatus.Text & "')"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            Dim ms, ms1 As New MemoryStream()
            Dim bmpImage As New Bitmap(Picture.Image)
            Dim bmpImage1 As New Bitmap(PictureBox2.Image)
            bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            bmpImage1.Save(ms1, System.Drawing.Imaging.ImageFormat.Jpeg)
            Dim data As Byte() = ms.GetBuffer()
            Dim data1 As Byte() = ms1.GetBuffer()
            Dim p As New OleDbParameter("@d1", OleDbType.VarBinary)
            Dim p1 As New OleDbParameter("@d2", OleDbType.VarBinary)
            p.Value = data
            p1.Value = data1
            cmd.Parameters.Add(p)
            cmd.Parameters.Add(p1)
            cmd.ExecuteNonQuery()
            con.Close()
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "update room set BedsAvailable = BedsAvailable - 1 where HostelName= '" & cmbHostelName.Text & "' and RoomNo='" & cmbRoomNo.Text & "'"
            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            Autocomplete()
            con.Close()
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct1 As String = "insert into CheckIn (HostelerID,HostelerName,HostelName,RoomNo,AcadYear,CheckInDate) values('" & txtHostelerID.Text & "','" & txtHostelerName.Text & "','" & cmbHostelName.Text & "','" & cmbRoomNo.Text & "','" & cmbAcadYear.Text & "',#" & dtpDateOfJoining.Text & "#)"
            cmd = New OleDbCommand(ct1)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
           
            MessageBox.Show("Hosteler Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            btnPrint.Enabled = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try

            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptHostelers 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New Hostelers_DBDataSet 'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT Hostelers.HostelerID, Hostelers.HostelerName,Hostelers.USN, Hostelers.DOB, Hostelers.Gender, Hostelers.RoomNo, Hostelers.HostelName, Hostelers.DateOfJoining,Hostelers.Purpose, Hostelers.FatherName, Hostelers.MobNo1, Hostelers.Phone1, Hostelers.MotherName, Hostelers.MobNo2, Hostelers.City,Hostelers.Address, Hostelers.Email, Hostelers.ContactNo, Hostelers.InstOfcDetails, Hostelers.Phone2, Hostelers.Agreement, Hostelers.GuardianName,Hostelers.GuardianAddress, Hostelers.MobNo3, Hostelers.Phone3, Hostelers.Photo, Hostelers.DocsPic, Hostelers.CompletionDate,Hostelers.Relegion,Hostelers.Caste,Hostelers.Subcaste,Hostelers.Category, Hostel.HostelName AS Expr1, Hostel.Hostel_Address, Hostel.Hostel_Phone, Hostel.ManagedBy, Hostel.Hostel_ContactNo FROM (Hostelers INNER JOIN Hostel ON Hostelers.HostelName = Hostel.HostelName) where HostelerID='" & txtHostelerID.Text & "'"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Hostelers")
            myDA.Fill(myDS, "Hostel")
            rpt.SetDataSource(myDS)
            frmHostelerReport.CrystalReportViewer1.ReportSource = rpt
            frmHostelerReport.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If Len(Trim(txtHostelerID.Text)) = 0 Then
            MessageBox.Show("Please Select Hosteler", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtHostelerName.Focus()
            Exit Sub
        End If
        con = New OleDbConnection(cs)
        con.Open()
        Dim ct3 As String = "select TotalDueAmount from DueAmount where HostelerID = '" & txtHostelerID.Text & "'"
        cmd = New OleDbCommand(ct3)
        cmd.Connection = con
        rdr = cmd.ExecuteReader()
        Try
            rdr.Read()
            txtTotalDueAmount.Visible = True
            txtTotalDueAmount.Text = rdr.GetValue(0)
            rdr.Close()
        Catch ex As Exception
            MsgBox("", MessageBoxButtons.OK, MessageBoxIcon.Information)
            rdr.Close()
        End Try
        con.Close()
        Try

            If txtTotalDueAmount.Text > 0 Then
                MessageBox.Show("Due Amount Pending Can not Check out!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtHostelerID.Focus()
                Exit Sub
            End If
            If Len(Trim(txtHostelerID.Text)) = 0 Then
                MessageBox.Show("Please retrieve hosteler id", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtHostelerID.Focus()
                Exit Sub
            End If
            If Len(Trim(txtHostelerName.Text)) = 0 Then
                MessageBox.Show("Please retrieve hosteler name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtHostelerName.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbHostelName.Text)) = 0 Then
                MessageBox.Show("Please select hostel name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbHostelName.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbRoomNo.Text)) = 0 Then
                MessageBox.Show("Please select room no.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbRoomNo.Focus()
                Exit Sub
            End If


            If MessageBox.Show("Are you sure want to check out?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then

                ' selection of hosteler ID from Checkout table if it exists not ID already exists

                ' con = New OleDbConnection(cs)
                ' con.Open()
                ' cmd = con.CreateCommand()
                ' cmd.CommandText = "SELECT HostelerID FROM Checkout where HostelerID='" & txtHostelerID.Text & "'"
                ' rdr = cmd.ExecuteReader()
                ' If rdr.Read() Then
                'MessageBox.Show("Hosteler already checked out", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ' If (rdr IsNot Nothing) Then
                'rdr.Close()
                'End If
                'Exit Sub
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT NoOfBeds,bedsavailable FROM Room WHERE RoomNo= '" & txtRoomno.Text & "' and HostelName='" & txtHostelname.Text & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtNoOfBeds.Text = rdr.GetInt32(0)
                txtBedAvailable.Text = rdr.GetInt32(1)
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            If (Val(txtBedAvailable.Text) > Val(txtNoOfBeds.Text)) Then
                MessageBox.Show("Unable to allocate room", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "update room set BedsAvailable =  BedsAvailable + 1 where HostelName= '" & cmbHostelName.Text & "' and RoomNo='" & cmbRoomNo.Text & "'"
            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()

            'making beds available = no of beds if 
            ' con = New OleDbConnection(cs)
            ' con.Open()
            ' cmd = con.CreateCommand()
            'cmd.CommandText = "SELECT NoOfBeds,bedsavailable FROM Room WHERE RoomNo= '" & txtRoomno.Text & "' and HostelName='" & txtHostelname.Text & "'"
            ' rdr = cmd.ExecuteReader()
            'If rdr.Read() Then
            ' txtNoOfBeds.Text = rdr.GetInt32(0)
            'txtBedAvailable.Text = rdr.GetInt32(1)
            ' End If
            ' If con.State = ConnectionState.Open Then
            'con.Close()
            ' End If
            ' making TotatalDueamount =30000 for next acadamic year for the same student
            'con = New OleDbConnection(cs)
            'con.Open()
            'Dim cd As String = "update DueAmount set TotalDueAmount = (select TotalCharges from RegCharges where HostelerId='" & txtHostelerID.Text & "')"
            'cmd = New OleDbCommand(cd)
            'cmd.Connection = con
            'cmd.ExecuteNonQuery()
            'con.Close()
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "insert into CheckOut(HostelerID,HostelerName,HostelName,RoomNo,CheckOutDate) VALUES('" & txtHostelerID.Text & "','" & txtHostelerName.Text & "','" & cmbHostelName.Text & "','" & cmbRoomNo.Text & "',#" & DateTime.Now.ToShortDateString() & "#)"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb1 As String = "insert into CheckOut_History(HostelerID,HostelerName,HostelName,RoomNo,CheckOutDate) VALUES('" & txtHostelerID.Text & "','" & txtHostelerName.Text & "','" & cmbHostelName.Text & "','" & cmbRoomNo.Text & "',#" & DateTime.Now.ToShortDateString() & "#)"
            cmd = New OleDbCommand(cb1)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            MessageBox.Show("Successfully Check out", "Record updated", MessageBoxButtons.OK, MessageBoxIcon.Information)
            'End If
            Reset()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtAgreement_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAgreement.TextChanged
        dtpCompletionDate.Text = dtpDateOfJoining.Value.Date.AddMonths(Val(txtAgreement.Text))
    End Sub

    Private Sub txtAgreement_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAgreement.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try

            If Len(Trim(txtHostelerName.Text)) = 0 Then
                MessageBox.Show("Please enter hosteler name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtHostelerName.Focus()
                Exit Sub
            End If
            If Len(Trim(txtHostelerName.Text)) = 0 Then
                MessageBox.Show("Please select hostel name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtHostelerName.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbRoomNo.Text)) = 0 Then
                MessageBox.Show("Please select room no.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbRoomNo.Focus()
                Exit Sub
            End If

            If Len(Trim(txtFatherName.Text)) = 0 Then
                MessageBox.Show("Please enter father's name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtFatherName.Focus()
                Exit Sub
            End If
            If Len(Trim(txtMotherName.Text)) = 0 Then
                MessageBox.Show("Please enter mother's name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtMotherName.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbPurpose.Text)) = 0 Then
                MessageBox.Show("Please select purpose", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbPurpose.Focus()
                Exit Sub
            End If
            If Len(Trim(txtCity.Text)) = 0 Then
                MessageBox.Show("Please enter city", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtCity.Focus()
                Exit Sub
            End If
            If Len(Trim(txtAddress.Text)) = 0 Then
                MessageBox.Show("Please enter address", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtAddress.Focus()
                Exit Sub
            End If
            If Len(Trim(txtInOfcDeatils.Text)) = 0 Then
                MessageBox.Show("Please enter institute/office details", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtInOfcDeatils.Focus()
                Exit Sub
            End If
            If Len(Trim(txtContactNo.Text)) = 0 Then
                MessageBox.Show("Please enter contact no.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtContactNo.Focus()
                Exit Sub
            End If
            If MessageBox.Show("Are you sure want to re-allocate?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                con = New OleDbConnection(cs)
                con.Open()
                cmd = con.CreateCommand()
                cmd.CommandText = "SELECT hostelerID FROM Checkout where HostelerID='" & txtHostelerID.Text & "'"
                rdr = cmd.ExecuteReader()
                If Not rdr.Read() Then
                    MessageBox.Show("hosteler is not checked out" & vbCrLf & "re-allocation is not allowed", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    cmbRoomNo.Focus()
                    If (rdr IsNot Nothing) Then
                        rdr.Close()
                    End If
                    Exit Sub
                End If
                con = New OleDbConnection(cs)
                con.Open()
                cmd = con.CreateCommand()
                cmd.CommandText = "SELECT BedsAvailable FROM Room WHERE RoomNo= '" & cmbRoomNo.Text & "' and HostelName='" & cmbHostelName.Text & "' and BedsAvailable <= 0"
                rdr = cmd.ExecuteReader()
                If rdr.Read() Then
                    MessageBox.Show("Bed not Available in selected room no.", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    cmbRoomNo.Focus()
                    If (rdr IsNot Nothing) Then
                        rdr.Close()
                    End If
                    Exit Sub
                End If
                con = New OleDbConnection(cs)
                con.Open()
                Dim cb As String = "update Hostelers set HostelerName='" & txtHostelerName.Text & "',USN='" & txtUsn.Text & "',DOB=#" & dtpDOB.Text & "#,RoomNo='" & cmbRoomNo.Text & "',Purpose='" & cmbPurpose.Text & "',HostelName='" & cmbHostelName.Text & "',FatherName='" & txtFatherName.Text & "',MobNo1='" & txtMobNo1.Text & "',Phone1='" & txtPhone.Text & "',MotherName='" & txtMotherName.Text & "',MobNo2='" & txtMobNo2.Text & "',Address='" & txtAddress.Text & "',Email='" & txtEmail.Text & "',ContactNo='" & txtContactNo.Text & "',InstOfcDetails='" & txtInOfcDeatils.Text & "',Phone2='" & txtInstPhoneNo.Text & "',Agreement='" & txtAgreement.Text & "',GuardianName='" & txtGuardianName.Text & "',GuardianAddress='" & txtGuardianAddress.Text & "',MobNo3='" & txtGuardianContactNo.Text & "',Phone3='" & txtGuardianPhoneNo.Text & "',Photo=@d1,DocsPic=@d2,City='" & txtCity.Text & "',CompletionDate=#" & dtpCompletionDate.Text & "#,Gender='" & cmbGender.Text & "',Relegion='" & cmbRelegion.Text & "',Caste='" & txtCaste.Text & "',Subcaste='" & txtSubCaste.Text & "',Category='" & cmbCategory.Text & "' where HostelerID='" & txtHostelerID.Text & "'"
                cmd = New OleDbCommand(cb)
                cmd.Connection = con
                Dim ms, ms1 As New MemoryStream()
                Dim bmpImage As New Bitmap(Picture.Image)
                Dim bmpImage1 As New Bitmap(PictureBox2.Image)
                bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
                bmpImage1.Save(ms1, System.Drawing.Imaging.ImageFormat.Jpeg)
                Dim data As Byte() = ms.GetBuffer()
                Dim data1 As Byte() = ms1.GetBuffer()
                Dim p As New OleDbParameter("@d1", OleDbType.VarBinary)
                Dim p1 As New OleDbParameter("@d2", OleDbType.VarBinary)
                p.Value = data
                p1.Value = data1
                cmd.Parameters.Add(p)
                cmd.Parameters.Add(p1)
                cmd.ExecuteNonQuery()
                con = New OleDbConnection(cs)
                con.Open()
                Dim ct As String = "update room set BedsAvailable = BedsAvailable - 1 where HostelName= '" & cmbHostelName.Text & "' and RoomNo='" & cmbRoomNo.Text & "'"
                cmd = New OleDbCommand(ct)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                Autocomplete()
                MessageBox.Show("Successfully re-allocated", "Hosteler", MessageBoxButtons.OK, MessageBoxIcon.Information)
                btnUpdate_record.Enabled = False
                Reset()
                con.Close()
                For i = 1 To Val(txtAgreement.Text)
                    con = New OleDbConnection(cs)
                    con.Open()
                    Dim ct2 As String = "insert into DueDate(HostelerID,DueDate) values('" & txtHostelerID.Text & "',#" & dtpDateOfJoining.Value.Date.AddMonths(i) & "#)"
                    cmd = New OleDbCommand(ct2)
                    cmd.Connection = con
                    cmd.ExecuteNonQuery()
                    con.Close()
                Next
                con = New OleDbConnection(cs)
                con.Open()
                Dim ct3 As String = "delete from Checkout where HostelerID='" & txtHostelerID.Text & "'"
                cmd = New OleDbCommand(ct3)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            If Len(Trim(txtHostelerID.Text)) = 0 Then
                MessageBox.Show("Please retrieve hosteler id", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtHostelerID.Focus()
                Exit Sub
            End If
            For i = 1 To Val(txtAgreement.Text)
                con = New OleDbConnection(cs)
                con.Open()
                Dim ct2 As String = "insert into DueDate(HostelerID,DueDate) values('" & txtHostelerID.Text & "',#" & dtpDateOfJoining.Value.Date.AddMonths(i) & "#)"
                cmd = New OleDbCommand(ct2)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                con.Close()
            Next
            MessageBox.Show("Successfully extended", "Agreement", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnEA.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Browse_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Browse.Click
        Try
            With OpenFileDialog1
                .Filter = ("Images |*.png; *.bmp; *.jpg;*.jpeg; *.gif;")
                .FilterIndex = 4
            End With
            'Clear the file name
            OpenFileDialog1.FileName = ""
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                Picture.Image = Image.FromFile(OpenFileDialog1.FileName)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    Private Sub btnNewRecord_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNewRecord.Click
        Reset()
    End Sub

    Private Sub btnUpdate_record_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpdate_record.Click
        Try
            If Len(Trim(txtHostelerName.Text)) = 0 Then
                MessageBox.Show("Please enter hosteler name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtHostelerName.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbHostelName.Text)) = 0 Then
                MessageBox.Show("Please select hostel name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbHostelName.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbRoomNo.Text)) = 0 Then
                MessageBox.Show("Please select room no.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbRoomNo.Focus()
                Exit Sub
            End If

            If Len(Trim(txtFatherName.Text)) = 0 Then
                MessageBox.Show("Please enter father's name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtFatherName.Focus()
                Exit Sub
            End If
            
            If Len(Trim(cmbPurpose.Text)) = 0 Then
                MessageBox.Show("Please select purpose", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbPurpose.Focus()
                Exit Sub
            End If
           
            If Len(Trim(txtAddress.Text)) = 0 Then
                MessageBox.Show("Please enter address", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtAddress.Focus()
                Exit Sub
            End If
            
            If Len(Trim(txtContactNo.Text)) = 0 Then
                MessageBox.Show("Please enter contact no.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtContactNo.Focus()
                Exit Sub
            End If
            If Len(Trim(txtCaste.Text)) = 0 Then
                MessageBox.Show("Please enter Caste.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtCaste.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbCategory.Text)) = 0 Then
                MessageBox.Show("Please Select Category.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbCategory.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbRelegion.Text)) = 0 Then
                MessageBox.Show("Please Select Relegion.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbRelegion.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbAcadYear.Text)) = 0 Then
                MessageBox.Show("Please Select Academic Year.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbAcadYear.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbStatus.Text)) = 0 Then
                MessageBox.Show("Please Select Status.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbStatus.Focus()
                Exit Sub
            End If
            ' con = New OleDbConnection(cs)
            '  con.Open()
            '  cmd = con.CreateCommand()
            '  cmd.CommandText = "SELECT BedsAvailable FROM Room WHERE RoomNo= '" & cmbRoomNo.Text & "' and HostelName='" & cmbHostelName.Text & "' and BedsAvailable <= 0"
            '  rdr = cmd.ExecuteReader()
            '  If rdr.Read() Then
            'MessageBox.Show("Bed not Available in selected room no.", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            '  cmbRoomNo.Focus()
            ' If (rdr IsNot Nothing) Then
            'rdr.Close()
            '  End If
            '  Exit Sub
            '  End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con = New OleDbConnection(cs)
            con.Open()
            If (cmbHostelName.Text <> txtHostelerName.Text And cmbRoomNo.Text = txtRoomno.Text) Then
                Dim ct4 As String = "update room set BedsAvailable =  BedsAvailable - 1 where HostelName= '" & cmbHostelName.Text & "' and RoomNo='" & cmbRoomNo.Text & "'"
                Dim ct1 As String = "update room set BedsAvailable =  BedsAvailable + 1 where HostelName= '" & txtHostelname.Text & "' and RoomNo='" & txtRoomno.Text & "'"
                cmd = New OleDbCommand(ct4)
                Dim cmd1 As OleDbCommand
                cmd1 = New OleDbCommand(ct1)
                cmd.Connection = con
                cmd1.Connection = con
                cmd.ExecuteNonQuery()
                cmd1.ExecuteNonQuery()
                con.Close()
            End If
            con = New OleDbConnection(cs)
            con.Open()
            If (cmbHostelName.Text = txtHostelerName.Text And cmbRoomNo.Text <> txtRoomno.Text) Then
                Dim ct2 As String = "update room set  BedsAvailable =  BedsAvailable - 1 where HostelName= '" & cmbHostelName.Text & "' and RoomNo='" & cmbRoomNo.Text & "'"
                Dim ct1 As String = "update room set  BedsAvailable =  BedsAvailable + 1 where HostelName= '" & txtHostelname.Text & "' and RoomNo='" & txtRoomno.Text & "'"
                cmd = New OleDbCommand(ct2)
                Dim cmd1 As OleDbCommand
                cmd1 = New OleDbCommand(ct1)
                cmd.Connection = con
                cmd1.Connection = con
                cmd.ExecuteNonQuery()
                cmd1.ExecuteNonQuery()
                con.Close()
            End If
            con = New OleDbConnection(cs)
            con.Open()
            If (cmbHostelName.Text <> txtHostelerName.Text And cmbRoomNo.Text <> txtRoomno.Text) Then
                Dim ct3 As String = "update room set BedsAvailable =  BedsAvailable - 1 where HostelName= '" & cmbHostelName.Text & "' and RoomNo='" & cmbRoomNo.Text & "'"
                Dim ct1 As String = "update room set  BedsAvailable =  BedsAvailable + 1 where HostelName= '" & txtHostelname.Text & "' and RoomNo='" & txtRoomno.Text & "'"
                cmd = New OleDbCommand(ct3)
                Dim cmd1 As OleDbCommand
                cmd1 = New OleDbCommand(ct1)
                cmd.Connection = con
                cmd1.Connection = con
                cmd.ExecuteNonQuery()
                cmd1.ExecuteNonQuery()
                con.Close()
            End If
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT NoOfBeds,bedsavailable FROM Room WHERE RoomNo= '" & txtRoomno.Text & "' and HostelName='" & txtHostelname.Text & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtNoOfBeds.Text = rdr.GetInt32(0)
                txtBedAvailable.Text = rdr.GetInt32(1)
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con = New OleDbConnection(cs)
            con.Open()
            If (Val(txtBedAvailable.Text) > Val(txtNoOfBeds.Text)) Then
                Dim ct5 As String = "update room set BedsAvailable =  BedsAvailable + 1 where HostelName= '" & cmbHostelName.Text & "' and RoomNo='" & cmbRoomNo.Text & "'"
                Dim ct1 As String = "update room set  BedsAvailable =  BedsAvailable - 1 where HostelName= '" & txtHostelname.Text & "' and RoomNo='" & txtRoomno.Text & "'"
                cmd = New OleDbCommand(ct5)
                Dim cmd1 As OleDbCommand
                cmd1 = New OleDbCommand(ct1)
                cmd.Connection = con
                cmd1.Connection = con
                cmd.ExecuteNonQuery()
                cmd1.ExecuteNonQuery()
                MessageBox.Show("Unable to allocate room", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                con.Close()
                Exit Sub
            End If

            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "update Hostelers set HostelerName='" & txtHostelerName.Text & "',USN='" & txtUsn.Text & "',AcadYear='" & cmbAcadYear.Text & "',DOB=#" & dtpDOB.Text & "#,Gender='" & cmbGender.Text & "',RoomNo='" & cmbRoomNo.Text & "',HostelName='" & cmbHostelName.Text & "', DateOfJoining = #" & dtpDateOfJoining.Text & "# ,Purpose='" & cmbPurpose.Text & "',FatherName='" & txtFatherName.Text & "',MobNo1='" & txtMobNo1.Text & "',Phone1='" & txtPhone.Text & "',MotherName='" & txtMotherName.Text & "',MobNo2='" & txtMobNo2.Text & "',Address='" & txtAddress.Text & "',Email='" & txtEmail.Text & "',ContactNo='" & txtContactNo.Text & "',InstOfcDetails='" & txtInOfcDeatils.Text & "',Phone2='" & txtInstPhoneNo.Text & "',Agreement='" & txtAgreement.Text & "',GuardianName='" & txtGuardianName.Text & "',GuardianAddress='" & txtGuardianAddress.Text & "',MobNo3='" & txtGuardianContactNo.Text & "',Phone3='" & txtGuardianPhoneNo.Text & "',Photo=@d1,DocsPic=@d2,City='" & txtCity.Text & "',Relegion='" & cmbRelegion.Text & "',Caste='" & txtCaste.Text & "',Subcaste='" & txtSubCaste.Text & "',Category='" & cmbCategory.Text & "',Status='" & cmbStatus.Text & "' where HostelerID='" & txtHostelerID.Text & "'"
            Dim cb1 As String = "update CheckIn  set HostelerName='" & txtHostelerName.Text & "', HostelName='" & cmbHostelName.Text & "',RoomNo='" & cmbRoomNo.Text & "',Acadyear='" & cmbAcadYear.Text & "', CheckInDate= #" & dtpDateOfJoining.Text & "# where HostelerId= '" & txtHostelerID.Text & "' and Hostelname= '" & cmbHostelName.Text & "' and RoomNo='" & cmbRoomNo.Text & "'"
            cmd = New OleDbCommand(cb)
            Dim cmd2 As OleDbCommand
            cmd2 = New OleDbCommand(cb1)
            cmd.Connection = con
            cmd2.Connection = con

            Dim ms, ms1 As New MemoryStream()
            Dim bmpImage As New Bitmap(Picture.Image)
            Dim bmpImage1 As New Bitmap(PictureBox2.Image)
            bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            bmpImage1.Save(ms1, System.Drawing.Imaging.ImageFormat.Jpeg)
            Dim data As Byte() = ms.GetBuffer()
            Dim data1 As Byte() = ms1.GetBuffer()
            Dim p As New OleDbParameter("@d1", OleDbType.VarBinary)
            Dim p1 As New OleDbParameter("@d2", OleDbType.VarBinary)
            p.Value = data
            p1.Value = data1
            cmd.Parameters.Add(p)
            cmd.Parameters.Add(p1)
            cmd.ExecuteNonQuery()
            Autocomplete()
            cmd2.ExecuteNonQuery()
            con.Close()

            'con = New OleDbConnection(cs)
            ' con.Open()
            ' Dim cb1 As String = "update CheckIn set HostelerName='" & txtHostelerName.Text & "',HostelName='" & cmbHostelName.Text & "',RoomNo='" & cmbRoomNo.Text & "',acadyeaer='" & cmbAcadYear.Text & "', CheckInDate= #" & dtpDateOfJoining.Text & "# where HostelerId= '" & txtHostelerID.Text & "' and Acadyear='" & cmbAcadYear.Text & "'"

            ' cmd.Connection = con


            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate_record.Enabled = False
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtGuardianPhoneNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtGuardianPhoneNo.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            If MessageBox.Show("Do you really want to delete this record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                DeleteRecord()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If Len(Trim(cmbHostelName.Text)) = 0 Then
                MessageBox.Show("Please select hostel name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbHostelName.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbRoomNo.Text)) = 0 Then
                MessageBox.Show("Please select room no.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbRoomNo.Focus()
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT NoOfBeds FROM Room WHERE RoomNo= '" & cmbRoomNo.Text & "' and HostelName='" & cmbHostelName.Text & "' and BedsAvailable > 0"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                MessageBox.Show("Bed Available in selected room no.", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Bed not Available in selected room no.", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
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

    Private Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            With OpenFileDialog1
                .Filter = ("Images |*.png; *.bmp; *.jpg;*.jpeg; *.gif;")
                .FilterIndex = 4
            End With
            'Clear the file name
            OpenFileDialog1.FileName = ""
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                PictureBox2.Image = Image.FromFile(OpenFileDialog1.FileName)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    Private Sub txtGuardianContactNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtGuardianContactNo.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    

   
  
   
   
End Class