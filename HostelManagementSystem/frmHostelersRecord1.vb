Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO

Public Class frmHostelersRecord1
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"
    Sub Getdata()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],DOB,Gender,HostelName as [Hostel Name],RoomNo as [Room No],DateOfJoining as [Date Of Joining],Purpose as [Department],FatherName as [Father's Name],MobNo1 as [Mobile No],Phone1 as [Phone No],MotherName as [Mother's Name],MobNo2 as [Mobile No 2],City,Address,Email,ContactNo as [Contact No],InstOfcDetails as [Ins/Ofc Details],Phone2 as [Phone No 2],Agreement,GuardianName as  [Guardian Name],GuardianAddress as [Guardian Address],MobNo3 as [Guardian Mobile No],Phone3 as [Guardian Phone No],CompletionDate as [Completion Date],Photo,DocsPic as [Docs Pic],USN as [USN],Relegion as [Relegion],Caste as [Caste],Subcaste as [Sub Caste],Category as [Category], AcadYear as [Academic Year], Status as [Status] from Hostelers order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
   
    
    Sub fillCity()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct RTRIM(City) FROM Hostelers", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbCity.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbCity.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
   
    Sub fillBranch()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct RTRIM(Hostelname) FROM Hostelers", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbBranch.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbBranch.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmHostelersRecord1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub
    Private Sub frmHostelersRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Getdata()
        fillRoomNo()
        fillCity()

        fillBranch()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],DOB,Gender,HostelName as [Hostel Name],RoomNo as [Room No],DateOfJoining as [Date Of Joining],Purpose as [Department],FatherName as [Father's Name],MobNo1 as [Mobile No],Phone1 as [Phone No],MotherName as [Mother's Name],MobNo2 as [Mobile No 2],City,Address,Email,ContactNo as [Contact No],InstOfcDetails as [Ins/Ofc Details],Phone2 as [Phone No 2],Agreement,GuardianName as  [Guardian Name],GuardianAddress as [Guardian Address],MobNo3 as [Guardian Mobile No],Phone3 as [Guardian Phone No],CompletionDate as [Completion Date],Photo,DocsPic as [Docs Pic],USN as [USN],Relegion as [Relegion],Caste as [Caste],Subcaste as [Sub Caste],Category as [Category], AcadYear as [Academic Year], Status as [Status] from Hostelers where DateOfJoining between #" & DateFrom.Text & " # and #" & DateTo.Text & "# order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        txtUSN.Text = ""
        txtHostelerName.Text = ""
        DataGridView1.DataSource = Nothing
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DateFrom.Text = Today
        DateTo.Text = Today
        DataGridView2.DataSource = Nothing
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        cmbRoomNo.Text = ""
        DataGridView3.DataSource = Nothing
    End Sub
    Sub fillRoomNo()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct RTRIM(RoomNo) FROM Hostelers", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbRoomNo.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbRoomNo.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbRoomNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbRoomNo.SelectedIndexChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],DOB,Gender,HostelName as [Hostel Name],RoomNo as [Room No],DateOfJoining as [Date Of Joining],Purpose as [Department],FatherName as [Father's Name],MobNo1 as [Mobile No],Phone1 as [Phone No],MotherName as [Mother's Name],MobNo2 as [Mobile No 2],City,Address,Email,ContactNo as [Contact No],InstOfcDetails as [Ins/Ofc Details],Phone2 as [Phone No 2],Agreement,GuardianName as  [Guardian Name],GuardianAddress as [Guardian Address],MobNo3 as [Guardian Mobile No],Phone3 as [Guardian Phone No],CompletionDate as [Completion Date],Photo,DocsPic as [Docs Pic],USN as [USN],Relegion as [Relegion],Caste as [Caste],Subcaste as [Sub Caste],Category as [Category], AcadYear as [Academic Year], Status as [Status] from Hostelers where RoomNo='" & cmbRoomNo.Text & "' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView3.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbBranch_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbBranch.SelectedIndexChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],DOB,Gender,HostelName as [Hostel Name],RoomNo as [Room No],DateOfJoining as [Date Of Joining],Purpose as [Department],FatherName as [Father's Name],MobNo1 as [Mobile No],Phone1 as [Phone No],MotherName as [Mother's Name],MobNo2 as [Mobile No 2],City,Address,Email,ContactNo as [Contact No],InstOfcDetails as [Ins/Ofc Details],Phone2 as [Phone No 2],Agreement,GuardianName as  [Guardian Name],GuardianAddress as [Guardian Address],MobNo3 as [Guardian Mobile No],Phone3 as [Guardian Phone No],CompletionDate as [Completion Date],Photo,DocsPic as [Docs Pic],USN as [USN],Relegion as [Relegion],Caste as [Caste],Subcaste as [Sub Caste],Category as [Category], AcadYear as [Academic Year], Status as [Status] from Hostelers where HostelName='" & cmbBranch.Text & "' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView4.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        cmbBranch.Text = ""
        DataGridView4.DataSource = Nothing
    End Sub

    Private Sub TabControl1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.Click
        cmbBranch.Text = ""
        DataGridView4.DataSource = Nothing
        cmbRoomNo.Text = ""
        DataGridView3.DataSource = Nothing
        DateFrom.Text = Today
        DateTo.Text = Today
        DataGridView2.DataSource = Nothing
        txtUSN.Text = ""
        txtHostelerName.Text = ""
        DataGridView1.DataSource = Nothing
        DataGridView6.DataSource = Nothing
        DateTimePicker1.Text = Today
        DataGridView7.DataSource = Nothing
        cmbCity.Text = ""
        
    End Sub

    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            Me.Hide()
            frmHostelers.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmHostelers.txtHostelerID.Text = dr.Cells(0).Value.ToString()
            frmHostelers.txtHostelerName.Text = dr.Cells(1).Value.ToString()
            frmHostelers.dtpDOB.Text = dr.Cells(2).Value.ToString()
            frmHostelers.cmbGender.Text = dr.Cells(3).Value.ToString()
            frmHostelers.cmbHostelName.Text = dr.Cells(4).Value.ToString()
            frmHostelers.cmbRoomNo.Text = dr.Cells(5).Value.ToString()
            frmHostelers.txtHostelname.Text = dr.Cells(4).Value.ToString()
            frmHostelers.txtRoomno.Text = dr.Cells(5).Value.ToString()
            frmHostelers.dtpDateOfJoining.Text = dr.Cells(6).Value.ToString()
            frmHostelers.cmbPurpose.Text = dr.Cells(7).Value.ToString()
            frmHostelers.txtFatherName.Text = dr.Cells(8).Value.ToString()
            frmHostelers.txtMobNo1.Text = dr.Cells(9).Value.ToString()
            frmHostelers.txtPhone.Text = dr.Cells(10).Value.ToString()
            frmHostelers.txtMotherName.Text = dr.Cells(11).Value.ToString()
            frmHostelers.txtMobNo2.Text = dr.Cells(12).Value.ToString()
            frmHostelers.txtCity.Text = dr.Cells(13).Value.ToString()
            frmHostelers.txtAddress.Text = dr.Cells(14).Value.ToString()
            frmHostelers.txtEmail.Text = dr.Cells(15).Value.ToString()
            frmHostelers.txtContactNo.Text = dr.Cells(16).Value.ToString()
            frmHostelers.txtInOfcDeatils.Text = dr.Cells(17).Value.ToString()
            frmHostelers.txtInstPhoneNo.Text = dr.Cells(18).Value.ToString()
            frmHostelers.txtAgreement.Text = dr.Cells(19).Value.ToString()
            frmHostelers.txtGuardianName.Text = dr.Cells(20).Value.ToString()
            frmHostelers.txtGuardianAddress.Text = dr.Cells(21).Value.ToString()
            frmHostelers.txtGuardianContactNo.Text = dr.Cells(22).Value.ToString()
            frmHostelers.txtGuardianPhoneNo.Text = dr.Cells(23).Value.ToString()
            frmHostelers.dtpCompletionDate.Text = dr.Cells(24).Value.ToString()
            Dim data As Byte() = DirectCast(dr.Cells(25).Value, Byte())
            Dim ms As New MemoryStream(data)
            frmHostelers.Picture.Image = Image.FromStream(ms)
            Dim data1 As Byte() = DirectCast(dr.Cells(26).Value, Byte())
            Dim ms1 As New MemoryStream(data1)
            frmHostelers.PictureBox2.Image = Image.FromStream(ms1)
            frmHostelers.txtUsn.Text = dr.Cells(27).Value.ToString()
            frmHostelers.cmbRelegion.Text = dr.Cells(28).Value.ToString()
            frmHostelers.txtCaste.Text = dr.Cells(29).Value.ToString()
            frmHostelers.txtSubCaste.Text = dr.Cells(30).Value.ToString()
            frmHostelers.cmbCategory.Text = dr.Cells(31).Value.ToString()
            frmHostelers.cmbAcadYear.Text = dr.Cells(32).Value.ToString()
            frmHostelers.cmbStatus.Text = dr.Cells(33).Value.ToString()
            frmHostelers.btnUpdate_record.Enabled = True
            frmHostelers.btnDelete.Enabled = True
            frmHostelers.btnPrint.Enabled = True
            frmHostelers.btnSave.Enabled = False
            frmHostelers.btnReallocation.Enabled = True
            frmHostelers.txtAgreement.ReadOnly = True
            If Today > frmHostelers.dtpCompletionDate.Value Then
                frmHostelers.txtAgreement.ReadOnly = False
                frmHostelers.btnEA.Enabled = True
            End If
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT bedsavailable,NoOfBeds FROM Room WHERE RoomNo= '" & dr.Cells(4).Value.ToString() & "' and HostelName='" & dr.Cells(3).Value.ToString() & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                frmHostelers.txtNoOfBeds.Text = rdr.GetInt32(1)
                frmHostelers.txtBedAvailable.Text = rdr.GetInt32(0)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView2_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView2.SelectedRows(0)
            Me.Hide()
            frmHostelers.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmHostelers.txtHostelerID.Text = dr.Cells(0).Value.ToString()
            frmHostelers.txtHostelerName.Text = dr.Cells(1).Value.ToString()
            frmHostelers.dtpDOB.Text = dr.Cells(2).Value.ToString()
            frmHostelers.cmbGender.Text = dr.Cells(3).Value.ToString()
            frmHostelers.cmbHostelName.Text = dr.Cells(4).Value.ToString()
            frmHostelers.cmbRoomNo.Text = dr.Cells(5).Value.ToString()
            frmHostelers.txtHostelname.Text = dr.Cells(4).Value.ToString()
            frmHostelers.txtRoomno.Text = dr.Cells(5).Value.ToString()
            frmHostelers.dtpDateOfJoining.Text = dr.Cells(6).Value.ToString()
            frmHostelers.cmbPurpose.Text = dr.Cells(7).Value.ToString()
            frmHostelers.txtFatherName.Text = dr.Cells(8).Value.ToString()
            frmHostelers.txtMobNo1.Text = dr.Cells(9).Value.ToString()
            frmHostelers.txtPhone.Text = dr.Cells(10).Value.ToString()
            frmHostelers.txtMotherName.Text = dr.Cells(11).Value.ToString()
            frmHostelers.txtMobNo2.Text = dr.Cells(12).Value.ToString()
            frmHostelers.txtCity.Text = dr.Cells(13).Value.ToString()
            frmHostelers.txtAddress.Text = dr.Cells(14).Value.ToString()
            frmHostelers.txtEmail.Text = dr.Cells(15).Value.ToString()
            frmHostelers.txtContactNo.Text = dr.Cells(16).Value.ToString()
            frmHostelers.txtInOfcDeatils.Text = dr.Cells(17).Value.ToString()
            frmHostelers.txtInstPhoneNo.Text = dr.Cells(18).Value.ToString()
            frmHostelers.txtAgreement.Text = dr.Cells(19).Value.ToString()
            frmHostelers.txtGuardianName.Text = dr.Cells(20).Value.ToString()
            frmHostelers.txtGuardianAddress.Text = dr.Cells(21).Value.ToString()
            frmHostelers.txtGuardianContactNo.Text = dr.Cells(22).Value.ToString()
            frmHostelers.txtGuardianPhoneNo.Text = dr.Cells(23).Value.ToString()
            frmHostelers.dtpCompletionDate.Text = dr.Cells(24).Value.ToString()
            Dim data As Byte() = DirectCast(dr.Cells(25).Value, Byte())
            Dim ms As New MemoryStream(data)
            frmHostelers.Picture.Image = Image.FromStream(ms)
            Dim data1 As Byte() = DirectCast(dr.Cells(26).Value, Byte())
            Dim ms1 As New MemoryStream(data1)
            frmHostelers.PictureBox2.Image = Image.FromStream(ms1)
            frmHostelers.txtUsn.Text = dr.Cells(27).Value.ToString()
            frmHostelers.cmbRelegion.Text = dr.Cells(28).Value.ToString()
            frmHostelers.txtCaste.Text = dr.Cells(29).Value.ToString()
            frmHostelers.txtSubCaste.Text = dr.Cells(30).Value.ToString()
            frmHostelers.cmbCategory.Text = dr.Cells(31).Value.ToString()
            frmHostelers.cmbAcadYear.Text = dr.Cells(32).Value.ToString()
            frmHostelers.cmbStatus.Text = dr.Cells(33).Value.ToString()
            frmHostelers.btnUpdate_record.Enabled = True
            frmHostelers.btnDelete.Enabled = True
            frmHostelers.btnPrint.Enabled = True
            frmHostelers.btnSave.Enabled = False
            frmHostelers.btnReallocation.Enabled = True
            frmHostelers.txtAgreement.ReadOnly = True
            If Today > frmHostelers.dtpCompletionDate.Value Then
                frmHostelers.txtAgreement.ReadOnly = False
                frmHostelers.btnEA.Enabled = True
            End If
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT bedsavailable,NoOfBeds FROM Room WHERE RoomNo= '" & dr.Cells(4).Value.ToString() & "' and HostelName='" & dr.Cells(3).Value.ToString() & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                frmHostelers.txtNoOfBeds.Text = rdr.GetInt32(1)
                frmHostelers.txtBedAvailable.Text = rdr.GetInt32(0)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView3_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView3.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView3.SelectedRows(0)
            Me.Hide()
            frmHostelers.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmHostelers.txtHostelerID.Text = dr.Cells(0).Value.ToString()
            frmHostelers.txtHostelerName.Text = dr.Cells(1).Value.ToString()
            frmHostelers.dtpDOB.Text = dr.Cells(2).Value.ToString()
            frmHostelers.cmbGender.Text = dr.Cells(3).Value.ToString()
            frmHostelers.cmbHostelName.Text = dr.Cells(4).Value.ToString()
            frmHostelers.cmbRoomNo.Text = dr.Cells(5).Value.ToString()
            frmHostelers.txtHostelname.Text = dr.Cells(4).Value.ToString()
            frmHostelers.txtRoomno.Text = dr.Cells(5).Value.ToString()
            frmHostelers.dtpDateOfJoining.Text = dr.Cells(6).Value.ToString()
            frmHostelers.cmbPurpose.Text = dr.Cells(7).Value.ToString()
            frmHostelers.txtFatherName.Text = dr.Cells(8).Value.ToString()
            frmHostelers.txtMobNo1.Text = dr.Cells(9).Value.ToString()
            frmHostelers.txtPhone.Text = dr.Cells(10).Value.ToString()
            frmHostelers.txtMotherName.Text = dr.Cells(11).Value.ToString()
            frmHostelers.txtMobNo2.Text = dr.Cells(12).Value.ToString()
            frmHostelers.txtCity.Text = dr.Cells(13).Value.ToString()
            frmHostelers.txtAddress.Text = dr.Cells(14).Value.ToString()
            frmHostelers.txtEmail.Text = dr.Cells(15).Value.ToString()
            frmHostelers.txtContactNo.Text = dr.Cells(16).Value.ToString()
            frmHostelers.txtInOfcDeatils.Text = dr.Cells(17).Value.ToString()
            frmHostelers.txtInstPhoneNo.Text = dr.Cells(18).Value.ToString()
            frmHostelers.txtAgreement.Text = dr.Cells(19).Value.ToString()
            frmHostelers.txtGuardianName.Text = dr.Cells(20).Value.ToString()
            frmHostelers.txtGuardianAddress.Text = dr.Cells(21).Value.ToString()
            frmHostelers.txtGuardianContactNo.Text = dr.Cells(22).Value.ToString()
            frmHostelers.txtGuardianPhoneNo.Text = dr.Cells(23).Value.ToString()
            frmHostelers.dtpCompletionDate.Text = dr.Cells(24).Value.ToString()
            Dim data As Byte() = DirectCast(dr.Cells(25).Value, Byte())
            Dim ms As New MemoryStream(data)
            frmHostelers.Picture.Image = Image.FromStream(ms)
            Dim data1 As Byte() = DirectCast(dr.Cells(26).Value, Byte())
            Dim ms1 As New MemoryStream(data1)
            frmHostelers.PictureBox2.Image = Image.FromStream(ms1)
            frmHostelers.txtUsn.Text = dr.Cells(27).Value.ToString()
            frmHostelers.cmbRelegion.Text = dr.Cells(28).Value.ToString()
            frmHostelers.txtCaste.Text = dr.Cells(29).Value.ToString()
            frmHostelers.txtSubCaste.Text = dr.Cells(30).Value.ToString()
            frmHostelers.cmbCategory.Text = dr.Cells(31).Value.ToString()
            frmHostelers.cmbAcadYear.Text = dr.Cells(32).Value.ToString()
            frmHostelers.cmbStatus.Text = dr.Cells(33).Value.ToString()
            frmHostelers.btnUpdate_record.Enabled = True
            frmHostelers.btnDelete.Enabled = True
            frmHostelers.btnPrint.Enabled = True
            frmHostelers.btnSave.Enabled = False
            frmHostelers.btnReallocation.Enabled = True
            frmHostelers.txtAgreement.ReadOnly = True
            If Today > frmHostelers.dtpCompletionDate.Value Then
                frmHostelers.txtAgreement.ReadOnly = False
                frmHostelers.btnEA.Enabled = True
            End If
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT bedsavailable,NoOfBeds FROM Room WHERE RoomNo= '" & dr.Cells(4).Value.ToString() & "' and HostelName='" & dr.Cells(3).Value.ToString() & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                frmHostelers.txtNoOfBeds.Text = rdr.GetInt32(1)
                frmHostelers.txtBedAvailable.Text = rdr.GetInt32(0)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView4_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView4.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView4.SelectedRows(0)
            Me.Hide()
            frmHostelers.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmHostelers.txtHostelerID.Text = dr.Cells(0).Value.ToString()
            frmHostelers.txtHostelerName.Text = dr.Cells(1).Value.ToString()
            frmHostelers.dtpDOB.Text = dr.Cells(2).Value.ToString()
            frmHostelers.cmbGender.Text = dr.Cells(3).Value.ToString()
            frmHostelers.cmbHostelName.Text = dr.Cells(4).Value.ToString()
            frmHostelers.cmbRoomNo.Text = dr.Cells(5).Value.ToString()
            frmHostelers.txtHostelname.Text = dr.Cells(4).Value.ToString()
            frmHostelers.txtRoomno.Text = dr.Cells(5).Value.ToString()
            frmHostelers.dtpDateOfJoining.Text = dr.Cells(6).Value.ToString()
            frmHostelers.cmbPurpose.Text = dr.Cells(7).Value.ToString()
            frmHostelers.txtFatherName.Text = dr.Cells(8).Value.ToString()
            frmHostelers.txtMobNo1.Text = dr.Cells(9).Value.ToString()
            frmHostelers.txtPhone.Text = dr.Cells(10).Value.ToString()
            frmHostelers.txtMotherName.Text = dr.Cells(11).Value.ToString()
            frmHostelers.txtMobNo2.Text = dr.Cells(12).Value.ToString()
            frmHostelers.txtCity.Text = dr.Cells(13).Value.ToString()
            frmHostelers.txtAddress.Text = dr.Cells(14).Value.ToString()
            frmHostelers.txtEmail.Text = dr.Cells(15).Value.ToString()
            frmHostelers.txtContactNo.Text = dr.Cells(16).Value.ToString()
            frmHostelers.txtInOfcDeatils.Text = dr.Cells(17).Value.ToString()
            frmHostelers.txtInstPhoneNo.Text = dr.Cells(18).Value.ToString()
            frmHostelers.txtAgreement.Text = dr.Cells(19).Value.ToString()
            frmHostelers.txtGuardianName.Text = dr.Cells(20).Value.ToString()
            frmHostelers.txtGuardianAddress.Text = dr.Cells(21).Value.ToString()
            frmHostelers.txtGuardianContactNo.Text = dr.Cells(22).Value.ToString()
            frmHostelers.txtGuardianPhoneNo.Text = dr.Cells(23).Value.ToString()
            frmHostelers.dtpCompletionDate.Text = dr.Cells(24).Value.ToString()
            Dim data As Byte() = DirectCast(dr.Cells(25).Value, Byte())
            Dim ms As New MemoryStream(data)
            frmHostelers.Picture.Image = Image.FromStream(ms)
            Dim data1 As Byte() = DirectCast(dr.Cells(26).Value, Byte())
            Dim ms1 As New MemoryStream(data1)
            frmHostelers.PictureBox2.Image = Image.FromStream(ms1)
            frmHostelers.txtUsn.Text = dr.Cells(27).Value.ToString()
            frmHostelers.cmbRelegion.Text = dr.Cells(28).Value.ToString()
            frmHostelers.txtCaste.Text = dr.Cells(29).Value.ToString()
            frmHostelers.txtSubCaste.Text = dr.Cells(30).Value.ToString()
            frmHostelers.cmbCategory.Text = dr.Cells(31).Value.ToString()
            frmHostelers.cmbAcadYear.Text = dr.Cells(32).Value.ToString()
            frmHostelers.cmbStatus.Text = dr.Cells(33).Value.ToString()
            frmHostelers.btnUpdate_record.Enabled = True
            frmHostelers.btnDelete.Enabled = True
            frmHostelers.btnPrint.Enabled = True
            frmHostelers.btnSave.Enabled = False
            frmHostelers.btnReallocation.Enabled = True
            frmHostelers.txtAgreement.ReadOnly = True
            If Today > frmHostelers.dtpCompletionDate.Value Then
                frmHostelers.txtAgreement.ReadOnly = False
                frmHostelers.btnEA.Enabled = True
            End If
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT bedsavailable,NoOfBeds FROM Room WHERE RoomNo= '" & dr.Cells(4).Value.ToString() & "' and HostelName='" & dr.Cells(3).Value.ToString() & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                frmHostelers.txtNoOfBeds.Text = rdr.GetInt32(1)
                frmHostelers.txtBedAvailable.Text = rdr.GetInt32(0)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
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

    Private Sub DataGridView4_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView4.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView4.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView4.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))
    End Sub
    Private Sub txtUSN_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUSN.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],DOB,Gender,HostelName as [Hostel Name],RoomNo as [Room No],DateOfJoining as [Date Of Joining],Purpose as [Department],FatherName as [Father's Name],MobNo1 as [Mobile No],Phone1 as [Phone No],MotherName as [Mother's Name],MobNo2 as [Mobile No 2],City,Address,Email,ContactNo as [Contact No],InstOfcDetails as [Ins/Ofc Details],Phone2 as [Phone No 2],Agreement,GuardianName as  [Guardian Name],GuardianAddress as [Guardian Address],MobNo3 as [Guardian Mobile No],Phone3 as [Guardian Phone No],CompletionDate as [Completion Date],Photo,DocsPic as [Docs Pic],USN as [USN],Relegion as [Relegion],Caste as [Caste],Subcaste as [Sub Caste],Category as [Category], AcadYear as [Academic Year], Status as [Status] from Hostelers where USN like '%" & txtUSN.Text & "%' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub txtHostelerName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHostelerName.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],DOB,Gender,HostelName as [Hostel Name],RoomNo as [Room No],DateOfJoining as [Date Of Joining],Purpose as [Department],FatherName as [Father's Name],MobNo1 as [Mobile No],Phone1 as [Phone No],MotherName as [Mother's Name],MobNo2 as [Mobile No 2],City,Address,Email,ContactNo as [Contact No],InstOfcDetails as [Ins/Ofc Details],Phone2 as [Phone No 2],Agreement,GuardianName as  [Guardian Name],GuardianAddress as [Guardian Address],MobNo3 as [Guardian Mobile No],Phone3 as [Guardian Phone No],CompletionDate as [Completion Date],Photo,DocsPic as [Docs Pic],USN as [USN],Relegion as [Relegion],Caste as [Caste],Subcaste as [Sub Caste],Category as [Category], AcadYear as [Academic Year], Status as [Status] from Hostelers where HostelerName like '%" & txtHostelerName.Text & "%' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        DataGridView7.DataSource = Nothing
        cmbCity.Text = ""
    End Sub

    Private Sub cmbCity_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCity.SelectedIndexChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],DOB,Gender,HostelName as [Hostel Name],RoomNo as [Room No],DateOfJoining as [Date Of Joining],Purpose as [Department],FatherName as [Father's Name],MobNo1 as [Mobile No],Phone1 as [Phone No],MotherName as [Mother's Name],MobNo2 as [Mobile No 2],City,Address,Email,ContactNo as [Contact No],InstOfcDetails as [Ins/Ofc Details],Phone2 as [Phone No 2],Agreement,GuardianName as  [Guardian Name],GuardianAddress as [Guardian Address],MobNo3 as [Guardian Mobile No],Phone3 as [Guardian Phone No],CompletionDate as [Completion Date],Photo,DocsPic as [Docs Pic],USN as [USN],Relegion as [Relegion],Caste as [Caste],Subcaste as [Sub Caste],Category as [Category], AcadYear as [Academic Year], Status as [Status] from Hostelers where City='" & cmbCity.Text & "' order by HostelerName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView7.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        DataGridView6.DataSource = Nothing
        DateTimePicker1.Text = Today
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],DOB,Gender,HostelName as [Hostel Name],RoomNo as [Room No],DateOfJoining as [Date Of Joining],Purpose as [Department],FatherName as [Father's Name],MobNo1 as [Mobile No],Phone1 as [Phone No],MotherName as [Mother's Name],MobNo2 as [Mobile No 2],City,Address,Email,ContactNo as [Contact No],InstOfcDetails as [Ins/Ofc Details],Phone2 as [Phone No 2],Agreement,GuardianName as  [Guardian Name],GuardianAddress as [Guardian Address],MobNo3 as [Guardian Mobile No],Phone3 as [Guardian Phone No],CompletionDate as [Completion Date],Photo,DocsPic as [Docs Pic],USN as [USN],Relegion as [Relegion],Caste as [Caste],Subcaste as [Sub Caste],Category as [Category], AcadYear as [Academic Year], Status as [Status] from Hostelers WHERE MONTH(DOB) = " & DateTimePicker1.Value.Month & " AND DAY(DOB) = " & DateTimePicker1.Value.Day & "", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            DataGridView6.DataSource = myDataSet.Tables("Hostelers").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView6_RowHeaderMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView6.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView6.SelectedRows(0)
            Me.Hide()
            frmHostelers.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmHostelers.txtHostelerID.Text = dr.Cells(0).Value.ToString()
            frmHostelers.txtHostelerName.Text = dr.Cells(1).Value.ToString()
            frmHostelers.dtpDOB.Text = dr.Cells(2).Value.ToString()
            frmHostelers.cmbGender.Text = dr.Cells(3).Value.ToString()
            frmHostelers.cmbHostelName.Text = dr.Cells(4).Value.ToString()
            frmHostelers.cmbRoomNo.Text = dr.Cells(5).Value.ToString()
            frmHostelers.txtHostelname.Text = dr.Cells(4).Value.ToString()
            frmHostelers.txtRoomno.Text = dr.Cells(5).Value.ToString()
            frmHostelers.dtpDateOfJoining.Text = dr.Cells(6).Value.ToString()
            frmHostelers.cmbPurpose.Text = dr.Cells(7).Value.ToString()
            frmHostelers.txtFatherName.Text = dr.Cells(8).Value.ToString()
            frmHostelers.txtMobNo1.Text = dr.Cells(9).Value.ToString()
            frmHostelers.txtPhone.Text = dr.Cells(10).Value.ToString()
            frmHostelers.txtMotherName.Text = dr.Cells(11).Value.ToString()
            frmHostelers.txtMobNo2.Text = dr.Cells(12).Value.ToString()
            frmHostelers.txtCity.Text = dr.Cells(13).Value.ToString()
            frmHostelers.txtAddress.Text = dr.Cells(14).Value.ToString()
            frmHostelers.txtEmail.Text = dr.Cells(15).Value.ToString()
            frmHostelers.txtContactNo.Text = dr.Cells(16).Value.ToString()
            frmHostelers.txtInOfcDeatils.Text = dr.Cells(17).Value.ToString()
            frmHostelers.txtInstPhoneNo.Text = dr.Cells(18).Value.ToString()
            frmHostelers.txtAgreement.Text = dr.Cells(19).Value.ToString()
            frmHostelers.txtGuardianName.Text = dr.Cells(20).Value.ToString()
            frmHostelers.txtGuardianAddress.Text = dr.Cells(21).Value.ToString()
            frmHostelers.txtGuardianContactNo.Text = dr.Cells(22).Value.ToString()
            frmHostelers.txtGuardianPhoneNo.Text = dr.Cells(23).Value.ToString()
            frmHostelers.dtpCompletionDate.Text = dr.Cells(24).Value.ToString()
            Dim data As Byte() = DirectCast(dr.Cells(25).Value, Byte())
            Dim ms As New MemoryStream(data)
            frmHostelers.Picture.Image = Image.FromStream(ms)
            Dim data1 As Byte() = DirectCast(dr.Cells(26).Value, Byte())
            Dim ms1 As New MemoryStream(data1)
            frmHostelers.PictureBox2.Image = Image.FromStream(ms1)
            frmHostelers.txtUsn.Text = dr.Cells(27).Value.ToString()
            frmHostelers.cmbRelegion.Text = dr.Cells(28).Value.ToString()
            frmHostelers.txtCaste.Text = dr.Cells(29).Value.ToString()
            frmHostelers.txtSubCaste.Text = dr.Cells(30).Value.ToString()
            frmHostelers.cmbCategory.Text = dr.Cells(31).Value.ToString()
            frmHostelers.cmbAcadYear.Text = dr.Cells(32).Value.ToString()
            frmHostelers.cmbStatus.Text = dr.Cells(33).Value.ToString()
            frmHostelers.btnUpdate_record.Enabled = True
            frmHostelers.btnDelete.Enabled = True
            frmHostelers.btnPrint.Enabled = True
            frmHostelers.btnSave.Enabled = False
            frmHostelers.btnReallocation.Enabled = True
            frmHostelers.txtAgreement.ReadOnly = True
            If Today > frmHostelers.dtpCompletionDate.Value Then
                frmHostelers.txtAgreement.ReadOnly = False
                frmHostelers.btnEA.Enabled = True
            End If
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT bedsavailable,NoOfBeds FROM Room WHERE RoomNo= '" & dr.Cells(4).Value.ToString() & "' and HostelName='" & dr.Cells(3).Value.ToString() & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                frmHostelers.txtNoOfBeds.Text = rdr.GetInt32(1)
                frmHostelers.txtBedAvailable.Text = rdr.GetInt32(0)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView7_RowHeaderMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView7.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView7.SelectedRows(0)
            Me.Hide()
            frmHostelers.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmHostelers.txtHostelerID.Text = dr.Cells(0).Value.ToString()
            frmHostelers.txtHostelerName.Text = dr.Cells(1).Value.ToString()
            frmHostelers.dtpDOB.Text = dr.Cells(2).Value.ToString()
            frmHostelers.cmbGender.Text = dr.Cells(3).Value.ToString()
            frmHostelers.cmbHostelName.Text = dr.Cells(4).Value.ToString()
            frmHostelers.cmbRoomNo.Text = dr.Cells(5).Value.ToString()
            frmHostelers.txtHostelname.Text = dr.Cells(4).Value.ToString()
            frmHostelers.txtRoomno.Text = dr.Cells(5).Value.ToString()
            frmHostelers.dtpDateOfJoining.Text = dr.Cells(6).Value.ToString()
            frmHostelers.cmbPurpose.Text = dr.Cells(7).Value.ToString()
            frmHostelers.txtFatherName.Text = dr.Cells(8).Value.ToString()
            frmHostelers.txtMobNo1.Text = dr.Cells(9).Value.ToString()
            frmHostelers.txtPhone.Text = dr.Cells(10).Value.ToString()
            frmHostelers.txtMotherName.Text = dr.Cells(11).Value.ToString()
            frmHostelers.txtMobNo2.Text = dr.Cells(12).Value.ToString()
            frmHostelers.txtCity.Text = dr.Cells(13).Value.ToString()
            frmHostelers.txtAddress.Text = dr.Cells(14).Value.ToString()
            frmHostelers.txtEmail.Text = dr.Cells(15).Value.ToString()
            frmHostelers.txtContactNo.Text = dr.Cells(16).Value.ToString()
            frmHostelers.txtInOfcDeatils.Text = dr.Cells(17).Value.ToString()
            frmHostelers.txtInstPhoneNo.Text = dr.Cells(18).Value.ToString()
            frmHostelers.txtAgreement.Text = dr.Cells(19).Value.ToString()
            frmHostelers.txtGuardianName.Text = dr.Cells(20).Value.ToString()
            frmHostelers.txtGuardianAddress.Text = dr.Cells(21).Value.ToString()
            frmHostelers.txtGuardianContactNo.Text = dr.Cells(22).Value.ToString()
            frmHostelers.txtGuardianPhoneNo.Text = dr.Cells(23).Value.ToString()
            frmHostelers.dtpCompletionDate.Text = dr.Cells(24).Value.ToString()
            Dim data As Byte() = DirectCast(dr.Cells(25).Value, Byte())
            Dim ms As New MemoryStream(data)
            frmHostelers.Picture.Image = Image.FromStream(ms)
            Dim data1 As Byte() = DirectCast(dr.Cells(26).Value, Byte())
            Dim ms1 As New MemoryStream(data1)
            frmHostelers.PictureBox2.Image = Image.FromStream(ms1)
            frmHostelers.txtUsn.Text = dr.Cells(27).Value.ToString()
            frmHostelers.cmbRelegion.Text = dr.Cells(28).Value.ToString()
            frmHostelers.txtCaste.Text = dr.Cells(29).Value.ToString()
            frmHostelers.txtSubCaste.Text = dr.Cells(30).Value.ToString()
            frmHostelers.cmbCategory.Text = dr.Cells(31).Value.ToString()
            frmHostelers.cmbAcadYear.Text = dr.Cells(32).Value.ToString()
            frmHostelers.cmbStatus.Text = dr.Cells(33).Value.ToString()
            frmHostelers.btnUpdate_record.Enabled = True
            frmHostelers.btnDelete.Enabled = True
            frmHostelers.btnPrint.Enabled = True
            frmHostelers.btnSave.Enabled = False
            frmHostelers.btnReallocation.Enabled = True
            frmHostelers.txtAgreement.ReadOnly = True
            If Today > frmHostelers.dtpCompletionDate.Value Then
                frmHostelers.txtAgreement.ReadOnly = False
                frmHostelers.btnEA.Enabled = True
            End If
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT bedsavailable,NoOfBeds FROM Room WHERE RoomNo= '" & dr.Cells(4).Value.ToString() & "' and HostelName='" & dr.Cells(3).Value.ToString() & "'"
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                frmHostelers.txtNoOfBeds.Text = rdr.GetInt32(1)
                frmHostelers.txtBedAvailable.Text = rdr.GetInt32(0)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView6_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView6.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView6.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView6.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub DataGridView7_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView7.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView7.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView7.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

   


   


   

   
    
End Class