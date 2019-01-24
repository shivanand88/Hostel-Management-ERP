Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class frmAccountsRecord
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"

    Sub fillHostelName()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT HostelName FROM Hostel", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            For Each drow As DataRow In dtable.Rows
                cmbHostelName.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub cmbHostelName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbHostelName.SelectedIndexChanged
        cmbRoomNo1.Items.Clear()
        cmbRoomNo2.Items.Clear()
        cmbRoomNo1.Text = ""
        cmbRoomNo2.Text = ""
        cmbAcadYear2.Text = ""
        cmbRoomNo1.Enabled = True
        cmbRoomNo2.Enabled = True
        cmbAcadYear2.Enabled = True
        cmbRoomNo1.Focus()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select distinct RTRIM(RoomNo) from Room where HostelName= '" & cmbHostelName.Text & "'"
            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()
            While rdr.Read()
                cmbRoomNo1.Items.Add(rdr(0))
                cmbRoomNo2.Items.Add(rdr(0))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub frmAccountsRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillHostelName()

    End Sub
    Private Sub frmAccountsRecord_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub

    

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If DataGridView1.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = DataGridView1.RowCount - 1
            colsTotal = DataGridView1.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView1.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView1.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        
        If Len(Trim(cmbHostelType2.Text)) = 0 Then
            MessageBox.Show("Please select Hostel Name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbHostelType2.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbAcadYear.Text)) = 0 Then
            MessageBox.Show("Please select Academic Year", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbAcadYear.Focus()
            Exit Sub
        End If
        Try
            
            GroupBox11.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("SELECT Hostelers.HostelerName as [Hosteler Name] , Hostelers.USN as [USN] , RoomAllotment.Hostelname, Hostel.HostelType as [Hostel Name] , RoomAllotment.RoomNo  as [Room No] , RoomAllotment.AcadYear as [Academic Year] ,  (FeePayment.ServiceCharges- RegCharges.PrevDue) as [Total] , RegCharges.PrevDue as[Prev Due] ,FeePayment.ServiceCharges as [Total + Prev Due],  sum(FeePayment.Totalpaid) as [Total paid] FROM (Hostel INNER JOIN ((Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID]) INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where   Hostel.HostelType ='" & cmbHostelType2.Text & "' and RoomAllotment.AcadYear = '" & cmbAcadYear.Text & "'  and FeePayment.AcadYear=RoomAllotment.AcadYear and FeePayment.AcadYear=RegCharges.AcadYear and RoomAllotment.AcadYear=RegCharges.AcadYear group by HostelerName,USn,purpose, RoomAllotment.Hostelname,RoomAllotment.RoomNo,ServiceCharges,RoomAllotment.AcadYear,RegCharges.PrevDue,HostelType order by Hostelername ", con)
            ' cmd = New OleDbCommand("SELECT DISTINCT  Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], RoomAllotment.Hostelname,Hostel as [Hostel Name].HostelType as [Hostel Type] , RoomAllotment.RoomNo as [Room No], RoomAllotment.AcadYear as [Academic Year], RegCharges.CautionMoney as [Deposit], RegCharges.RentalCharges as [Room Rent], RegCharges.FormFee as [Form Fee],  ( RegCharges.CautionMoney + RegCharges.RentalCharges + RegCharges.FormFee )as [Total ], RegCharges.PrevDue as [Prev Due],  RegCharges.TotalCharges as [Total+Due], SUM(FeePayment.TotalPaid) as [Total Paid] FROM ((Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID])  INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where  Hostel.HostelType ='" & cmbHostelType2.Text & "' and RoomAllotment.AcadYear = '" & cmbAcadYear.Text & "'  and FeePayment.AcadYear=RoomAllotment.AcadYear and  FeePayment.AcadYear = RegCharges.AcadYear And RoomAllotment.AcadYear = RegCharges.AcadYear  group by Hostelers.HostelerName,Hostelers.HostelerName,  Hostelers.USN, RoomAllotment.Hostelname, RoomAllotment.RoomNo, RoomAllotment.AcadYear, RegCharges.CautionMoney,   RegCharges.RentalCharges,  RegCharges.FormFee,RegCharges.PrevDue, RegCharges.TotalCharges order by Hostelername ", con)

            cmd = New OleDbCommand("SELECT DISTINCT   Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], RoomAllotment.Hostelname as [Hostel Name],  RoomAllotment.RoomNo as [Room No], RoomAllotment.AcadYear as [Acad Year], RegCharges.CautionMoney as [Deposit], RegCharges.RentalCharges as [Room Rent], RegCharges.FormFee as [Form Fee], RegCharges.OtherFee as [Other Fee],( RegCharges.CautionMoney + RegCharges.RentalCharges + RegCharges.FormFee + RegCharges.OtherFee )as [Total ], RegCharges.PrevDue as [Prev Due], RegCharges.TotalCharges as [Total + Prev Due], sum(FeePayment.TotalPaid) as [Current Paid] FROM (Hostel INNER JOIN ((Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID]) INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where Hostel.HostelType ='" & cmbHostelType2.Text & "' and RoomAllotment.AcadYear = '" & cmbAcadYear.Text & "' and FeePayment.AcadYear=RoomAllotment.AcadYear and FeePayment.AcadYear=RegCharges.AcadYear and RoomAllotment.AcadYear=RegCharges.AcadYear  group by  Hostelers.HostelerName,Hostelers.HostelerName, Hostelers.USN, RoomAllotment.Hostelname, Hostel.HostelType,RoomAllotment.RoomNo,   RoomAllotment.AcadYear, RegCharges.CautionMoney, RegCharges.RentalCharges, RegCharges.FormFee,RegCharges.OtherFee,RegCharges.PrevDue, RegCharges.TotalCharges", con)

            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "FeePayment")
            myDA.Fill(myDataSet, "RoomAllotment")
            myDA.Fill(myDataSet, "RegCharges")
            myDA.Fill(myDataSet, "Hostel")
            DataGridView3.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView3.DataSource = myDataSet.Tables("FeePayment").DefaultView
            DataGridView3.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
            DataGridView3.DataSource = myDataSet.Tables("RegCharges").DefaultView
            DataGridView3.DataSource = myDataSet.Tables("Hostel").DefaultView
            Dim sum As Int64 = 0
            Dim sum1 As Int64 = 0

            For Each r As DataGridViewRow In Me.DataGridView3.Rows
                If (DataGridView3.Rows.Count <= 2) Then
                    If (r.Cells(10).Value > 0) Then
                        sum = sum + r.Cells(12).Value
                        sum1 = sum1 + ((r.Cells(9).Value + r.Cells(10).Value) - r.Cells(12).Value)
                    Else
                        sum = sum + r.Cells(12).Value
                        sum1 = sum1 + (r.Cells(9).Value - r.Cells(12).Value)
                    End If
                Else
                    sum = sum + r.Cells(12).Value
                    sum1 = sum1 + (r.Cells(9).Value - r.Cells(12).Value)
                End If

            Next
            TextBox3.Text = sum
            Label17.Text = frmDueCharges.GetInWords(sum)
            TextBox8.Text = sum1

            con.Close()
            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("SELECT Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN],Concession.AcadYear as [Acadeic Year], Concession.Con_Amount as [Concession Amount] FROM Hostelers INNER JOIN Concession ON Hostelers.[HostelerID] = Concession.[HostelerID] ", con)
            cmd = New OleDbCommand(" SELECT Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], RoomAllotment.Hostelname as [Hostel Name] ,  Concession.AcadYear as [Academic Year], Concession.Con_Amount as [Concession Amt], Concession.Con_Date as [Concession Date] FROM (Hostel INNER JOIN (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN Concession ON Hostelers.[HostelerID] = Concession.[HostelerID] where Hostel.HostelType ='" & cmbHostelType2.Text & "' and Concession.AcadYear = '" & cmbAcadYear.Text & "' and Hostelers.AcadYear=RoomAllotment.AcadYear ", con)
            Dim myDA1 As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet1 As DataSet = New DataSet()
            myDA1.Fill(myDataSet1, "Hostelers")
            myDA1.Fill(myDataSet1, "Hostel")
            myDA1.Fill(myDataSet1, "RoomAllotment")
            myDA1.Fill(myDataSet1, "Concession")
            DataGridView9.DataSource = myDataSet1.Tables("Hostelers").DefaultView
            DataGridView9.DataSource = myDataSet1.Tables("Hostel").DefaultView
            DataGridView9.DataSource = myDataSet1.Tables("RoomAllotment").DefaultView
            DataGridView9.DataSource = myDataSet1.Tables("Concession").DefaultView
            Dim sum5 As Int64 = 0
            For Each r1 As DataGridViewRow In Me.DataGridView9.Rows
                sum5 = sum5 + (r1.Cells(4).Value)
            Next
            TextBox14.Text = sum5
            Label36.Text = frmDueCharges.GetInWords(sum5)
            Dim sum6 As Int64 = 0
            sum6 = sum1 - sum5
            TextBox8.Text = sum6
            Label18.Text = frmDueCharges.GetInWords(sum6)
            con.Close()

        Catch ex As Exception
            MessageBox.Show("No Record Found!.")
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        If Len(Trim(cmbHostelType1.Text)) = 0 Then
            MessageBox.Show("Please select Hostel Name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbHostelType1.Focus()
            Exit Sub
        End If
        Try
            GroupBox7.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("SELECT Hostelers.HostelerName as [Hosteler Name] , Hostelers.USN as [USN] , RoomAllotment.Hostelname as [Hostel Name], Hostel.HostelType as [Hostel Type] , RoomAllotment.RoomNo  as [Room No] , RoomAllotment.AcadYear as [Academic Year] ,  (FeePayment.ServiceCharges- RegCharges.PrevDue) as [Total] , RegCharges.PrevDue as[Prev Due] ,FeePayment.ServiceCharges as [Total + Prev Due],  sum(FeePayment.Totalpaid) as [Total paid] FROM (Hostel INNER JOIN ((Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID]) INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where   Hostel.HostelType ='" & cmbHostelType1.Text & "' and FeePayment.Paymentdate between #" & DateFrom.Text & " # and #" & DateTo.Text & "# and FeePayment.AcadYear=RoomAllotment.AcadYear and FeePayment.AcadYear=RegCharges.AcadYear and RoomAllotment.AcadYear=RegCharges.AcadYear group by HostelerName,USn,purpose, RoomAllotment.Hostelname,RoomAllotment.RoomNo,ServiceCharges,RoomAllotment.AcadYear,RegCharges.PrevDue,HostelType order by Hostelername ", con)
            cmd = New OleDbCommand("SELECT DISTINCT   Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], RoomAllotment.Hostelname as [Hostel Name],  RoomAllotment.RoomNo as [Room No], RoomAllotment.AcadYear as [Acad Year], RegCharges.CautionMoney as [Deposit], RegCharges.RentalCharges as [Room Rent], RegCharges.FormFee as [Form Fee], RegCharges.OtherFee as [Other Fee],( RegCharges.CautionMoney + RegCharges.RentalCharges + RegCharges.FormFee + RegCharges.OtherFee )as [Total ], RegCharges.PrevDue as [Prev Due], RegCharges.TotalCharges as [Total + Prev Due], sum(FeePayment.TotalPaid) as [Current Paid] FROM (Hostel INNER JOIN ((Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID]) INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where Hostel.HostelType ='" & cmbHostelType1.Text & "' and FeePayment.Paymentdate between #" & DateFrom.Text & " # and #" & DateTo.Text & "# and FeePayment.AcadYear=RoomAllotment.AcadYear and FeePayment.AcadYear=RegCharges.AcadYear and RoomAllotment.AcadYear=RegCharges.AcadYear  group by  Hostelers.HostelerName,Hostelers.HostelerName, Hostelers.USN, RoomAllotment.Hostelname, Hostel.HostelType,RoomAllotment.RoomNo,   RoomAllotment.AcadYear, RegCharges.CautionMoney, RegCharges.RentalCharges, RegCharges.FormFee,RegCharges.OtherFee,RegCharges.PrevDue, RegCharges.TotalCharges", con)

            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "FeePayment")
            myDA.Fill(myDataSet, "RoomAllotment")
            myDA.Fill(myDataSet, "RegCharges")
            myDA.Fill(myDataSet, "Hostel")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView2.DataSource = myDataSet.Tables("FeePayment").DefaultView
            DataGridView2.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
            DataGridView2.DataSource = myDataSet.Tables("RegCharges").DefaultView
            DataGridView2.DataSource = myDataSet.Tables("Hostel").DefaultView
            Dim sum2 As Int64 = 0
            Dim sum3 As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView2.Rows
                If (DataGridView2.Rows.Count <= 2) Then
                    If (r.Cells(10).Value > 0) Then
                        sum2 = sum2 + r.Cells(12).Value
                        sum3 = sum3 + ((r.Cells(9).Value + r.Cells(10).Value) - r.Cells(12).Value)
                    Else
                        sum2 = sum2 + r.Cells(12).Value
                        sum3 = sum3 + (r.Cells(9).Value - r.Cells(12).Value)
                    End If
                Else
                    sum2 = sum2 + r.Cells(12).Value
                    sum3 = sum3 + (r.Cells(9).Value - r.Cells(12).Value)
                End If
            Next



            TextBox2.Text = sum2
            Label15.Text = frmDueCharges.GetInWords(sum2)
            TextBox6.Text = sum3

            con.Close()

            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("SELECT Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN],Concession.AcadYear as [Acadeic Year], Concession.Con_Amount as [Concession Amount] FROM Hostelers INNER JOIN Concession ON Hostelers.[HostelerID] = Concession.[HostelerID] ", con)
            cmd = New OleDbCommand(" SELECT Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], RoomAllotment.Hostelname as [Hostel Name] , Concession.AcadYear as [Acad Year], Concession.Con_Amount as [Concession Amt], Concession.Con_Date as [Concession Date] FROM (Hostel INNER JOIN (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN Concession ON Hostelers.[HostelerID] = Concession.[HostelerID] where Hostel.HostelType ='" & cmbHostelType1.Text & "' and Concession.Con_Date between #" & DateFrom.Text & " # and #" & DateTo.Text & "# and Hostelers.AcadYear=RoomAllotment.AcadYear ", con)
            Dim myDA1 As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet1 As DataSet = New DataSet()
            myDA1.Fill(myDataSet1, "Hostelers")
            myDA1.Fill(myDataSet1, "Hostel")
            myDA1.Fill(myDataSet1, "RoomAllotment")
            myDA1.Fill(myDataSet1, "Concession")
            DataGridView8.DataSource = myDataSet1.Tables("Hostelers").DefaultView
            DataGridView8.DataSource = myDataSet1.Tables("Hostel").DefaultView
            DataGridView8.DataSource = myDataSet1.Tables("RoomAllotment").DefaultView
            DataGridView8.DataSource = myDataSet1.Tables("Concession").DefaultView
            Dim sum5 As Int64 = 0
            For Each r1 As DataGridViewRow In Me.DataGridView8.Rows
                sum5 = sum5 + (r1.Cells(4).Value)
            Next
            TextBox13.Text = sum5
            Label33.Text = frmDueCharges.GetInWords(sum5)
            Dim sum6 As Int64 = 0
            sum6 = sum3 - sum5
            TextBox6.Text = sum6
            Label16.Text = frmDueCharges.GetInWords(sum6)
            con.Close()
       
        Catch ex As Exception
            MessageBox.Show("No Record Found!.")
            ' MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        If Len(Trim(cmbAcadYear1.Text)) = 0 Then
            MessageBox.Show("Please select Academic Year", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbAcadYear1.Focus()
            Exit Sub
        End If

        Try

            GroupBox18.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            '      cmd = New OleDbCommand("SELECT Hostelers.HostelerName as [Hosteler Name] , Hostelers.USN as [USN] , RoomAllotment.Hostelname, Hostel.HostelType as [Hostel Name] , RoomAllotment.RoomNo  as [Room No] , RoomAllotment.AcadYear as [Academic Year] ,  (FeePayment.ServiceCharges- RegCharges.PrevDue) as [Total] , RegCharges.PrevDue as[Prev Due] ,FeePayment.ServiceCharges as [Total + Prev Due],  sum(FeePayment.Totalpaid) as [Total paid] FROM (Hostel INNER JOIN ((Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID]) INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where   RoomAllotment.AcadYear = '" & cmbAcadYear1.Text & "'  and FeePayment.AcadYear=RoomAllotment.AcadYear and FeePayment.AcadYear=RegCharges.AcadYear and RoomAllotment.AcadYear=RegCharges.AcadYear group by HostelerName,USn,purpose, RoomAllotment.Hostelname,RoomAllotment.RoomNo,ServiceCharges,RoomAllotment.AcadYear,RegCharges.PrevDue,HostelType order by Hostelername ", con)
            cmd = New OleDbCommand("SELECT DISTINCT   Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], RoomAllotment.Hostelname as [Hostel Name], RoomAllotment.RoomNo as [Room No], RoomAllotment.AcadYear as [Acad Year], RegCharges.CautionMoney as [Deposit], RegCharges.RentalCharges as [Room Rent], RegCharges.FormFee as [Form Fee], RegCharges.OtherFee as [Other Fee],( RegCharges.CautionMoney + RegCharges.RentalCharges + RegCharges.FormFee + RegCharges.OtherFee)as [Total ], RegCharges.PrevDue as [Prev Due], RegCharges.TotalCharges as [Total + Prev Due], sum(FeePayment.TotalPaid) as [Current Paid] FROM (Hostel INNER JOIN ((Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID]) INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where  RoomAllotment.AcadYear = '" & cmbAcadYear1.Text & "' and FeePayment.AcadYear=RoomAllotment.AcadYear and FeePayment.AcadYear=RegCharges.AcadYear and RoomAllotment.AcadYear=RegCharges.AcadYear  group by  Hostelers.HostelerName,Hostelers.HostelerName, Hostelers.USN, RoomAllotment.Hostelname, Hostel.HostelType,RoomAllotment.RoomNo,   RoomAllotment.AcadYear, RegCharges.CautionMoney, RegCharges.RentalCharges, RegCharges.FormFee,RegCharges.OtherFee,RegCharges.PrevDue, RegCharges.TotalCharges", con)


            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "FeePayment")
            myDA.Fill(myDataSet, "RoomAllotment")
            myDA.Fill(myDataSet, "RegCharges")
            myDA.Fill(myDataSet, "Hostel")
            DataGridView5.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView5.DataSource = myDataSet.Tables("FeePayment").DefaultView
            DataGridView5.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
            DataGridView5.DataSource = myDataSet.Tables("RegCharges").DefaultView
            DataGridView5.DataSource = myDataSet.Tables("Hostel").DefaultView
            Dim sum As Int64 = 0
            Dim sum1 As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView5.Rows
                If (DataGridView5.Rows.Count <= 2) Then
                    If (r.Cells(10).Value > 0) Then

                        sum = sum + r.Cells(12).Value
                        sum1 = sum1 + ((r.Cells(9).Value + r.Cells(10).Value) - r.Cells(12).Value)
                    Else
                        sum = sum + r.Cells(12).Value
                        sum1 = sum1 + (r.Cells(9).Value - r.Cells(12).Value)
                    End If
                Else
                    If (r.Cells(10).Value > 0) Then

                        sum = sum + r.Cells(12).Value
                        sum1 = sum1 + ((r.Cells(9).Value + r.Cells(10).Value) - r.Cells(12).Value)
                    Else
                        sum = sum + r.Cells(12).Value
                        sum1 = sum1 + (r.Cells(9).Value - r.Cells(12).Value)
                    End If
                End If
            Next
            TextBox5.Text = sum
            Label19.Text = frmDueCharges.GetInWords(sum)
            TextBox7.Text = sum1

            con.Close()
            con.Open()
            'cmd = New OleDbCommand("SELECT Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN],Concession.AcadYear as [Acadeic Year], Concession.Con_Amount as [Concession Amount] FROM Hostelers INNER JOIN Concession ON Hostelers.[HostelerID] = Concession.[HostelerID] ", con)
            cmd = New OleDbCommand(" SELECT Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], RoomAllotment.Hostelname as [Hostel Name] ,Concession.AcadYear as [Acad Year], Concession.Con_Amount as [Concession Amt], Concession.Con_Date as [Concession Date] FROM (Hostel INNER JOIN (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN Concession ON Hostelers.[HostelerID] = Concession.[HostelerID] where Concession.AcadYear= '" & cmbAcadYear1.Text & "' and Hostelers.AcadYear=RoomAllotment.AcadYear ", con)
            Dim myDA1 As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet1 As DataSet = New DataSet()
            myDA1.Fill(myDataSet1, "Hostelers")
            myDA1.Fill(myDataSet1, "Hostel")
            myDA1.Fill(myDataSet1, "RoomAllotment")
            myDA1.Fill(myDataSet1, "Concession")
            DataGridView10.DataSource = myDataSet1.Tables("Hostelers").DefaultView
            DataGridView10.DataSource = myDataSet1.Tables("Hostel").DefaultView
            DataGridView10.DataSource = myDataSet1.Tables("RoomAllotment").DefaultView
            DataGridView10.DataSource = myDataSet1.Tables("Concession").DefaultView
            Dim sum5 As Int64 = 0
            For Each r1 As DataGridViewRow In Me.DataGridView10.Rows
                sum5 = sum5 + (r1.Cells(4).Value)
            Next
            TextBox15.Text = sum5
            Label39.Text = frmDueCharges.GetInWords(sum5)
            Dim sum6 As Int64 = 0
            sum6 = sum1 - sum5
            TextBox7.Text = sum6
            Label20.Text = frmDueCharges.GetInWords(sum6)
            con.Close()
        Catch ex As Exception
            MessageBox.Show("No Record Found!.")
            ' MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try
            If Len(Trim(cmbHostelType.Text)) = 0 Then
                MessageBox.Show("Please select Hostel Name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbHostelType.Focus()
                Exit Sub
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptAccountrecord 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = " SELECT Hostelers.HostelerName, Hostelers.USN, RoomAllotment.Hostelname, RoomAllotment.RoomNo, FeePayment.AcadYear, FeePayment.ServiceCharges, FeePayment.TotalPaid FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where  FeePayment.AcadYear=RoomAllotment.AcadYear"
            'MyCommand.CommandText = "SELECT distinct Hostelers.HostelerName , Hostelers.USN, RoomAllotment.Hostelname, RoomAllotment.RoomNo, RoomAllotment.AcadYear,  FeePayment.ServiceCharges ,FeePayment.Totalpaid FROM ((Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID]) INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where  RoomAllotment.HostelName='" & cmbHostelName.Text & "' FeePayment.AcadYear=RoomAllotment.AcadYear and FeePayment.AcadYear=RegCharges.AcadYear and RoomAllotment.AcadYear=RegCharges.AcadYear  order by Hostelername "

            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Hostelers")
            myDA.Fill(myDS, "RoomAllotment")
            myDA.Fill(myDS, "FeePayment")


            rpt.SetDataSource(myDS)
            frmAccountRecordReport.CrystalReportViewer1.ReportSource = rpt
            frmAccountRecordReport.Show()
        Catch ex As Exception
            MessageBox.Show("No Record Found!.")
            ' MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        If DataGridView5.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = DataGridView5.RowCount - 1
            colsTotal = DataGridView5.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView5.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView5.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        If DataGridView3.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = DataGridView3.RowCount - 1
            colsTotal = DataGridView3.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView3.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView3.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        cmbHostelType2.Text = ""
        cmbAcadYear.Text = ""
        DataGridView3.DataSource = Nothing
        DataGridView9.DataSource = Nothing
        GroupBox11.Visible = False
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        cmbHostelType.Text = ""
        DataGridView1.DataSource = Nothing
        DataGridView6.DataSource = Nothing
        GroupBox2.Visible = False
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        cmbHostelType1.Text = ""
        DateFrom.Text = Today
        DateTo.Text = Today
        DataGridView2.DataSource = Nothing
        DataGridView8.DataSource = Nothing
        GroupBox7.Visible = False
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        cmbAcadYear1.Text = ""
        DataGridView5.DataSource = Nothing
        DataGridView10.DataSource = Nothing
        GroupBox18.Visible = False
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If DataGridView2.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = DataGridView2.RowCount - 1
            colsTotal = DataGridView2.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView2.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView2.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub





    Private Sub btnGetData1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetData1.Click

        If Len(Trim(cmbHostelType.Text)) = 0 Then
            MessageBox.Show("Please select Hostel Name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbHostelType.Focus()
            Exit Sub
        End If

        Try
            GroupBox2.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("SELECT distinct Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], Hostelers.Purpose as [Department], RoomAllotment.Hostelname as [Hostel Name], RoomAllotment.RoomNo as [Room No] , RoomAllotment.AcadYear as [Acadeic Year], (FeePayment.ServiceCharges-RegCharges.PrevDue) as [ServiceCharges], RegCharges.PrevDue as[Prev Due],  FeePayment.ServiceCharges as [Total + Prev Due], sum(FeePayment.Totalpaid) as [Total paid] FROM ((Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID]) INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where   RoomAllotment.HostelName ='" & cmbHostelName.Text & "' and FeePayment.AcadYear=RoomAllotment.AcadYear and FeePayment.AcadYear=RegCharges.AcadYear and RoomAllotment.AcadYear=RegCharges.AcadYear group by HostelerName,USn,purpose, RoomAllotment.Hostelname,RoomAllotment.RoomNo,ServiceCharges,RoomAllotment.AcadYear,RegCharges.PrevDue  order by Hostelername ", con)
            '  cmd = New OleDbCommand("SELECT Hostelers.HostelerName as [Hosteler Name] , Hostelers.USN as [USN] , RoomAllotment.Hostelname as [Hostel Name], Hostel.HostelType as [Hostel Type] , RoomAllotment.RoomNo  as [Room No] , RoomAllotment.AcadYear as [Academic Year] ,  (FeePayment.ServiceCharges- RegCharges.PrevDue) as [Total] , RegCharges.PrevDue as[Prev Due] ,FeePayment.ServiceCharges as [service Charge+Due],  sum(FeePayment.Totalpaid) as [Total paid] FROM (Hostel INNER JOIN ((Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID]) INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where   Hostel.HostelType ='" & cmbHostelType.Text & "' and FeePayment.AcadYear=RoomAllotment.AcadYear and FeePayment.AcadYear=RegCharges.AcadYear and RoomAllotment.AcadYear=RegCharges.AcadYear group by HostelerName,USn,purpose, RoomAllotment.Hostelname,RoomAllotment.RoomNo,ServiceCharges,RoomAllotment.AcadYear,RegCharges.PrevDue,HostelType order by Hostelername ", con)
            cmd = New OleDbCommand("SELECT DISTINCT   Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], RoomAllotment.Hostelname as [Hostel Name], RoomAllotment.RoomNo as [Room No], RoomAllotment.AcadYear as [Acad Year], RegCharges.CautionMoney as [Deposit], RegCharges.RentalCharges as [Room Rent], RegCharges.FormFee as [Form Fee], RegCharges.OtherFee as [Other Fee],( RegCharges.CautionMoney + RegCharges.RentalCharges + RegCharges.FormFee + RegCharges.OtherFee )as [Total ], RegCharges.PrevDue as [Prev Due], RegCharges.TotalCharges as [Total + Prev Due], sum(FeePayment.TotalPaid) as [Current Paid] FROM (Hostel INNER JOIN ((Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID]) INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where Hostel.HostelType ='" & cmbHostelType.Text & "' and FeePayment.AcadYear=RoomAllotment.AcadYear and FeePayment.AcadYear=RegCharges.AcadYear and RoomAllotment.AcadYear=RegCharges.AcadYear  group by  Hostelers.HostelerName,Hostelers.HostelerName, Hostelers.USN, RoomAllotment.Hostelname, Hostel.HostelType,RoomAllotment.RoomNo,   RoomAllotment.AcadYear, RegCharges.CautionMoney, RegCharges.RentalCharges, RegCharges.FormFee,RegCharges.OtherFee,RegCharges.PrevDue, RegCharges.TotalCharges", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "FeePayment")
            myDA.Fill(myDataSet, "RoomAllotment")
            myDA.Fill(myDataSet, "RegCharges")
            myDA.Fill(myDataSet, "Hostel")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("FeePayment").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("RegCharges").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("Hostel").DefaultView
            Dim sum As Int64 = 0
            Dim sum1 As Int64 = 0

            For Each r As DataGridViewRow In Me.DataGridView1.Rows

                sum = sum + r.Cells(12).Value
                sum1 = sum1 + (r.Cells(9).Value - r.Cells(12).Value)

            Next

            TextBox1.Text = sum
            Label13.Text = frmDueCharges.GetInWords(sum)
            TextBox4.Text = sum1

            con.Close()


            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("SELECT Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN],Concession.AcadYear as [Acadeic Year], Concession.Con_Amount as [Concession Amount] FROM Hostelers INNER JOIN Concession ON Hostelers.[HostelerID] = Concession.[HostelerID] ", con)
            cmd = New OleDbCommand(" SELECT Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], RoomAllotment.Hostelname as [Hostel Name] ,Concession.AcadYear as [Acad Year], Concession.Con_Amount as [Concession Amt], Concession.Con_Date as [Concession Date] FROM (Hostel INNER JOIN (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN Concession ON Hostelers.[HostelerID] = Concession.[HostelerID] where Hostel.HostelType ='" & cmbHostelType.Text & "' and Hostelers.AcadYear=RoomAllotment.AcadYear ", con)
            Dim myDA1 As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet1 As DataSet = New DataSet()
            myDA1.Fill(myDataSet1, "Hostelers")
            myDA1.Fill(myDataSet1, "Hostel")
            myDA1.Fill(myDataSet1, "RoomAllotment")
            myDA1.Fill(myDataSet1, "Concession")
            DataGridView6.DataSource = myDataSet1.Tables("Hostelers").DefaultView
            DataGridView6.DataSource = myDataSet1.Tables("Hostel").DefaultView
            DataGridView6.DataSource = myDataSet1.Tables("RoomAllotment").DefaultView
            DataGridView6.DataSource = myDataSet1.Tables("Concession").DefaultView
            Dim sum2 As Int64 = 0
            For Each r1 As DataGridViewRow In Me.DataGridView6.Rows
                sum2 = sum2 + (r1.Cells(4).Value)
            Next
            TextBox9.Text = sum2
            Label23.Text = frmDueCharges.GetInWords(sum2)
            Dim sum3 As Int64 = 0
            sum3 = sum1 - sum2
            TextBox4.Text = sum3
            Label14.Text = frmDueCharges.GetInWords(sum3)
            con.Close()
        
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    
    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        Try

            GroupBox13.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            
            '  cmd = New OleDbCommand("SELECT Hostelers.HostelerName as [Hosteler Name] , Hostelers.USN as [USN] , RoomAllotment.Hostelname as [Hostel name], Hostel.HostelType as [Hostel Type] , RoomAllotment.RoomNo  as [Room No] , RoomAllotment.AcadYear as [Academic Year] ,  (FeePayment.ServiceCharges- RegCharges.PrevDue) as [Total] , RegCharges.PrevDue as[Prev Due] ,FeePayment.ServiceCharges as [service Charge+Due],  sum(FeePayment.Totalpaid) as [Total Paid] FROM (Hostel INNER JOIN ((Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID]) INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where  FeePayment.AcadYear=RoomAllotment.AcadYear and FeePayment.AcadYear=RegCharges.AcadYear and RoomAllotment.AcadYear=RegCharges.AcadYear group by HostelerName,USN,purpose, RoomAllotment.Hostelname,RoomAllotment.RoomNo,ServiceCharges,RoomAllotment.AcadYear,RegCharges.PrevDue,HostelType order by Hostelername ", con)
            cmd = New OleDbCommand("SELECT DISTINCT   Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], RoomAllotment.Hostelname as [Hostel Name], RoomAllotment.RoomNo as [Room No], RoomAllotment.AcadYear as [Acad Year], RegCharges.CautionMoney as [Deposit], RegCharges.RentalCharges as [Room Rent], RegCharges.FormFee as [Form Fee], RegCharges.OtherFee as [Other Fee],( RegCharges.CautionMoney + RegCharges.RentalCharges + RegCharges.FormFee + RegCharges.OtherFee)as [Total ], RegCharges.PrevDue as [Prev Due], RegCharges.TotalCharges as [Total + Prev Due], sum(FeePayment.TotalPaid) as [Current Paid] FROM (Hostel INNER JOIN ((Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID]) INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] Where FeePayment.AcadYear=RegCharges.AcadYear and RoomAllotment.AcadYear=RegCharges.AcadYear  group by  Hostelers.HostelerName,Hostelers.HostelerName, Hostelers.USN, RoomAllotment.Hostelname, Hostel.HostelType,RoomAllotment.RoomNo,   RoomAllotment.AcadYear, RegCharges.CautionMoney, RegCharges.RentalCharges, RegCharges.FormFee, RegCharges.OtherFee,RegCharges.PrevDue, RegCharges.TotalCharges", con)


            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "FeePayment")
            myDA.Fill(myDataSet, "RoomAllotment")
            myDA.Fill(myDataSet, "RegCharges")
            myDA.Fill(myDataSet, "Hostel")
            DataGridView4.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView4.DataSource = myDataSet.Tables("FeePayment").DefaultView
            DataGridView4.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
            DataGridView4.DataSource = myDataSet.Tables("RegCharges").DefaultView
            DataGridView4.DataSource = myDataSet.Tables("Hostel").DefaultView
            Dim sum As Int64 = 0
            Dim sum1 As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView4.Rows
         
                sum = sum + r.Cells(12).Value
                sum1 = sum1 + (r.Cells(9).Value - r.Cells(12).Value)
                   
            Next
            TextBox12.Text = sum
            Label27.Text = frmDueCharges.GetInWords(sum)
            TextBox11.Text = sum1

            con.Close()
     
            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("SELECT Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN],Concession.AcadYear as [Acadeic Year], Concession.Con_Amount as [Concession Amount] FROM Hostelers INNER JOIN Concession ON Hostelers.[HostelerID] = Concession.[HostelerID] ", con)
            cmd = New OleDbCommand(" SELECT Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], RoomAllotment.Hostelname as [Hostel Name] , Concession.AcadYear as [Acad Year], Concession.Con_Amount as [Concession Amt], Concession.Con_Date as [Concession Date] FROM (Hostel INNER JOIN (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN Concession ON Hostelers.[HostelerID] = Concession.[HostelerID] where  Hostelers.AcadYear=RoomAllotment.AcadYear ", con)
            Dim myDA1 As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet1 As DataSet = New DataSet()
            myDA1.Fill(myDataSet1, "Hostelers")
            myDA1.Fill(myDataSet1, "Hostel")
            myDA1.Fill(myDataSet1, "RoomAllotment")
            myDA1.Fill(myDataSet1, "Concession")
            DataGridView7.DataSource = myDataSet1.Tables("Hostelers").DefaultView
            DataGridView7.DataSource = myDataSet1.Tables("Hostel").DefaultView
            DataGridView7.DataSource = myDataSet1.Tables("RoomAllotment").DefaultView
            DataGridView7.DataSource = myDataSet1.Tables("Concession").DefaultView
            Dim sum2 As Double = 0
            For Each r1 As DataGridViewRow In Me.DataGridView7.Rows
                sum2 = sum2 + (r1.Cells(4).Value)
            Next
            TextBox10.Text = sum2
            Label31.Text = frmDueCharges.GetInWords(sum2)
            Dim sum3 As Int64 = 0
            sum3 = sum1 - sum2
            TextBox11.Text = sum3
            Label28.Text = frmDueCharges.GetInWords(sum3)
            con.Close()
           
       
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        If DataGridView4.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = DataGridView4.RowCount - 1
            colsTotal = DataGridView4.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView4.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView4.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        DataGridView4.DataSource = Nothing
        DataGridView7.DataSource = Nothing
        GroupBox13.Visible = False
        TextBox11.Text = ""
        TextBox12.Text = ""
    End Sub

   
    
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        If DataGridView6.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = DataGridView6.RowCount - 1
            colsTotal = DataGridView6.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView6.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView6.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        If DataGridView8.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = DataGridView8.RowCount - 1
            colsTotal = DataGridView8.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView8.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView8.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        If DataGridView9.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = DataGridView9.RowCount - 1
            colsTotal = DataGridView9.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView9.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView9.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        If DataGridView10.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = DataGridView10.RowCount - 1
            colsTotal = DataGridView10.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView10.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView10.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        If DataGridView7.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = DataGridView7.RowCount - 1
            colsTotal = DataGridView7.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView7.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView7.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        If Len(Trim(cmbHostelName.Text)) = 0 Then
            MessageBox.Show("Please select Hostel", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbHostelName.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbRoomNo1.Text)) = 0 Then
            MessageBox.Show("Please select Room Number", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbRoomNo1.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbRoomNo2.Text)) = 0 Then
            MessageBox.Show("Please select Room Number", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbRoomNo2.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbAcadYear2.Text)) = 0 Then
            MessageBox.Show("Please select Academica Year", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbAcadYear2.Focus()
            Exit Sub
        End If
        Try

            GroupBox3.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            '         cmd = New OleDbCommand("SELECT Hostelers.HostelerName as [Hosteler Name] , Hostelers.USN as [USN] , RoomAllotment.Hostelname as [Hostel Name], Hostel.HostelType as [Hostel Type] , RoomAllotment.RoomNo  as [Room No] , RoomAllotment.AcadYear as [Academic Year] ,  (FeePayment.ServiceCharges- RegCharges.PrevDue) as [Total] , RegCharges.PrevDue as[Prev Due] ,FeePayment.ServiceCharges as [Total + Prev Due],  sum(FeePayment.Totalpaid) as [Total paid] FROM (Hostel INNER JOIN ((Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID]) INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where   RoomAllotment.HostelName ='" & cmbHostelName.Text & "'and RoomAllotment.RoomNo between '" & cmbRoomNo1.Text & "' and '" & cmbRoomNo2.Text & "' and RoomAllotment.AcadYear = '" & cmbAcadYear2.Text & "' and FeePayment.AcadYear=RoomAllotment.AcadYear and FeePayment.AcadYear=RegCharges.AcadYear and RoomAllotment.AcadYear=RegCharges.AcadYear group by HostelerName,USn,purpose, RoomAllotment.Hostelname,RoomAllotment.RoomNo,ServiceCharges,RoomAllotment.AcadYear,RegCharges.PrevDue,HostelType order by Hostelername ", con)
            cmd = New OleDbCommand("SELECT DISTINCT   Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], RoomAllotment.Hostelname as [Hostel Name],  RoomAllotment.RoomNo as [Room No], RoomAllotment.AcadYear as [Acad Year], RegCharges.CautionMoney as [Deposit], RegCharges.RentalCharges as [Room Rent], RegCharges.FormFee as [Form Fee], RegCharges.OtherFee as [Othr Fee],( RegCharges.CautionMoney + RegCharges.RentalCharges + RegCharges.FormFee + RegCharges.OtherFee )as [Total ], RegCharges.PrevDue as [Prev Due], RegCharges.TotalCharges as [Total + Prev Due], sum(FeePayment.TotalPaid) as [Current Paid] FROM (Hostel INNER JOIN ((Hostelers INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID]) INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN FeePayment ON Hostelers.[HostelerID] = FeePayment.[HostelerID] where  RoomAllotment.HostelName ='" & cmbHostelName.Text & "'and RoomAllotment.RoomNo between '" & cmbRoomNo1.Text & "' and '" & cmbRoomNo2.Text & "' and FeePayment.AcadYear=RoomAllotment.AcadYear and FeePayment.AcadYear=RegCharges.AcadYear and RoomAllotment.AcadYear=RegCharges.AcadYear  group by  Hostelers.HostelerName,Hostelers.HostelerName, Hostelers.USN, RoomAllotment.Hostelname, Hostel.HostelType,RoomAllotment.RoomNo,   RoomAllotment.AcadYear, RegCharges.CautionMoney, RegCharges.RentalCharges, RegCharges.FormFee,RegCharges.OtherFee,RegCharges.PrevDue, RegCharges.TotalCharges", con)


            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "FeePayment")
            myDA.Fill(myDataSet, "RoomAllotment")
            myDA.Fill(myDataSet, "RegCharges")
            myDA.Fill(myDataSet, "Hostel")
            DataGridView12.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView12.DataSource = myDataSet.Tables("FeePayment").DefaultView
            DataGridView12.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
            DataGridView12.DataSource = myDataSet.Tables("RegCharges").DefaultView
            DataGridView12.DataSource = myDataSet.Tables("Hostel").DefaultView
            Dim sum As Int64 = 0
            Dim sum1 As Int64 = 0

            For Each r As DataGridViewRow In Me.DataGridView12.Rows
                If (DataGridView12.Rows.Count <= 2) Then
                    If (r.Cells(10).Value > 0) Then
                        sum = sum + r.Cells(12).Value
                        sum1 = sum1 + ((r.Cells(9).Value + r.Cells(10).Value) - r.Cells(12).Value)
                    Else
                        sum = sum + r.Cells(12).Value
                        sum1 = sum1 + (r.Cells(9).Value - r.Cells(12).Value)
                    End If
                Else
                    sum = sum + r.Cells(12).Value
                    sum1 = sum1 + (r.Cells(9).Value - r.Cells(12).Value)
                End If

            Next
            TextBox16.Text = sum
            Label45.Text = frmDueCharges.GetInWords(sum)
            TextBox18.Text = sum1

            con.Close()
            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("SELECT Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN],Concession.AcadYear as [Acadeic Year], Concession.Con_Amount as [Concession Amount] FROM Hostelers INNER JOIN Concession ON Hostelers.[HostelerID] = Concession.[HostelerID] ", con)
            cmd = New OleDbCommand(" SELECT Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], RoomAllotment.Hostelname as [Hostel Name] ,RoomAllotment.RoomNo as [Room No],Concession.AcadYear as [Acad Year], Concession.Con_Amount as [Concession Amt], Concession.Con_Date as [Concession Date] FROM (Hostel INNER JOIN (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) ON Hostel.[HostelName] = Hostelers.[HostelName]) INNER JOIN Concession ON Hostelers.[HostelerID] = Concession.[HostelerID] where RoomAllotment.HostelName ='" & cmbHostelName.Text & "'and RoomAllotment.RoomNo between '" & cmbRoomNo1.Text & "' and '" & cmbRoomNo2.Text & "' and Concession.AcadYear = '" & cmbAcadYear2.Text & "' and Hostelers.AcadYear=RoomAllotment.AcadYear ", con)
            Dim myDA1 As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet1 As DataSet = New DataSet()
            myDA1.Fill(myDataSet1, "Hostelers")
            myDA1.Fill(myDataSet1, "Hostel")
            myDA1.Fill(myDataSet1, "RoomAllotment")
            myDA1.Fill(myDataSet1, "Concession")
            DataGridView11.DataSource = myDataSet1.Tables("Hostelers").DefaultView
            DataGridView11.DataSource = myDataSet1.Tables("Hostel").DefaultView
            DataGridView11.DataSource = myDataSet1.Tables("RoomAllotment").DefaultView
            DataGridView11.DataSource = myDataSet1.Tables("Concession").DefaultView
            Dim sum5 As Int64 = 0
            For Each r1 As DataGridViewRow In Me.DataGridView11.Rows
                sum5 = sum5 + (r1.Cells(5).Value)
            Next
            TextBox17.Text = sum5
            Label43.Text = frmDueCharges.GetInWords(sum5)
            Dim sum6 As Int64 = 0
            sum6 = sum1 - sum5
            TextBox18.Text = sum6
            Label46.Text = frmDueCharges.GetInWords(sum6)
            con.Close()

        Catch ex As Exception
            MessageBox.Show("No Record Found!.")
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

   
    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        cmbHostelName.Text = ""
        cmbRoomNo1.Text = ""
        cmbRoomNo2.Text = ""
        cmbAcadYear2.Text = ""
    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        If DataGridView12.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = DataGridView12.RowCount - 1
            colsTotal = DataGridView12.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView12.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView12.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        If DataGridView11.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = DataGridView11.RowCount - 1
            colsTotal = DataGridView11.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView11.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView11.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click

    End Sub
End Class

