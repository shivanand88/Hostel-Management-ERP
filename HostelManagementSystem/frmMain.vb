Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class frmMain

    Dim rdr As OleDbDataReader = Nothing
    Dim con As OleDbConnection = Nothing
    Dim cmd As OleDbCommand = Nothing
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"

    Private Sub CalculatorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CalculatorToolStripMenuItem.Click
        System.Diagnostics.Process.Start("Calc.exe")
    End Sub

    Private Sub NotepadToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NotepadToolStripMenuItem.Click
        System.Diagnostics.Process.Start("Notepad.exe")
    End Sub

    Private Sub TaskManagerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TaskManagerToolStripMenuItem.Click
        System.Diagnostics.Process.Start("TaskMgr.exe")
    End Sub

    Private Sub MSWordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MSWordToolStripMenuItem.Click
        System.Diagnostics.Process.Start("WinWord.exe")
    End Sub

    Private Sub WordPadToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WordPadToolStripMenuItem.Click
        System.Diagnostics.Process.Start("Wordpad.exe")
    End Sub

    Private Sub SystemInfoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SystemInfoToolStripMenuItem.Click
        frmSystemInfo.Show()
    End Sub

    Private Sub LoginDetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoginDetailsToolStripMenuItem.Click
        frmLoginDetails.Show()
    End Sub

    Private Sub RegistrationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistrationToolStripMenuItem.Click
        Me.Hide()
        frmRegistration.Show()
    End Sub

    Private Sub ProfileEntryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProfileEntryToolStripMenuItem.Click
        Me.Hide()
        frmHostelers.Show()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        frmAbout.Show()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        lblDateTime.Text = Now
    End Sub

    Private Sub HostelFeesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HostelFeesToolStripMenuItem.Click
        Me.Hide()
        frmServices.Show()
    End Sub

    Private Sub RoomToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RoomToolStripMenuItem.Click
        frmRoom.Show()
    End Sub



    Private Sub ProductsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frmPurchaseInventory.Show()
    End Sub

    Private Sub HostelersToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HostelersToolStripMenuItem1.Click
        Me.Hide()
        frmHostelers.Show()
    End Sub

    Private Sub LogOutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogOutToolStripMenuItem.Click
        Me.Hide()
        frmLogin.Show()
        frmLogin.UserName.Text = ""
        frmLogin.Password.Text = ""
        frmLogin.ProgressBar1.Visible = False
        frmLogin.UserName.Focus()

    End Sub

    Private Sub HostelersToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HostelersToolStripMenuItem2.Click
        Me.Hide()
        frmHostelersRecord.Show()
    End Sub

    Private Sub PurchaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PurchaseToolStripMenuItem.Click
        frmPurchaseInventory.Show()
    End Sub

    Private Sub FeePaymentToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FeePaymentToolStripMenuItem1.Click
        Me.Hide()
        frmFeePayment.Show()
    End Sub

    Private Sub FeePaymentToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FeePaymentToolStripMenuItem3.Click
        Me.Hide()
        frmFeePayment.Show()
    End Sub

    Private Sub RegistrationToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistrationToolStripMenuItem1.Click
        Me.Hide()
        frmRegistration.Show()
    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        End
        ' Me.Hide()
        ' frmLogin.Show()
        ' frmLogin.reset()
    End Sub

   

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelName as [Hostel Name], RoomNO as [Room No], RoomType as [Room Type], NoOfBeds as [Total Beds], BedsAvailable as[Available Beds] from Room  order by hostelname,roomno,roomtype", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()

            myDA.Fill(myDataSet, "Room")

            DataGridView1.DataSource = myDataSet.Tables("Room").DefaultView

            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Me.frmMain_Load(New Object, New System.EventArgs)
        'btnRefreshForm.PerformClick()
    End Sub
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT * from Room where bedsavailable > 0 order by hostelname,roomno,roomtype", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()

            myDA.Fill(myDataSet, "Room")

            DataGridView1.DataSource = myDataSet.Tables("Room").DefaultView

            con.Close()
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

    Private Sub AdvanceEntryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AdvanceEntryToolStripMenuItem.Click

        frmTransaction.Show()
    End Sub

    Private Sub AdvanceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AdvanceToolStripMenuItem.Click

        frmTransactionRecord.Show()
    End Sub

    Private Sub FeePaymentToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FeePaymentToolStripMenuItem2.Click

        frmFeePaymentRecord1.Show()
    End Sub

    Private Sub ExtraServicesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExtraServicesToolStripMenuItem.Click

        frmExtraServices.Show()
    End Sub

    Private Sub ExtraServicesToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExtraServicesToolStripMenuItem1.Click

        frmExtraServicesRecord.Show()
    End Sub


    Private Sub CheckOutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckOutToolStripMenuItem.Click

        frmCheckOutRecord.Show()
    End Sub

    'Private Sub ServicesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ServicesToolStripMenuItem.Click
    'Me.Hide()
    ' frmServices.Show()
    ' End Sub

    Private Sub DueChargesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DueChargesToolStripMenuItem.Click
        ' Me.Hide()
        ' frmDueCharges.Show()
        frmDueChargesReport.Show()
    End Sub

    Private Sub HostelToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HostelToolStripMenuItem.Click
        frmHostel.Show()
    End Sub

    Private Sub QuotationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QuotationToolStripMenuItem.Click
        frmQuotation.Show()
    End Sub

    Private Sub RegistrationChargesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistrationChargesToolStripMenuItem.Click
        frmRegCharges.Show()
    End Sub

    Private Sub PaymentVoucherToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PaymentVoucherToolStripMenuItem.Click

        frmPaymentVoucher.Show()
    End Sub

    Private Sub DueDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DueDateToolStripMenuItem.Click
        frmDueDate.Show()
    End Sub

    Private Sub RegistrationFormToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistrationFormToolStripMenuItem.Click
        frmHostelerForm.Show()
    End Sub


    Private Sub RegToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegToolStripMenuItem1.Click

        frmRegCharges.Show()
    End Sub


    Private Sub AccountsRecordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AccountsRecordToolStripMenuItem.Click

        frmAccountsRecord.Show()
    End Sub

    Private Sub CheckOutToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckOutToolStripMenuItem1.Click

        frmCheckOut.Show()
    End Sub


    Private Sub ReAllocateToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ReAllocateToolStripMenuItem.Click
        Me.Hide()
        frmReAllocate.Show()
    End Sub

    Private Sub BackUpToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BackUpToolStripMenuItem1.Click

        frmDataBackUp.Show()
    End Sub

    Private Sub RoomChangeToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RoomChangeToolStripMenuItem1.Click
        frmRoomChange.Show()
    End Sub

    Private Sub RefundToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefundToolStripMenuItem1.Click

        frmRefund.Show()
    End Sub

    Private Sub HobbyToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HobbyToolStripMenuItem1.Click

        frmHobbies.Show()
    End Sub

    Private Sub HobbyReportToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HobbyReportToolStripMenuItem1.Click

        frmHobbiesRecord.Show()

    End Sub

    Private Sub ConcessionToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConcessionToolStripMenuItem1.Click

        frmConcession.Show()
    End Sub


    Private Sub ConcessionRecordToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConcessionRecordToolStripMenuItem1.Click

        frmConcessionRecord.Show()
    End Sub

    Private Sub RefundRecordToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefundRecordToolStripMenuItem1.Click

        frmRefundRecord.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
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

    Private Sub GuestRegistrationMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GuestRegistrationMenuItem.Click
        frmGuestRegistration.Show()
    End Sub

    Private Sub GuestCheckOutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GuestCheckOutToolStripMenuItem.Click
        frmGuestCheckOut.Show()
    End Sub

    Private Sub GuestreallocateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GuestreallocateToolStripMenuItem.Click
        frmGuestReAllocate.Show()
    End Sub

    
    Private Sub RoomAllotmentDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RoomAllotmentDetails.Click
        Me.Hide()
        frmRoomAllotmentDetails.Show()
    End Sub
End Class
