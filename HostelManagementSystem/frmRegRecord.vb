Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmRegRecord
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"
    Private Sub txtUSN_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUSN.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("Select ReceiptNumber as [Receipt No],Hostelers.HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],HostelName as [Hostel Name],RoomNo as [Room No],Hostelers.AcadYear as [Academic Year],CautionMoney as [Deposit Money],RentalCharges as [Rental Charges],FormFee as [Form Fee],TotalCharges as [Total Charges], PaymentDate as [Payment Date] from RegCharges,Hostelers where Hostelers.HostelerID=RegCharges.HostelerID and Hostelers.AcadYear=RegCharges.AcadYear and HostelerName like '" & txtHostelerName.Text & "%' order by HostelerName", con)
            cmd = New OleDbCommand("SELECT distinct RegCharges.ReceiptNumber as [Receipt No], RegCharges.HostelerID as [Hosteler ID], Hostelers.HostelerName as [Hosteler Name], RoomAllotment.Hostelname as [Hostel Name], RoomAllotment.RoomNo as [Room No], RegCharges.AcadYear as [Academic Year], RegCharges.CautionMoney as [Deposit Money], RegCharges.RentalCharges as [Rental Charges], RegCharges.FormFee as [Form Fee], RegCharges.OtherFee as [Other Fee], RegCharges.PrevDue as [Prev Due], RegCharges.TotalCharges as [Total Charges], RegCharges.PaymentDate as [Payment Date],RegCharges.Remarks as [Remarks] FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID]  where USN like '%" & txtUSN.Text & "%' and RegCharges.AcadYear=RoomAllotment.AcadYear order by HostelerName,PaymentDate Desc", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "RegCharges")
            myDA.Fill(myDataSet, "RoomAllotment")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("regCharges").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub txtHostelerName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHostelerName.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("Select ReceiptNumber as [Receipt No],Hostelers.HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],HostelName as [Hostel Name],RoomNo as [Room No],Hostelers.AcadYear as [Academic Year],CautionMoney as [Deposit Money],RentalCharges as [Rental Charges],FormFee as [Form Fee],TotalCharges as [Total Charges], PaymentDate as [Payment Date] from RegCharges,Hostelers where Hostelers.HostelerID=RegCharges.HostelerID and Hostelers.AcadYear=RegCharges.AcadYear and HostelerName like '" & txtHostelerName.Text & "%' order by HostelerName", con)
            cmd = New OleDbCommand("SELECT distinct RegCharges.ReceiptNumber as [Receipt No], RegCharges.HostelerID as [Hosteler ID], Hostelers.HostelerName as [Hosteler Name], RoomAllotment.Hostelname as [Hostel Name], RoomAllotment.RoomNo as [Room No], RegCharges.AcadYear as [Academic Year], RegCharges.CautionMoney as [Deposit Money], RegCharges.RentalCharges as [Rental Charges], RegCharges.FormFee as [Form Fee], RegCharges.OtherFee as [Other Fee], RegCharges.PrevDue as [Prev Due], RegCharges.TotalCharges as [Total Charges], RegCharges.PaymentDate as [Payment Date],RegCharges.Remarks as [Remarks] FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID]  where HostelerName like '%" & txtHostelerName.Text & "%' and RegCharges.AcadYear=RoomAllotment.AcadYear order by HostelerName,PaymentDate Desc", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "RegCharges")
            myDA.Fill(myDataSet, "RoomAllotment")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("regCharges").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub GetData()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("Select ReceiptNumber as [Receipt No],Hostelers.HostelerID as [Hosteler ID],HostelerName as [Hosteler Name],HostelName as [Hostel Name],RoomNo as [Room No],AcadYear as [Academic Year],CautionMoney as [Caution Money],RentalCharges as [Rental+Foods Charges],FormFee as [Form Fee],TotalCharges as [Total Charges], PaymentDate as [Payment Date] from RegCharges,Hostelers where Hostelers.HostelerID=RegCharges.HostelerID and HostelerName like '" & txtHostelerName.Text & "%' order by HostelerName", con)
            cmd = New OleDbCommand("SELECT distinct RegCharges.ReceiptNumber as [Receipt No], RegCharges.HostelerID as [Hosteler ID], Hostelers.HostelerName as [Hosteler Name], RoomAllotment.Hostelname as [Hostel Name], RoomAllotment.RoomNo as [Room No], RegCharges.AcadYear as [Academic Year], RegCharges.CautionMoney as [Deposit Money], RegCharges.RentalCharges as [Rental Charges], RegCharges.FormFee as [Form Fee], RegCharges.OtherFee as [Other Fee], RegCharges.PrevDue as [Prev Due], RegCharges.TotalCharges as [Total Charges], RegCharges.PaymentDate as [Payment Date],RegCharges.Remarks as [Remarks] FROM (Hostelers INNER JOIN RoomAllotment ON Hostelers.[HostelerID] = RoomAllotment.[HostelerID]) INNER JOIN RegCharges ON Hostelers.[HostelerID] = RegCharges.[HostelerID]   where RegCharges.AcadYear=RoomAllotment.AcadYear order by HostelerName,PaymentDate Desc", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "RegCharges")
            myDA.Fill(myDataSet, "RoomAllotment")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("regCharges").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("RoomAllotment").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub frmRegRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GetData()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        txtHostelerName.Text = ""
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
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

    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            Me.Hide()
            frmRegCharges.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            'frmRegCharges.txtPrevDue.Text = "0"

            frmRegCharges.cmbHostelerID.Text = dr.Cells(1).Value.ToString()
            frmRegCharges.txtHostelerName.Text = dr.Cells(2).Value.ToString()
            frmRegCharges.txtHostelName.Text = dr.Cells(3).Value.ToString()
            frmRegCharges.txtRoomNo.Text = dr.Cells(4).Value.ToString()
            frmRegCharges.cmbAcadYear.Text = dr.Cells(5).Value.ToString()
            frmRegCharges.txtCautionMoney.Text = dr.Cells(6).Value.ToString()
            frmRegCharges.txtRentalCharges.Text = dr.Cells(7).Value.ToString()
            frmRegCharges.txtFormFee.Text = dr.Cells(8).Value.ToString()
            frmRegCharges.txtOtherFee.Text = dr.Cells(9).Value.ToString()
            frmRegCharges.txtPrevDue.Text = dr.Cells(10).Value.ToString()
            frmRegCharges.txtTotalCharges.Text = dr.Cells(11).Value.ToString()
            frmRegCharges.dtpPaymentDate.Text = dr.Cells(12).Value.ToString()
            frmRegCharges.txtRemarks.Text = dr.Cells(13).Value.ToString()
            frmRegCharges.txtReceiptNo.Text = dr.Cells(0).Value.ToString()
            frmRegCharges.btnUpdate_record.Enabled = True
            frmRegCharges.btnDelete.Enabled = True
            frmRegCharges.btnSave.Enabled = False
            frmRegCharges.btnPrint.Enabled = True

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub DataGridView1_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView1.RowHeadersWidth < Convert.ToInt32((size.Width + 18)) Then
            DataGridView1.RowHeadersWidth = Convert.ToInt32((size.Width + 18))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub


End Class