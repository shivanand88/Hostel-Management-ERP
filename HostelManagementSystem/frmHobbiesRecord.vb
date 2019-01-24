Imports System.Data.OleDb
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmHobbiesRecord
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"

    Private Sub frmHobbiesRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        fillHostelerName()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID],HostelerName as [Hosteler Name], from Hostelers where HostelerName='" & cmbHostelerName.Text & "' order by HostelerName", con)
            cmd = New OleDbCommand("SELECT Hostelers.HostelerID as [Hosteler ID],Hostelers.HostelerName as [Hosteler Name],Hostelers.USN as [USN],Hobbies.Hobby1 as [Hobby 1], Hobbies.Hobby2 as [Hobby 2], Hobbies.Hobby3 as [Hobby 3], Hobbies.Hobby4 as [Hobby 4], Hobbies.Hobby5 as [Hobby 5], Hobbies.Hobby6 as [Hobby 6], Hobbies.Hobby7 as [Hobby 7], Hobbies.Hobby8 as [Hobby 8],Hobbies.HobbyList as [Hobby List] FROM Hostelers INNER JOIN Hobbies ON Hostelers.[HostelerID] = Hobbies.[HostelerID]  order by HostelerName ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "hobbies")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("Hobbies").DefaultView
            DataGridView3.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView3.DataSource = myDataSet.Tables("Hobbies").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbHostelerName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbHostelerName.SelectedIndexChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID],HostelerName as [Hosteler Name], from Hostelers where HostelerName='" & cmbHostelerName.Text & "' order by HostelerName", con)
            cmd = New OleDbCommand("SELECT Hostelers.HostelerID as [Hosteler ID],Hostelers.HostelerName as [Hosteler Name],Hostelers.USN as [USN],Hobbies.Hobby1 as [Hobby 1], Hobbies.Hobby2 as [Hobby 2], Hobbies.Hobby3 as [Hobby 3], Hobbies.Hobby4 as [Hobby 4], Hobbies.Hobby5 as [Hobby 5], Hobbies.Hobby6 as [Hobby 6], Hobbies.Hobby7 as [Hobby 7], Hobbies.Hobby8 as [Hobby 8],Hobbies.HobbyList as [Hobby List] FROM Hostelers INNER JOIN Hobbies ON Hostelers.[HostelerID] = Hobbies.[HostelerID] where HostelerName='" & cmbHostelerName.Text & "' and Hobbies.HostelerID=Hostelers.HostelerID order by HostelerName ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "hobbies")
            DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("Hobbies").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
   
    Sub fillHostelerName()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct RTRIM(HostelerName) FROM Hostelers", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbHostelerName.Items.Clear()

            For Each drow As DataRow In dtable.Rows
                cmbHostelerName.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
   
   

    Private Sub txtHobby_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHobby.TextChanged
        Try

            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID],HostelerName as [Hosteler Name], from Hostelers where HostelerName='" & cmbHostelerName.Text & "' order by HostelerName", con)
            cmd = New OleDbCommand("SELECT Hostelers.HostelerID as [Hosteler ID],Hostelers.HostelerName as [Hosteler Name],Hostelers.USN as [USN],Hobbies.Hobby1 as [Hobby 1], Hobbies.Hobby2 as [Hobby 2], Hobbies.Hobby3 as [Hobby 3], Hobbies.Hobby4 as [Hobby 4], Hobbies.Hobby5 as [Hobby 5], Hobbies.Hobby6 as [Hobby 6], Hobbies.Hobby7 as [Hobby 7], Hobbies.Hobby8 as [Hobby 8],Hobbies.HobbyList as [Hobby List] FROM Hostelers INNER JOIN Hobbies ON Hostelers.[HostelerID] = Hobbies.[HostelerID] where Hobbies.HobbyList like '%" & txtHobby.Text & "%' and Hobbies.HostelerID=Hostelers.HostelerID order by HostelerName ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Hobbies")
            DataGridView3.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView3.DataSource = myDataSet.Tables("Hobbies").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        txtUSN.Text = ""
        cmbHostelerName.Text = ""
        txtHostelerName.Text = ""
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
   
    
    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            Me.Hide()
            frmHobbies.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmHobbies.cmbHostelerID.Text = dr.Cells(0).Value.ToString()
            frmHobbies.txtHostelerName.Text = dr.Cells(1).Value.ToString()
            frmHobbies.txtHobby1.Text = dr.Cells(3).Value.ToString()
            frmHobbies.txtHobby2.Text = dr.Cells(4).Value.ToString()
            frmHobbies.txtHobby3.Text = dr.Cells(5).Value.ToString()
            frmHobbies.txtHobby4.Text = dr.Cells(6).Value.ToString()
            frmHobbies.txtHobby5.Text = dr.Cells(7).Value.ToString()
            frmHobbies.txtHobby6.Text = dr.Cells(8).Value.ToString()
            frmHobbies.txtHobby7.Text = dr.Cells(9).Value.ToString()
            frmHobbies.txtHobby8.Text = dr.Cells(10).Value.ToString()
            frmHobbies.txtHobbyList.Text = dr.Cells(11).Value.ToString()
            frmHobbies.cmbHostelerID.Enabled = False
            frmHobbies.txtHostelerName.Enabled = False
            frmHobbies.txtHobbyList.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmHobbiesRecord_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmHobbies.Show()
    End Sub

    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
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
    Private Sub btnExport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport1.Click
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

    
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
      
        Try
            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID],HostelerName as [Hosteler Name], from Hostelers where HostelerName='" & cmbHostelerName.Text & "' order by HostelerName", con)
            cmd = New OleDbCommand("SELECT Hostelers.HostelerID as [Hosteler ID],Hostelers.HostelerName as [Hosteler Name],Hobbies.Hobby1 as [Hobby 1], Hobbies.Hobby2 as [Hobby 2], Hobbies.Hobby3 as [Hobby 3], Hobbies.Hobby4 as [Hobby 4], Hobbies.Hobby5 as [Hobby 5], Hobbies.Hobby6 as [Hobby 6], Hobbies.Hobby7 as [Hobby 7], Hobbies.Hobby8 as [Hobby 8],Hobbies.HobbyList as [Hobby List] FROM Hostelers INNER JOIN Hobbies ON Hostelers.[HostelerID] = Hobbies.[HostelerID] where  Hobbies.HostelerID=Hostelers.HostelerID order by HostelerName ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Hobbies")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView2.DataSource = myDataSet.Tables("Hobbies").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

 
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        txtHobby.Text = ""
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try
            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID],HostelerName as [Hosteler Name], from Hostelers where HostelerName='" & cmbHostelerName.Text & "' order by HostelerName", con)
            cmd = New OleDbCommand("SELECT Hostelers.HostelerID as [Hosteler ID],Hostelers.HostelerName as [Hosteler Name],Hobbies.Hobby1 as [Hobby 1], Hobbies.Hobby2 as [Hobby 2], Hobbies.Hobby3 as [Hobby 3], Hobbies.Hobby4 as [Hobby 4], Hobbies.Hobby5 as [Hobby 5], Hobbies.Hobby6 as [Hobby 6], Hobbies.Hobby7 as [Hobby 7], Hobbies.Hobby8 as [Hobby 8],Hobbies.HobbyList as [Hobby List] FROM Hostelers INNER JOIN Hobbies ON Hostelers.[HostelerID] = Hobbies.[HostelerID] where Hobby1 = '" & txtHobby.Text & "' or Hobby2 = '" & txtHobby.Text & "' or Hobby3 = '" & txtHobby.Text & "' or Hobby4 = '" & txtHobby.Text & "' or Hobby5 = '" & txtHobby.Text & "' or Hobby6 = '" & txtHobby.Text & "' or Hobby7 = '" & txtHobby.Text & "'  or Hobby8 = '" & txtHobby.Text & "'   and Hobbies.HostelerID=Hostelers.HostelerID order by HostelerName ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Hobbies")
            DataGridView2.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView2.DataSource = myDataSet.Tables("Hobbies").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

   
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
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

    Private Sub txtHostelerName_TextChanged1(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHostelerName.TextChanged
        txtUSN.Text = ""
        Try

            con = New OleDbConnection(cs)
            con.Open()
            'cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID],HostelerName as [Hosteler Name], from Hostelers where HostelerName='" & cmbHostelerName.Text & "' order by HostelerName", con)
            cmd = New OleDbCommand("SELECT Hostelers.HostelerID as [Hosteler ID],Hostelers.HostelerName as [Hosteler Name],Hostelers.USN as [USN],Hobbies.Hobby1 as [Hobby 1], Hobbies.Hobby2 as [Hobby 2], Hobbies.Hobby3 as [Hobby 3], Hobbies.Hobby4 as [Hobby 4], Hobbies.Hobby5 as [Hobby 5], Hobbies.Hobby6 as [Hobby 6], Hobbies.Hobby7 as [Hobby 7], Hobbies.Hobby8 as [Hobby 8],Hobbies.HobbyList as [Hobby List] FROM Hostelers INNER JOIN Hobbies ON Hostelers.[HostelerID] = Hobbies.[HostelerID] where Hostelers.HostelerName like '%" & txtHostelerName.Text & "%' and Hobbies.HostelerID=Hostelers.HostelerID order by HostelerName ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostelers")
            myDA.Fill(myDataSet, "Hobbies")
            DataGridView3.DataSource = myDataSet.Tables("Hostelers").DefaultView
            DataGridView3.DataSource = myDataSet.Tables("Hobbies").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub txtUSN_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUSN.TextChanged
        txtHostelerName.Text = ""
        Try
        con = New OleDbConnection(cs)
        con.Open()
        'cmd = New OleDbCommand("SELECT HostelerID as [Hosteler ID],HostelerName as [Hosteler Name], from Hostelers where HostelerName='" & cmbHostelerName.Text & "' order by HostelerName", con)
        cmd = New OleDbCommand("SELECT Hostelers.HostelerID as [Hosteler ID],Hostelers.HostelerName as [Hosteler Name],Hostelers.USN as [USN],Hobbies.Hobby1 as [Hobby 1], Hobbies.Hobby2 as [Hobby 2], Hobbies.Hobby3 as [Hobby 3], Hobbies.Hobby4 as [Hobby 4], Hobbies.Hobby5 as [Hobby 5], Hobbies.Hobby6 as [Hobby 6], Hobbies.Hobby7 as [Hobby 7], Hobbies.Hobby8 as [Hobby 8],Hobbies.HobbyList as [Hobby List] FROM Hostelers INNER JOIN Hobbies ON Hostelers.[HostelerID] = Hobbies.[HostelerID] where Hostelers.USN like '%" & txtUSN.Text & "%' and Hobbies.HostelerID=Hostelers.HostelerID order by HostelerName ", con)
        Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
        Dim myDataSet As DataSet = New DataSet()
        myDA.Fill(myDataSet, "Hostelers")
        myDA.Fill(myDataSet, "Hobbies")
        DataGridView3.DataSource = myDataSet.Tables("Hostelers").DefaultView
        DataGridView3.DataSource = myDataSet.Tables("Hobbies").DefaultView
        con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class