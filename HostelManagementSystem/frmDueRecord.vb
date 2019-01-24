Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class frmDueRecord
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Try
            GroupBox2.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Hostelers.HostelerName as [Hosteler Name], Hostelers.USN as [USN], Hostelers.Purpose as [Department], DueAmount.TotalCharges as[Total Charges] , DueAmount.TotalDueAmount as[Due Amount], DueAmount.AcadYear as [Academic Year] FROM Hostelers INNER JOIN DueAmount ON Hostelers.[HostelerID] = DueAmount.[HostelerID]  where  TotalDueAmount > '0' order by Hostelername", con) ' Hostelers,FeePayment where Hostelers.HostelerID=FeePayment.HostelerID group by Hostelers.HostelerID,Hostelername,Hostelname,RoomNo having sum(DuePayment >0) 
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "DueAmount")
            'myDA.Fill(myDataSet, "Feepayment")
            DataGridView3.DataSource = myDataSet.Tables("DueAmount").DefaultView
            'DataGridView1.DataSource = myDataSet.Tables("Hostelers").DefaultView
            Dim sum As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                sum = sum + r.Cells(4).Value
            Next
            TextBox8.Text = sum
            Label1.Text = frmDueCharges.GetInWords(TextBox8.Text.Trim)
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class