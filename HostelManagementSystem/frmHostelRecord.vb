Imports System.Data.OleDb
Public Class frmHostelRecord
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"
    Sub GetData()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT HostelName as [Hostel Name],HostelType as [Hostel Type],Hostel_Address as [Address],Hostel_Phone as [Phone No],ManagedBy as [Managed By], Hostel_ContactNo as [Contact No] from Hostel ", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Hostel")
            DataGridView1.DataSource = myDataSet.Tables("Hostel").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub frmHostelRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GetData()
    End Sub

    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            Me.Hide()
            frmHostel.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmHostel.txtHostelName.Text = dr.Cells(0).Value.ToString()
            frmHostel.TextBox1.Text = dr.Cells(0).Value.ToString()
            frmHostel.cmbHostelType.Text = dr.Cells(1).Value.ToString()
            frmHostel.txtAddress.Text = dr.Cells(2).Value.ToString()
            frmHostel.txtPhoneNo.Text = dr.Cells(3).Value.ToString()
            frmHostel.txtManagedBy.Text = dr.Cells(4).Value.ToString()
            frmHostel.txtContactNo.Text = dr.Cells(5).Value.ToString()
            frmHostel.btnUpdate_record.Enabled = True
            frmHostel.btnDelete.Enabled = True
            frmHostel.btnSave.Enabled = False
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
End Class