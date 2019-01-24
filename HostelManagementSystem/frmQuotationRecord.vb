Imports System.Data.OleDb
Public Class frmQuotationRecord
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\HMS_DB.accdb;Persist Security Info=False;"

    Private Sub frmQuotationRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GetData()
    End Sub

    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            Me.Hide()
            frmQuotation.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmQuotation.txtQuotationID.Text = dr.Cells(0).Value.ToString()
            frmQuotation.cmbHostelName.Text = dr.Cells(1).Value.ToString()
            frmQuotation.txtStudentName.Text = dr.Cells(2).Value.ToString()
            frmQuotation.txtContactNo.Text = dr.Cells(3).Value.ToString()
            frmQuotation.txtChargesPerMonth1.Text = dr.Cells(4).Value.ToString()
            frmQuotation.txtMonth1.Text = dr.Cells(5).Value.ToString()
            frmQuotation.txtTotalCharges1.Text = dr.Cells(6).Value.ToString()
            frmQuotation.txtChargesPerMonth2.Text = dr.Cells(7).Value.ToString()
            frmQuotation.txtMonth2.Text = dr.Cells(8).Value.ToString()
            frmQuotation.txtTotalCharges2.Text = dr.Cells(9).Value.ToString()
            frmQuotation.txtGrandTotal.Text = dr.Cells(10).Value.ToString()
            frmQuotation.dtpQuotationDate.Text = dr.Cells(11).Value.ToString()
            frmQuotation.btnUpdate_record.Enabled = True
            frmQuotation.btnDelete.Enabled = True
            frmQuotation.btnSave.Enabled = False
            frmQuotation.btnPrint.Enabled = True
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
    Sub GetData()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT QuotationID as [Quotation Id] , Hostelname as [Hostel Name],StudentName as [Student Name],ContactNo as [Contact No],ChargesPerMonth1 as [Rental Charges Per Month],NoOfMonth1 as [Total Months],TotalCharges1 as [Total Rental Charges],ChargesPerMonth2 as [Mess Charges Per Month],NoOfMonth2 as [Total Months1],TotalCharges2 as [Total Mess Charges],GrandTotal as [Grand Total],QuotationDate as [Quotation Date] from Quotation  order by quotationdate", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Quotation")
            DataGridView1.DataSource = myDataSet.Tables("Quotation").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class