

Partial Public Class HMS_DBDataSet
    Partial Class FeePaymentDataTable

        Private Sub FeePaymentDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.DuePaymentColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

End Class

Namespace HMS_DBDataSetTableAdapters
    
    Partial Public Class FeePaymentTableAdapter
    End Class
End Namespace
