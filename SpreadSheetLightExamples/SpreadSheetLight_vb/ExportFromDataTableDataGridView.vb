Public Class ExportFromDataTableDataGridView
    Private ops As New Operations
    Private Sub ExportFromDataTableDataGridView_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.DataSource = ops.ReadCustomersFromXml()
    End Sub
    Private Sub exportToExcelButton_Click(sender As Object, e As EventArgs) Handles exportToExcelButton.Click
        If ops.ExportDataTable(CType(DataGridView1.DataSource, DataTable)) Then
            MessageBox.Show("Exported")
        Else
            MessageBox.Show($"Failed{Environment.NewLine}{ops.Exception.Message}")
        End If
    End Sub
End Class