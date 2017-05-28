'
' Needed for several properties setters 
'
Imports DOS = DocumentFormat.OpenXml.Spreadsheet
Imports SpreadsheetLight
Public Class Operations
    ''' <summary>
    ''' date used to show in DataGridView which in turn gets exported to Excel
    ''' </summary>
    ''' <returns></returns>
    Public Function ReadCustomersFromXml() As DataTable
        Dim ds As New DataSet
        ds.ReadXmlSchema(IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Customers.xsd"))
        ds.ReadXml(IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Customers.xml"))
        ds.Tables("Customer").Columns.Add(New DataColumn With {.ColumnName = "MyDate", .DataType = GetType(DateTime), .DefaultValue = Now})

        Dim counter As Integer = 1
        For Each row As DataRow In ds.Tables("Customer").Rows
            row.SetField(Of DateTime)("MyDate", row.Field(Of DateTime)("MyDate").AddDays(counter))
            counter += 1
        Next

        Return ds.Tables("Customer")

    End Function
    Private theExportFileName As String = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExportSample.xlsx")
    Public ReadOnly Property ExportFileName As String
        Get
            Return theExportFileName
        End Get
    End Property
    Private theException As Exception
    Public ReadOnly Property Exception As Exception
        Get
            Return theException
        End Get
    End Property
    ''' <summary>
    ''' Basic export of DataTable to a worksheet
    ''' </summary>
    ''' <param name="table"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' How did I figure this out? 
    ''' Used Google to search for operations I wanted.
    ''' Total time to write 99.99% of code, 30 minutes first time
    ''' How long to do the column header (LOL), 15 minutes on top of the 30 to figure out the colors.
    ''' Why 15 minutes, because I was exploring and having fun.
    ''' </remarks>
    Public Function ExportDataTable(ByVal table As DataTable) As Boolean
        Try
            Using sl As New SLDocument(ExportFileName)
                Dim startRow As Integer = 1
                Dim startColumn As Integer = 1

                ' redundent
                sl.SelectWorksheet("Sheet1")

                ' clear cells if this is ran more than once and the row or column count changes
                sl.ClearCellContent()

                ' import DataTable with column headers
                sl.ImportDataTable(startRow, startColumn, table, True)

                ' set the Date style
                Dim dateStyle = sl.CreateStyle
                dateStyle.FormatCode = "mm-dd-yyyy"

                sl.SetCellStyle(2, table.Columns("MyDate").Ordinal + 1, table.Rows.Count - 1, table.Columns("MyDate").Ordinal + 1, dateStyle)

                ' set the column header stype
                Dim headerSyle = sl.CreateStyle
                headerSyle.Font.FontColor = Color.White
                headerSyle.Font.Strike = False
                headerSyle.Fill.SetPattern(DOS.PatternValues.Solid, Color.Green, Color.White)
                headerSyle.Font.Underline = DOS.UnderlineValues.None
                headerSyle.Font.Bold = True
                headerSyle.Font.Italic = False
                sl.SetCellStyle(1, 1, 1, table.Columns.Count, headerSyle)

                ' auto-fit the columns
                sl.AutoFitColumn(1, table.Columns.Count)

                ' save back to the Excel file
                sl.Save()

            End Using
            Return True
        Catch ex As Exception
            theException = ex
            Return False
        End Try
    End Function
End Class
