Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form1
    Dim xlApp As New Excel.Application
    Sub OpenFile()
        OpenFileDialog1.ShowDialog()
        TextBox1.Text = OpenFileDialog1.FileName

        Dim sr As New IO.StreamReader(TextBox1.Text)
        Dim dt As New DataTable
        Dim newline() As String = sr.ReadLine.Split(","c)
        dt.Columns.AddRange({New DataColumn(newline(4)), _
                             New DataColumn(newline(7)), _
                             New DataColumn(newline(8)), _
                             New DataColumn(newline(9)), _
                             New DataColumn(newline(28)), _
                             New DataColumn(newline(29)), _
                             New DataColumn(newline(47)), _
                             New DataColumn(newline(48)), _
                             New DataColumn(newline(49)), _
                             New DataColumn(newline(51)), _
                             New DataColumn(newline(52)), _
                             New DataColumn(newline(53)), _
                             New DataColumn(newline(54)), _
                             New DataColumn(newline(55)), _
                             New DataColumn(newline(56)), _
                             New DataColumn(newline(57)), _
                             New DataColumn(newline(80)), _
                             New DataColumn(newline(81))})
        While (Not sr.EndOfStream)
            newline = sr.ReadLine.Split(","c)
            Dim newrow As DataRow = dt.NewRow
            newrow.ItemArray = {newline(4), newline(7), newline(8), newline(9), newline(28), newline(29), newline(47), newline(48), newline(49), newline(51), newline(52), newline(53), newline(54), newline(55), newline(56), newline(57), newline(80), newline(81)}
            dt.Rows.Add(newrow)
        End While
        DataGridView1.DataSource = dt
    End Sub

    Sub ConvertFile()
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True
            rowsTotal = DataGridView1.RowCount - 1
            colsTotal = DataGridView1.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                'For iC = 0 To colsTotal
                '.Cells(1, iC + 1).Value = DataGridView1.Columns(iC).HeaderText
                'Next
                For I = 0 To rowsTotal
                    For j = 0 To colsTotal
                        .Cells(I + 1, j + 1).value = DataGridView1.Rows(I).Cells(j).Value

                    Next j
                Next I
                '.Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 10
                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
                '.Range("A1:A8").Formula = "=(((" + DataGridView1.Rows(I).Cells(j).Value + "+(3*3600))/86400)+25569)"

            End With
        Catch ex As Exception
            MsgBox("Export Excel Error " & ex.Message)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call OpenFile()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call ConvertFile()


    End Sub
End Class
