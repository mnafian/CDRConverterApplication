Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form1
    Dim xlApp As New Excel.Application
    Sub OpenFile()
        OpenFileDialog1.ShowDialog()
        TextBox1.Text = OpenFileDialog1.FileName

        Dim sr As New IO.StreamReader(TextBox1.Text)
        Dim dt As New DataTable

        Dim newline() As String = sr.ReadLine.Split(","c)
        dt.Columns.AddRange({New DataColumn(newline(4).Trim(New Char() {""""})), _
                             New DataColumn(newline(7).Trim(New Char() {""""})), _
                             New DataColumn(newline(8).Trim(New Char() {""""})), _
                             New DataColumn(newline(9).Trim(New Char() {""""})), _
                             New DataColumn(newline(28).Trim(New Char() {""""})), _
                             New DataColumn(newline(29).Trim(New Char() {""""})), _
                             New DataColumn(newline(47).Trim(New Char() {""""})), _
                             New DataColumn(newline(48).Trim(New Char() {""""})), _
                             New DataColumn(newline(49).Trim(New Char() {""""})), _
                             New DataColumn(newline(51).Trim(New Char() {""""})), _
                             New DataColumn(newline(52).Trim(New Char() {""""})), _
                             New DataColumn(newline(53).Trim(New Char() {""""})), _
                             New DataColumn(newline(54).Trim(New Char() {""""})), _
                             New DataColumn(newline(55).Trim(New Char() {""""})), _
                             New DataColumn(newline(56).Trim(New Char() {""""})), _
                             New DataColumn(newline(57).Trim(New Char() {""""})), _
                             New DataColumn(newline(80).Trim(New Char() {""""})), _
                             New DataColumn(newline(81).Trim(New Char() {""""}))})
        sr.ReadLine.Remove(2)
        While (Not sr.EndOfStream)
            newline = sr.ReadLine.Split(",")
            Dim newrow As DataRow = dt.NewRow
            newrow.ItemArray = {(((Convert.ToInt32(newline(4)) + (3 * 3600)) / 86400) + 25569), newline(7), newline(8).Trim(New Char() {""""}), newline(9).Trim(New Char() {""""}), newline(28), newline(29).Trim(New Char() {""""}), newline(47), newline(48), newline(49).Trim(New Char() {""""}), newline(51), newline(52).Trim(New Char() {""""}), newline(53).Trim(New Char() {""""}), newline(54).Trim(New Char() {""""}), newline(55), newline(56).Trim(New Char() {""""}), newline(57).Trim(New Char() {""""}), newline(80).Trim(New Char() {""""}), newline(81).Trim(New Char() {""""})}
            dt.Rows.Add(newrow)
        End While

        DataGridView1.DataSource = dt
        DataGridView1.Columns(0).DefaultCellStyle.Format = "MM'/'dd'/'yyyy"
    End Sub

    Sub CopyData()

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
                        '.Range("A2:A").Formula = "=(((" + DataGridView1.Rows(I).Cells(j).Value + "+(3*3600))/86400)+25569)"
                    Next j
                Next I
                '.Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 10
                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()


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
