Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form1
    Dim xlApp As New Excel.Application
    Public OrigIp As String
    Public DesIp As String
    Dim listIP1 As New ArrayList
    Dim listIP2 As New ArrayList
    Public num1 As String
    Public num2 As String
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
        sr.ReadLine.Remove(0)
       
        While (Not sr.EndOfStream)
            newline = sr.ReadLine.Split(",")
            Dim newrow As DataRow = dt.NewRow

            For i As Integer = 1 To 4
                If (Hex(newline(7)).Length) = 8 Then
                    OrigIp = Convert.ToInt32(Hex(newline(7)).Substring(Hex(newline(7)).Length - (2 * i), 2), 16)
                ElseIf (Hex(newline(7)).Length) = 7 Then
                    Dim kar As String = "0" + Hex(newline(7))
                    OrigIp = Convert.ToInt32(kar.Substring(kar.Length - (2 * i), 2), 16)
                End If
                listIP1.Add(OrigIp)
            Next

            For i As Integer = 1 To 4
                If (Hex(newline(28)).Length) = 8 Then
                    DesIp = Convert.ToInt32(Hex(newline(28)).Substring(Hex(newline(28)).Length - (2 * i), 2), 16)
                ElseIf (Hex(newline(28)).Length) = 7 Then
                    Dim kar As String = "0" + Hex(newline(28))
                    DesIp = Convert.ToInt32(kar.Substring(kar.Length - (2 * i), 2), 16)
                ElseIf (Hex(newline(28)).Length) = 1 Then
                    DesIp = 0
                End If
                listIP2.Add(DesIp)
            Next

            num1 = (Join(listIP1.ToArray, "."))
            num2 = (Join(listIP2.ToArray, "."))
            newrow.ItemArray = {Date.FromOADate((((Convert.ToInt32(newline(4)) + (3 * 3600)) / 86400) + 25569)), num1, newline(8).Trim(New Char() {""""}), newline(9).Trim(New Char() {""""}), num2, newline(29).Trim(New Char() {""""}), Date.FromOADate((((Convert.ToInt32(newline(47)) + (3 * 3600)) / 86400) + 25569)), Date.FromOADate((((Convert.ToInt32(newline(48)) + (3 * 3600)) / 86400) + 25569)), newline(49).Trim(New Char() {""""}), newline(51), newline(52).Trim(New Char() {""""}), newline(53).Trim(New Char() {""""}), newline(54).Trim(New Char() {""""}), newline(55), newline(56).Trim(New Char() {""""}), newline(57).Trim(New Char() {""""}), newline(80).Trim(New Char() {""""}), newline(81).Trim(New Char() {""""})}
            dt.Rows.Add(newrow)
            listIP1.Clear()
            listIP2.Clear()
        End While
        DataGridView1.DataSource = dt
    End Sub

    Sub line7()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call OpenFile()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

    End Sub
End Class
