  Imports System.IO
  Imports Microsoft.Office.Interop
  Imports System.Text

   Public Class Form1
    Dim folder As String = Directory.GetCurrentDirectory()
    
   End Sub
    
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click'
    
    Dim oexcel As Object
        oexcel = CreateObject("Excel.Application")
        Dim obook As Excel.Workbook
        Dim osheet As Excel.Worksheet
        Dim fnExc As String = (folder + "\example.xlsx")
        Dim curLine As Integer = 0

        ' Only use once
        obook = oexcel.Workbooks.Add

        If oexcel.Application.Sheets.Count() < 1 Then
            osheet = CType(obook.Worksheets.Add(), Excel.Worksheet)
        Else
            osheet = oexcel.Worksheets(1)
        End If
        osheet.Name = "Example"
        osheet.Range("A1").Value = TextBox1.Text
        ' Jump one row
        osheet.Range("A1", "A1").Insert(Shift:=Excel.XlDirection.xlDown)
        osheet.Range("A1").Value = Textbox2.Text

        obook.SaveAs(fnExc)
        obook.Close()
        obook = Nothing
    End Sub
