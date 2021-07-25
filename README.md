1. Create a Excel file
3. Create e new project
4. Choose Windows Form Application
5. Add two Textboxes (if you want)
6. Add one Button
7. Add reference (COM) Microsoft Excel 16.0 Object Library (Depends on your Microsoft Office version, if you're using 2013 use 13.0)
8. Import (System.IO, Microsoft.Office.Interop and System.Text)

If you want to use a different folder remove the "Dim folder" line near to the Public Class Form1 and change the:
Dim fnExc As String = "C:\Temp\ExcelTest\Example.xlsx"
To
Dim fnExc As String = "Your Directory"

Else if you want to use the current folder (Where the app is compiled) add the following line above teh Public Class Form1:
    Dim folder As String = Directory.GetCurrentDirectory()
    
    Change the
  Dim fnExc As String = "C:\Temp\ExcelTest\Example.xlsx"
    To
    Dim fnExc As String = (folder + "\Example.xlsx")
    
