Dim path
path = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\"

Dim fileArray
fileArray = Array("rollingsales_manhattan","rollingsales_brooklyn","rollingsales_queens","rollingsales_bronx","rollingsales_statenisland")

For Each xlsFile In fileArray
  Dim oExcel
  Set oExcel = CreateObject("Excel.Application")
  Dim oBook
  Set oBook = oExcel.Workbooks.Open(path & xlsFile & ".xls")
  oBook.SaveAs path & xlsFile & ".csv", 6
  oBook.Close False
  oExcel.Quit
Next
