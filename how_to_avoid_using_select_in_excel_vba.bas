Attribute VB_Name = "Module1"
'* From: https://stackoverflow.com/questions/10714251/how-to-avoid-using-select-in-excel-vba
'* Author: chris neilsen

'*****************************************
'* How to avoid using Select in Excel VBA*
'*****************************************

'* Some examples of how to avoid select
'* Use Dim'd variables
Dim rng As Range
'* Set the variable to the required range. There are many ways to refer to a single-cell range
Set rng = Range("A1")
Set rng = Cells(1, 1)
Set rng = Range("NamedRange")
'* or a multi-cell range
Set rng = Range("A1:B10")
Set rng = Range("A1", "B10")
Set rng = Range(Cells(1, 1), Cells(10, 2))
Set rng = Range("AnotherNamedRange")
Set rng = Range("A1").Resize(10, 2)
'* You can use the shortcut to the Evaluate method, but this is less efficient and should generally be avoided in production code.
Set rng = [A1]
Set rng = [A1:B10]
'* All the above examples refer to cells on the active sheet. Unless you specifically want to work only with the active sheet, it is better to Dim a Worksheet variable too
Dim ws As Worksheet
Set ws = Worksheets("Sheet1")
Set rng = ws.Cells(1, 1)
With ws
    Set rng = .Range(.Cells(1, 1), .Cells(2, 10))
End With
'* If you do want to work with the ActiveSheet, for clarity it's best to be explicit. But take care, as some Worksheet methods change the active sheet.
Set rng = ActiveSheet.Range("A1")
'* Again, this refers to the active workbook. Unless you specifically want to work only with the ActiveWorkbook or ThisWorkbook, it is better to Dim a Workbook variable too.
Dim wb As Workbook
Set wb = Application.Workbooks("Book1")
Set rng = wb.Worksheets("Sheet1").Range("A1")
'* If you do want to work with the ActiveWorkbook, for clarity it's best to be explicit. But take care, as many WorkBook methods change the active book.
Set rng = ActiveWorkbook.Worksheets("Sheet1").Range("A1")
'* You can also use the ThisWorkbook object to refer to the book containing the running code.
Set rng = ThisWorkbook.Worksheets("Sheet1").Range("A1")


'* A common (bad) piece of code is to open a book, get some data then close again
'* This is bad:
Sub foo()
    Dim v As Variant
    Workbooks("Book1.xlsx").Sheets(1).Range("A1").Clear
    Workbooks.Open ("C:\Path\To\SomeClosedBook.xlsx")
    v = ActiveWorkbook.Sheets(1).Range("A1").Value
    Workbooks("SomeAlreadyOpenBook.xlsx").Activate
    ActiveWorkbook.Sheets("SomeSheet").Range("A1").Value = v
    Workbooks(2).Activate
    ActiveWorkbook.Close
End Sub
'* And would be better like:
Sub foo()
    Dim v As Variant
    Dim wb1 As Workbook
    Dim wb2 As Workbook
    Set wb1 = Workbooks("SomeAlreadyOpenBook.xlsx")
    Set wb2 = Workbooks.Open("C:\Path\To\SomeClosedBook.xlsx")
    v = wb2.Sheets("SomeSheet").Range("A1").Value
    wb1.Sheets("SomeOtherSheet").Range("A1").Value = v
    wb2.Close
End Sub



'* Pass ranges to your Sub's and Function's as Range variables
Sub ClearRange(r As Range)
    r.ClearContents
    '....
End Sub
Sub MyMacro()
    Dim rng As Range
    Set rng = ThisWorkbook.Worksheets("SomeSheet").Range("A1:B10")
    ClearRange rng
End Sub


'* You should also apply Methods (such as Find and Copy) to variables
Dim rng1 As Range
Dim rng2 As Range
Set rng1 = ThisWorkbook.Worksheets("SomeSheet").Range("A1:A10")
Set rng2 = ThisWorkbook.Worksheets("SomeSheet").Range("B1:B10")
rng1.Copy rng2


'* If you are looping over a range of cells it is often better (faster) to copy the range values to a variant array first and loop over that
Dim dat As Variant
Dim rng As Range
Dim i As Long

Set rng = ThisWorkbook.Worksheets("SomeSheet").Range("A1:A10000")
dat = rng.Value  ' dat is now array (1 to 10000, 1 to 1)
For i = LBound(dat, 1) To UBound(dat, 1)
    dat(i, 1) = dat(i, 1) * 10 'or whatever operation you need to perform
Next
rng.Value = dat ' put new values back on sheet
'* This is a small taster for what's possible.
