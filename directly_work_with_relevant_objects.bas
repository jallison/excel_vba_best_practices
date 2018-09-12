Attribute VB_Name = "Module1"
'* From: https://stackoverflow.com/questions/10714251/how-to-avoid-using-select-in-excel-vba?rq=1
'* Author: Siddharth Rout

'* Two Main reasons why .Select/.Activate/Selection/Activecell/Activesheet/Activeworkbook etc... should be avoided
'* It slows down your code.
'* It is usually the main cause of runtime errors.

'* How do we avoid it?
'* 1) Directly work with the relevant objects

'* Consider this code
Sheets("Sheet1").Activate
Range("A1").Select
Selection.Value = "Blah"
Selection.NumberFormat = "@"

'* This code can also be written as
With Sheets("Sheet1").Range("A1")
    .Value = "Blah"
    .NumberFormat = "@"
End With

'* 2) If required declare your variables. The same code above can be written as
Dim ws As Worksheet

Set ws = Sheets("Sheet1")

With ws.Range("A1")
    .Value = "Blah"
    .NumberFormat = "@"
End With
