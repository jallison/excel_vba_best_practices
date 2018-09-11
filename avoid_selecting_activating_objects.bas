VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'* From: https://www.microsoft.com/en-us/microsoft-365/blog/2009/03/12/excel-vba-performance-coding-best-practices/
'* Author: Diego Oppenheimer

'***************************************
'* Avoid Selecting / Activating Objects*
'***************************************

'* Slow code

For i = 0 To ActiveSheet.Shapes.Count
   ActiveSheet.Shapes(i).Select
   Selection.Text = “Hello”
Next i

'* Fast code

For i = 0 To ActiveSheet.Shapes.Count
   ActiveSheet.Shapes(i).TextEffect.Text = “Hello”
Next i

