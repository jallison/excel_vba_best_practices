Attribute VB_Name = "Module1"
'* From: https://stackoverflow.com/questions/10714251/how-to-avoid-using-select-in-excel-vba?rq=1
'* Author: Rick Teachey

'* Named ranges make your code easier to read and understand.
'* Example:

Dim Months As Range
Dim MonthlySales As Range

Set Months = Range("Months")
'e.g, "Months" might be a named range referring to A1:A12

Set MonthlySales = Range("MonthlySales")
'e.g, "Monthly Sales" might be a named range referring to B1:B12

Dim Month As Range
For Each Month In Months
    Debug.Print MonthlySales(Month.Row)
Next Month

'* It is pretty obvious what the named ranges Months and MonthlySales
'* contain, and what the procedure is doing. Why is this important?
'* Partially because it is easier for other people to understand it,
'* but even if you are the only person who will ever see or use your
'* code, you should still use named ranges and good variable names
'* because YOU WILL FORGET what you meant to do with it a year later,
'* and you will waste 30 minutes just figuring out what your code is doing.

'* Named ranges ensure that your macros do not break when (not if!) the
'* configuration of the spreadsheet changes.

'*Consider, if the above example had been written like this:
Dim rng1 As Range
Dim rng2 As Range

Set rng1 = Range("A1:A12")
Set rng2 = Range("B1:B12")

Dim rng3 As Range
For Each rng3 In rng1
    Debug.Print rng2(rng3.Row)
Next rng3

'* This code will work just fine at first - that is until you or a future
'* user decides "gee wiz, I think I'm going to add a new column with the
'* year in Column A!", or put an expenses column between the months and
'* sales columns, or add a header to each column. Now, your code is broken.
'* And because you used terrible variable names, it will take you a lot more
'* time to figure out how to fix it than it should take. If you had used named
'* ranges to begin with, the Months and Sales columns could be moved around
'* all you like, and your code will continue working just fine.
