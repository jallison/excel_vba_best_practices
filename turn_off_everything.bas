VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'* From https://www.microsoft.com/en-us/microsoft-365/blog/2009/03/12/excel-vba-performance-coding-best-practices/
'* Author: Diego Oppenheimer

'*******************************************************************
'* Turn Off Everything But the Essentials While Your Code is Running
'*******************************************************************


'* Get current state of various Excel settings; put this at the beginning of your code
screenUpdateState = Application.ScreenUpdating
statusBarState = Application.DisplayStatusBar
calcState = Application.Calculation
eventsState = Application.EnableEvents
displayPageBreakState = ActiveSheet.DisplayPageBreaks '* note this is a sheet-level setting

'* turn off some Excel functionality so your code runs faster
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False '*note this is a sheet-level setting

'* >>your code goes here<<

'* after your code runs, restore state; put this at the end of your code
Application.ScreenUpdating = screenUpdateState
Application.DisplayStatusBar = statusBarState
Application.Calculation = calcState
Application.EnableEvents = eventsState
ActiveSheet.DisplayPageBreaks = displayPageBreaksState '*note this is a sheet-level setting

'Application.ScreenUpdating: This setting tells Excel to not redraw the screen while False.
'    The benefit here is that you probably don’t need Excel using up resources trying to draw
'    the screen since it’s changing faster than the user can perceive. Since it requires lots
'    of resources to draw the screen so frequently, just turn off drawing the screen until the
'    end of your code execution. Be sure to turn it back on right before your code ends.
'
'Application.DisplayStatusBar: This setting tells Excel to stop showing status while False. For
'    example, if you use VBA to copy/paste a range, while the paste is completing Excel will show
'    the progress of that operation on the status bar. Turning off screen updating is separate from
'    turning off the status bar display so that you can disable screen updating but still provide
'    feedback to the user, if desired. Again, turn it back on right before your code ends execution.
'
'Application.Calculation: This setting allows you to programmatically set Excel’s calculation mode.
'    “Manual” (xlCalculationManual) mode means Excel waits for the user (or your code) to explicitly
'    initiate calculation. “Automatic” is the default and means that Excel decides when to recalculate
'    the workbook (e.g. when you enter a new formula on the sheet). Since recalculating your workbook
'    can be time and resource intensive, you might not want Excel triggering a recalc every time you
'    change a cell value. Turn off calculation while your code executes, then set the mode back.
'    Note: setting the mode back to “Automatic” (xlCalculationAutomatic) will trigger a recalc.
'
'Application.EnableEvents: This setting tells Excel to not fire events while False. While looking into
'    Excel VBA performance issues I learned that some desktop search tools implement event listeners
'    (probably to better track document contents as it changes). You might not want Excel firing an
'    event for every cell you’re changing via code, and turning off events will speed up your VBA code
'    performance if there is a COM Add-In listening in on Excel events. (Thanks to Doug Jenkins for
'    pointing this out in my earlier post).
'
'ActiveSheet.DisplayPageBreaks: A good description of this setting already exists: http://support.microsoft.com/kb/199505
'    (Thanks to David McRitchie for pointing this out).
