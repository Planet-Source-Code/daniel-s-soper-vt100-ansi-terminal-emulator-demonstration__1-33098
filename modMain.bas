Attribute VB_Name = "modMain"
' This is my attempt at creating an ANSI-enabled VT100
' terminal emulator in Visual Basic. It is admittedly
' incomplete, but it seems to be relatively stable in its
' current form.

' Please send me feedback regarding any bugs, additions, or
' other problems that you encounter.

' If you like this humble attempt, then please vote for it!

' Thanks,

' Daniel S. Soper
' Creator


Option Explicit


Sub Main()
    Call TermInitialize
    frmMain.frameConnection.Top = frmMain.picTerm.Top + frmMain.picTerm.Height + 5
    frmMain.Width = frmMain.picTerm.Width * Screen.TwipsPerPixelX
    frmMain.Height = (frmMain.frameConnection.Top + 110) * Screen.TwipsPerPixelY
    frmMain.frameConnection.Left = (frmMain.picTerm.Width - frmMain.frameConnection.Width) / 2
    frmMain.Show
End Sub
