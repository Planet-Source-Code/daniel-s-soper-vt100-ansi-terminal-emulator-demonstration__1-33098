Attribute VB_Name = "modVT100ANSI"
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

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
    Public Const PATCOPY = &HF00021
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function Beep Lib "kernel32.dll" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
    Public Const DSTINVERT = &H550009
    Public Const SRCCOPY = &HCC0020
Type Cur
    X As Integer
    Y As Integer
    Visible As Boolean
    Reverse As Boolean
End Type
Type Char
    Width As Integer
    Height As Integer
End Type
Type TermAttr
    XOFF As Boolean
    LNM As Boolean 'True = Set, False = Reset
    Wordwrap As Boolean 'True = Set, False = Reset
    Blink As Boolean
    Bold As Boolean
    Invisible As Boolean
    Reverse As Boolean
    Underline As Boolean
    ForeColor As Long
    BackColor As Long
End Type
Type CharAttr
    Text As Byte
    Bold As Boolean
    Blink As Boolean
    Underline As Boolean
    Reverse As Boolean
    Invisible As Boolean
    ForeColor As Long
    BackColor As Long
End Type
Type Brush
    Black As Long
    Red As Long
    Green As Long
    Yellow As Long
    Blue As Long
    Magenta As Long
    Cyan As Long
    White As Long
    Gray As Long
End Type

Public Terminal As PictureBox
Public Character As Char
Public Cursor As Cur, CursorSave As Cur
Public TerminalAttribute As TermAttr
Public InEscape As Boolean, ReceivingData As Boolean
Public EscString As String
Public Tabs(19) As Integer, BlinkSentry As Integer
Public TabNum As Integer
Public TerminalChar(24, 79) As CharAttr
Public Color As Brush

Public DebugText As String

Public Sub TermInitialize()
    Dim X As Integer, Y As Integer
    Set Terminal = frmMain.picTerm
    frmMain.ScaleMode = 3 'Pixel
    Terminal.ScaleMode = 3 'Pixel
    Terminal.FontName = "Terminal"
    Terminal.FontSize = 9
    Character.Height = Terminal.TextHeight("O")
    Character.Width = Terminal.TextWidth("O")
    BlinkSentry = 0
    BlinkSentry = 0
    Terminal.Cls
    Terminal.Refresh
    With Cursor
        .X = 0
        .Y = 0
        .Visible = True
        .Reverse = False
    End With
    With CursorSave
        .X = 0
        .Y = 0
        .Visible = True
        .Reverse = False
    End With
    InEscape = False
    ReceivingData = False
    With Terminal
        .Top = 0
        .Left = 0
        .Width = (80 * Character.Width) + 8
        .Height = (25 * Character.Height) + 3
        .ForeColor = QBColor(8)
        .BackColor = QBColor(0)
    End With
    With TerminalAttribute
        .XOFF = False
        .LNM = False
        .Wordwrap = False
        .Blink = False
        .Bold = False
        .Invisible = False
        .Reverse = False
        .Underline = False
        .ForeColor = Terminal.ForeColor
        .BackColor = Terminal.BackColor
    End With
    For Y = 0 To 24
        For X = 0 To 79
            With TerminalChar(Y, X)
                .Blink = False
                .Bold = False
                .Invisible = False
                .Reverse = False
                .Text = 32
                .Underline = False
                .ForeColor = Terminal.ForeColor
                .BackColor = Terminal.BackColor
            End With
        Next
    Next
    For X = 0 To 19
        Tabs(X) = X * 4
    Next
    TabNum = 0
    With Color
        .Black = CreateSolidBrush(vbBlack)
        .Blue = CreateSolidBrush(vbBlue)
        .Cyan = CreateSolidBrush(vbCyan)
        .Gray = CreateSolidBrush(QBColor(8))
        .Green = CreateSolidBrush(vbGreen)
        .Magenta = CreateSolidBrush(vbMagenta)
        .Red = CreateSolidBrush(vbRed)
        .White = CreateSolidBrush(vbWhite)
        .Yellow = CreateSolidBrush(vbYellow)
    End With
End Sub

Public Sub TermProcessInput(ByteIn As Byte)
    If TerminalAttribute.XOFF = False Then
        Cursor.Visible = False
        If InEscape = True Then
            Call TermProcessEscapeSequence(ByteIn)
        Else
            EscString = ""
            Select Case ByteIn
                Case 0 'NUL - Ignored on input
                Case 4 'EOT - End of Transmission
                    Call frmMain.Disconnect
                Case 5 'ENQ - Transmit answerback message
                    'Debug.Print "ENQ - Transmit answerback message"
                Case 7 'BEL - Sound bell tone
                    Call Beep(700, 250)
                Case 8 'BS - Move cursor left one character position
                    Cursor.X = Cursor.X - 1
                    Call TermCheckCursor
                Case 9 'HT - Move to next tab stop, or to right margin if no more tab stops
                    If TabNum = 19 Then
                        Cursor.X = 79
                    Else
                        TabNum = TabNum + 1
                        Cursor.X = Tabs(TabNum)
                    End If
                Case 10, 11, 12 'LF, VT, FF, - line feed or new line operation when in new line mode
                    If Cursor.Y >= 24 Then
                        Call TermScrollUp
                        Cursor.Y = 24
                        If TerminalAttribute.LNM = True Then 'If new line mode
                            Cursor.X = 0
                        End If
                    Else
                        Cursor.Y = Cursor.Y + 1
                        If TerminalAttribute.LNM = True Then 'If new line mode
                            Cursor.X = 0
                        End If
                    End If
                    Call TermCheckCursor
                Case 13 'CR - Move the cursor to the left margin on the current line
                    Cursor.X = 0
                Case 14 'SO - Invoke G1 character set, as designated by SCS escape sequence
                    'Debug.Print "SO - Invoke G1 character set, as designated by SCS escape sequence"
                Case 15 'SI - Select G0 character set, as selected by escape sequence
                    'Debug.Print "SI - Select G0 character set, as selected by escape sequence"
                Case 17 'XON - OK to Resume transmission
                    TerminalAttribute.XOFF = False
                Case 19 'XOFF - Stop transmission except for XOFF and XON escape sequences
                    TerminalAttribute.XOFF = True
                Case 24 'CAN - Immediately terminate control sequence, and display error character
                    'Debug.Print "CAN - Immediately terminate control sequence, and display error character"
                Case 26 'SUB - Interpreted as CAN
                    'Debug.Print "SUB - Interpreted as CAN"
                Case 27 'ESC - Invoke an escape sequence
                    InEscape = True
                Case Else 'Bytes to display on terminal screen
                    If CInt(ByteIn) >= 32 Then
                        Call TermCheckCursor
                        With TerminalChar(Cursor.Y, Cursor.X)
                            .BackColor = TerminalAttribute.BackColor
                            .Blink = TerminalAttribute.Blink
                            .Bold = TerminalAttribute.Bold
                            .ForeColor = TerminalAttribute.ForeColor
                            .Invisible = TerminalAttribute.Invisible
                            .Reverse = TerminalAttribute.Reverse
                            .Text = ByteIn
                            .Underline = TerminalAttribute.Underline
                        End With
                        Cursor.X = Cursor.X + 1
                        Call TermCheckCursor
                        DebugText = DebugText & Chr$(ByteIn)
                    Else
                        'Debug.Print "Unhandled: " & ByteIn
                    End If
            End Select
        End If
        Cursor.Visible = True
    Else
        If ByteIn = 17 Then 'XON - OK to Resume transmission
            TerminalAttribute.XOFF = False
        End If
    End If
End Sub

Private Sub TermProcessEscapeSequence(ByteIn As Byte)
    Dim X As Integer, Y As Integer, Z As Integer
    If EscString = "" Then 'If no Control Sequence Introducer exists...
        Select Case Chr$(ByteIn)
            Case "[", "(", ")", "#"
                EscString = EscString & Chr$(ByteIn)
            Case "Z" 'DECID - Identify Terminal
                Call TermSendData(Chr(27) & "[?1;0c")
                InEscape = False
            Case "=" 'DECKPAM - Keypad Application Mode
                InEscape = False
            Case ">" 'DECKPNM - Keypad Numeric Mode
                InEscape = False
            Case "8" 'DECRC - Restore Cursor
                With Cursor
                    .X = CursorSave.X
                    .Y = CursorSave.Y
                    .Reverse = CursorSave.Reverse
                    .Visible = CursorSave.Visible
                End With
                Call TermCheckCursor
                InEscape = False
            Case "7" 'DECSC - Save Cursor
                With CursorSave
                    .X = Cursor.X
                    .Y = Cursor.Y
                    .Reverse = Cursor.Reverse
                    .Visible = Cursor.Visible
                End With
                InEscape = False
            Case "H" 'HTS - Horizontal Tabulation Set
                If TabNum >= 19 Then
                    Tabs(19) = Cursor.X
                    TabNum = 19
                Else
                    Tabs(TabNum) = Cursor.X
                    TabNum = TabNum + 1
                End If
                InEscape = False
            Case "D" 'IND - Index
                If Cursor.Y >= 24 Then
                    Call TermScrollUp
                    Cursor.Y = 24
                Else
                    Cursor.Y = Cursor.Y + 1
                End If
                Call TermCheckCursor
                InEscape = False
            Case "E" 'NEL - Next Line
                If Cursor.Y >= 24 Then
                    Call TermScrollUp
                    Cursor.Y = 24
                Else
                    Cursor.Y = Cursor.Y + 1
                End If
                Cursor.X = 0
                Call TermCheckCursor
                InEscape = False
            Case "M" 'RI - Reverse Index
                If Cursor.Y <= 0 Then
                    'Call TermScrollDown
                    Cursor.Y = 0
                Else
                    Cursor.Y = Cursor.Y - 1
                End If
                Call TermCheckCursor
                InEscape = False
            Case "c" 'RIS - Reset to Initial State
                Call TermInitialize
                InEscape = False
            Case Else
                InEscape = False
        End Select
    Else 'If a Control Sequence Introducer exists...
        EscString = EscString & Chr$(ByteIn)
        Select Case Left(EscString, 1)
            Case "["
                Select Case Right(EscString, 1)
                    Case "D" 'CUB - Cursor Backward
                        If Len(EscString) = 2 Then 'if no parameter value
                            Cursor.X = Cursor.X - 1
                        Else
                            X = CInt(Mid(EscString, 2, Len(EscString) - 2))
                            If X = 0 Or X = 1 Then
                                Cursor.X = Cursor.X - 1
                            Else
                                Cursor.X = Cursor.X - X
                            End If
                        End If
                        Call TermCheckCursor
                        InEscape = False
                    Case "B" 'CUD - Cursor Down
                        If Len(EscString) = 2 Then 'if no parameter value
                            Cursor.Y = Cursor.Y + 1
                        Else
                            X = CInt(Mid(EscString, 2, Len(EscString) - 2))
                            If X = 0 Or X = 1 Then
                                Cursor.Y = Cursor.Y + 1
                            Else
                                Cursor.Y = Cursor.Y + X
                            End If
                        End If
                        Call TermCheckCursor
                        InEscape = False
                    Case "C" 'CUF - Cursor Forward
                        If Len(EscString) = 2 Then 'if no parameter value
                            Cursor.X = Cursor.X + 1
                        Else
                            X = CInt(Mid(EscString, 2, Len(EscString) - 2))
                            If X = 0 Or X = 1 Then
                                Cursor.X = Cursor.X + 1
                            Else
                                Cursor.X = Cursor.X + X
                            End If
                        End If
                        Call TermCheckCursor
                        InEscape = False
                    Case "H" 'CUP - Cursor Position
                        If Len(EscString) = 2 Then 'if no parameter values
                            Cursor.X = 0
                            Cursor.Y = 0
                        ElseIf Len(EscString) = 3 Then 'if no parameter values
                            If InStr(1, EscString, ";", vbTextCompare) <> 0 Then
                                Cursor.X = 0
                                Cursor.Y = 0
                            End If
                        Else
                            EscString = Mid(EscString, 2, Len(EscString) - 2)
                            Cursor.Y = CInt(TermSeparateParameters(EscString)) - 1
                            Cursor.X = CInt(EscString) - 1
                        End If
                        Call TermCheckCursor
                        InEscape = False
                    Case "A" 'CUU - Cursor Up
                        If Len(EscString) = 2 Then 'if no parameter value
                            Cursor.Y = Cursor.Y - 1
                        Else
                            X = CInt(Mid(EscString, 2, Len(EscString) - 2))
                            If X = 0 Or X = 1 Then
                                Cursor.Y = Cursor.Y - 1
                            Else
                                Cursor.Y = Cursor.Y - X
                            End If
                        End If
                        Call TermCheckCursor
                        InEscape = False
                    Case "c" 'DA - Device Attributes
                        If Len(EscString) = 2 Then 'if no parameter value
                            Call TermSendData(Chr(27) & "[?1;0c")
                        Else
                            X = CInt(Mid(EscString, 2, Len(EscString) - 2))
                            If X = 0 Then
                                Call TermSendData(Chr(27) & "[?1;0c")
                            End If
                        End If
                        InEscape = False
                    Case "n" 'DSR - Device Status Report
                        X = CInt(Mid(EscString, 2, Len(EscString) - 2))
                        Select Case X
                            Case 5 'Please report status
                                Call TermSendData(Chr(27) & "[0n")
                            Case 6 'Please report active position (CPR)
                                Call TermSendData(Chr(27) & "[" & Cursor.Y + 1 & ";" & Cursor.X + 1 & "R")
                        End Select
                        InEscape = False
                    Case "J" 'ED - Erase In Display
                        If Len(EscString) = 2 Then 'if no parameter value
                            For Z = Cursor.X To 79
                                TerminalChar(Cursor.Y, Z).Text = 32
                            Next
                            For Y = (Cursor.Y + 1) To 24
                                For Z = 0 To 79
                                    TerminalChar(Y, Z).Text = 32
                                Next
                            Next
                        Else
                            X = CInt(Mid(EscString, 2, Len(EscString) - 2))
                            Select Case X
                                Case 0 'Erase from active position to end of screen, inclusive
                                    For Z = Cursor.X To 79
                                        TerminalChar(Cursor.Y, Z).Text = 32
                                    Next
                                    For Y = (Cursor.Y + 1) To 24
                                        For Z = 0 To 79
                                            TerminalChar(Y, Z).Text = 32
                                        Next
                                    Next
                                Case 1 'Erase from start of screen to active position, inclusive
                                    For Z = 0 To Cursor.X
                                        TerminalChar(Cursor.Y, Z).Text = 32
                                    Next
                                    For Y = 0 To (Cursor.Y - 1)
                                        For Z = 0 To 79
                                            TerminalChar(Y, Z).Text = 32
                                        Next
                                    Next
                                Case 2 'Erase entire display, cursor does not move
                                    For Y = 0 To 24
                                        For Z = 0 To 79
                                            TerminalChar(Y, Z).Text = 32
                                        Next
                                    Next
                            End Select
                            Terminal.Refresh
                        End If
                        InEscape = False
                    Case "K" 'EL - Erase In Line
                        If Len(EscString) = 2 Then 'if no parameter value
                            For Z = Cursor.X To 79
                                TerminalChar(Cursor.Y, Z).Text = 32
                            Next
                        Else
                            X = CInt(Mid(EscString, 2, Len(EscString) - 2))
                            Select Case X
                                Case 0 'Erase from active position to end of line, inclusive
                                    For Z = Cursor.X To 79
                                        TerminalChar(Cursor.Y, Z).Text = 32
                                    Next
                                Case 1 'Erase from start of line to active position, inclusive
                                    For Z = 0 To Cursor.X
                                        TerminalChar(Cursor.Y, Z).Text = 32
                                    Next
                                Case 2 'Erase entire line
                                    For Z = 0 To 79
                                        TerminalChar(Cursor.Y, Z).Text = 32
                                    Next
                            End Select
                            Terminal.Refresh
                        End If
                        InEscape = False
                    Case "f" 'HVP - Horizontal and Vertical Position
                        If Len(EscString) = 2 Then 'if no parameter values
                            Cursor.X = 0
                            Cursor.Y = 0
                        ElseIf Len(EscString) = 3 Then 'if no parameter values
                            If InStr(1, EscString, ";", vbTextCompare) <> 0 Then
                                Cursor.X = 0
                                Cursor.Y = 0
                            End If
                        Else
                            EscString = Mid(EscString, 2, Len(EscString) - 2)
                            Cursor.Y = CInt(TermSeparateParameters(EscString)) - 1
                            Cursor.X = CInt(EscString) - 1
                        End If
                        Call TermCheckCursor
                        InEscape = False
                    Case "l" 'RM - Reset Mode
                        EscString = Mid(EscString, 2, Len(EscString) - 2)
                        Do
                            Select Case TermSeparateParameters(EscString)
                                Case "20" 'LNM = Reset
                                    TerminalAttribute.LNM = False
                                Case "7", "=7" 'Wordwrap = Reset
                                    TerminalAttribute.Wordwrap = False
                            End Select
                        Loop Until EscString = ""
                        InEscape = False
                    Case "m" 'SGR - Select Graphic Rendition
                        EscString = Mid(EscString, 2, Len(EscString) - 2)
                        Do
                            Select Case TermSeparateParameters(EscString)
                                Case "0" 'Normal Display
                                    DebugText = ""
                                    
                                    With TerminalAttribute
                                        .Blink = False
                                        .Bold = False
                                        .Invisible = False
                                        .Reverse = False
                                        .Underline = False
                                    End With
                                    
                                Case "1" 'Bold
                                    TerminalAttribute.Bold = True
                                Case "4" 'Underscore
                                    TerminalAttribute.Underline = True
                                Case "5" 'Blink
                                    TerminalAttribute.Blink = True
                                Case "7" 'Reverse Video
                                    TerminalAttribute.Reverse = True
                                Case "8" 'Invisible
                                    TerminalAttribute.Invisible = True
                                Case "30" 'Black Foreground
                                    TerminalAttribute.ForeColor = QBColor(8)
                                Case "31" 'Red Foreground
                                    TerminalAttribute.ForeColor = vbRed
                                Case "32" 'Green Foreground
                                    TerminalAttribute.ForeColor = vbGreen
                                Case "33" 'Yellow Foreground
                                    TerminalAttribute.ForeColor = vbYellow
                                Case "34" 'Blue Foreground
                                    TerminalAttribute.ForeColor = vbBlue
                                Case "35" 'Magenta Foreground
                                    TerminalAttribute.ForeColor = vbMagenta
                                Case "36" 'Cyan Foreground
                                    TerminalAttribute.ForeColor = vbCyan
                                Case "37" 'White Foreground
                                    TerminalAttribute.ForeColor = vbWhite
                                Case "40" 'Black Background
                                    TerminalAttribute.BackColor = vbBlack
                                Case "41" 'Red Background
                                    TerminalAttribute.BackColor = vbRed
                                Case "42" 'Green Background
                                    TerminalAttribute.BackColor = vbGreen
                                Case "43" 'Yellow Background
                                    TerminalAttribute.BackColor = vbYellow
                                Case "44" 'Blue Background
                                    TerminalAttribute.BackColor = vbBlue
                                Case "45" 'Magenta Background
                                    TerminalAttribute.BackColor = vbMagenta
                                Case "46" 'Cyan Background
                                    TerminalAttribute.BackColor = vbCyan
                                Case "47" 'White Background
                                    TerminalAttribute.BackColor = vbWhite
                            End Select
                        Loop Until EscString = ""
                        InEscape = False
                        
                    Case "h" 'SM - Set Mode
                        EscString = Mid(EscString, 2, Len(EscString) - 2)
                        Do
                            Select Case TermSeparateParameters(EscString)
                                Case "20" 'LNM =Set
                                    TerminalAttribute.LNM = True
                                Case "7", "=7" 'Wordwrap = Set
                                    TerminalAttribute.Wordwrap = True
                            End Select
                        Loop Until EscString = ""
                        InEscape = False
                    Case "g" 'TBC - Tabulation Clear
                        If Len(EscString) = 2 Then 'if no parameter value
                            For X = 0 To 19
                                If Tabs(X) = Cursor.X Then
                                    Tabs(X) = ""
                                End If
                            Next
                        Else
                            EscString = Mid(EscString, 2, Len(EscString) - 2)
                            Select Case EscString
                                Case 0 'Clear horizontal tab stop at the active position
                                    For X = 0 To 19
                                        If Tabs(X) = Cursor.X Then
                                            Tabs(X) = ""
                                        End If
                                    Next
                                Case 3 'Clear all horizontal tab stops
                                    For X = 0 To 19
                                        Tabs(X) = ""
                                    Next
                                    TabNum = 0
                            End Select
                        End If
                        InEscape = False
                    Case "s" 'ANSI - Save Cursor Position
                        With CursorSave
                            .X = Cursor.X
                            .Y = Cursor.Y
                            .Reverse = Cursor.Reverse
                            .Visible = Cursor.Visible
                        End With
                        InEscape = False
                    Case "u" 'ANSI - Restore Cursor Position
                        With Cursor
                            .X = CursorSave.X
                            .Y = CursorSave.Y
                            .Reverse = CursorSave.Reverse
                            .Visible = CursorSave.Visible
                        End With
                        Call TermCheckCursor
                        InEscape = False
                End Select
            Case "(", ")" 'SCS - Select Character Set
                InEscape = False
            Case "#"
                Select Case Right(EscString, 1)
                    Case "8" 'DECALN - Screen Alignment Display
                        Cursor.X = 0
                        Cursor.Y = 0
                        For Y = 0 To 24
                            'Cursor.Y = Y
                            For X = 0 To 79
                                TerminalChar(Y, X).Text = Asc("E")
                            Next
                        Next
                        Cursor.X = 0
                        Cursor.Y = 0
                        Terminal.Refresh
                        InEscape = False
                    Case Else
                        InEscape = False
                End Select
        End Select
    End If
End Sub

Private Sub TermCheckCursor()
    If Cursor.X < 0 Then
        Cursor.X = 0
    ElseIf Cursor.X > 79 Then
        If TerminalAttribute.Wordwrap = False Then
            Cursor.X = 79
        Else
            Cursor.X = 0
            Cursor.Y = Cursor.Y + 1
            Call TermCheckCursor
        End If
    End If
    If Cursor.Y < 0 Then
        Cursor.Y = 0
    ElseIf Cursor.Y > 24 Then
        Cursor.Y = 24
    End If
End Sub

Public Sub TermSendData(Data As String)
    Dim X As Integer
    Dim bData() As Byte
    
    For X = 1 To Len(Data)
        If frmMain.Winsock1.State = 7 Then '7 = connected
            frmMain.Winsock1.SendData CByte(Asc(Mid(Data, X, 1)))
        End If
    Next
End Sub

Private Function TermSeparateParameters(EscStringIn As String) As String
    Dim X As Integer

    X = InStr(EscStringIn, ";")
    If X = 0 Then
        TermSeparateParameters = EscStringIn
        EscString = ""
    Else
        TermSeparateParameters = Left$(EscStringIn, X - 1)
        EscString = Mid$(EscStringIn, X + 1)
    End If
End Function

Private Sub TermScrollUp()
    Dim X As Integer, Y As Integer
    
    For Y = 0 To 23
        For X = 0 To 79
            With TerminalChar(Y, X)
                .BackColor = TerminalChar(Y + 1, X).BackColor
                .Blink = TerminalChar(Y + 1, X).Blink
                .Bold = TerminalChar(Y + 1, X).Bold
                .ForeColor = TerminalChar(Y + 1, X).ForeColor
                .Invisible = TerminalChar(Y + 1, X).Invisible
                .Reverse = TerminalChar(Y + 1, X).Reverse
                .Text = TerminalChar(Y + 1, X).Text
                .Underline = TerminalChar(Y + 1, X).Underline
            End With
        Next
    Next
    
    
    For X = 0 To 79
        With TerminalChar(24, X)
            .BackColor = TerminalAttribute.BackColor
            .Blink = TerminalAttribute.Blink
            .Bold = TerminalAttribute.Bold
            .ForeColor = TerminalAttribute.ForeColor
            .Invisible = TerminalAttribute.Invisible
            .Reverse = TerminalAttribute.Reverse
            .Text = 32
            .Underline = TerminalAttribute.Underline
        End With
        
    Next
End Sub

Public Sub TermFreeMemory()
    Call DeleteObject(Color.Black)
    Call DeleteObject(Color.Blue)
    Call DeleteObject(Color.Cyan)
    Call DeleteObject(Color.Gray)
    Call DeleteObject(Color.Green)
    Call DeleteObject(Color.Magenta)
    Call DeleteObject(Color.Red)
    Call DeleteObject(Color.White)
    Call DeleteObject(Color.Yellow)
End Sub

Public Sub TermRefresh()
    Dim Y As Integer, X As Integer
    Dim tempBrush As Long
    
    Terminal.Cls
    
    If Cursor.Visible = True Then
        BlinkSentry = BlinkSentry - 1
        If BlinkSentry <= 0 Then
            BlinkSentry = 1
            If Cursor.Reverse = True Then
                Cursor.Reverse = False
            Else
                Cursor.Reverse = True
                Call BitBlt(Terminal.hdc, Cursor.X * Character.Width, Cursor.Y * Character.Height, Character.Width, Character.Height, Terminal.hdc, Cursor.X * Character.Width, Cursor.Y * Character.Height, DSTINVERT)
            End If
        End If
    End If
    
    For Y = 0 To 24
        For X = 0 To 79
            Call SetTextColor(Terminal.hdc, TerminalChar(Y, X).ForeColor)
            
            If TerminalChar(Y, X).Blink = True Then
                If Cursor.Reverse = False Then
                    Call TextOut(Terminal.hdc, X * Character.Width, Y * Character.Height, Chr$(32), 1)
                Else
                    Call TextOut(Terminal.hdc, X * Character.Width, Y * Character.Height, Chr$(TerminalChar(Y, X).Text), 1)
                End If
            Else
                Call TextOut(Terminal.hdc, X * Character.Width, Y * Character.Height, Chr$(TerminalChar(Y, X).Text), 1)
            End If
        Next
    Next
    
    Terminal.Refresh
End Sub
