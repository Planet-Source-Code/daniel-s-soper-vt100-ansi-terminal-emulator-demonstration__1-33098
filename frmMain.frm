VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dan's VT100 ANSI Emulator..."
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   297
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   672
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4560
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer timCursor 
      Interval        =   500
      Left            =   4560
      Top             =   3600
   End
   Begin VB.Frame frameConnection 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Connection Options..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1140
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   4455
      Begin VB.CommandButton cmdConnect 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   6
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtPort 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Text            =   "23"
         Top             =   720
         Width           =   540
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Text            =   "bbs.frogland.net"
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "Remote Port:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   760
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "Remote Address:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   400
         Width           =   1335
      End
   End
   Begin VB.PictureBox picTerm 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   0
      ScaleHeight     =   191
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   311
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Private Sub cmdConnect_Click() 'Connect / Disconnect button
    If cmdConnect.Caption = "Connect" Then
        cmdConnect.Caption = "Disconnect"
        txtAddress.Enabled = False
        txtPort.Enabled = False
        Call TermInitialize
        Winsock1.Close
        Winsock1.Connect txtAddress.Text, txtPort.Text
        timCursor.Enabled = True
        Terminal.SetFocus
    Else
        Call Disconnect
    End If
End Sub

Public Sub Disconnect() 'Disconnect from remote host
    cmdConnect.Caption = "Connect"
    Winsock1.Close
    txtAddress.Enabled = True
    txtPort.Enabled = True
    timCursor.Enabled = False
    Call TermFreeMemory
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Winsock1.State = 7 Then '7 = connected
        Winsock1.SendData Chr$(KeyAscii)
    End If
End Sub

Private Sub Form_Paint()
    Terminal.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Disconnect
End Sub

Private Sub timCursor_Timer()
    If ReceivingData = False Then
        Call TermRefresh
    End If
End Sub

Private Sub Winsock1_Close()
    Call Disconnect
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim DataIn() As Byte
    Dim DataLength As Long
    Dim X As Integer
    
    ReceivingData = True
    
    DataLength = Winsock1.BytesReceived
    Winsock1.GetData DataIn, vbArray + vbByte
    For X = 0 To DataLength - 1
        Call TermProcessInput(DataIn(X))
    Next
    
    ReceivingData = False
End Sub
