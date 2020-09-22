VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "WINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anonymous Emailer - C0ded by JadaCyrus"
   ClientHeight    =   3630
   ClientLeft      =   8850
   ClientTop       =   1680
   ClientWidth     =   3525
   Icon            =   "anonmail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   3525
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   360
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00404040&
      Height          =   615
      Left            =   2400
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   2520
      Picture         =   "anonmail.frx":08CA
      ToolTipText     =   "C0ded in Visual Basic By JadaCyrus: cell.s5.com"
      Top             =   360
      Width           =   525
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Left            =   600
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "  Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   3000
      Width           =   855
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Left            =   1920
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "       Send"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mail Server:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Body:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rcpt To:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   3375
      Left            =   120
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub Pause(duration)
'Pause for the specified duration
'Duration is in seconds
Dim Current As Long
Current = Timer
Do Until Timer - Current >= duration
    DoEvents
Loop
End Sub
Private Sub Form_Load()
Label2.BackColor = RGB(192, 192, 192)
Label3.BackColor = RGB(192, 192, 192)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Winsock1.Close

Label2.BackColor = RGB(32, 32, 32)
Label2.ForeColor = vbWhite
Form2.Show
Winsock1.RemoteHost = Text5.Text
Winsock1.RemotePort = 25
Winsock1.Connect

End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = RGB(192, 192, 192)
Label2.ForeColor = vbBlack

End Sub

Private Sub Label3_Click()
Winsock1.Close
Form2.Text1.Text = Form2.Text1.Text + vbNewLine + "Connection Closed"
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackColor = RGB(32, 32, 32)
Label3.ForeColor = vbWhite
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackColor = RGB(192, 192, 192)
Label3.ForeColor = vbBlack

End Sub

Private Sub Winsock1_Close()
Winsock1.Close
Form2.Text1.Text = Form2.Text1.Text + vbNewLine + "Connection Closed"

End Sub

Private Sub Winsock1_Connect()
Form2.Text1.Text = Form2.Text1.Text + vbNewLine + "Connected.."
Pause 0.9
Pause 0.9
Winsock1.SendData "HELO " & Winsock1.LocalIP

Pause 0.3
Pause 0.9
Pause 3

Pause 0.9

Winsock1.SendData "MAIL FROM:" & Text1.Text & vbCrLf
Winsock1.SendData "MAIL FROM:" & Text1.Text & vbCrLf
Winsock1.SendData "MAIL FROM:" & Text1.Text & vbCrLf
Form2.Text1.Text = Form2.Text1.Text + vbNewLine + "MAIL FROM:" & Text1.Text
Pause 0.5
Pause 0.9
Pause 4

Pause 0.9

Winsock1.SendData "RCPT TO:" & Text2.Text & vbCrLf
Form2.Text1.Text = Form2.Text1.Text + vbNewLine + "RCPT TO:" & Text2.Text
Pause 0.5
Pause 0.9

Form2.Text1.Text = Form2.Text1.Text + vbNewLine + "DATA"
Pause 0.9
Winsock1.SendData "DATA" & vbCrLf

Pause 0.3
Winsock1.SendData "SUBJECT:" & Text3.Text & vbCrLf
Pause 0.9

Pause 0.9

Pause 0.3
Winsock1.SendData Text4.Text & vbCrLf
Pause 0.9
Winsock1.SendData vbCrLf & "." & vbCrLf
Pause 3
Pause 3
Pause 3
Winsock1.Close
Form2.Text1.Text = Form2.Text1.Text + vbNewLine + "Connection Closed"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim data As String
Winsock1.GetData data, vbString
Form2.Text1.Text = Form2.Text1.Text + vbNewLine + data

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1.Close
Form2.Text1.Text = Form2.Text1.Text + vbNewLine + "ERROR: " & Number & ": " & Description

End Sub
