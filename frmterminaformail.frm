VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Terminal"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3855
   Icon            =   "frmterminaformail.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2535
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
Text1.SelStart = Len(Text1)

End Sub
