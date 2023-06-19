VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Clock2"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12375
   Icon            =   "Clock2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Clock2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   1605
      Left            =   -3975
      TabIndex        =   0
      Top             =   2760
      Width           =   20415
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Timer1.Interval = 1000
Label1.AutoSize = True
End Sub

Private Sub Form_Resize()
Dim NewFont As Single
Label1.Left = (Form1.ScaleWidth / 2) - (Label1.Width / 2)
Label1.Top = (Form1.ScaleHeight / 2) - (Label1.Height / 2)
NewFont = 5 * (Me.ScaleWidth * 0.001)
Label1.FontSize = NewFont
End Sub

Private Sub Label1_Click()
Clipboard.SetText (Label1.Caption)
End Sub

Private Sub Timer1_Timer()
Label1 = Now
Label1.Caption = Format(Now, "hh:mm:ss" & vbCrLf & "mm/dd/yyyy")
End Sub
