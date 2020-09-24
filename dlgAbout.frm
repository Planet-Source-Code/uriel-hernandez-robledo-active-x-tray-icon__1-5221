VERSION 5.00
Begin VB.Form dlgAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acerca de ..."
   ClientHeight    =   2565
   ClientLeft      =   2340
   ClientTop       =   1890
   ClientWidth     =   4125
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1763.359
   ScaleMode       =   0  'User
   ScaleWidth      =   3865.687
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2520
      TabIndex        =   0
      Top             =   2040
      Width           =   1260
   End
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   1680
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Uriel Hernández Robledo Ing. Sistemas. 3er. Semestre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1575
      TabIndex        =   3
      Top             =   1080
      Width           =   2040
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "System Tray"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0.0"
      Height          =   225
      Left            =   1575
      TabIndex        =   1
      Top             =   840
      Width           =   2205
   End
   Begin VB.Line Line2 
      X1              =   224.913
      X2              =   3486.147
      Y1              =   1154.949
      Y2              =   1154.949
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   0
      Picture         =   "dlgAbout.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   915
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   1
      X1              =   224.913
      X2              =   3486.147
      Y1              =   1154.949
      Y2              =   1154.949
   End
End
Attribute VB_Name = "dlgAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DeltaX, DeltaY As Integer   ' Declare variables.
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Timer1_Timer()
    Image1.Move Image1.Left + DeltaX, Image1.Top + DeltaY
    If Image1.Left < ScaleLeft Then DeltaX = 100
    If Image1.Left + Image1.Width > ScaleWidth + ScaleLeft Then
        DeltaX = -100
    End If
    If Image1.Top < ScaleTop Then DeltaY = 100
    If Image1.Top + Image1.Height > ScaleHeight + ScaleTop Then
        DeltaY = -100
    End If
End Sub

Private Sub Form_Load()

Timer1.Interval = 100  ' Set Interval.
    DeltaX = 100    ' Initialize variables.
    DeltaY = 100
End Sub
