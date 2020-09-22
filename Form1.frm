VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "HD44780 PARALLEL LCD (0.3.0)"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   12285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Write String"
      Height          =   510
      Left            =   -90
      TabIndex        =   4
      Top             =   630
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Left            =   990
      Top             =   3195
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show time"
      Height          =   510
      Left            =   45
      TabIndex        =   3
      Top             =   2925
      Width           =   1185
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cursor Off"
      Height          =   600
      Left            =   1125
      TabIndex        =   2
      Top             =   0
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cursor On"
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Write Screen"
      Height          =   510
      Left            =   1125
      TabIndex        =   0
      Top             =   630
      Width           =   1140
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim data(0 To 3) As String
    data(0) = "  This is my first  "
    data(1) = "       screen       "
    data(2) = "Time:               "
    data(3) = "left  center   right"
    
    
    LCD_Write20x4_Screen data
End Sub

Private Sub Command2_Click()
    LCD_CursorOff
End Sub

Private Sub Command3_Click()
    LCD_CursorOn
End Sub

Private Sub Command4_Click()
    Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Command5_Click()
    LCD_WriteString "This is my string", 2, 1
End Sub

Private Sub Form_Load()
    LCD_Init
End Sub

Private Sub Timer1_Timer()
    LCD_WriteString Format(Now, "hh:mm:ss"), 3, 12
    LCD_WriteString Format(Now, "hh:mm:ss"), 1, 12
End Sub
