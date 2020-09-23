VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Image Menu "
   ClientHeight    =   4800
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10080
   FillColor       =   &H00FF8080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.bioscopedvd.com"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   2760
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Right click to see popupmenu"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   2115
   End
   Begin VB.Image menu 
      Height          =   4680
      Left            =   4920
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   3300
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Option"
      Begin VB.Menu mnuFile 
         Caption         =   "&File"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuView 
         Caption         =   "&View"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuPro 
         Caption         =   "&Project"
      End
      Begin VB.Menu mnuFormate 
         Caption         =   "F&ormate"
      End
      Begin VB.Menu mnuDeb 
         Caption         =   "&Debug"
      End
      Begin VB.Menu mnuRun 
         Caption         =   "&Run"
      End
      Begin VB.Menu mnuQue 
         Caption         =   "Q&uery"
      End
      Begin VB.Menu mnuDia 
         Caption         =   "D&iagram"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Tools"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add-Ins"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuWin 
         Caption         =   "&Window"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Width = 3000
Me.Left = (Screen.Width - Me.Width) / 2 'center the form
Me.Top = (Screen.Height - Me.Height) / 2 'center the form
SetCustomMenus
End Sub
Private Sub SetCustomMenus()
    startODMenus Me, True
    With CustomMenu
        .Texture = False
        Set .Picture = menu.Picture
        .UseCustomFonts = True
'        .FontUnderline = False
        .FontName = "arial" '"Comic Sans MS" '
'        .FontItalic = False
'        .FontStrikeOut = False
         .PosX = 30
    End With
    MenuMode = VBLook 'XP look menu
    With CustomColor
        .ForeColor = RGB(16, 0, 16) 'Magic color
        .DefTextColor = vbRed  ' Red
        .HilightColor = RGB(182, 189, 210)
        .NormalColor = RGB(186, 186, 204)
        .BackColor = RGB(58, 110, 165)
        .SelectedTextColor = RGB(0, 0, 255)
        .MenuTextColor = vbBlack
        .BorderColor = RGB(240, 72, 72)
        .RECTColor = RGB(10, 36, 106)
    End With
    With CustomMenu
        .Icon.Add Array(101), "&File"
            End With
  End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu mnuOption 'call popupmenu
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu mnuOption 'call popupmenu
End Sub

Private Sub mnuAdd_Click()
MsgBox "Write code for this menu"

End Sub

Private Sub mnuDeb_Click()
MsgBox "Write code for this menu"

End Sub

Private Sub mnuDia_Click()
MsgBox "Write code for this menu"

End Sub

Private Sub mnuEdit_Click()
MsgBox "Write code for this menu"
End Sub

Private Sub mnuExit_Click()
stopODMenus Me
Unload Me
End Sub

Private Sub mnuFormate_Click()
MsgBox "Write code for this menu"

End Sub

Private Sub mnuPro_Click()
MsgBox "Write code for this menu"

End Sub

Private Sub mnuQue_Click()
MsgBox "Write code for this menu"

End Sub

Private Sub mnuRun_Click()
MsgBox "Write code for this menu"

End Sub

Private Sub mnuTool_Click()
MsgBox "Write code for this menu"

End Sub

Private Sub mnuView_Click()
MsgBox "Write code for this menu"

End Sub

Private Sub mnuWin_Click()
MsgBox "Write code for this menu"

End Sub
