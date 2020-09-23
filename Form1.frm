VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Mapa"
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   705
      Top             =   1665
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Coded by : Safo
' Email : safo@zoznam.sk
' Program : Test LCD

Dim Click As Boolean
Private Step As Integer
Dim Color As Long
Dim i As Integer

Private Sub Form_Click()

If Click = True Then

  Select Case Step
  Case 1
    Click = False
    Timer1.Interval = 10
    Timer1.Enabled = True
    i = 0
  Case 2
    Click = False
    Timer1.Interval = 10
    Timer1.Enabled = True
    i = 0
  Case 3
    Click = False
    Timer1.Interval = 10
    Timer1.Enabled = True
    i = 0
  Case 4
    Click = False
    Timer1.Interval = 10
    Timer1.Enabled = True
    i = 0
  Case 5
    Unload Me
  End Select
  
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()

ShowCursor False
Step = 1
Click = True
Me.Left = 1
Me.Top = 1
Me.Width = Screen.Width
Me.Height = Screen.Height

End Sub

Private Sub Form_Unload(Cancel As Integer)
  ShowCursor True
End Sub

Private Sub Timer1_Timer()

Select Case Step

Case 1
  
  If i >= 254 Then
    Timer1.Enabled = False
    Step = 2
    Click = True
  End If
  
  Me.BackColor = RGB(255, 255 - i, 255 - i)

Case 2

  If i >= 254 Then
    Timer1.Enabled = False
    Step = 3
    Click = True
  End If
  
  Me.BackColor = RGB(255 - i, i, 0)
  
Case 3
  
  If i >= 254 Then
    Timer1.Enabled = False
    Step = 4
    Click = True
  End If
  
  Me.BackColor = RGB(0, 255 - i, i)
  
Case 4
  
  If i >= 254 Then
    Timer1.Enabled = False
    Step = 5
    Click = True
  End If
  
  Me.BackColor = RGB(0, 0, 255 - i)
  
End Select

i = i + 2

End Sub
