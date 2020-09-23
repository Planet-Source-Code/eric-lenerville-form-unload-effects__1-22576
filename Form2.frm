VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00404040&
   Caption         =   "Effects"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Unload Me"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Resize()
' I center the command button to unload the form so it
'looks neat
Me.Command1.Top = (Me.Height - Me.Command1.Height) / 2
Me.Command1.Left = (Me.Width - Me.Command1.Width) / 2


'here we only center the form when we use certain styles
'to unload it
If Form1.Combo2.Text = "Expand Out 1" Then
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Else
If Form1.Combo2.Text = "Expand Out 2" Then
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Else
If Form1.Combo2.Text = "Expand Out 3" Then
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Else
If Form1.Combo2.Text = "Shrink 1" Then
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Else
If Form1.Combo2.Text = "Shrink 2" Then
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Else
If Form1.Combo2.Text = "Shrink 3" Then
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Else

End If
End If
End If
End If
End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'we dim num as an integer so later we can use num = num + 1
'this will make the following code easyer
Dim num As Integer
num = 1

'this will expand the form outwards in all directions
If Form1.Combo2.Text = "Expand Out 1" Then
For i = 0 To Screen.Height And Screen.Width
Me.Height = Me.Height + num
num = num + 1
Me.Width = Me.Width + num
num = num + 1
DoEvents
Next i
End If

'this will expand the form horizontal then vertical
If Form1.Combo2.Text = "Expand Out 2" Then
For i = 0 To Screen.Width
Me.Width = Me.Width + num
num = num + 1
DoEvents
Next i

'we must set num = 1 again or else the effect will end
num = 1

For i = 0 To Screen.Height
Me.Height = Me.Height + num
num = num + 1
DoEvents
Next i
End If

'this will expand the form Vertical then Horizontal
If Form1.Combo2.Text = "Expand Out 3" Then
For i = 0 To Screen.Height
Me.Height = Me.Height + num
num = num + 1
DoEvents
Next i

'we must set num = 1 again or else the effect will end
num = 1

For i = 0 To Screen.Width
Me.Width = Me.Width + num
num = num + 1
DoEvents
Next i
End If

'moves form from center to upper right
If Form1.Combo2.Text = "Diagonal UpperRight" Then
For i = 0 To Screen.Height And Screen.Width
Me.Left = Me.Left + num
Me.Top = Me.Top - num
' to make this go slower remove the following code
'num = num + 1
num = num + 1
DoEvents
Next i
End If

'moves form from center to upper left
If Form1.Combo2.Text = "Diagonal UpperLeft" Then
For i = 0 To Screen.Height And Screen.Width
Me.Left = Me.Left - num
Me.Top = Me.Top - num
' to make this go slower remove the following code
'num = num + 1
num = num + 1
DoEvents
Next i
End If

'moves form from center to lower right
If Form1.Combo2.Text = "Diagonal LowerRight" Then
For i = 0 To Screen.Height And Screen.Width
Me.Left = Me.Left + num
Me.Top = Me.Top + num
' to make this go slower remove the following code
'num = num + 1
num = num + 1
DoEvents
Next i
End If

'moves form from center to lower left
If Form1.Combo2.Text = "Diagonal LowerLeft" Then
For i = 0 To Screen.Height And Screen.Width
Me.Left = Me.Left - num
Me.Top = Me.Top + num
' to make this go slower remove the following code
'num = num + 1
num = num + 1
DoEvents
Next i
End If

'will shink the entire form
If Form1.Combo2.Text = "Shrink 1" Then
On Error GoTo done

For i = 0 To Me.Height And Me.Width
Me.Height = Me.Height - num
Me.Width = Me.Width - num
num = num + 1
DoEvents
Next i
done:
Exit Sub
End If

'we shrink horizontal then vertical
If Form1.Combo2.Text = "Shrink 2" Then
On Error GoTo crap

For i = 0 To Me.Height
Me.Height = Me.Height - num
DoEvents
Next i

For i = 0 To Me.Width
Me.Width = Me.Width - num
num = num + 1
DoEvents
Next i

crap:
Exit Sub
End If

'will shrink vertical then horizontal
If Form1.Combo2.Text = "Shrink 3" Then
On Error GoTo hell

For i = 0 To Me.Width
Me.Width = Me.Width - num
DoEvents
Next i

For i = 0 To Me.Height
Me.Height = Me.Height - num
num = num + 1
DoEvents
Next i

hell:
Exit Sub
End If


'will roll bottom up then squeez the titlebar shut
'then finaly will disapear
If Form1.Combo2.Text = "Roll Up" Then
On Error GoTo die
For i = 0 To Me.Height
Me.Height = Me.Height - num
DoEvents
Next i
For i = 0 To Me.Width
Me.Width = Me.Width - num
DoEvents
Next i

die:
Unload Me
End If

'moves form from upper left to the right
'then from the upper right straight down
If Form1.Combo2.Text = "Left-Right Down" Then
On Error GoTo fuck
Me.Top = 0
Me.Left = 0

For i = 0 To Screen.Width - Me.Width
Me.Left = Me.Left + num
DoEvents
Next i


For i = 0 To Screen.Height - Me.Height
Me.Top = Me.Top + num
num = num + 1
DoEvents
Next i

fuck:
Unload Me
End If

'moves the form from the upper left straight down
'then moves it to the lower right
If Form1.Combo2.Text = "Down Left-Right" Then
On Error GoTo upyours
Me.Top = 0
Me.Left = 0

For i = 0 To Screen.Height - Me.Height
Me.Top = Me.Top + num
DoEvents
Next i

For i = 0 To Screen.Width - Me.Width
Me.Left = Me.Left + num
num = num + 1
DoEvents
Next i

upyours:
Unload Me
End If

'slides form off screen up
If Form1.Combo2.Text = "Up" Then

For i = 0 To Screen.Height - Me.Height
Me.Top = Me.Top - num
' to make this go faster just add the following code
'num = num + 1
DoEvents
Next i
End If

'slides form off screen down
If Form1.Combo2.Text = "Down" Then

For i = 0 To Screen.Height - Me.Height
Me.Top = Me.Top + num
' to make this go faster just add the following code
'num = num + 1
DoEvents
Next i
End If

'slides form off screen left
If Form1.Combo2.Text = "Left" Then

For i = 0 To Screen.Width - Me.Width
Me.Left = Me.Left - num
' to make this go faster just add the following code
'num = num + 1
DoEvents
Next i
End If

'slides form off screen right
If Form1.Combo2.Text = "Right" Then
For i = 0 To Screen.Width - Me.Width
Me.Left = Me.Left + num
' to make this go faster just add the following code
'num = num + 1
DoEvents
Next i
End If

'slides form off screen left then bounces back and moves
'right.. after it reaches the center again it disapears
If Form1.Combo2.Text = "Bounce Left" Then
For i = 0 To Screen.Width - Me.Width
Me.Left = Me.Left - num
DoEvents
Next i

For i = 0 To Screen.Width - Me.Width
Me.Left = Me.Left + num
DoEvents
Next i
End If

'moves form off screen right then bounces back
'and moves left till it reaches the center again
If Form1.Combo2.Text = "Bounce Right" Then
For i = 0 To Screen.Width - Me.Width
Me.Left = Me.Left + num
DoEvents
Next i

For i = 0 To Screen.Width - Me.Width
Me.Left = Me.Left - num
DoEvents
Next i
End If

'moves off screen up then bounces back and moves down
' till it reaches the center
If Form1.Combo2.Text = "Bounce Up" Then
For i = 0 To Screen.Height - Me.Height
Me.Top = Me.Top - num
DoEvents
Next i

For i = 0 To Screen.Height - Me.Height
Me.Top = Me.Top + num
DoEvents
Next i
End If

'moves off screen down then bounces back and moves up
' till it reaches the center
If Form1.Combo2.Text = "Bounce Down" Then
For i = 0 To Screen.Height - Me.Height
Me.Top = Me.Top + num
DoEvents
Next i

For i = 0 To Screen.Height - Me.Height
Me.Top = Me.Top - num
DoEvents
Next i
End If

End Sub

