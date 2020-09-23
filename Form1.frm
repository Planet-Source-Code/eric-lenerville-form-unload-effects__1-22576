VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "Form Unload Effects  v 0.1"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Go"
      Height          =   255
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   5175
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Caption         =   "Form Unload Effect:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Form2.Visible = True
End Sub

Private Sub Command2_Click()


End Sub

Private Sub Form_Load()


Combo2.AddItem "Shrink 1"
Combo2.AddItem "Shrink 2"
Combo2.AddItem "Shrink 3"
Combo2.AddItem "Roll Up"
Combo2.AddItem "Expand Out 1"
Combo2.AddItem "Expand Out 2"
Combo2.AddItem "Expand Out 3"
Combo2.AddItem "Up"
Combo2.AddItem "Down"
Combo2.AddItem "Left"
Combo2.AddItem "Right"
Combo2.AddItem "Diagonal UpperLeft"
Combo2.AddItem "Diagonal LowerLeft"
Combo2.AddItem "Diagonal UpperRight"
Combo2.AddItem "Diagonal LowerRight"
Combo2.AddItem "Bounce Up"
Combo2.AddItem "Bounce Down"
Combo2.AddItem "Bounce Left"
Combo2.AddItem "Bounce Right"
Combo2.AddItem "Left-Right Down"
Combo2.AddItem "Down Left-Right"
Me.Top = 0
Me.Left = 0
End Sub
