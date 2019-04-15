VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7545
   ClientLeft      =   5700
   ClientTop       =   2355
   ClientWidth     =   6330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Football.frx":0000
   ScaleHeight     =   7696.811
   ScaleMode       =   0  'User
   ScaleWidth      =   6938.055
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   240
      Top             =   5520
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   7080
      Width           =   6495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   735
      Left            =   2640
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   690
      Left            =   2640
      Picture         =   "Football.frx":8E14
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   690
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GAME OVER"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Menu new 
      Caption         =   "NEW GAME"
   End
   Begin VB.Menu exit 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gravity, yspeed, xspeed, score As Integer

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
gravity = 15
xspeed = 0
yspeed = 0
score = 0
Timer1.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
score = score + 1
If X < Image1.Left And Y >= (Image1.Top + Image1.Height) Then
xspeed = 100
yspeed = -150
ElseIf X > (Image1.Left + Image1.Width) And Y >= (Image1.Top + Image1.Height) Then
xspeed = -100
yspeed = -150
ElseIf (X >= Image1.Left And X <= (Image1.Left + Image1.Width)) And (Y >= (Image1.Top + Image1.Height)) Then
xspeed = 0
yspeed = -150
End If
End If
End Sub

Private Sub new_Click()
Unload Me
Me.Show
End Sub

Private Sub Timer1_Timer()
Label2.Caption = score
yspeed = yspeed + gravity
Image1.Top = Image1.Top + yspeed
Image1.Left = Image1.Left + xspeed
If collide(Image1, Shape1) = True Then
Timer1.Enabled = False
Label1.Visible = True
End If
If Image1.Left <= 10 Then
score = score - 1
xspeed = 80
yspeed = 120
ElseIf Image1.Left >= 6100 Then
score = score - 1
xspeed = -80
yspeed = 120
End If
End Sub
Private Function collide(ByVal ball As Object, grnd As Object) As Boolean
If ball.Top + ball.Height >= grnd.Top Then
collide = True
Else
collide = False
End If
End Function

