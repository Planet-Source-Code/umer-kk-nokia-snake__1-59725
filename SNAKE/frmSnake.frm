VERSION 5.00
Begin VB.Form frmSnake 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Snake White Python"
   ClientHeight    =   7485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8280
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmSnake"
   ScaleHeight     =   499
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   552
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timGameLoop 
      Interval        =   250
      Left            =   4080
      Top             =   360
   End
   Begin VB.PictureBox picInfo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   0
      Picture         =   "frmSnake.frx":0000
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   260
      TabIndex        =   3
      Top             =   4200
      Width           =   3900
      Begin Snake_By_Umer.Button buttNewGame 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   390
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
      End
      Begin Snake_By_Umer.Button buttExit 
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   390
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         lbl             =   "Exit"
      End
      Begin Snake_By_Umer.Button buttOptions 
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   390
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         lbl             =   "Options"
      End
      Begin VB.Label lblScore 
         BackStyle       =   0  'Transparent
         Caption         =   "Score comes Here"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   5
         Top             =   120
         Width           =   3615
      End
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      Picture         =   "frmSnake.frx":989A
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   260
      TabIndex        =   2
      Top             =   0
      Width           =   3900
      Begin VB.Label lblBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Snake"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   60
         Width           =   3615
      End
   End
   Begin VB.PictureBox picGameField 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3900
      Left            =   0
      Picture         =   "frmSnake.frx":D5CC
      ScaleHeight     =   260
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   260
      TabIndex        =   0
      Top             =   300
      Width           =   3900
      Begin VB.PictureBox picFood 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   150
         Left            =   75
         Picture         =   "frmSnake.frx":3EE22
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   8
         Top             =   225
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.PictureBox picSnake 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   150
         Index           =   0
         Left            =   75
         Picture         =   "frmSnake.frx":3EFA4
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   1
         Top             =   75
         Visible         =   0   'False
         Width           =   150
      End
   End
End
Attribute VB_Name = "frmSnake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#########################################
'#         Developed By Umer KK          #
'#        Umerkhan_63@Hotmail.com        #
'#########################################

Option Explicit
Dim intScore As Integer
Dim intSnakeSize As Integer
Dim lngX As Long
Dim lngY As Long
Dim strDirection As String
Dim blnKeyAcces As Boolean

Const dimX1 As Byte = 5
Const dimX2 As Byte = 245
Const dimY1 As Byte = 5
Const dimY2 As Byte = 245

Private Sub buttExit_Click()
 'Exit the program
  Unload Me
End Sub

Private Sub buttNewGame_Click()
 'Start a New Game
 'Calls all the sequences to start a new game...
  Call NewGame
 'Sets focus to player's screen
  picGameField.SetFocus
End Sub

Sub NewGame()
 'Starts a New Game
  'declarations
   Dim t As Integer 'teller
  
  'Unload previous snake
   If Not intSnakeSize = 0 Then
    For t = intSnakeSize To 1 Step -1
     Unload picSnake(t)
    Next t
    intSnakeSize = 0
   End If
  
  'Place the head of the snake, and make a body
   picSnake(0).Move dimX1 + 40, dimY1
   picSnake(0).Visible = True
  'Make a body
   For t = 1 To 4
    Load picSnake(t)
    picSnake(t).Move picSnake(t - 1).Left - 10, picSnake(0).Top
    picSnake(t).Visible = True
    intSnakeSize = intSnakeSize + 1
   Next t
  'Place Food
   picFood.Visible = True
   Call PlaceFood
  'The snake moves right!
   strDirection = "right"
  'Start the GameLoop
   timGameLoop.Enabled = True
End Sub
'#########################################
'#         Developed By Umer KK          #
'#        Umerkhan_63@Hotmail.com        #
'#########################################
Sub PlaceFood()
 'This sub will handle the placing of the food
  'declarations
   'none

  'Calculate where it should be placed (randomly)
   Do
    Call CalculateFood
   Loop Until CalculateFood = True

  'Place the food
   picFood.Move lngX, lngY
End Sub

Function CalculateFood() As Boolean
 'This function will calculate a place for the food
 'When you cannot divide it by 10 or 5 it will return a false, else a true
  'declarations
   Dim temp As String
   Dim t As Integer 'teller
  
  'calc
   lngX = Int((dimX2 - dimX1 + 1) * Rnd + dimX1)
   lngY = Int((dimY2 - dimY1 + 1) * Rnd + dimY1)
  'check
   temp = lngX
   If Not Right(temp, 1) = 5 Then
    CalculateFood = False
   Else
    temp = lngY
    If Not Right(temp, 1) = 5 Then
     CalculateFood = False
    Else
     'Now we're going to check wheter the food is'nt placed on the snake
      For t = 0 To intSnakeSize
       If picSnake(t).Left = lngX And picSnake(t).Top = lngY Then
        CalculateFood = False
       Else
        CalculateFood = True
       End If
      Next t
    End If
   End If
End Function

Private Sub buttOptions_click()
 frmOptions.Show
 Me.Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 'This sub Checks which key is pressed
 
  If blnKeyAcces = False Then Exit Sub
 
  If KeyCode = vbKeyLeft And Not strDirection = "right" Then strDirection = "left"
  If KeyCode = vbKeyRight And Not strDirection = "left" Then strDirection = "right"
  If KeyCode = vbKeyUp And Not strDirection = "down" Then strDirection = "up"
  If KeyCode = vbKeyDown And Not strDirection = "up" Then strDirection = "down"
  
  blnKeyAcces = False
End Sub

Private Sub Form_Load()
 'Sets the CaptionTitle
  lblBar.Caption = "Snake v" & App.Major & "." & App.Minor & "." & "."
  Me.Caption = lblBar.Caption
 'Resize the Form
  Me.Height = 4950
  Me.Width = 3900
 'Name Buttons
  buttExit.Tekst = "Exit"
  buttNewGame.Tekst = "New Game"
  buttOptions.Tekst = "Options"
 'Score is 0
  intScore = 0
  lblScore = "Score: " & intScore
 'Going through walls is enabled
  blnGoThroughWalls = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Please Vote Me On mY This Submission, I Worked Very Hard IN Making This Game, Please Vote Me.", vbInformation, "Snake"
End Sub

Private Sub lblBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'This Let's the Form Move, like a real titlebar
  Call MouseMove(Me)
End Sub

Private Sub picBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'This Let's the Form Move, like a real titlebar
  Call MouseMove(Me)
End Sub

Private Sub timGameLoop_Timer()
 'This Timer Loops the Game
  'Makes the snake move..
   Call MoveSnake
  'Checks any Collision
   Call CheckCollision
  
  'Grant acces to keypress
   blnKeyAcces = True
End Sub

Sub MoveSnake()
 'This Sub moves the snake and it's body
  'declarations
   Dim t As Integer 'teller
   
  'Move the Body
   For t = intSnakeSize To 1 Step -1
    picSnake(t).Move picSnake(t - 1).Left, picSnake(t - 1).Top
   Next t
  'Move the Head
   Select Case strDirection
    Case "left"
     picSnake(0).Left = picSnake(0).Left - 10
    Case "right"
     picSnake(0).Left = picSnake(0).Left + 10
    Case "up"
     picSnake(0).Top = picSnake(0).Top - 10
    Case "down"
     picSnake(0).Top = picSnake(0).Top + 10
   End Select
End Sub

Sub CheckCollision()
 'This sub checks whether the snake has hit something
  'declarations
   Dim t As Integer 'teller

  'Check for hitting walls
   If picSnake(0).Left < dimX1 Or picSnake(0).Left > dimX2 Or _
      picSnake(0).Top < dimY1 Or picSnake(0).Top > dimY2 Then
    If blnGoThroughWalls = False Then
     picSnake(0).Visible = False
     timGameLoop.Enabled = False
     lblScore = "Score: " & intScore & "  Eaten: " & intScore / 10
    Else
     Call GoThroughWalls
    End If
   End If
  'check for eating itself
   For t = 1 To intSnakeSize
    If picSnake(0).Left = picSnake(t).Left And _
       picSnake(0).Top = picSnake(t).Top Then
     timGameLoop.Enabled = False
     lblScore = "Score: " & intScore & "  Eaten: " & intScore / 10
    End If
   Next t

  'check for eating food
   If picSnake(0).Left = picFood.Left And _
      picSnake(0).Top = picFood.Top Then
    intScore = intScore + 10
    lblScore.Caption = "Score: " & intScore
    Call GrowSnake
    Call PlaceFood
   End If
End Sub

Sub GoThroughWalls()
 'This sub makes the snake move through the walls
  
  'Locate where he the head is, and then move it to the opposite
   'Left
    If picSnake(0).Left < dimX1 Then
     picSnake(0).Left = dimX2
    End If
   'Right
    If picSnake(0).Left > dimX2 Then
     picSnake(0).Left = dimX1
    End If
   'Up
    If picSnake(0).Top < dimY1 Then
     picSnake(0).Top = dimY2
    End If
   'Down
    If picSnake(0).Top > dimY2 Then
     picSnake(0).Top = dimY1
    End If
End Sub

Sub GrowSnake()
 'This sub let's the snake grow by one
  'Add one
   intSnakeSize = intSnakeSize + 1
  'Load it
   Load picSnake(intSnakeSize)
   picSnake(intSnakeSize).Visible = True
End Sub
