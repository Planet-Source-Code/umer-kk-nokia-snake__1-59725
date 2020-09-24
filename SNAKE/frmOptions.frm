VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   0  'None
   ClientHeight    =   3045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   203
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   389
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOptionsOne 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   0
      Picture         =   "frmOptions.frx":0000
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   260
      TabIndex        =   2
      Top             =   300
      Width           =   3900
      Begin VB.ListBox lstSkin 
         Enabled         =   0   'False
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
         ItemData        =   "frmOptions.frx":16DE2
         Left            =   360
         List            =   "frmOptions.frx":16DEC
         TabIndex        =   7
         Top             =   960
         Width           =   3375
      End
      Begin Snake_By_Umer.Button buttOK 
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         lbl             =   "OK"
      End
      Begin Snake_By_Umer.CheckBox chkboxWalls 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   255
         _extentx        =   450
         _extenty        =   450
      End
      Begin Snake_By_Umer.Button buttAbout 
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         lbl             =   "About"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   " Skins"
         Enabled         =   0   'False
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
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label lblGoThroughWalls 
         BackStyle       =   0  'Transparent
         Caption         =   " Snake can go through walls."
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
         Left            =   360
         TabIndex        =   4
         Top             =   255
         Width           =   3375
      End
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      Picture         =   "frmOptions.frx":16E0B
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   260
      TabIndex        =   0
      Top             =   0
      Width           =   3900
      Begin VB.Label lblBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
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
         TabIndex        =   1
         Top             =   60
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#########################################
'#         Developed By Umer KK          #
'#        Umerkhan_63@Hotmail.com        #
'#########################################

Option Explicit

Private Sub buttAbout_click()
 'show about form
  frmAbout.Show
 'disbale me
  Me.Enabled = False
End Sub

Private Sub buttOK_click()
 'unload the form
  Unload Me
 'enable frmsnake
  frmSnake.Enabled = True
  frmSnake.picGameField.SetFocus
 
  'If lstSkin.Text = "White Python" Then
  'Call SkinAll("whitepython")
 'Else
  'Call SkinAll("blackadder")
 'End If
End Sub

Private Sub chkboxWalls_click()
'Sets option to go through walls or not
 If chkboxWalls.Value = Checked Then
  blnGoThroughWalls = False
  chkboxWalls.Value = Unchecked
 Else
  blnGoThroughWalls = True
  chkboxWalls.Value = Checked
 End If
End Sub

Private Sub Form_Load()
 'Set width and height
  Me.Height = 2100
  Me.Width = 3915
 'set chkboxes
  If blnGoThroughWalls = True Then
   chkboxWalls.Value = Checked
  Else
   chkboxWalls.Value = Unchecked
  End If
 'Set buttons
  buttOK.Tekst = "OK"
  buttAbout.Tekst = "About"
End Sub

Private Sub lblBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'This Let's the Form Move, like a real titlebar
  Call MouseMove(Me)
End Sub

Private Sub picBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'This Let's the Form Move, like a real titlebar
  Call MouseMove(Me)
End Sub
