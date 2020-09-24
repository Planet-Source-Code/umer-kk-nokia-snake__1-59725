VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   562
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLogo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   -30
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   1800
      ScaleWidth      =   3930
      TabIndex        =   5
      Top             =   300
      Width           =   3930
      Begin VB.Label lblExplenations 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   3735
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   0
      Picture         =   "frmAbout.frx":171A2
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   260
      TabIndex        =   2
      Top             =   2100
      Width           =   3900
      Begin Snake_By_Umer.Button buttOK 
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   390
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         lbl             =   "OK"
      End
      Begin VB.Label lblExplenations2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      Picture         =   "frmAbout.frx":20A3C
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   260
      TabIndex        =   0
      Top             =   0
      Width           =   3900
      Begin VB.Label lblBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "About"
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
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#########################################
'#         Developed By Umer KK          #
'#        Umerkhan_63@Hotmail.com        #
'#########################################
Option Explicit

Private Sub buttOK_click()
'unload the form
 Unload Me
'enabled the options form
 frmOptions.Enabled = True
End Sub

Private Sub Form_Load()
 'Set width and height
  Me.Height = 2850
  Me.Width = 3915
 'Set buttons
  buttOK.Tekst = "OK"
 'add text
  lblExplenations.Caption = "Coded and Designed by UmEr KK."
  lblExplenations2.Caption = "Umerkhan_63@Hotmail.com"
End Sub

Private Sub lblBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'This Let's the Form Move, like a real titlebar
  Call MouseMove(Me)
End Sub

Private Sub picBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'This Let's the Form Move, like a real titlebar
  Call MouseMove(Me)
End Sub

