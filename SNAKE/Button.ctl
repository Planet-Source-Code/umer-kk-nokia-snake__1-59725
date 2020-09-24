VERSION 5.00
Begin VB.UserControl Button 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox picRender 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      Picture         =   "Button.ctx":0000
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   2
      Top             =   0
      Width           =   1200
      Begin VB.Label lblTekst 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "New Game"
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
         Left            =   0
         TabIndex        =   3
         Top             =   45
         Width           =   1215
      End
   End
   Begin VB.PictureBox picDown 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   360
      Left            =   0
      Picture         =   "Button.ctx":1302
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   1
      Top             =   2520
      Width           =   1260
   End
   Begin VB.PictureBox picUp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   360
      Left            =   0
      Picture         =   "Button.ctx":2604
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   0
      Top             =   2880
      Width           =   1260
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#########################################
'#         Developed By Umer KK          #
'#        Umerkhan_63@Hotmail.com        #
'#########################################

Option Explicit
Event click()

'Private Sub lblTekst_Click()
' picRender.Picture = picDown.Picture
' lblTekst.Move 1, 4
'End Sub

Private Sub lblTekst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 picRender.Picture = picDown.Picture
 lblTekst.Move 1, 4
End Sub

Private Sub lblTekst_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 picRender.Picture = picUp.Picture
 lblTekst.Move 0, 3
 RaiseEvent click
End Sub

'Private Sub picRender_Click()
' picRender.Picture = picUp.Picture
' lblTekst.Move 1, 4
'End Sub

Private Sub picRender_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 picRender.Picture = picUp.Picture
 lblTekst.Move 1, 4
End Sub

Private Sub picRender_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 picRender.Picture = picUp.Picture
 lblTekst.Move 0, 3
 RaiseEvent click
End Sub

Private Sub UserControl_Initialize()
 picRender.Picture = picUp.Picture
 UserControl.Width = picRender.Width
 UserControl.Height = picRender.Height
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTekst,lblTekst,-1,Caption
Public Property Get Tekst() As String
Attribute Tekst.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Tekst = lblTekst.Caption
End Property

Public Property Let Tekst(ByVal New_Tekst As String)
    lblTekst.Caption() = New_Tekst
    PropertyChanged "Tekst"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picUp,picUp,-1,Picture
Public Property Get UP() As Picture
Attribute UP.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set UP = picUp.Picture
End Property

Public Property Set UP(ByVal New_UP As Picture)
    Set picUp.Picture = New_UP
    PropertyChanged "UP"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=lblTekst,lblTekst,-1,Caption
'Public Property Get DOWN() As String
'    DOWN = lblTekst.Caption
'End Property
'
'Public Property Let DOWN(ByVal New_DOWN As String)
'    lblTekst.Caption() = New_DOWN
'    PropertyChanged "DOWN"
'End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'    lblTekst.Caption = PropBag.ReadProperty("Tekst", "New Game")
    Set Picture = PropBag.ReadProperty("UP", Nothing)
    lblTekst.Caption = PropBag.ReadProperty("LBL", "New Game")
    'Set Picture = PropBag.ReadProperty("DOWN", Nothing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'    Call PropBag.WriteProperty("Tekst", lblTekst.Caption, "New Game")
    Call PropBag.WriteProperty("UP", Picture, Nothing)
    Call PropBag.WriteProperty("LBL", lblTekst.Caption, "New Game")
    Call PropBag.WriteProperty("DOWN", Picture, Nothing)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picDown,picDown,-1,Picture
Public Property Get DOWN() As Picture
Attribute DOWN.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set DOWN = picDown.Picture
End Property

Public Property Set DOWN(ByVal New_DOWN As Picture)
    Set picDown.Picture = New_DOWN
    PropertyChanged "DOWN"
End Property

