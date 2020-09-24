VERSION 5.00
Begin VB.UserControl CheckBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.CheckBox checkboxii 
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.PictureBox picChecked 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   960
      Picture         =   "CheckBox.ctx":0000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   1320
      Width           =   225
   End
   Begin VB.PictureBox picUnChecked 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   720
      Picture         =   "CheckBox.ctx":0312
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   1320
      Width           =   225
   End
   Begin VB.PictureBox picRender 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   0
      Picture         =   "CheckBox.ctx":0624
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   0
      Width           =   225
   End
End
Attribute VB_Name = "CheckBox"
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
Dim blnChecked As String

Private Sub picRender_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If checkboxii.Value = Unchecked Then
  picRender.Picture = picUnChecked.Picture
 Else
  picRender.Picture = picChecked.Picture
 End If
End Sub

Private Sub picRender_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent click
 If checkboxii.Value = Unchecked Then
  picRender.Picture = picUnChecked.Picture
 Else
  picRender.Picture = picChecked.Picture
 End If
End Sub

Private Sub UserControl_Initialize()
 If checkboxii.Value = Unchecked Then
  picRender.Picture = picUnChecked.Picture
 Else
  picRender.Picture = picChecked.Picture
 End If
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=checkboxii,checkboxii,-1,Value
Public Property Get Value() As Integer
Attribute Value.VB_Description = "Returns/sets the value of an object."
    Value = checkboxii.Value
     If checkboxii.Value = Unchecked Then
      picRender.Picture = picUnChecked.Picture
     Else
      picRender.Picture = picChecked.Picture
     End If
End Property

Public Property Let Value(ByVal New_Value As Integer)
    checkboxii.Value() = New_Value
    PropertyChanged "Value"
     If checkboxii.Value = Unchecked Then
      picRender.Picture = picUnChecked.Picture
     Else
      picRender.Picture = picChecked.Picture
     End If
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    checkboxii.Value = PropBag.ReadProperty("Value", 0)
     If checkboxii.Value = Unchecked Then
      picRender.Picture = picUnChecked.Picture
     Else
      picRender.Picture = picChecked.Picture
     End If
    Set Picture = PropBag.ReadProperty("CheckedPic", Nothing)
    Set Picture = PropBag.ReadProperty("UncheckedPic", Nothing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Value", checkboxii.Value, 0)
     If checkboxii.Value = Unchecked Then
      picRender.Picture = picUnChecked.Picture
     Else
      picRender.Picture = picChecked.Picture
     End If
    Call PropBag.WriteProperty("CheckedPic", Picture, Nothing)
    Call PropBag.WriteProperty("UncheckedPic", Picture, Nothing)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picChecked,picChecked,-1,Picture
Public Property Get CheckedPic() As Picture
Attribute CheckedPic.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set CheckedPic = picChecked.Picture
End Property

Public Property Set CheckedPic(ByVal New_CheckedPic As Picture)
    Set picChecked.Picture = New_CheckedPic
    PropertyChanged "CheckedPic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picUnChecked,picUnChecked,-1,Picture
Public Property Get UncheckedPic() As Picture
Attribute UncheckedPic.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set UncheckedPic = picUnChecked.Picture
End Property

Public Property Set UncheckedPic(ByVal New_UncheckedPic As Picture)
    Set picUnChecked.Picture = New_UncheckedPic
    PropertyChanged "UncheckedPic"
End Property

