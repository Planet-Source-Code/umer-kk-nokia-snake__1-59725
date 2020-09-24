Attribute VB_Name = "modAlphaBlend"
Option Explicit

 'This AlphaBlending Code isn't mine,
 'I used it with permission of
 'Alexander Anikin, aka@i.com.ua
 
 'But i did adapt it to my needs,
 'I made a function, so it's like an engine that
 'Can be called througout the whole project.
 'Function Blend (picSrc as Picturebox, frm as Form, Speed as Long)

 Public Declare Function AlphaBlend _
  Lib "msimg32" ( _
  ByVal hDestDC As Long, _
  ByVal X As Long, ByVal Y As Long, _
  ByVal nWidth As Long, _
  ByVal nHeight As Long, _
  ByVal hSrcDC As Long, _
  ByVal xSrc As Long, _
  ByVal ySrc As Long, _
  ByVal widthSrc As Long, _
  ByVal heightSrc As Long, _
  ByVal dreamAKA As Long) _
  As Boolean 'only Windows 98 or Later
 Dim Num As Byte, nN%, nBlend&
 Dim FormWidth As Long
 Dim FormHeight As Long
 
 Public Function BlendIn(picSrc As PictureBox, frm As Form, SetX As Long, SetY As Long, Speed As Long)
 Num = 255
 nN = Speed
 Do
  DoEvents
  '***********************************************
   nBlend = vbBlue - CLng(Num) * (vbYellow + 1)
  'It's Magic Formula is
  'Alchemical Mixture of Elements of Gold & Sky
  'It's obtained by an almost mystical way
  '***********************************************
  '(above, the original text added with the code,
  ' I have no idea what he means by this...)
  Num = Num - nN
  If Num = 0 Then
    'nN = -1
    Exit Function
  ElseIf Num = 255 Then
    nN = Speed
  End If
  frm.Cls
  
  AlphaBlend frm.hDC, SetX, SetY, picSrc.ScaleWidth, picSrc.ScaleHeight, picSrc.hDC, 0, 0, picSrc.ScaleWidth, picSrc.ScaleHeight, nBlend
 Loop
End Function

 Public Function BlendOut(picSrc As PictureBox, frm As PictureBox, SetX As Long, SetY As Long, Speed As Long)
 Num = 0
 nN = -Speed
 Do
  DoEvents
  '***********************************************
   nBlend = vbBlue - CLng(Num) * (vbYellow + 1)
  'It's Magic Formula is
  'Alchemical Mixture of Elements of Gold & Sky
  'It's obtained by an almost mystical way
  '***********************************************
  '(above, the original text added with the code,
  ' I have no idea what he means by this...)
  Num = Num - nN
  If Num = 0 Then
    nN = -Speed
  ElseIf Num = 255 Then
    'nN = 1
    Exit Function
  End If
  frm.Cls
  
  AlphaBlend frm.hDC, SetX, SetY, picSrc.ScaleWidth, picSrc.ScaleHeight, picSrc.hDC, 0, 0, picSrc.ScaleWidth, picSrc.ScaleHeight, nBlend
 Loop
End Function

