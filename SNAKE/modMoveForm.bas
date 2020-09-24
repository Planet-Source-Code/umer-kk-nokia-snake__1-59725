Attribute VB_Name = "modMoveForm"
'#########################################
'#         Developed By Umer KK          #
'#        Umerkhan_63@Hotmail.com        #
'#########################################

'This MOD sets declares the API's and Const's used to move a form

 Option Explicit

 Public Declare Function ReleaseCapture Lib "user32" () As Long
 Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 Public Const WM_NCLBUTTONDOWN = &HA1
 Public Const HTCAPTION = 2
 
 Public Function MouseMove(frm As Form)
  ReleaseCapture
  SendMessage frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
 End Function
