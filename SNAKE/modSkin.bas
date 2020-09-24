Attribute VB_Name = "modSkin"
'#########################################
'#         Developed By Umer KK          #
'#        Umerkhan_63@Hotmail.com        #
'#########################################


'This MOD handles the skinning,
'It's in a HUGE
'Becase skinning snake is now very heay, and lots of bitmaps,
'it could be reduced, and the users hould be able to ADD skins

'dont mind this code, still under contruction,will be changed, pictureboxes will be replaced by i think PRINT, and graphics won't be
'created for every meny anymore, we'll see...

 Option Explicit
 Public strLoadedSkin As String
 
 Const pSKINS As String = "\art\ingame\"
 
 Const pBAR As String = "\bar.bmp"
 Const pOPTIONS As String = "\options.bmp"
 Const pGAMEFIELD As String = "\gamefield.bmp"
 Const pBODY As String = "\body.bmp"
 Const pFOOD As String = "\food.bmp"
 Const pINFO As String = "\info.bmp"
 Const pButtUp As String = "\butt_up.bmp"
 Const pButtDown As String = "\butt_dn.bmp"
 Const pChecked As String = "checkbox.bmp"
 Const pUnChecked As String = "\checkboxc.bmp"
 
 Public Function SkinAll(Skin As String)
  'this function will skin the forms
    
   'skin the frmSnake form
    Call SkinFrM(Skin, frmSnake)
 End Function
 
 Public Function SkinFrM(Skin As String, frm As Form)
  'skins an individual form
   frm.picBar.Picture = App.Path & pSKINS & Skin & pBAR
 End Function
 
