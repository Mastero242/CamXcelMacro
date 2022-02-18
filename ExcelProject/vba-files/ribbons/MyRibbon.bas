Attribute VB_Name = "MyRibbon"

'namespace=vba-files/ribbons

'/*
'[Ribbon Menu Action]
'Ribbon Buttom 1 example call code
'
'*/
Public Sub btn1(ByRef control As Office.IRibbonControl)

 Call ImagesController.SetRootPath

End Sub

'/*
'[Ribbon Menu Action]
'Ribbon Buttom 2 example call code
'
'*/
Public Sub btn2(ByRef control As Office.IRibbonControl)

 Call ImagesController.CreateCatalogSheet
 
End Sub


'/*
'[Ribbon Menu Action]
'Ribbon Buttom 2 example call code
'
'*/
Public Sub btn3(ByRef control As Office.IRibbonControl)

 Call ReferencesController.ExcecuteRefLink
 
End Sub