VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' Insert the code from line 13 to your worksheet so that the VBA code is added to the context menu automatically.

Private Sub Workbook_Open()
  ' Rainer Winkler CubeServ 13.March.2012
  ' Add menut items to the context menu
  RequirementFunctions.AddSimpleReqMenuToConextMenu
  RequirementFunctions.NeuesMenue
  ' Der Schutz muss so sein, dass das Makro die ID ändern darf
  ' Darum wird der Schutz hier auch dann gesetzt, wenn er dieses bereits ist
  RequirementFunctions.BlattEinstellungenSchuetzen
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
  ' Rainer Winkler CubeServ 13.March.2012
  ' Remove the menu items when the workbook is closed
  ' It is save to run this method manually. But it may remove context menu items from other projects too
  RequirementFunctions.ResetContextMenuChanges
  RequirementFunctions.NeuesMenueLoeschen
End Sub
