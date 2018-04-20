VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True



Private Sub Workbook_Open()
  ' Rainer Winkler CubeServ 13.03.2012
  RequirementFunctions.AddSimpleReqMenuToConextMenu
  RequirementFunctions.NeuesMenue
  ' Der Schutz muss so sein, dass das Makro die ID ändern darf
  ' Darum wird der Schutz hier auch dann gesetzt, wenn er dieses bereits ist
  RequirementFunctions.BlattEinstellungenSchuetzen
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
  ' Rainer Winkler CubeServ 13.03.2012
  ' Wenn das Workbook geschlossen wird sollen die Menus wieder
  ' entfernt werden.
  RequirementFunctions.KontextmenueZuruecksetzen
  RequirementFunctions.NeuesMenueLoeschen
End Sub
