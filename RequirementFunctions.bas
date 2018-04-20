Attribute VB_Name = "RequirementFunctions"
' Rainer Winkler 02.10.2015
' Version 3.0

' MIT License
'
' Copyright (c) 2018
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Const lngHeaderRow As Long = 10
Const lngIDColumn As Long = 1
Const lngGliederungColumn As Long = 2
Const lngAnforderungColumn As Long = 3
Const lngDetailColumn As Long = 4
Const lngEbeneColumn As Long = 6
Const lngStatusColumn As Long = 7
' New Rainer Winkler 28.07.2015
Const lngReqStatusColumn As Long = 8
' End Rainer Winkler 28.07.2015
' Rainer Winkler 17.10.2014 2.1
' Readable constants for status
Const strStatusOpen = ""
Const strStatusCheck = "check" ' Rainer Winkler 28.07.2015: Obsolete
' New Rainer Winkler 28.07.2015
Const strStatusToBeDeveloped = "to be developed"
' End Rainer Winkler 28.07.2015
Const strStatusDevelopmentStarted = "in development"
Const strStatusDeveloped = "developed"
Const strStatusDeleted = "deleted" ' Rainer Winkler 28.07.2015: Obsolete
Const strStatusDeveloperTestOK = "Developer Test passed"
Const strStatusExternalTestOK = "External Test passed"
' New Rainer Winkler 28.07.2015
Const strStatusReqOpen = "Requirement Open"
Const strStatusReqConf = "Requirement Confirmed"
Const strStatusReqCheck = "Check Requirement"
Const strStatusReqDeleted = "Requirement deleted"
' End Rainer Winkler 28.07.2015
Const strStatusErrors = "Errors reported"
Const strEbeneHeader1 = "Header Level 1"
Const strEbeneHeader2 = "Header Level 2"
Const strEbeneHeader3 = "Header Level 3"
Const strEbeneHeader4 = "Header Level 4"
Const strEbeneHeader5 = "Header Level 5"
Const strEbeneHeader6 = "Header Level 6"
Const strEbeneUndefined = ""
Const strEbeneRequirement = "Requirement"
Const strEbeneComment = "Comment"

Sub KontextmenueErgaenzen()
' Rainer Winkler 09.04.2013
' Kontextmenu setzen

' Aus Monika Weber "Excel VBA"

Dim cb As CommandBar
Dim cbc As CommandBarControl

Set cb = Application.CommandBars("Cell")
cb.Reset

Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=5)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = True
    .Caption = "Generate ID"
    .OnAction = "GetNewID"
    End With
    
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=6)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = True
    .Caption = "Header level 1"
    .OnAction = "HeaderLevel1"
    End With
    
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=7)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = False
    .Caption = "Header level 2"
    .OnAction = "HeaderLevel2"
    End With
    
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=8)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = False
    .Caption = "Header level 3"
    .OnAction = "HeaderLevel3"
    End With
    
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=9)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = False
    .Caption = "Header level 4"
    .OnAction = "HeaderLevel4"
    End With
   
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=10)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = False
    .Caption = "Header level 5"
    .OnAction = "HeaderLevel5"
    End With
   
   
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=11)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = True
    .Caption = "Requirement"
    .OnAction = "ReqLevel"
    End With
    
' Neu am 17.10.2014 2.1
    
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=12)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = True
    .Caption = "Comment"
    .OnAction = "CommentLevel"
    End With
    
' New Rainer Winkler 28.07.2015
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=13)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = True
    .Caption = "Open Requirement"
    .OnAction = "RequiremementOpen"
    End With
    
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=14)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = False
    .Caption = "Confirmed Requirement"
    .OnAction = "RequiremementConfirmed"
    End With
    
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=15)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = False
    .Caption = "Requirement to be checked"
    .OnAction = "CheckRequiremement"
    End With
    
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=16)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = False
    .Caption = "Requirement deleted"
    .OnAction = "RequiremementDeleted"
    End With
' End Rainer Winkler 28.07.2015
    
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=17)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = True
    .Caption = "Open"
    .OnAction = "Opened"
    End With
    
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=18)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = False
    .Caption = "To be developed"
    .OnAction = "ToBeDeveloped"
    End With
    
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=19)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = False
    .Caption = "development started"
    .OnAction = "DevelopmentStarted"
    End With

Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=20)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = False
    .Caption = "Developed"
    .OnAction = "Developed"
    End With
    

    

    
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=21)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = False
    .Caption = "Errors reported"
    .OnAction = "ContainsErrors"
    End With
    

   

    
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=22)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = False
    .Caption = "Developer test passed"
    .OnAction = "Tested"
    End With
    
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=23)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = False
    .Caption = "External test passed"
    .OnAction = "PassedExternalTest"
    End With
    
' Neu am 28.02.2013
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=24)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = True
    .Caption = "Deleted"
    .OnAction = "Deleted"
    End With
    
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=25)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = False
    .Caption = "Check (Obsolete)"
    .OnAction = "Check"
    End With
    
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=26)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = True
    .Caption = "Caption numbering refresh"
    .OnAction = "GliederungErstellen"
    End With
    
Set cbc = cb.Controls.Add(Type:=msoControlButton, Before:=27)
  '(Type:=msoControlButton, Before:=5)
  With cbc
    .BeginGroup = True
    .Caption = "Export requirements into file"
    .OnAction = "Export"
    End With
    
End Sub

Sub KontextmenueZuruecksetzen()
' Rainer Winkler 13.03.2012
' Aus Monika Weber "Excel VBA"
  Application.CommandBars("Cell").Reset
End Sub

Sub NeuesMenue()
' Rainer Winkler 13.03.2012
' Menu setzen

' Aus Monika Weber "Excel VBA"
  Dim cb As CommandBar, cbc As CommandBarControl
  Set cb = Application.CommandBars("Worksheet Menu Bar")
  Set cbc = cb.Controls.Add(Type:=msoControlPopup)
  cbc.Caption = "Requirements"
  
  With cbc.Controls.Add
    .Caption = "ID ziehen"
    .OnAction = "GetNewID"
    End With
    
  With cbc.Controls.Add
    .BeginGroup = True
    .Caption = "Request Level 1"
    .OnAction = "ReqLevel1"
  End With

  With cbc.Controls.Add
    .BeginGroup = False
    .Caption = "Request Level 2"
    .OnAction = "ReqLevel2"
  End With
  
  With cbc.Controls.Add
    .BeginGroup = False
    .Caption = "Request Level 3"
    .OnAction = "ReqLevel3"
  End With
  
  With cbc.Controls.Add
    .BeginGroup = False
    .Caption = "Request Level 4"
    .OnAction = "ReqLevel4"
  End With
  
  With cbc.Controls.Add
    .BeginGroup = False
    .Caption = "Request Level 5"
    .OnAction = "ReqLevel5"
  End With
  
  With cbc.Controls.Add
    .BeginGroup = False
    .Caption = "Request Level 6"
    .OnAction = "ReqLevel6"
  End With
  
  
End Sub

Sub NeuesMenueLoeschen()
' Rainer Winkler 13.03.2012
' Aus Monika Weber "Excel VBA"
  Dim cbc As CommandBarControl
  For Each cbc In Application.CommandBars("Worksheet Menu Bar").Controls
    If cbc.Caption = "Requirements" Then
       cbc.Delete
    End If
  Next cbc
End Sub

Sub GetNewID()
' Rainer Winkler 13.03.2012
' Ziehen einer neuen "eindeutigen" ID
' In einer Zelle mit dem Namen LastID muss die letzte gültige Nummer stehen
' In einer Zelle mit dem Namen prefixForID muss ein Prefix stehen (darf leer sein)
' Die Eindeutigkeit wird hier nicht wirklich geprüft. Es ist in der Verantwortung des Nutzers
' das die Eindeutigkeit nicht durch Fehlbedienung verschwindet.

Dim id As Integer

  If ActiveCell.Value <> "" Then
    MsgBox "Only empty cells can be filled wit an ID"
  Else
    If Range("prefixForID").Value = "" Then
    MsgBox "Enter a searchable prefix for the ID on table Einstellungen first"
    Else
    'Range("lastID").Value = Range("lastID").Value + 1
    id = Range("lastID").Value + 1
    Range("lastID").Value = id

    ActiveCell.Value = Range("prefixForID").Value & Range("lastID").Value
    
    Call ReqLevel
    End If
  End If
End Sub

Sub HeaderLevel1()
' Rainer Winkler 07.03.2013
ActiveCell.IndentLevel = 0
Cells(ActiveCell.Row, lngEbeneColumn).Value = strEbeneHeader1
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Italic = False
ActiveCell.Font.Size = Range("fontSizeLevel1").Value
End Sub

Sub HeaderLevel2()
' Rainer Winkler 07.03.2013
ActiveCell.IndentLevel = 1
Cells(ActiveCell.Row, lngEbeneColumn).Value = strEbeneHeader2
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Italic = False
ActiveCell.Font.Size = Range("fontSizeLevel2").Value
End Sub

Sub HeaderLevel3()
' Rainer Winkler 07.03.2013
ActiveCell.IndentLevel = 2
Cells(ActiveCell.Row, lngEbeneColumn).Value = strEbeneHeader3
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Italic = False
ActiveCell.Font.Size = Range("fontSizeLevel3").Value
End Sub

Sub HeaderLevel4()
' Rainer Winkler 07.03.2013
ActiveCell.IndentLevel = 3
Cells(ActiveCell.Row, lngEbeneColumn).Value = strEbeneHeader4
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Italic = False
ActiveCell.Font.Size = Range("fontSizeLevel4").Value
End Sub

Sub HeaderLevel5()
' Rainer Winkler 07.03.2013
ActiveCell.IndentLevel = 4
Cells(ActiveCell.Row, lngEbeneColumn).Value = strEbeneHeader5
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Italic = False
ActiveCell.Font.Size = Range("fontSizeLevel5").Value
End Sub


Sub ReqLevel()
' Rainer Winkler 09.04.2013
ActiveCell.IndentLevel = 0
Cells(ActiveCell.Row, lngEbeneColumn).Value = strEbeneRequirement
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Italic = False
ActiveCell.Font.Size = Range("fontSizeLevel6").Value
End Sub

Sub CommentLevel()
' Rainer Winkler 17.10.2014 2.1
ActiveCell.IndentLevel = 0
Cells(ActiveCell.Row, lngEbeneColumn).Value = strEbeneComment
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Italic = True
ActiveCell.Font.Size = Range("fontSizeLevel6").Value
End Sub
' New Rainer Winkler 28.07.2015
Sub RequiremementOpen()
Cells(ActiveCell.Row, lngReqStatusColumn).Value = strStatusReqOpen
Cells(ActiveCell.Row, lngAnforderungColumn).Interior.Pattern = xlNone
Cells(ActiveCell.Row, lngAnforderungColumn).Interior.TintAndShade = 0
Cells(ActiveCell.Row, lngAnforderungColumn).Interior.PatternTintAndShade = 0
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Strikethrough = False
End Sub

Sub RequiremementConfirmed()
Cells(ActiveCell.Row, lngReqStatusColumn).Value = strStatusReqConf
Cells(ActiveCell.Row, lngAnforderungColumn).Interior.Color = 5296274
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Strikethrough = False
End Sub
Sub CheckRequiremement()
Cells(ActiveCell.Row, lngReqStatusColumn).Value = strStatusReqCheck
Cells(ActiveCell.Row, lngAnforderungColumn).Interior.Pattern = xlSolid
Cells(ActiveCell.Row, lngAnforderungColumn).Interior.PatternTintAndShade = 0
Cells(ActiveCell.Row, lngAnforderungColumn).Interior.Color = 65535
Cells(ActiveCell.Row, lngAnforderungColumn).Interior.PatternColorIndex = xlAutomatic
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Strikethrough = False
End Sub

Sub RequiremementDeleted()
Cells(ActiveCell.Row, lngReqStatusColumn).Value = strStatusReqDeleted
Cells(ActiveCell.Row, lngAnforderungColumn).Interior.ThemeColor = xlThemeColorDark1
Cells(ActiveCell.Row, lngAnforderungColumn).Interior.TintAndShade = -0.149998474074526
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Strikethrough = True
End Sub

' End Rainer Winkler 28.07.2015

Sub Deleted()
' Rainer Winkler 28.02.2013
Cells(ActiveCell.Row, lngStatusColumn).Value = strStatusDeleted
Cells(ActiveCell.Row, lngIDColumn).Interior.ThemeColor = xlThemeColorDark1
Cells(ActiveCell.Row, lngIDColumn).Interior.TintAndShade = -0.149998474074526
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Strikethrough = True
End Sub

Sub DevelopmentStarted()
' Rainer Winkler 24.07.2015 Local change
Cells(ActiveCell.Row, lngStatusColumn).Value = strStatusDevelopmentStarted
Cells(ActiveCell.Row, lngIDColumn).Interior.Pattern = xlSolid
Cells(ActiveCell.Row, lngIDColumn).Interior.PatternColorIndex = xlAutomatic
Cells(ActiveCell.Row, lngIDColumn).Interior.ThemeColor = xlThemeColorAccent3
Cells(ActiveCell.Row, lngIDColumn).Interior.TintAndShade = 0.799981688894314
Cells(ActiveCell.Row, lngIDColumn).Interior.PatternTintAndShade = 0
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Strikethrough = False
End Sub

Sub Developed()
' Rainer Winkler 28.02.2013
Cells(ActiveCell.Row, lngStatusColumn).Value = strStatusDeveloped
Cells(ActiveCell.Row, lngIDColumn).Interior.Color = 5296274
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Strikethrough = False
End Sub


Sub ToBeDeveloped()
' Rainer Winkler 17.10.2014 2.1
Cells(ActiveCell.Row, lngStatusColumn).Value = strStatusToBeDeveloped
Cells(ActiveCell.Row, lngIDColumn).Interior.Pattern = xlSolid
Cells(ActiveCell.Row, lngIDColumn).Interior.PatternColorIndex = xlAutomatic
Cells(ActiveCell.Row, lngIDColumn).Interior.Color = 49407
Cells(ActiveCell.Row, lngIDColumn).Interior.TintAndShade = 0
Cells(ActiveCell.Row, lngIDColumn).Interior.PatternTintAndShade = 0
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Strikethrough = False
End Sub

Sub Check()
' Rainer Winkler 07.03.2013 (28.7.2015: Obsolete)
Cells(ActiveCell.Row, lngStatusColumn).Value = strStatusCheck
Cells(ActiveCell.Row, lngIDColumn).Interior.Pattern = xlSolid
Cells(ActiveCell.Row, lngIDColumn).Interior.PatternTintAndShade = 0
Cells(ActiveCell.Row, lngIDColumn).Interior.Color = 65535
Cells(ActiveCell.Row, lngIDColumn).Interior.PatternColorIndex = xlAutomatic
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Strikethrough = False
End Sub

Sub Opened()
' Rainer Winkler 28.02.2013
Cells(ActiveCell.Row, lngStatusColumn).Value = strStatusOpen
Cells(ActiveCell.Row, lngIDColumn).Interior.Pattern = xlNone
Cells(ActiveCell.Row, lngIDColumn).Interior.TintAndShade = 0
Cells(ActiveCell.Row, lngIDColumn).Interior.PatternTintAndShade = 0
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Strikethrough = False
End Sub
Sub ContainsErrors()
' Rainer Winkler 17.10.2014 2.1
Cells(ActiveCell.Row, lngStatusColumn).Value = strStatusErrors
Cells(ActiveCell.Row, lngIDColumn).Interior.Pattern = xlSolid
Cells(ActiveCell.Row, lngIDColumn).Interior.PatternColorIndex = xlAutomatic
Cells(ActiveCell.Row, lngIDColumn).Interior.Color = 255
Cells(ActiveCell.Row, lngIDColumn).Interior.TintAndShade = 0
Cells(ActiveCell.Row, lngIDColumn).Interior.PatternTintAndShade = 0
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Strikethrough = False
End Sub
Sub Tested()
' Only Developer Test
' Rainer Winkler 28.02.2013
Cells(ActiveCell.Row, lngStatusColumn).Value = strStatusDeveloperTestOK
Cells(ActiveCell.Row, lngIDColumn).Interior.Pattern = xlSolid
Cells(ActiveCell.Row, lngIDColumn).Interior.PatternColorIndex = xlAutomatic
Cells(ActiveCell.Row, lngIDColumn).Interior.Color = 15773696
Cells(ActiveCell.Row, lngIDColumn).Interior.TintAndShade = 0
Cells(ActiveCell.Row, lngIDColumn).Interior.PatternTintAndShade = 0
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Strikethrough = False
End Sub

Sub PassedExternalTest()
' Rainer Winkler 17.10.2014 2.1
Cells(ActiveCell.Row, lngStatusColumn).Value = strStatusExternalTestOK
Cells(ActiveCell.Row, lngIDColumn).Interior.Pattern = xlSolid
Cells(ActiveCell.Row, lngIDColumn).Interior.PatternColorIndex = xlAutomatic
Cells(ActiveCell.Row, lngIDColumn).Interior.Color = 12611584
Cells(ActiveCell.Row, lngIDColumn).Interior.TintAndShade = 0
Cells(ActiveCell.Row, lngIDColumn).Interior.PatternTintAndShade = 0
Cells(ActiveCell.Row, lngAnforderungColumn).Font.Strikethrough = False
End Sub
Sub Export()

    ' Relevante Spalten ermitteln

    Dim lngFirstRow As Long
    Dim lngLastRow As Long
    Dim lngLastRowGliederung As Long
    Dim lngLastRowAnforderung As Long
    lngFirstRow = lngHeaderRow + 1
    lngLastRowAnforderung = Cells(65536, lngAnforderungColumn).End(xlUp).Row
    lngLastRowGliederung = Cells(65536, 2).End(xlUp).Row
    If lngLastRowAnforderung > lngLastRowGliederung Then
        lngLastRow = lngLastRowAnforderung
    Else
        lngLastRow = lngLastRowGliederung
    End If
    
    ' Zieldokument öffnen
    ' Quelle: http://vba1.de/vba/032html.php
    Dim sInput As String
    sInput = InputBox("Filename with complete path:", "Pfad")
    
    Dim iAntwortIDAusgabe As Integer
    Dim iAntwortCloseWithDot As Integer
    iAntwortIDAusgabe = MsgBox("Add ID of requirement?", vbYesNo, "")
    iAntwortCloseWithDot = MsgBox("Finish requirements and comments with dot at the end?", vbYesNo, "")
    
    If sInput = "" Then
        MsgBox "No Filename with complete path entered", vbExclamation, "Info"
        sInput = InputBox("Filename with complete path:", "Pfad")
    Else
    'Hier deine Prozedur bzw. dein Speichern eintragen z.B
    
        Dim fsDatei As Object
        Dim fsinhalt As Object
        'html-Datei erstellen
        Set fsDatei = CreateObject("Scripting.FileSystemObject")
        fsDatei.CreateTextFile sInput
    
        Set fsDatei = fsDatei.getfile(sInput)

        Set fsinhalt = fsDatei.OpenAsTextStream(2, -2)
    
        ' html-Grundgerüst erstellen
        ' & vbCrLf erzeugt im html-Quelltext einen Zeilenumbruch mit Cariage Return und Linefeed
        fsinhalt.write "<html>" & vbCrLf
        fsinhalt.write "<!-- diese Datei wurde über eine VBA-Prozedur erstellt-->" & vbCrLf
        fsinhalt.write "<head>" & vbCrLf
        fsinhalt.write "</head>" & vbCrLf
        fsinhalt.write "<body>" & vbCrLf
  
        For i = lngFirstRow To lngLastRow
  
            ' New 22.07.2015 regard also the case of deleted to prevent export in that case
            ' New Rainer Winkler 28.07.2015 Check Request Deleted
            If Cells(i, lngStatusColumn).Value = "-" Or Cells(i, lngStatusColumn).Value = "deleted" Or Cells(i, lngReqStatusColumn).Value = strStatusReqDeleted Or Cells(i, lngAnforderungColumn).Value = "" Then
            ' Zeile ignorieren
            Else
            Dim sOpen As String
            Dim sClose As String
                Select Case Cells(i, lngEbeneColumn).Value
                    Case Is = strEbeneHeader1
                        sOpen = "<h1>"
                        sClose = "</h1>"
                    Case Is = strEbeneHeader2
                        sOpen = "<h2>"
                        sClose = "</h2>"
                    Case Is = strEbeneHeader3
                        sOpen = "<h3>"
                        sClose = "</h3>"
                    Case Is = strEbeneHeader4
                        sOpen = "<h4>"
                        sClose = "</h4>"
                    Case Is = strEbeneHeader5
                        sOpen = "<h5>"
                        sClose = "</h5>"
                    Case Is = strEbeneHeader6
                        sOpen = "<h6>"
                        sClose = "</h6>"
                    ' BEGIN NEW 2.1
                    Case Is = strEbeneComment
                        sOpen = "<em>"
                        sClose = "</em>"
                    ' END NEW 2.1
                    Case Is = strEbeneRequirement
                        sOpen = "<p>"
                        sClose = "</p>"
                    ' Für den Fall, dass die Anforderung nicht im Kontextmenu gekennzeichnet ist, ist Spalte
                    ' Ebene leer. Auch in diesem Fall als Anforderung behandeln
                    Case Is = strEbeneUndefined
                        sOpen = "<p>"
                        sClose = "</p>"
                End Select
                
                fsinhalt.write sOpen
                fsinhalt.write Cells(i, lngAnforderungColumn).Value
                If Cells(i, lngDetailColumn).Value <> "" Then
                fsinhalt.write " (" & Cells(i, lngDetailColumn).Value & ")"
                End If
                If iAntwortIDAusgabe = vbYes Then
                    If Cells(i, lngIDColumn).Value <> "" Then
                        fsinhalt.write " [" & Cells(i, lngIDColumn).Value & "]"
                    End If
                End If
                ' Wenn gewünscht werden Anforderungen und Kommentare, aber nicht Überschriften und auch nicht Anforderungen
                ' die Überschriften sind mit einem Punkt abgeschlossen
                If iAntwortCloseWithDot = vbYes And (sOpen = "<p>" Or sOpen = "<em>") Then
                    fsinhalt.write "."
                End If
                fsinhalt.write sClose & vbCrLf
        
            End If
  
        Next i
        
        ' alle html-tags wieder schließen

        fsinhalt.write "</body>" & vbCrLf
        fsinhalt.write "</html>"
        fsinhalt.Close
        
        
    End If
End Sub

Sub GliederungErstellen()

    Dim lngFirstRow As Long
    Dim lngLastRow As Long
    Dim lngLastRowGliederung As Long
    Dim lngLastRowAnforderung As Long
    lngFirstRow = lngHeaderRow + 1
    lngLastRowAnforderung = Cells(65536, lngAnforderungColumn).End(xlUp).Row
    lngLastRowGliederung = Cells(65536, 2).End(xlUp).Row
    If lngLastRowAnforderung > lngLastRowGliederung Then
        lngLastRow = lngLastRowAnforderung
    Else
        lngLastRow = lngLastRowGliederung
    End If
  
    Dim lngLevel1Count As Long
    Dim lngLevel2Count As Long
    Dim lngLevel3Count As Long
    Dim lngLevel4Count As Long
    Dim lngLevel5Count As Long
    Dim lngLevel6Count As Long
    lngLevel1Count = 0
    lngLevel2Count = 0
    lngLevel3Count = 0
    lngLevel4Count = 0
    lngLevel5Count = 0
    lngLevel6Count = 0
    Dim i As Long
    For i = lngFirstRow To lngLastRow
        If Cells(i, lngStatusColumn).Value = strStatusDeleted _
            Or Cells(i, lngAnforderungColumn).Value = strStatusOpen _
            Or Cells(i, lngEbeneColumn).Value = strEbeneRequirement _
            Or Cells(i, lngEbeneColumn).Value = strEbeneComment Then
            Cells(i, lngGliederungColumn).Value = ""
        Else
            If Cells(i, lngEbeneColumn).Value = strEbeneHeader1 And Cells(i, lngAnforderungColumn).Value <> "" Then
            lngLevel1Count = lngLevel1Count + 1
            lngLevel2Count = 0
            lngLevel3Count = 0
            lngLevel4Count = 0
            lngLevel5Count = 0
            lngLevel6Count = 0
            Cells(i, lngGliederungColumn).Value = lngLevel1Count
            End If
            If Cells(i, lngEbeneColumn).Value = strEbeneHeader2 And Cells(i, lngAnforderungColumn).Value <> "" Then
                lngLevel2Count = lngLevel2Count + 1
                Cells(i, lngGliederungColumn).Value = CStr(lngLevel1Count) + "." + CStr(lngLevel2Count)
                lngLevel3Count = 0
                lngLevel4Count = 0
                lngLevel5Count = 0
                lngLevel6Count = 0
            End If
            If Cells(i, lngEbeneColumn).Value = strEbeneHeader3 And Cells(i, lngAnforderungColumn).Value <> "" Then
                lngLevel3Count = lngLevel3Count + 1
                Cells(i, lngGliederungColumn).Value = CStr(lngLevel1Count) + "." + CStr(lngLevel2Count) _
                + "." + CStr(lngLevel3Count)
                lngLevel4Count = 0
                lngLevel5Count = 0
                lngLevel6Count = 0
            End If
            If Cells(i, lngEbeneColumn).Value = strEbeneHeader4 And Cells(i, lngAnforderungColumn).Value <> "" Then
                lngLevel4Count = lngLevel4Count + 1
                Cells(i, lngGliederungColumn).Value = CStr(lngLevel1Count) + "." + CStr(lngLevel2Count) _
                + "." + CStr(lngLevel3Count) + "." + CStr(lngLevel4Count)
                lngLevel5Count = 0
                lngLevel6Count = 0
            End If
            If Cells(i, lngEbeneColumn).Value = strEbeneHeader5 And Cells(i, lngAnforderungColumn).Value <> "" Then
                lngLevel5Count = lngLevel5Count + 1
                Cells(i, lngGliederungColumn).Value = CStr(lngLevel1Count) + "." + CStr(lngLevel2Count) _
                + "." + CStr(lngLevel3Count) + "." + CStr(lngLevel4Count) _
                + "." + CStr(lngLevel5Count)
                lngLevel6Count = 0
            End If
            If Cells(i, lngEbeneColumn).Value = strEbeneHeader6 And Cells(i, lngAnforderungColumn).Value <> "" Then
                lngLevel6Count = lngLevel6Count + 1
                Cells(i, lngGliederungColumn).Value = CStr(lngLevel1Count) + "." + CStr(lngLevel2Count) _
                + "." + CStr(lngLevel3Count) + "." + CStr(lngLevel4Count) _
                + "." + CStr(lngLevel5Count) + "." + CStr(lngLevel6Count)
            End If
        End If
    Next i
End Sub

Sub BlattEinstellungenSchuetzen()
  ' Rainer Winkler CubeServ 13.03.2012
  ' Die ID wird vom Makro hochgezählt, darum darf der Schutz nur für den Anwender gelten
  ActiveWorkbook.Sheets("Einstellungen").Protect UserInterfaceOnly:=True
  ' ActiveSheet.Protect UserInterfaceOnly:=True
End Sub

Sub BlattEinstellungenSchutzAufheben()
  ' Rainer Winkler CubeServ 13.03.2012
  ActiveSheet.Unprotect
End Sub


