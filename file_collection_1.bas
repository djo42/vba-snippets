Attribute VB_Name = "zusammenfassen"
Sub zusammenfassen()
    Dim Zeile&, letzteZ&
    
    'Auswertungsblatt einfügen
    Worksheets.Add.Name = "Summary"
    ActiveSheet.Move Before:=Worksheets(1)
    
    'Von Blatt 1 bis Blatt 10 zusammenfassen
    For i = 2 To 7
        With Worksheets(i)
            letzteZ = .Cells(Rows.Count, 1).End(xlUp).Row
            Zeile = Worksheets("Summary").Cells(Rows.Count, 1).End(xlUp).Row + 1
            .Range("A2:AJ" & letzteZ).Copy Worksheets("Summary").Range("A" & Zeile)
        End With
    Next
    
End Sub