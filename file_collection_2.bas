Attribute VB_Name = "file_collection"

Sub file_collection()

    Dim Ziel As Object
    Dim Quelle As Object
    Dim Pfad As String
    Dim Datei As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set Ziel = ActiveWorkbook

    Pfad = InputBox("Pfad eingeben", "Pfad")
    Datei = Dir(CStr(Pfad & "*.xl*"))

    Do While Datei <> ""

        Set Quelle = Workbooks.Open(Pfad & Datei, False, True)
        Quelle.Sheets().Copy After:=Ziel.Sheets(Ziel.Sheets.Count)

        Ziel.Sheets(Ziel.Sheets.Count).Name = Datei

        Quelle.Close

        Datei = Dir()
    Loop

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Dateien wurden zusammengeführt"

    Set Ziel = Nothing
    Set Quelle = Nothing

End Sub