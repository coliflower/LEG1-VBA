Attribute VB_Name = "LEG1_GitHub_Sync"
Option Explicit

' ============================================================
' LEG1_GitHub_Sync
'
' Synchronisiert:
'   GitHub -> LEG1_Spieler_Abgleich
'
' Dieses Modul selbst wird niemals ersetzt.
'
' Mac + Windows kompatibel
'
' WICHTIG:
' Es werden keine neuen VBA-Module importiert.
' Das bestehende Modul LEG1_Spieler_Abgleich bleibt erhalten.
' Nur dessen Codeinhalt wird aktualisiert.
'
' BESONDERHEIT:
' Der GitHub-Download wird als UTF-8 eingelesen.
' Zusätzlich werden überflüssige CR-Zeichen aus den einzelnen
' Download-Zeilen entfernt, bevor AddFromString verwendet wird.
' ============================================================

Private Const GITHUB_RAW_URL As String = _
    "https://raw.githubusercontent.com/coliflower/LEG1-VBA/main/VBA/LEG1_Spieler_Abgleich.bas"

Private Const TARGET_MODULE As String = _
    "LEG1_Spieler_Abgleich"

Private Const SYNC_MODULE As String = _
    "LEG1_GitHub_Sync"

Private Const DOWNLOAD_SHEET As String = _
    "__LEG1_GitHub_Download"


' ============================================================
' HAUPTPROZEDUR
' ============================================================

Public Sub LEG1_GitHub_Synchronisieren()

    Dim wb As Workbook
    Dim wsDownload As Worksheet

    Dim sourceCode As String
    Dim oldCode As String

    Dim vbProj As Object
    Dim targetComp As Object

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean

    Dim fehlerNummer As Long
    Dim fehlerBeschreibung As String
    Dim schritt As String

    Dim oldCodeGesichert As Boolean
    Dim neuerCodeGeschrieben As Boolean

    On Error GoTo Fehler

    Set wb = ThisWorkbook

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    schritt = "VBA-Projekt öffnen"

    Set vbProj = GetVBProject()

    If vbProj Is Nothing Then
        Err.Raise vbObjectError + 1000, _
                  "LEG1_GitHub_Sync", _
                  "Das VBA-Projekt konnte nicht geöffnet werden."
    End If

    schritt = "LEG1_Spieler_Abgleich-Modul prüfen"

    Set targetComp = GetVBComponent(vbProj, TARGET_MODULE)

    If targetComp Is Nothing Then
        Err.Raise vbObjectError + 1001, _
                  "LEG1_GitHub_Sync", _
                  "Das Modul " & TARGET_MODULE & _
                  " ist im VBA-Projekt nicht vorhanden."
    End If

    If StrComp(targetComp.Name, SYNC_MODULE, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 1002, _
                  "LEG1_GitHub_Sync", _
                  "Sicherheitsfehler: " & _
                  SYNC_MODULE & _
                  " wurde als Zielmodul erkannt."
    End If

    schritt = "Aktuellen LEG1-Code sichern"

    oldCode = CodeModulLesen(targetComp)

    If Len(oldCode) = 0 Then
        Err.Raise vbObjectError + 1003, _
                  "LEG1_GitHub_Sync", _
                  "Der bisherige Code von " & _
                  TARGET_MODULE & _
                  " konnte nicht gelesen werden."
    End If

    oldCodeGesichert = True

    schritt = "GitHub-Datei herunterladen"

    Set wsDownload = DownloadGitHubSource(wb)

    If wsDownload Is Nothing Then
        Err.Raise vbObjectError + 1004, _
                  "LEG1_GitHub_Sync", _
                  "Die GitHub-Datei konnte nicht heruntergeladen werden."
    End If

    schritt = "GitHub-Inhalt aus Download-Tabelle lesen"

    sourceCode = DownloadTabelleZuCode(wsDownload)

    If Len(Trim$(sourceCode)) = 0 Then
        Err.Raise vbObjectError + 1005, _
                  "LEG1_GitHub_Sync", _
                  "Die heruntergeladene GitHub-Datei ist leer."
    End If

    schritt = "GitHub-VBA-Code prüfen"

    If Not BasDateiIstGueltig(sourceCode) Then
        Err.Raise vbObjectError + 1006, _
                  "LEG1_GitHub_Sync", _
                  "Die heruntergeladene Datei sieht nicht wie " & _
                  "ein gültiges VBA-Modul aus."
    End If

    schritt = "Zeichenkodierung prüfen"

    If Not ZeichenkodierungIstKorrekt(sourceCode) Then
        Err.Raise vbObjectError + 1007, _
                  "LEG1_GitHub_Sync", _
                  "Der GitHub-Code wurde beim Download offenbar " & _
                  "mit einer falschen Zeichenkodierung eingelesen." & _
                  vbCrLf & vbCrLf & _
                  "Der vorhandene LEG1-Code wurde deshalb nicht verändert."
    End If

    schritt = "GitHub-Code normalisieren"

    sourceCode = VBAQuelltextNormalisieren(sourceCode)

    If Len(Trim$(sourceCode)) = 0 Then
        Err.Raise vbObjectError + 1008, _
                  "LEG1_GitHub_Sync", _
                  "Der normalisierte GitHub-Code ist leer."
    End If

    schritt = "GitHub-Code in LEG1_Spieler_Abgleich schreiben"

    CodeModulErsetzen targetComp, sourceCode

    neuerCodeGeschrieben = True

    schritt = "Neue LEG1-Version prüfen"

    If targetComp.CodeModule.CountOfLines = 0 Then
        Err.Raise vbObjectError + 1009, _
                  "LEG1_GitHub_Sync", _
                  "Der neue GitHub-Code wurde nicht korrekt " & _
                  "in das Zielmodul geschrieben."
    End If

    schritt = "LEG1-Modulnamen prüfen"

    If StrComp(targetComp.Name, TARGET_MODULE, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1010, _
                  "LEG1_GitHub_Sync", _
                  "Der Name des Zielmoduls hat sich unerwartet geändert."
    End If

    schritt = "LEG1_GitHub_Sync-Schutz prüfen"

    If GetVBComponent(vbProj, SYNC_MODULE) Is Nothing Then
        Err.Raise vbObjectError + 1011, _
                  "LEG1_GitHub_Sync", _
                  "Das Sync-Modul " & _
                  SYNC_MODULE & _
                  " ist nicht mehr vorhanden."
    End If

    schritt = "Temporäre Download-Tabelle entfernen"

    DeleteDownloadSheet wb, wsDownload

    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    MsgBox _
        "LEG1_Spieler_Abgleich wurde erfolgreich von GitHub aktualisiert." & _
        vbCrLf & vbCrLf & _
        "Das vorhandene Modul wurde direkt aktualisiert." & _
        vbCrLf & _
        "Es wurde kein neues Modul erzeugt." & _
        vbCrLf & _
        "LEG1_GitHub_Sync wurde nicht verändert.", _
        vbInformation, _
        "LEG1 GitHub Sync"

    Exit Sub

Fehler:

    fehlerNummer = Err.Number
    fehlerBeschreibung = Err.Description

    On Error Resume Next

    If Not wsDownload Is Nothing Then
        DeleteDownloadSheet wb, wsDownload
    End If

    If neuerCodeGeschrieben And oldCodeGesichert Then
        If Not targetComp Is Nothing Then
            CodeModulErsetzen targetComp, oldCode
        End If
    End If

    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    On Error GoTo 0

    MsgBox _
        "Die GitHub-Synchronisierung wurde abgebrochen." & _
        vbCrLf & vbCrLf & _
        "Schritt:" & vbCrLf & _
        schritt & _
        vbCrLf & vbCrLf & _
        "Fehlernummer: " & CStr(fehlerNummer) & _
        vbCrLf & _
        "Fehlerbeschreibung:" & vbCrLf & _
        fehlerBeschreibung & _
        vbCrLf & vbCrLf & _
        "Der bisherige Code von " & _
        TARGET_MODULE & _
        " wurde wiederhergestellt.", _
        vbExclamation, _
        "LEG1 GitHub Sync"

End Sub


' ============================================================
' CODE EINES VBA-MODULS LESEN
' ============================================================

Private Function CodeModulLesen( _
    ByVal vbComp As Object) As String

    Dim anzahlZeilen As Long

    anzahlZeilen = vbComp.CodeModule.CountOfLines

    If anzahlZeilen <= 0 Then
        CodeModulLesen = vbNullString
        Exit Function
    End If

    CodeModulLesen = _
        vbComp.CodeModule.lines(1, anzahlZeilen)

End Function


' ============================================================
' CODE EINES VBA-MODULS KOMPLETT ERSETZEN
' ============================================================

Private Sub CodeModulErsetzen( _
    ByVal vbComp As Object, _
    ByVal neuerCode As String)

    Dim anzahlZeilen As Long

    On Error GoTo Fehler

    anzahlZeilen = vbComp.CodeModule.CountOfLines

    If anzahlZeilen > 0 Then
        vbComp.CodeModule.DeleteLines 1, anzahlZeilen
    End If

    vbComp.CodeModule.AddFromString neuerCode

    Exit Sub

Fehler:

    Err.Raise Err.Number, _
              "LEG1_GitHub_Sync / CodeModulErsetzen", _
              Err.Description

End Sub


' ============================================================
' GITHUB DOWNLOAD
'
' Der Download wird explizit als UTF-8 angefordert.
'
' WICHTIG:
' TextFileParseType wird NICHT gesetzt, weil genau diese
' Eigenschaft auf Excel/Mac zuvor Fehler 1004 verursacht hat.
' ============================================================

Private Function DownloadGitHubSource( _
    ByVal wb As Workbook) As Worksheet

    Dim ws As Worksheet
    Dim qt As QueryTable

    Dim downloadErrorNumber As Long
    Dim downloadErrorDescription As String

    On Error GoTo Fehler

    On Error Resume Next
    wb.Worksheets(DOWNLOAD_SHEET).Delete
    On Error GoTo Fehler

    Set ws = wb.Worksheets.Add( _
        After:=wb.Worksheets(wb.Worksheets.Count))

    ws.Name = DOWNLOAD_SHEET

    Set qt = ws.QueryTables.Add( _
        Connection:="URL;" & GITHUB_RAW_URL, _
        Destination:=ws.Range("A1"))

    With qt

        .BackgroundQuery = False
        .RefreshStyle = xlOverwriteCells
        .AdjustColumnWidth = False

        ' ----------------------------------------------------
        ' UTF-8 erzwingen.
        ' Falls Excel diese Eigenschaft auf der jeweiligen
        ' Plattform nicht unterstützt, wird der Fehler
        ' kontrolliert zurückgegeben.
        ' ----------------------------------------------------

        On Error Resume Next

        Err.Clear
        .TextFilePlatform = 65001

        downloadErrorNumber = Err.Number
        downloadErrorDescription = Err.Description

        On Error GoTo Fehler

        If downloadErrorNumber <> 0 Then

            Err.Raise downloadErrorNumber, _
                      "LEG1_GitHub_Sync / UTF-8", _
                      "Die Excel-Version konnte den GitHub-Download " & _
                      "nicht als UTF-8 konfigurieren." & _
                      vbCrLf & vbCrLf & _
                      downloadErrorDescription

        End If

        On Error Resume Next

        Err.Clear
        .Refresh

        downloadErrorNumber = Err.Number
        downloadErrorDescription = Err.Description

        On Error GoTo Fehler

        If downloadErrorNumber <> 0 Then

            Err.Raise downloadErrorNumber, _
                      "LEG1_GitHub_Sync / QueryTable.Refresh", _
                      downloadErrorDescription

        End If

    End With

    If Len(Trim$(CStr(ws.Range("A1").Value))) = 0 Then

        Err.Raise vbObjectError + 1100, _
                  "LEG1_GitHub_Sync", _
                  "GitHub wurde erreicht, aber es wurde kein Inhalt " & _
                  "in A1 geladen."

    End If

    qt.Delete

    Set DownloadGitHubSource = ws

    Exit Function

Fehler:

    downloadErrorNumber = Err.Number
    downloadErrorDescription = Err.Description

    On Error Resume Next

    If Not qt Is Nothing Then
        qt.Delete
    End If

    If Not ws Is Nothing Then
        ws.Delete
    End If

    On Error GoTo 0

    Err.Raise downloadErrorNumber, _
              "LEG1_GitHub_Sync / DownloadGitHubSource", _
              downloadErrorDescription

End Function


' ============================================================
' DOWNLOAD-TABELLE -> VBA SOURCECODE
'
' Excel kann beim Einlesen einer Textdatei ein CR am Ende
' einer Zelle belassen. Dieses wird hier entfernt.
'
' Dadurch entstehen keine CRCRLF-Sequenzen mehr.
' ============================================================

Private Function DownloadTabelleZuCode( _
    ByVal ws As Worksheet) As String

    Dim lastRow As Long
    Dim r As Long

    Dim lineText As String
    Dim result As String

    lastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            1).End(xlUp).Row

    If lastRow < 1 Then
        DownloadTabelleZuCode = vbNullString
        Exit Function
    End If

    result = vbNullString

    For r = 1 To lastRow

        lineText = CStr(ws.Cells(r, 1).Value)

        ' ----------------------------------------------------
        ' Alle vorhandenen Zeilenende-Zeichen am Ende
        ' der einzelnen Excel-Zelle entfernen.
        ' ----------------------------------------------------

        Do While Len(lineText) > 0

            If Right$(lineText, 1) = vbCr Or _
               Right$(lineText, 1) = vbLf Then

                lineText = Left$( _
                    lineText, _
                    Len(lineText) - 1)

            Else

                Exit Do

            End If

        Loop

        If r > 1 Then
            result = result & vbCrLf
        End If

        result = result & lineText

    Next r

    DownloadTabelleZuCode = result

End Function


' ============================================================
' VBA-QUELLTEXT NORMALISIEREN
' ============================================================

Private Function VBAQuelltextNormalisieren( _
    ByVal sourceCode As String) As String

    Dim t As String

    t = sourceCode

    ' --------------------------------------------------------
    ' Alle Zeilenenden zunächst vereinheitlichen.
    ' --------------------------------------------------------

    t = Replace(t, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)

    ' --------------------------------------------------------
    ' Danach exakt auf CRLF für das VBA-CodeModule bringen.
    ' --------------------------------------------------------

    t = Replace(t, vbLf, vbCrLf)

    VBAQuelltextNormalisieren = t

End Function


' ============================================================
' VALIDIERUNG DES HERUNTERGELADENEN VBA-CODES
' ============================================================

Private Function BasDateiIstGueltig( _
    ByVal sourceCode As String) As Boolean

    Dim t As String

    t = Trim$(sourceCode)

    If Len(t) = 0 Then
        BasDateiIstGueltig = False
        Exit Function
    End If

    If InStr(1, t, "<html", vbTextCompare) > 0 Then
        BasDateiIstGueltig = False
        Exit Function
    End If

    If InStr(1, t, "<!DOCTYPE", vbTextCompare) > 0 Then
        BasDateiIstGueltig = False
        Exit Function
    End If

    If InStr(1, t, "Option Explicit", vbTextCompare) = 0 Then

        If InStr(1, t, "Sub ", vbTextCompare) = 0 And _
           InStr(1, t, "Function ", vbTextCompare) = 0 Then

            BasDateiIstGueltig = False
            Exit Function

        End If

    End If

    If InStr(1, t, "Public Sub LEG1_Spieler_Abgleich", _
             vbTextCompare) = 0 Then

        BasDateiIstGueltig = False
        Exit Function

    End If

    BasDateiIstGueltig = True

End Function


' ============================================================
' ZEICHENKODIERUNG PRÜFEN
'
' Die GitHub-Datei enthält deutsche Umlaute.
' Insbesondere "über" ist im aktuellen GitHub-Code vorhanden.
'
' Typische UTF-8/ANSI-Fehlinterpretationen werden abgefangen.
' ============================================================

Private Function ZeichenkodierungIstKorrekt( _
    ByVal sourceCode As String) As Boolean

    Dim t As String

    t = sourceCode

    ' --------------------------------------------------------
    ' Bekannte Fehlinterpretationen erkennen.
    ' --------------------------------------------------------

    If InStr(1, t, "√º", vbBinaryCompare) > 0 Then
        ZeichenkodierungIstKorrekt = False
        Exit Function
    End If

    If InStr(1, t, "√?", vbBinaryCompare) > 0 Then
        ZeichenkodierungIstKorrekt = False
        Exit Function
    End If

    If InStr(1, t, "√¶", vbBinaryCompare) > 0 Then
        ZeichenkodierungIstKorrekt = False
        Exit Function
    End If

    If InStr(1, t, "√„", vbBinaryCompare) > 0 Then
        ZeichenkodierungIstKorrekt = False
        Exit Function
    End If

    If InStr(1, t, "√–", vbBinaryCompare) > 0 Then
        ZeichenkodierungIstKorrekt = False
        Exit Function
    End If

    If InStr(1, t, "√œ", vbBinaryCompare) > 0 Then
        ZeichenkodierungIstKorrekt = False
        Exit Function
    End If

    If InStr(1, t, "√Ÿ", vbBinaryCompare) > 0 Then
        ZeichenkodierungIstKorrekt = False
        Exit Function
    End If

    If InStr(1, t, "Ã", vbBinaryCompare) > 0 Then
        ZeichenkodierungIstKorrekt = False
        Exit Function
    End If

    ' --------------------------------------------------------
    ' Die aktuelle GitHub-Datei enthält "über".
    ' Wenn es fehlt, gehen wir sicherheitshalber davon aus,
    ' dass die Kodierung nicht korrekt übernommen wurde.
    ' --------------------------------------------------------

    If InStr(1, t, "über", vbBinaryCompare) = 0 Then
        ZeichenkodierungIstKorrekt = False
        Exit Function
    End If

    ZeichenkodierungIstKorrekt = True

End Function


' ============================================================
' VBA-PROJEKT
' ============================================================

Private Function GetVBProject() As Object

    On Error Resume Next
    Set GetVBProject = ThisWorkbook.VBProject
    On Error GoTo 0

End Function


' ============================================================
' MODUL SUCHEN
' ============================================================

Private Function GetVBComponent( _
    ByVal vbProj As Object, _
    ByVal moduleName As String) As Object

    Dim comp As Object

    On Error Resume Next
    Set comp = vbProj.VBComponents(moduleName)
    On Error GoTo 0

    Set GetVBComponent = comp

End Function


' ============================================================
' TEMPORÄRES DOWNLOAD-BLATT LÖSCHEN
' ============================================================

Private Sub DeleteDownloadSheet( _
    ByVal wb As Workbook, _
    ByVal ws As Worksheet)

    On Error Resume Next

    If Not ws Is Nothing Then

        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True

    End If

    On Error GoTo 0

End Sub

