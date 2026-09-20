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
' DOWNLOAD:
' Die GitHub-API liefert den Inhalt der BAS-Datei als Base64.
' Dadurch wird das UTF-8-Problem des direkten QueryTable-
' Downloads vermieden.
'
' Es wird NICHT verwendet:
'   - TextFilePlatform
'   - Windows API
'   - FileSystemObject
'   - Scripting.Dictionary
'   - WScript
'   - CreateObject
' ============================================================

Private Const GITHUB_API_URL As String = _
    "https://api.github.com/repos/coliflower/LEG1-VBA/contents/VBA/LEG1_Spieler_Abgleich.bas?ref=main"

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

    schritt = "VBA-Projekt oeffnen"

    Set vbProj = GetVBProject()

    If vbProj Is Nothing Then
        Err.Raise vbObjectError + 1000, _
                  "LEG1_GitHub_Sync", _
                  "Das VBA-Projekt konnte nicht geoeffnet werden."
    End If

    schritt = "LEG1_Spieler_Abgleich-Modul pruefen"

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

    schritt = "GitHub-VBA-Code pruefen"

    If Not BasDateiIstGueltig(sourceCode) Then
        Err.Raise vbObjectError + 1006, _
                  "LEG1_GitHub_Sync", _
                  "Die heruntergeladene Datei sieht nicht wie " & _
                  "ein gueltiges VBA-Modul aus."
    End If

    schritt = "GitHub-Code normalisieren"

    sourceCode = VBAQuelltextNormalisieren(sourceCode)

    If Len(Trim$(sourceCode)) = 0 Then
        Err.Raise vbObjectError + 1007, _
                  "LEG1_GitHub_Sync", _
                  "Der normalisierte GitHub-Code ist leer."
    End If

    schritt = "GitHub-Code in LEG1_Spieler_Abgleich schreiben"

    CodeModulErsetzen targetComp, sourceCode

    neuerCodeGeschrieben = True

    schritt = "Neue LEG1-Version pruefen"

    If targetComp.CodeModule.CountOfLines = 0 Then
        Err.Raise vbObjectError + 1008, _
                  "LEG1_GitHub_Sync", _
                  "Der neue GitHub-Code wurde nicht korrekt " & _
                  "in das Zielmodul geschrieben."
    End If

    schritt = "LEG1-Modulnamen pruefen"

    If StrComp(targetComp.Name, TARGET_MODULE, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1009, _
                  "LEG1_GitHub_Sync", _
                  "Der Name des Zielmoduls hat sich unerwartet geaendert."
    End If

    schritt = "LEG1_GitHub_Sync-Schutz pruefen"

    If GetVBComponent(vbProj, SYNC_MODULE) Is Nothing Then
        Err.Raise vbObjectError + 1010, _
                  "LEG1_GitHub_Sync", _
                  "Das Sync-Modul " & _
                  SYNC_MODULE & _
                  " ist nicht mehr vorhanden."
    End If

    schritt = "Temporaere Download-Tabelle entfernen"

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
        "LEG1_GitHub_Sync wurde nicht veraendert.", _
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
        vbComp.CodeModule.Lines(1, anzahlZeilen)

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
' Es wird die GitHub-API aufgerufen.
'
' Die API liefert:
'
' {
'   "name": "...",
'   "content": "BASE64...",
'   ...
' }
'
' Der eigentliche VBA-Code wird daher nicht als UTF-8-Text
' durch Excel transportiert.
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
        Connection:="URL;" & GITHUB_API_URL, _
        Destination:=ws.Range("A1"))

    With qt

        .BackgroundQuery = False
        .RefreshStyle = xlOverwriteCells
        .AdjustColumnWidth = False

        ' Die GitHub-API-Antwort ist JSON. Sie darf von Excel
        ' nicht an Kommas/anderen Trennzeichen in viele Zellen
        ' zerlegt werden. Die Antwort bleibt daher in A1.
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileOtherDelimiter = Chr$(1)

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

    If LetzteBelegteZeile(ws) < 1 Then

        Err.Raise vbObjectError + 1100, _
                  "LEG1_GitHub_Sync", _
                  "GitHub wurde erreicht, aber es wurde kein Inhalt " & _
                  "geladen."

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
' DOWNLOAD-TABELLE -> BASE64 -> VBA SOURCECODE
' ============================================================

Private Function DownloadTabelleZuCode( _
    ByVal ws As Worksheet) As String

    Dim jsonText As String
    Dim base64Text As String

    jsonText = DownloadTabelleAlsText(ws)

    If Len(jsonText) = 0 Then
        DownloadTabelleZuCode = vbNullString
        Exit Function
    End If

    base64Text = GitHubBase64AusJSON(jsonText)

    If Len(base64Text) = 0 Then
        Err.Raise vbObjectError + 1200, _
                  "LEG1_GitHub_Sync", _
                  "Der GitHub-API-Antwort konnte kein " & _
                  "Base64-Dateiinhalt entnommen werden."
    End If

    DownloadTabelleZuCode = Base64UTF8Dekodieren(base64Text)

End Function


' ============================================================
' KOMPLETTE DOWNLOAD-TABELLE ALS TEXT LESEN
' ============================================================

Private Function DownloadTabelleAlsText( _
    ByVal ws As Worksheet) As String

    Dim jsonText As String

    ' Die API-Antwort wird bewusst als ein einziger Textblock
    ' aus A1 gelesen. Dadurch kann Excel keine riesige Schleife
    ' ueber eine versehentlich aufgeteilte JSON-Tabelle erzeugen.
    jsonText = CStr(ws.Range("A1").Value2)

    DownloadTabelleAlsText = jsonText

End Function


' ============================================================
' BASE64-INHALT AUS GITHUB-JSON HOLEN
' ============================================================

Private Function GitHubBase64AusJSON( _
    ByVal jsonText As String) As String

    Dim marker As String
    Dim startPos As Long
    Dim endPos As Long

    Dim contentText As String

    marker = """content"":"""

    startPos = InStr(1, jsonText, marker, vbBinaryCompare)

    If startPos = 0 Then
        GitHubBase64AusJSON = vbNullString
        Exit Function
    End If

    startPos = startPos + Len(marker)

    endPos = InStr(startPos, jsonText, """", vbBinaryCompare)

    If endPos = 0 Then
        GitHubBase64AusJSON = vbNullString
        Exit Function
    End If

    contentText = Mid$( _
        jsonText, _
        startPos, _
        endPos - startPos)

    contentText = Replace(contentText, vbCr, vbNullString)
    contentText = Replace(contentText, vbLf, vbNullString)
    contentText = Replace(contentText, " ", vbNullString)
    contentText = Replace(contentText, vbTab, vbNullString)

    GitHubBase64AusJSON = contentText

End Function


' ============================================================
' BASE64 -> UTF-8 -> VBA STRING
' ============================================================

Private Function Base64UTF8Dekodieren( _
    ByVal base64Text As String) As String

    Dim data() As Byte
    Dim dataLength As Long

    data = Base64Dekodieren(base64Text, dataLength)

    If dataLength <= 0 Then
        Base64UTF8Dekodieren = vbNullString
        Exit Function
    End If

    Base64UTF8Dekodieren = UTF8BytesZuString(data, dataLength)

End Function


' ============================================================
' BASE64 DEKODIEREN
' ============================================================

Private Function Base64Dekodieren( _
    ByVal base64Text As String, _
    ByRef dataLength As Long) As Byte()

    Dim output() As Byte

    Dim i As Long
    Dim n As Long

    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long

    Dim ch As String

    ReDim output(0 To 0)

    dataLength = 0

    base64Text = Replace(base64Text, vbCr, vbNullString)
    base64Text = Replace(base64Text, vbLf, vbNullString)
    base64Text = Replace(base64Text, " ", vbNullString)
    base64Text = Replace(base64Text, vbTab, vbNullString)

    n = Len(base64Text)

    If n = 0 Then
        Base64Dekodieren = output
        Exit Function
    End If

    ReDim output(0 To ((n \ 4) * 3) + 2)

    For i = 1 To n Step 4

        a = Base64Wert(Mid$(base64Text, i, 1))

        If i + 1 <= n Then
            b = Base64Wert(Mid$(base64Text, i + 1, 1))
        Else
            b = 0
        End If

        If i + 2 <= n Then
            ch = Mid$(base64Text, i + 2, 1)
            If ch = "=" Then
                c = 0
            Else
                c = Base64Wert(ch)
            End If
        Else
            c = 0
        End If

        If i + 3 <= n Then
            ch = Mid$(base64Text, i + 3, 1)
            If ch = "=" Then
                d = 0
            Else
                d = Base64Wert(ch)
            End If
        Else
            d = 0
        End If

        output(dataLength) = _
            CByte((a * 64) + (b \ 4))

        dataLength = dataLength + 1

        If i + 2 <= n Then
            If Mid$(base64Text, i + 2, 1) <> "=" Then
                output(dataLength) = _
                    CByte(((b And 3) * 64) + (c \ 16))
                dataLength = dataLength + 1
            End If
        End If

        If i + 3 <= n Then
            If Mid$(base64Text, i + 3, 1) <> "=" Then
                output(dataLength) = _
                    CByte(((c And 15) * 16) + d)
                dataLength = dataLength + 1
            End If
        End If

    Next i

    If dataLength > 0 Then
        ReDim Preserve output(0 To dataLength - 1)
    Else
        ReDim output(0 To 0)
    End If

    Base64Dekodieren = output

End Function


' ============================================================
' BASE64 ZEICHEN -> WERT
' ============================================================

Private Function Base64Wert( _
    ByVal ch As String) As Long

    Dim n As Long

    If Len(ch) = 0 Then
        Base64Wert = 0
        Exit Function
    End If

    n = AscW(ch)

    Select Case n

        Case 65 To 90
            Base64Wert = n - 65

        Case 97 To 122
            Base64Wert = n - 97 + 26

        Case 48 To 57
            Base64Wert = n - 48 + 52

        Case 43
            Base64Wert = 62

        Case 47
            Base64Wert = 63

        Case Else
            Err.Raise vbObjectError + 1201, _
                      "LEG1_GitHub_Sync", _
                      "Ungueltiges Base64-Zeichen: " & ch

    End Select

End Function


' ============================================================
' UTF-8 BYTE ARRAY -> VBA UNICODE STRING
' ============================================================

Private Function UTF8BytesZuString( _
    ByRef data() As Byte, _
    ByVal dataLength As Long) As String

    Dim i As Long

    Dim b1 As Long
    Dim b2 As Long
    Dim b3 As Long
    Dim b4 As Long

    Dim codePoint As Long

    Dim result As String

    result = vbNullString

    i = 0

    Do While i < dataLength

        b1 = data(i)

        If b1 < 128 Then

            result = result & ChrW$(b1)
            i = i + 1

        ElseIf b1 >= 192 And b1 <= 223 Then

            If i + 1 >= dataLength Then Exit Do

            b2 = data(i + 1)

            codePoint = _
                ((b1 And 31) * 64) + _
                (b2 And 63)

            result = result & ChrW$(codePoint)

            i = i + 2

        ElseIf b1 >= 224 And b1 <= 239 Then

            If i + 2 >= dataLength Then Exit Do

            b2 = data(i + 1)
            b3 = data(i + 2)

            codePoint = _
                ((b1 And 15) * 4096) + _
                ((b2 And 63) * 64) + _
                (b3 And 63)

            result = result & ChrW$(codePoint)

            i = i + 3

        ElseIf b1 >= 240 And b1 <= 247 Then

            If i + 3 >= dataLength Then Exit Do

            b2 = data(i + 1)
            b3 = data(i + 2)
            b4 = data(i + 3)

            codePoint = _
                ((b1 And 7) * 262144) + _
                ((b2 And 63) * 4096) + _
                ((b3 And 63) * 64) + _
                (b4 And 63)

            codePoint = codePoint - &H10000

            result = result & _
                     ChrW$(&HD800 Or (codePoint \ 1024))

            result = result & _
                     ChrW$(&HDC00 Or (codePoint And 1023))

            i = i + 4

        Else

            Err.Raise vbObjectError + 1202, _
                      "LEG1_GitHub_Sync", _
                      "Ungueltige UTF-8-Daten im GitHub-Code."

        End If

    Loop

    UTF8BytesZuString = result

End Function


' ============================================================
' VBA-QUELLTEXT NORMALISIEREN
' ============================================================

Private Function VBAQuelltextNormalisieren( _
    ByVal sourceCode As String) As String

    Dim t As String

    t = sourceCode

    t = Replace(t, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
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
' LETZTE BELEGTE ZEILE
' ============================================================

Private Function LetzteBelegteZeile( _
    ByVal ws As Worksheet) As Long

    Dim letzteZelle As Range

    On Error Resume Next

    Set letzteZelle = ws.Cells.Find( _
        What:="*", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)

    On Error GoTo 0

    If letzteZelle Is Nothing Then
        LetzteBelegteZeile = 0
    Else
        LetzteBelegteZeile = letzteZelle.Row
    End If

End Function


' ============================================================
' LETZTE BELEGTE SPALTE
' ============================================================

Private Function LetzteBelegteSpalte( _
    ByVal ws As Worksheet) As Long

    Dim letzteZelle As Range

    On Error Resume Next

    Set letzteZelle = ws.Cells.Find( _
        What:="*", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)

    On Error GoTo 0

    If letzteZelle Is Nothing Then
        LetzteBelegteSpalte = 0
    Else
        LetzteBelegteSpalte = letzteZelle.Column
    End If

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
' TEMPORAERES DOWNLOAD-BLATT LOESCHEN
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
