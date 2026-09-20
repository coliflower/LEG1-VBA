Option Explicit

' ============================================================
' KONSTANTEN
' ============================================================

Private Const DashboardBlattName As String = "dashboard"
Private Const UpdateBlattName As String = "update"

Private Const CHESTS_DATEI As String = "TB__Chests.xlsx"

' ------------------------------------------------------------
' Dashboard - definierte Namen / Name Manager
' ------------------------------------------------------------

Private Const NAME_LEG1_SPIELER As String = "LEG1_spieler"
Private Const NAME_LEG1_RANG As String = "LEG1_rang"
Private Const NAME_LEG1_H As String = "LEG1_H"
Private Const NAME_LEG1_MACHT As String = "LEG1_macht"
Private Const NAME_LEG1_T9 As String = "LEG1_T9"
Private Const NAME_LEG1_G As String = "LEG1_G"
Private Const NAME_LEG1_M As String = "LEG1_M"
Private Const NAME_LEG1_S As String = "LEG1_S"
Private Const NAME_LEG1_E As String = "LEG1_E"
Private Const NAME_LEG1_JOINED As String = "LEG1_joined"
Private Const NAME_LEG1_ALIASE As String = "LEG1_aliasse"
Private Const NAME_LEG1_INTERNNEU As String = "LEG1_InternNeu"
Private Const NAME_LEG1_INTERNSORTIERUNG As String = "LEG1_InternSortierung"
Private Const NAME_LEG1_T9_SEIT As String = "LEG1_T9_seit"

Private Const LEG1_Basisdaten_Kopf As String = "LEG1_basisdaten"

Private Const CARTER_NOK_Kopf As String = "Carter_NOK"
Private Const CHESTS_NOK_Kopf As String = "Chests_NOK"

Private Const NAME_DASHBOARD_SUMMEN As String = "Dashboard_1_summen"
Private Const NAME_DASHBOARD_SCHWELLE1 As String = "Dashboard_2_schwelle_1"
Private Const NAME_DASHBOARD_SCHWELLE2 As String = "Dashboard_3_schwelle_2"
Private Const NAME_DASHBOARD_BEZUEGE As String = "Dashboard_4_bezüge"
Private Const NAME_DASHBOARD_SPIELER1 As String = "Dashboard_spieler_1"

Private Const UPDATE_SPIELER_SPALTE As Long = 1
Private Const UPDATE_RANG_SPALTE As Long = 2
Private Const UPDATE_ALIASE_SPALTE As Long = 3
Private Const UPDATE_H_SPALTE As Long = 4
Private Const UPDATE_MACHT_SPALTE As Long = 5
Private Const UPDATE_G_SPALTE As Long = 6
Private Const UPDATE_M_SPALTE As Long = 7
Private Const UPDATE_S_SPALTE As Long = 8
Private Const UPDATE_E_SPALTE As Long = 9
Private Const UPDATE_JOINED_SPALTE As Long = 12

Private Const CHESTS_SPIELER_SPALTE As Long = 3
Private Const CHESTS_WERT_SPALTE As Long = 5
Private Const CHESTS_KOPFZEILE As Long = 3
Private Const CHESTS_DATENSTART As Long = 3

' ============================================================
' HAUPTPROZEDUR
' ============================================================

Public Sub LEG1_Spieler_Abgleich()

    Dim wsDash As Worksheet
    Dim wsUpdate As Worksheet
    Dim wsLog As Worksheet

    Dim cSpieler As Long
    Dim cRang As Long
    Dim cH As Long
    Dim cMacht As Long
    Dim cT9 As Long
    Dim cG As Long
    Dim cM As Long
    Dim cS As Long
    Dim cE As Long
    Dim cJoined As Long
    Dim cAliasse As Long
    Dim cInternNeu As Long
    Dim cInternSortierung As Long
    Dim cT9Seit As Long

    Dim cBasisdaten As Long
    Dim cLetzterNOK As Long

    Dim zeileSummen As Long
    Dim zeileSchwelle1 As Long
    Dim zeileSchwelle2 As Long
    Dim zeileBezuege As Long
    Dim zeileSpieler1 As Long

    Dim letzteZeileDash As Long
    Dim letzteZeileUpdate As Long

    Dim neueSpieler As Long
    Dim wiederAktiv As Long
    Dim inaktiveSpieler As Long
    Dim datenGeaendert As Long

    Dim i As Long
    Dim spieler As String
    Dim key As String
    Dim dashZeile As Long

    Dim dashboardSpieler As Collection
    Dim updateSpieler As Collection
    Dim verarbeitetSpieler As Collection

    Dim neueZeilen As Collection

    Dim alterCalc As XlCalculation
    Dim alterScreenUpdating As Boolean
    Dim alterEnableEvents As Boolean
    Dim alterDisplayAlerts As Boolean

    Dim errNum As Long
    Dim errDesc As String
    Dim schritt As String

    On Error GoTo Fehler

    alterCalc = Application.Calculation
    alterScreenUpdating = Application.ScreenUpdating
    alterEnableEvents = Application.EnableEvents
    alterDisplayAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    schritt = "Arbeitsblätter setzen"

    Set wsDash = ThisWorkbook.Worksheets(DashboardBlattName)
    Set wsUpdate = ThisWorkbook.Worksheets(UpdateBlattName)

    Set wsLog = LogBlattErstellen()

    schritt = "Dashboard-Struktur über Namensmanager ermitteln"

    cSpieler = DashboardSpalte(wsDash, NAME_LEG1_SPIELER)
    cRang = DashboardSpalte(wsDash, NAME_LEG1_RANG)
    cH = DashboardSpalte(wsDash, NAME_LEG1_H)
    cMacht = DashboardSpalte(wsDash, NAME_LEG1_MACHT)
    cT9 = DashboardSpalte(wsDash, NAME_LEG1_T9)
    cG = DashboardSpalte(wsDash, NAME_LEG1_G)
    cM = DashboardSpalte(wsDash, NAME_LEG1_M)
    cS = DashboardSpalte(wsDash, NAME_LEG1_S)
    cE = DashboardSpalte(wsDash, NAME_LEG1_E)
    cJoined = DashboardSpalte(wsDash, NAME_LEG1_JOINED)
    cAliasse = DashboardSpalte(wsDash, NAME_LEG1_ALIASE)
    cInternNeu = DashboardSpalte(wsDash, NAME_LEG1_INTERNNEU)
    cInternSortierung = DashboardSpalte(wsDash, NAME_LEG1_INTERNSORTIERUNG)
    cT9Seit = DashboardSpalte(wsDash, NAME_LEG1_T9_SEIT)

    cBasisdaten = DashboardSpalte(wsDash, LEG1_Basisdaten_Kopf)

    cLetzterNOK = LetzterNOKMarker(wsDash)

    zeileSummen = DashboardZeileErmitteln(wsDash, NAME_DASHBOARD_SUMMEN)
    zeileSchwelle1 = DashboardZeileErmitteln(wsDash, NAME_DASHBOARD_SCHWELLE1)
    zeileSchwelle2 = DashboardZeileErmitteln(wsDash, NAME_DASHBOARD_SCHWELLE2)
    zeileBezuege = DashboardZeileErmitteln(wsDash, NAME_DASHBOARD_BEZUEGE)
    zeileSpieler1 = DashboardZeileErmitteln(wsDash, NAME_DASHBOARD_SPIELER1)

    If cSpieler = 0 Then Err.Raise 1004, , "Der definierte Name '" & NAME_LEG1_SPIELER & "' wurde nicht gefunden."
    If cRang = 0 Then Err.Raise 1004, , "Der definierte Name '" & NAME_LEG1_RANG & "' wurde nicht gefunden."
    If cH = 0 Then Err.Raise 1004, , "Der definierte Name '" & NAME_LEG1_H & "' wurde nicht gefunden."
    If cMacht = 0 Then Err.Raise 1004, , "Der definierte Name '" & NAME_LEG1_MACHT & "' wurde nicht gefunden."
    If cT9 = 0 Then Err.Raise 1004, , "Der definierte Name '" & NAME_LEG1_T9 & "' wurde nicht gefunden."
    If cG = 0 Then Err.Raise 1004, , "Der definierte Name '" & NAME_LEG1_G & "' wurde nicht gefunden."
    If cM = 0 Then Err.Raise 1004, , "Der definierte Name '" & NAME_LEG1_M & "' wurde nicht gefunden."
    If cS = 0 Then Err.Raise 1004, , "Der definierte Name '" & NAME_LEG1_S & "' wurde nicht gefunden."
    If cE = 0 Then Err.Raise 1004, , "Der definierte Name '" & NAME_LEG1_E & "' wurde nicht gefunden."
    If cJoined = 0 Then Err.Raise 1004, , "Der definierte Name '" & NAME_LEG1_JOINED & "' wurde nicht gefunden."
    If cAliasse = 0 Then Err.Raise 1004, , "Der definierte Name '" & NAME_LEG1_ALIASE & "' wurde nicht gefunden."
    If cInternNeu = 0 Then Err.Raise 1004, , "Der definierte Name '" & NAME_LEG1_INTERNNEU & "' wurde nicht gefunden."
    If cInternSortierung = 0 Then Err.Raise 1004, , "Der definierte Name '" & NAME_LEG1_INTERNSORTIERUNG & "' wurde nicht gefunden."
    If cT9Seit = 0 Then Err.Raise 1004, , "Der definierte Name '" & NAME_LEG1_T9_SEIT & "' wurde nicht gefunden."
    If cBasisdaten = 0 Then Err.Raise 1004, , "Der definierte Name '" & LEG1_Basisdaten_Kopf & "' wurde nicht gefunden."
    If cLetzterNOK = 0 Then Err.Raise 1004, , "Es wurde kein gültiger _NOK-Endmarker gefunden."
    If zeileSummen = 0 Then Err.Raise 1004, , "Der definierte Name '" & NAME_DASHBOARD_SUMMEN & "' wurde nicht gefunden."
    If zeileSchwelle1 = 0 Then Err.Raise 1004, , "Der definierte Name '" & NAME_DASHBOARD_SCHWELLE1 & "' wurde nicht gefunden."
    If zeileSchwelle2 = 0 Then Err.Raise 1004, , "Der definierte Name '" & NAME_DASHBOARD_SCHWELLE2 & "' wurde nicht gefunden."
    If zeileBezuege = 0 Then Err.Raise 1004, , "Der definierte Name '" & NAME_DASHBOARD_BEZUEGE & "' wurde nicht gefunden."
    If zeileSpieler1 = 0 Then Err.Raise 1004, , "Der definierte Name '" & NAME_DASHBOARD_SPIELER1 & "' wurde nicht gefunden."

    schritt = "Dashboard-Struktur prüfen"

    DashboardStrukturPruefen _
        wsDash, _
        cSpieler, _
        cRang, _
        cH, _
        cMacht, _
        cT9, _
        cG, _
        cM, _
        cS, _
        cE, _
        cJoined, _
        cAliasse, _
        cInternNeu, _
        cInternSortierung, _
        cT9Seit, _
        cBasisdaten, _
        cLetzterNOK, _
        zeileSummen, _
        zeileSchwelle1, _
        zeileSchwelle2, _
        zeileBezuege, _
        zeileSpieler1

    schritt = "Sicherung erstellen"
    SicherungErstellen wsLog

    schritt = "Autofilter zurücksetzen"
    AutofilterZuruecksetzen wsDash

    schritt = "Letzte Zeilen bestimmen"

    letzteZeileDash = LetzteSpielerZeile( _
        wsDash, _
        cSpieler, _
        zeileSpieler1)

    letzteZeileUpdate = LetzteSpielerZeile( _
        wsUpdate, _
        UPDATE_SPIELER_SPALTE, _
        2)

    NeueSpielerFarbeZuruecksetzen _
        wsDash, _
        cInternNeu, _
        cSpieler, _
        cLetzterNOK, _
        letzteZeileDash, _
        zeileSpieler1

    Set neueZeilen = New Collection

    schritt = "Sortierung prüfen"

    If SortierungErforderlich( _
            wsDash, _
            cInternSortierung, _
            zeileSpieler1, _
            letzteZeileDash) Then

        DashboardSortieren _
            wsDash, _
            cSpieler, _
            cBasisdaten, _
            letzteZeileDash, _
            zeileSpieler1, _
            cLetzterNOK

        SortierungskennzeichenLoeschen _
            wsDash, _
            cInternSortierung, _
            zeileSpieler1, _
            letzteZeileDash

    End If

    schritt = "Update-Daten prüfen"

    UpdateDatenPruefen _
        wsUpdate, _
        letzteZeileUpdate

    schritt = "Spielerdaten einlesen"

    Set dashboardSpieler = New Collection
    Set updateSpieler = New Collection
    Set verarbeitetSpieler = New Collection

    For i = zeileSpieler1 To letzteZeileDash

        spieler = SichererText( _
            wsDash.Cells(i, cSpieler).Value)

        If Len(spieler) > 0 Then

            key = SpielerKey(spieler)

            If Not CollectionKeyExistiert( _
                    dashboardSpieler, _
                    key) Then

                dashboardSpieler.Add i, key

            End If

        End If

    Next i

    For i = 2 To letzteZeileUpdate

        spieler = SichererText( _
            wsUpdate.Cells( _
                i, _
                UPDATE_SPIELER_SPALTE).Value)

        If Len(spieler) = 0 Then _
            GoTo NaechsterUpdateSpielerEinlesen

        key = SpielerKey(spieler)

        If CollectionKeyExistiert( _
                updateSpieler, _
                key) Then

            LogEintrag _
                wsLog, _
                "DUPLIKAT UPDATE", _
                spieler, _
                "Spieler kommt mehrfach im Update vor."

            GoTo NaechsterUpdateSpielerEinlesen

        End If

        updateSpieler.Add i, key

NaechsterUpdateSpielerEinlesen:

    Next i

    schritt = "Dashboard-Duplikate bereinigen"

    DashboardDuplikateBereinigen _
        wsDash, _
        wsLog, _
        cSpieler, _
        cBasisdaten, _
        letzteZeileDash, _
        zeileSpieler1

    Set dashboardSpieler = New Collection

    For i = zeileSpieler1 To letzteZeileDash

        spieler = SichererText( _
            wsDash.Cells(i, cSpieler).Value)

        If Len(spieler) > 0 Then

            key = SpielerKey(spieler)

            If Not CollectionKeyExistiert( _
                    dashboardSpieler, _
                    key) Then

                dashboardSpieler.Add i, key

            End If

        End If

    Next i

    schritt = "Bestehende Spieler aktualisieren"

    For i = 2 To letzteZeileUpdate

        spieler = SichererText( _
            wsUpdate.Cells( _
                i, _
                UPDATE_SPIELER_SPALTE).Value)

        If Len(spieler) = 0 Then _
            GoTo NaechsterUpdateSpieler

        key = SpielerKey(spieler)

        If CollectionKeyExistiert( _
                dashboardSpieler, _
                key) Then

            dashZeile = CLng( _
                CollectionWert( _
                    dashboardSpieler, _
                    key))

            If Not CollectionKeyExistiert( _
                    verarbeitetSpieler, _
                    key) Then

                verarbeitetSpieler.Add _
                    dashZeile, _
                    key

            End If

            If Not IstAktiv( _
                    wsDash.Cells( _
                        dashZeile, _
                        cBasisdaten).Value) Then

                wsDash.Cells( _
                    dashZeile, _
                    cBasisdaten).Value = 1

                wiederAktiv = wiederAktiv + 1

                LogEintrag _
                    wsLog, _
                    "WIEDER AKTIV", _
                    spieler, _
                    "Spieler ist wieder im aktuellen Update vorhanden."

            End If

            If DatenAenderungenErmitteln( _
                    wsDash, _
                    wsUpdate, _
                    dashZeile, _
                    i, _
                    cSpieler, _
                    cRang, _
                    cH, _
                    cMacht, _
                    cG, _
                    cM, _
                    cS, _
                    cE, _
                    cJoined, _
                    cAliasse) Then

                DatenUebernehmen _
                    wsDash, _
                    wsUpdate, _
                    dashZeile, _
                    i, _
                    cSpieler, _
                    cRang, _
                    cH, _
                    cMacht, _
                    cG, _
                    cM, _
                    cS, _
                    cE, _
                    cJoined, _
                    cAliasse

                datenGeaendert = datenGeaendert + 1

            End If

            ZeileAutomatischeFarbe _
                wsDash, _
                dashZeile, _
                cSpieler, _
                cBasisdaten, _
                cLetzterNOK

        End If

NaechsterUpdateSpieler:

    Next i

    schritt = "Neue Spieler anlegen"

    For i = 2 To letzteZeileUpdate

        spieler = SichererText( _
            wsUpdate.Cells( _
                i, _
                UPDATE_SPIELER_SPALTE).Value)

        If Len(spieler) = 0 Then _
            GoTo NaechsterUpdateSpielerNeu

        key = SpielerKey(spieler)

        If Not CollectionKeyExistiert( _
                dashboardSpieler, _
                key) Then

            letzteZeileDash = letzteZeileDash + 1
            neueSpieler = neueSpieler + 1

            DatenUebernehmen _
                wsDash, _
                wsUpdate, _
                letzteZeileDash, _
                i, _
                cSpieler, _
                cRang, _
                cH, _
                cMacht, _
                cG, _
                cM, _
                cS, _
                cE, _
                cJoined, _
                cAliasse

            wsDash.Cells( _
                letzteZeileDash, _
                cBasisdaten).Value = 1

            wsDash.Cells( _
                letzteZeileDash, _
                cInternNeu).Value = GetUTCNow()

            wsDash.Cells( _
                letzteZeileDash, _
                cInternNeu).NumberFormat = _
                "dd.mm.yyyy hh:mm"

            wsDash.Cells( _
                letzteZeileDash, _
                cInternSortierung).Value = 1

            neueZeilen.Add _
                letzteZeileDash, _
                CStr(letzteZeileDash)

            dashboardSpieler.Add _
                letzteZeileDash, _
                key

            NeueSpielerFarbeSetzen _
                wsDash, _
                letzteZeileDash, _
                cSpieler, _
                cLetzterNOK

            LogEintrag _
                wsLog, _
                "NEUER SPIELER", _
                spieler, _
                "Spieler wurde neu angelegt."

        End If

NaechsterUpdateSpielerNeu:

    Next i

    schritt = "Fehlende Spieler deaktivieren"

    For i = zeileSpieler1 To letzteZeileDash

        spieler = SichererText( _
            wsDash.Cells(i, cSpieler).Value)

        If Len(spieler) > 0 Then

            key = SpielerKey(spieler)

            If Not CollectionKeyExistiert( _
                    verarbeitetSpieler, _
                    key) Then

                If Not CollectionKeyExistiert( _
                        updateSpieler, _
                        key) Then

                    If IstAktiv( _
                            wsDash.Cells( _
                                i, _
                                cBasisdaten).Value) Then

                        wsDash.Cells( _
                            i, _
                            cBasisdaten).Value = 0

                        inaktiveSpieler = _
                            inaktiveSpieler + 1

                        LogEintrag _
                            wsLog, _
                            "INAKTIV", _
                            spieler, _
                            "Spieler ist im aktuellen Update nicht vorhanden."

                    End If

                End If

            End If

        End If

    Next i

    schritt = "Aktive Spieler zählen"

    AktiveSpielerZaehlenUndAnzeigen _
        wsDash, _
        cBasisdaten, _
        zeileSpieler1, _
        letzteZeileDash

    schritt = "T9 berechnen"

    T9AlleSpielerNeuBerechnen _
        wsDash, _
        cG, _
        cS, _
        cT9, _
        zeileSpieler1, _
        letzteZeileDash

    schritt = "T9_seit aktualisieren"

    For i = zeileSpieler1 To letzteZeileDash

        T9SeitAktualisieren _
            wsDash, _
            i, _
            cT9, _
            cT9Seit

    Next i

    schritt = "Carter-Bereich prüfen"

    CarterBereichPruefen _
        wsDash, _
        cBasisdaten, _
        zeileSpieler1, _
        letzteZeileDash

    schritt = "Chests-Spalten und Chests-Daten prüfen"

    ChestsSpaltenPruefen _
        wsDash, _
        cBasisdaten, _
        zeileSummen, _
        zeileSchwelle2, _
        zeileBezuege, _
        cSpieler, _
        zeileSpieler1, _
        letzteZeileDash, _
        wsLog

    schritt = "Neue TB__Bereiche automatisch erkennen"

    NeueTBBereicheAutomatischAnlegen _
        wsDash, _
        cBasisdaten, _
        zeileSummen, _
        zeileSchwelle1, _
        zeileSchwelle2, _
        zeileBezuege, _
        cSpieler, _
        cT9, _
        zeileSpieler1, _
        letzteZeileDash, _
        wsLog

    cLetzterNOK = LetzterNOKMarker(wsDash)

    schritt = "Formatierung der Bereichsspalten setzen"

    TB_BereichsformatierungSetzen _
        wsDash, _
        cBasisdaten, _
        zeileSchwelle1, _
        zeileSchwelle2, _
        zeileBezuege

    schritt = "Datenspalten der Bereiche gruppieren"

    BereichsDatenspaltenGruppieren _
        wsDash, _
        cBasisdaten

    schritt = "Endgültige Schriftfarben setzen"

    EndgueltigeSchriftfarbenSetzen _
        wsDash, _
        cBasisdaten, _
        letzteZeileDash, _
        neueZeilen, _
        cSpieler, _
        cLetzterNOK, _
        zeileSpieler1

    wsDash.Columns(cInternNeu).Hidden = True
    wsDash.Columns(cInternSortierung).Hidden = True
    wsDash.Columns(cT9Seit).Hidden = True

    schritt = "Berechnung"

    Application.Calculate

    Application.Calculation = alterCalc
    Application.ScreenUpdating = alterScreenUpdating
    Application.EnableEvents = alterEnableEvents
    Application.DisplayAlerts = alterDisplayAlerts

    MsgBox _
        "LEG1 Spieler-Abgleich abgeschlossen." & _
        vbCrLf & vbCrLf & _
        "Neue Spieler: " & CStr(neueSpieler) & vbCrLf & _
        "Wieder aktiv: " & CStr(wiederAktiv) & vbCrLf & _
        "Inaktiv: " & CStr(inaktiveSpieler) & vbCrLf & _
        "Daten geändert: " & CStr(datenGeaendert)

    Exit Sub

Fehler:

    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next

    Application.Calculation = alterCalc
    Application.ScreenUpdating = alterScreenUpdating
    Application.EnableEvents = alterEnableEvents
    Application.DisplayAlerts = alterDisplayAlerts

    LogEintrag _
        wsLog, _
        "FEHLER", _
        "", _
        "Schritt: " & schritt & _
        " | Fehler " & CStr(errNum) & _
        ": " & errDesc

    MsgBox _
        "Fehler beim LEG1 Spieler-Abgleich." & _
        vbCrLf & vbCrLf & _
        "Schritt: " & schritt & _
        vbCrLf & _
        "Fehler " & CStr(errNum) & _
        ": " & errDesc, _
        vbCritical

End Sub

' ============================================================
' COLLECTION-HILFSFUNKTIONEN
' ============================================================

Private Function CollectionKeyExistiert( _
    ByVal col As Collection, _
    ByVal key As String) As Boolean

    Dim dummy As Variant

    CollectionKeyExistiert = False

    If col Is Nothing Then Exit Function

    On Error GoTo NichtVorhanden

    dummy = col.item(key)

    CollectionKeyExistiert = True

    Exit Function

NichtVorhanden:

    CollectionKeyExistiert = False

End Function

Private Function CollectionWert( _
    ByVal col As Collection, _
    ByVal key As String) As Variant

    If col Is Nothing Then

        CollectionWert = Empty
        Exit Function

    End If

    On Error GoTo NichtVorhanden

    CollectionWert = col.item(key)

    Exit Function

NichtVorhanden:

    CollectionWert = Empty

End Function

' ============================================================
' SICHERER TEXT
' ============================================================

Private Function SichererText(ByVal wert As Variant) As String

    If IsError(wert) Then

        SichererText = ""

    ElseIf IsNull(wert) Then

        SichererText = ""

    ElseIf IsEmpty(wert) Then

        SichererText = ""

    Else

        SichererText = Trim$(CStr(wert))

    End If

End Function

' ============================================================
' NAMENSMANAGER - RANGE ERMITTELN
' ============================================================

Private Function NameManagerRangeErmitteln( _
    ByVal ws As Worksheet, _
    ByVal nameText As String) As Range

    Dim nm As name
    Dim rng As Range

    Set NameManagerRangeErmitteln = Nothing

    On Error Resume Next

    Set nm = ThisWorkbook.Names(nameText)

    If nm Is Nothing Then
        Set nm = ws.Names(nameText)
    End If

    On Error GoTo 0

    If nm Is Nothing Then Exit Function

    On Error Resume Next

    Set rng = nm.RefersToRange

    On Error GoTo 0

    If rng Is Nothing Then Exit Function
    If Not rng.Parent Is ws Then Exit Function

    Set NameManagerRangeErmitteln = rng

End Function

Private Function DashboardSpalte( _
    ByVal ws As Worksheet, _
    ByVal nameText As String) As Long

    Dim rng As Range

    DashboardSpalte = 0

    Set rng = NameManagerRangeErmitteln(ws, nameText)

    If rng Is Nothing Then Exit Function

    DashboardSpalte = rng.Column

End Function

Private Function DashboardZeileErmitteln( _
    ByVal ws As Worksheet, _
    ByVal nameText As String) As Long

    Dim rng As Range

    DashboardZeileErmitteln = 0

    Set rng = NameManagerRangeErmitteln(ws, nameText)

    If rng Is Nothing Then Exit Function

    DashboardZeileErmitteln = rng.Row

End Function

' ============================================================
' _NOK-MARKER ERKENNEN
' ============================================================

Private Function IstNOKMarkerName( _
    ByVal nameText As String) As Boolean

    Dim kurzerName As String

    IstNOKMarkerName = False

    kurzerName = nameText

    If InStr(1, kurzerName, "!", vbTextCompare) > 0 Then

        kurzerName = Mid$( _
            kurzerName, _
            InStrRev(kurzerName, "!") + 1)

    End If

    If Left$(kurzerName, 1) = "'" Then

        kurzerName = Replace( _
            kurzerName, _
            "'", _
            "")

    End If

    If Len(kurzerName) < 4 Then Exit Function

    IstNOKMarkerName = _
        (Right$(kurzerName, 4) = "_NOK")

End Function

Private Function LetzterNOKMarker( _
    ByVal ws As Worksheet) As Long

    Dim nm As name
    Dim rng As Range
    Dim nameText As String
    Dim maxSpalte As Long

    LetzterNOKMarker = 0
    maxSpalte = 0

    On Error Resume Next

    For Each nm In ThisWorkbook.Names

        nameText = nm.name
        If IstNOKMarkerName(nameText) Then

            Set rng = Nothing
            Set rng = nm.RefersToRange

            If Not rng Is Nothing Then

                If rng.Parent Is ws Then

                    If rng.Column > maxSpalte Then
                        maxSpalte = rng.Column
                    End If

                End If

            End If

        End If

    Next nm

    On Error GoTo 0

    LetzterNOKMarker = maxSpalte

End Function

' ============================================================
' VORHERIGEN _NOK-MARKER ERMITTELN
' ============================================================

Private Function VorherigerNOKMarker( _
    ByVal ws As Worksheet, _
    ByVal zielSpalte As Long, _
    ByVal untergrenze As Long) As Long

    Dim nm As name
    Dim rng As Range
    Dim nameText As String
    Dim besteSpalte As Long

    VorherigerNOKMarker = 0
    besteSpalte = untergrenze

    On Error Resume Next

    For Each nm In ThisWorkbook.Names

        nameText = nm.name

        If IstNOKMarkerName(nameText) Then

            Set rng = Nothing
            Set rng = nm.RefersToRange

            If Not rng Is Nothing Then

                If rng.Parent Is ws Then

                    If rng.Column < zielSpalte Then

                        If rng.Column >= untergrenze Then

                            If rng.Column > besteSpalte Then
                                besteSpalte = rng.Column
                            End If

                        End If
                    End If

                End If

            End If

        End If

    Next nm

    On Error GoTo 0

    If besteSpalte > untergrenze Then
        VorherigerNOKMarker = besteSpalte
    Else
        VorherigerNOKMarker = untergrenze
    End If

End Function

' ============================================================
' ALLE _NOK-MARKER DYNAMISCH ERMITTELN
' ============================================================

Private Function NOKMarkerSpaltenErmitteln( _
    ByVal ws As Worksheet, _
    ByRef markerSpalten() As Long) As Long

    Dim nm As name
    Dim rng As Range
    Dim nameText As String
    Dim anzahl As Long
    Dim i As Long
    Dim j As Long
    Dim temp As Long

    ReDim markerSpalten(1 To 1)

    anzahl = 0

    On Error Resume Next

    For Each nm In ThisWorkbook.Names

        nameText = nm.name

        If IstNOKMarkerName(nameText) Then

            Set rng = Nothing
            Set rng = nm.RefersToRange

            If Not rng Is Nothing Then

                If rng.Parent Is ws Then

                    anzahl = anzahl + 1

                    ReDim Preserve markerSpalten(1 To anzahl)

                    markerSpalten(anzahl) = rng.Column

                End If

            End If

        End If

    Next nm

    On Error GoTo 0

    If anzahl <= 1 Then

        NOKMarkerSpaltenErmitteln = anzahl
        Exit Function

    End If

    For i = 1 To anzahl - 1

        For j = i + 1 To anzahl

            If markerSpalten(j) < markerSpalten(i) Then

                temp = markerSpalten(i)
                markerSpalten(i) = markerSpalten(j)
                markerSpalten(j) = temp

            End If

        Next j

    Next i

    NOKMarkerSpaltenErmitteln = anzahl

End Function

' ============================================================
' PRÜFEN, OB EINE SPALTE EIN ERKANNTER _NOK-MARKER IST
' ============================================================

Private Function IstErkannterNOKMarker( _
    ByRef markerSpalten() As Long, _
    ByVal markerAnzahl As Long, _
    ByVal zielSpalte As Long) As Boolean

    Dim i As Long

    IstErkannterNOKMarker = False

    If markerAnzahl <= 0 Then Exit Function

    For i = LBound(markerSpalten) To UBound(markerSpalten)

        If markerSpalten(i) = zielSpalte Then

            IstErkannterNOKMarker = True
            Exit Function

        End If

    Next i

End Function

' ============================================================
' DASHBOARD-STRUKTUR PRÜFEN
' ============================================================

Private Sub DashboardStrukturPruefen( _
    ByVal wsDash As Worksheet, _
    ByVal cSpieler As Long, _
    ByVal cRang As Long, _
    ByVal cH As Long, _
    ByVal cMacht As Long, _
    ByVal cT9 As Long, _
    ByVal cG As Long, _
    ByVal cM As Long, _
    ByVal cS As Long, _
    ByVal cE As Long, _
    ByVal cJoined As Long, _
    ByVal cAliasse As Long, _
    ByVal cInternNeu As Long, _
    ByVal cInternSortierung As Long, _
    ByVal cT9Seit As Long, _
    ByVal cBasisdaten As Long, _
    ByVal cLetzterNOK As Long, _
    ByVal zeileSummen As Long, _
    ByVal zeileSchwelle1 As Long, _
    ByVal zeileSchwelle2 As Long, _
    ByVal zeileBezuege As Long, _
    ByVal zeileSpieler1 As Long)

    Dim namePruefen As Variant
    Dim rng As Range

    Dim markerSpalten() As Long
    Dim markerAnzahl As Long

    Dim i As Long
    Dim ersterNOK As Long

    If cSpieler <= 0 Then _
        Err.Raise 1004, , "LEG1_spieler wurde nicht gefunden."

    If cRang <= 0 Then _
        Err.Raise 1004, , "LEG1_rang wurde nicht gefunden."

    If cH <= 0 Then _
        Err.Raise 1004, , "LEG1_H wurde nicht gefunden."

    If cMacht <= 0 Then _
        Err.Raise 1004, , "LEG1_macht wurde nicht gefunden."

    If cT9 <= 0 Then _
        Err.Raise 1004, , "LEG1_T9 wurde nicht gefunden."

    If cG <= 0 Then _
        Err.Raise 1004, , "LEG1_G wurde nicht gefunden."

    If cM <= 0 Then _
        Err.Raise 1004, , "LEG1_M wurde nicht gefunden."

    If cS <= 0 Then _
        Err.Raise 1004, , "LEG1_S wurde nicht gefunden."

    If cE <= 0 Then _
        Err.Raise 1004, , "LEG1_E wurde nicht gefunden."

    If cJoined <= 0 Then _
        Err.Raise 1004, , "LEG1_joined wurde nicht gefunden."

    If cAliasse <= 0 Then _
        Err.Raise 1004, , "LEG1_aliasse wurde nicht gefunden."

    If cInternNeu <= 0 Then _
        Err.Raise 1004, , "LEG1_InternNeu wurde nicht gefunden."

    If cInternSortierung <= 0 Then _
        Err.Raise 1004, , "LEG1_InternSortierung wurde nicht gefunden."

    If cT9Seit <= 0 Then _
        Err.Raise 1004, , "LEG1_T9_seit wurde nicht gefunden."

    If cBasisdaten <= 0 Then _
        Err.Raise 1004, , "LEG1_basisdaten wurde nicht gefunden."
    If cLetzterNOK <= 0 Then _
        Err.Raise 1004, , _
        "Kein gültiger _NOK-Endmarker wurde gefunden."

    If cSpieler >= cBasisdaten Then _
        Err.Raise 1004, , _
        "LEG1_spieler muss vor LEG1_basisdaten liegen."

    If cRang >= cBasisdaten Then _
        Err.Raise 1004, , _
        "LEG1_rang muss vor LEG1_basisdaten liegen."

    If cH >= cBasisdaten Then _
        Err.Raise 1004, , _
        "LEG1_H muss vor LEG1_basisdaten liegen."

    If cMacht >= cBasisdaten Then _
        Err.Raise 1004, , _
        "LEG1_macht muss vor LEG1_basisdaten liegen."

    If cT9 >= cBasisdaten Then _
        Err.Raise 1004, , _
        "LEG1_T9 muss vor LEG1_basisdaten liegen."

    If cG >= cBasisdaten Then _
        Err.Raise 1004, , _
        "LEG1_G muss vor LEG1_basisdaten liegen."

    If cM >= cBasisdaten Then _
        Err.Raise 1004, , _
        "LEG1_M muss vor LEG1_basisdaten liegen."

    If cS >= cBasisdaten Then _
        Err.Raise 1004, , _
        "LEG1_S muss vor LEG1_basisdaten liegen."

    If cE >= cBasisdaten Then _
        Err.Raise 1004, , _
        "LEG1_E muss vor LEG1_basisdaten liegen."

    If cJoined >= cBasisdaten Then _
        Err.Raise 1004, , _
        "LEG1_joined muss vor LEG1_basisdaten liegen."

    If cAliasse >= cBasisdaten Then _
        Err.Raise 1004, , _
        "LEG1_aliasse muss vor LEG1_basisdaten liegen."

    If cInternNeu >= cBasisdaten Then _
        Err.Raise 1004, , _
        "LEG1_InternNeu muss vor LEG1_basisdaten liegen."

    If cInternSortierung >= cBasisdaten Then _
        Err.Raise 1004, , _
        "LEG1_InternSortierung muss vor LEG1_basisdaten liegen."

    If cT9Seit >= cBasisdaten Then _
        Err.Raise 1004, , _
        "LEG1_T9_seit muss vor LEG1_basisdaten liegen."

    markerAnzahl = NOKMarkerSpaltenErmitteln( _
        wsDash, _
        markerSpalten)

    If markerAnzahl = 0 Then

        Err.Raise 1004, , _
            "Es wurde kein _NOK-Marker auf dem Dashboard gefunden."

    End If

    ersterNOK = markerSpalten(1)

    If ersterNOK <= cBasisdaten Then

        Err.Raise 1004, , _
            "Der erste _NOK-Marker muss rechts von " & _
            LEG1_Basisdaten_Kopf & _
            " liegen."

    End If

    For i = 1 To markerAnzahl

        If markerSpalten(i) <= cBasisdaten Then

            Err.Raise 1004, , _
                "Der _NOK-Marker in Spalte " & _
                CStr(markerSpalten(i)) & _
                " liegt nicht rechts von " & _
                LEG1_Basisdaten_Kopf & "."

        End If

        If i > 1 Then

            If markerSpalten(i) <= markerSpalten(i - 1) Then

                Err.Raise 1004, , _
                    "Die _NOK-Marker bilden keine eindeutige " & _
                    "aufsteigende Bereichsstruktur."

            End If

        End If

    Next i
    If zeileSummen <= 0 Then _
        Err.Raise 1004, , _
        "Dashboard_1_summen wurde nicht gefunden."

    If zeileSchwelle1 <= 0 Then _
        Err.Raise 1004, , _
        "Dashboard_2_schwelle_1 wurde nicht gefunden."

    If zeileSchwelle2 <= 0 Then _
        Err.Raise 1004, , _
        "Dashboard_3_schwelle_2 wurde nicht gefunden."

    If zeileBezuege <= 0 Then _
        Err.Raise 1004, , _
        "Dashboard_4_bezüge wurde nicht gefunden."

    If zeileSpieler1 <= 0 Then _
        Err.Raise 1004, , _
        "Dashboard_spieler_1 wurde nicht gefunden."

    If Not ( _
            zeileSummen < zeileSchwelle1 And _
            zeileSchwelle1 < zeileSchwelle2 And _
            zeileSchwelle2 < zeileBezuege And _
            zeileBezuege < zeileSpieler1) Then

        Err.Raise 1004, , _
            "Die Reihenfolge der Dashboard-Zeilenanker ist ungültig." & _
            vbCrLf & vbCrLf & _
            "Erwartet:" & vbCrLf & _
            "Dashboard_1_summen" & vbCrLf & _
            "Dashboard_2_schwelle_1" & vbCrLf & _
            "Dashboard_3_schwelle_2" & vbCrLf & _
            "Dashboard_4_bezüge" & vbCrLf & _
            "Dashboard_spieler_1"

    End If

    For Each namePruefen In Array( _
        NAME_LEG1_SPIELER, _
        NAME_LEG1_RANG, _
        NAME_LEG1_H, _
        NAME_LEG1_MACHT, _
        NAME_LEG1_T9, _
        NAME_LEG1_G, _
        NAME_LEG1_M, _
        NAME_LEG1_S, _
        NAME_LEG1_E, _
        NAME_LEG1_JOINED, _
        NAME_LEG1_ALIASE, _
        NAME_LEG1_INTERNNEU, _
        NAME_LEG1_INTERNSORTIERUNG, _
        NAME_LEG1_T9_SEIT, _
        LEG1_Basisdaten_Kopf, _
        NAME_DASHBOARD_SUMMEN, _
        NAME_DASHBOARD_SCHWELLE1, _
        NAME_DASHBOARD_SCHWELLE2, _
        NAME_DASHBOARD_BEZUEGE, _
        NAME_DASHBOARD_SPIELER1)

        Set rng = NameManagerRangeErmitteln( _
            wsDash, _
            CStr(namePruefen))

        If rng Is Nothing Then

            Err.Raise 1004, , _
                "Der definierte Name '" & _
                CStr(namePruefen) & _
                "' verweist nicht auf das Blatt '" & _
                wsDash.name & "'."

        End If

    Next namePruefen

End Sub

' ============================================================
' NEUE TB__-BEREICHE AUTOMATISCH ERKENNEN UND ANLEGEN
' ============================================================

Private Sub NeueTBBereicheAutomatischAnlegen( _
    ByVal wsDash As Worksheet, _
    ByVal cBasisdaten As Long, _
    ByVal zeileSummen As Long, _
    ByVal zeileSchwelle1 As Long, _
    ByVal zeileSchwelle2 As Long, _
    ByVal zeileBezuege As Long, _
    ByVal cSpieler As Long, _
    ByVal cT9 As Long, _
    ByVal zeileSpieler1 As Long, _
    ByVal letzteZeileDash As Long, _
    ByVal wsLog As Worksheet)

    Dim dateiname As String
    Dim dateipfad As String
    Dim bereichsName As String
    Dim excelBereichsName As String
    Dim nokName As String

    Dim wbQuelle As Workbook
    Dim wsQuelle As Worksheet

    Dim cLetzterNOK As Long
    Dim cStart As Long
    Dim cNOK As Long
    Dim c As Long
    Dim i As Long
    Dim blattIndex As Long
    Dim anzahlBlaetter As Long

    Dim warBereitsOffen As Boolean
    Dim spielerWerte As Collection

    Dim spieler As String
    Dim key As String
    Dim wert As Variant
    Dim dashboardWert As Variant
    Dim schwelle1 As Variant
    Dim schwelle2 As Variant
    Dim schwelleVerwenden As Variant

    Dim hatSchwelle1 As Boolean
    Dim hatSchwelle2 As Boolean
    Dim pruefen As Boolean
    Dim istBewertbar As Boolean
    Dim wertGefunden As Boolean

    Dim anzahlAktiveUnterSchwelle As Long
    Dim letzteZeileQuelle As Long

    Dim sheetName As String
    Dim datenName As String

    Dim namePruefen As String
    Dim namenKollision As Boolean

    If wsDash Is Nothing Then Exit Sub
    If cBasisdaten <= 0 Then Exit Sub
    If zeileSummen <= 0 Then Exit Sub
    If zeileSchwelle1 <= 0 Then Exit Sub
    If zeileSchwelle2 <= 0 Then Exit Sub
    If zeileBezuege <= 0 Then Exit Sub
    If cSpieler <= 0 Then Exit Sub
    If cT9 <= 0 Then Exit Sub
    If zeileSpieler1 <= 0 Then Exit Sub
    If letzteZeileDash < zeileSpieler1 Then Exit Sub

    If Len(ThisWorkbook.Path) = 0 Then

        LogEintrag _
            wsLog, _
            "TB__ WARNUNG", _
            "", _
            "Die LEG1-Arbeitsmappe wurde noch nicht gespeichert. " & _
            "Der Ordner für automatische TB__Bereiche konnte nicht ermittelt werden."

        Exit Sub

    End If

    ' Alle TB__*.xlsx-Dateien im Verzeichnis der LEG1-Arbeitsmappe
    ' werden nach derselben generischen Bereichslogik geprüft.
    ' Chests bleibt als Sonderformat ausgenommen, da dessen
    ' Datenstruktur datumsbezogen ist.
    On Error GoTo ScanFehler

    dateiname = Dir( _
        ThisWorkbook.Path & _
        Application.PathSeparator & _
        "TB__*.xlsx")

    Do While Len(dateiname) > 0

        If StrComp( _
                dateiname, _
                CHESTS_DATEI, _
                vbTextCompare) <> 0 Then

            bereichsName = dateiname

            If Left$(bereichsName, 4) = "TB__" Then
                bereichsName = Mid$(bereichsName, 5)
            End If

            If LCase$(Right$(bereichsName, 5)) = ".xlsx" Then
                bereichsName = Left$( _
                    bereichsName, _
                    Len(bereichsName) - 5)
            End If

            bereichsName = Trim$(bereichsName)
            excelBereichsName = ExcelNameBereinigen(bereichsName)

            If Len(bereichsName) = 0 Or _
               Len(excelBereichsName) = 0 Then

                LogEintrag _
                    wsLog, _
                    "TB__ WARNUNG", _
                    "", _
                    "Datei '" & _
                    dateiname & _
                    "' hat keinen gültigen Bereichsnamen."

            Else

                nokName = excelBereichsName & "_NOK"

                ' Bereits vorhandene Bereiche werden bewusst nicht
                ' neu angelegt. Dadurch bleiben bestehende Spezial-
                ' bereiche und bereits manuell eingerichtete Bereiche
                ' unangetastet.
                If DashboardSpalte(wsDash, nokName) > 0 Then

                    cNOK = DashboardSpalte(wsDash, nokName)

                    TBBereichVorhandenAktualisieren _
                        wsDash, _
                        cBasisdaten, _
                        cNOK, _
                        zeileSummen, _
                        zeileSchwelle1, _
                        zeileSchwelle2, _
                        zeileBezuege, _
                        cSpieler, _
                        cT9, _
                        zeileSpieler1, _
                        letzteZeileDash, _
                        wsLog, _
                        dateiname, _
                        bereichsName

                Else

                    dateipfad = _
                        ThisWorkbook.Path & _
                        Application.PathSeparator & _
                        dateiname

                    Set wbQuelle = Nothing
                    warBereitsOffen = False

                    On Error Resume Next

                    Set wbQuelle = Workbooks(dateiname)

                    On Error GoTo ScanFehler

                    If wbQuelle Is Nothing Then

                        On Error GoTo DateiOeffnenFehler

                        Set wbQuelle = Workbooks.Open( _
                            FileName:=dateipfad, _
                            UpdateLinks:=0, _
                            ReadOnly:=True, _
                            IgnoreReadOnlyRecommended:=True)

                        On Error GoTo ScanFehler

                    Else

                        warBereitsOffen = True

                    End If

                    If wbQuelle Is Nothing Then GoTo NaechsteTBDatei

                    anzahlBlaetter = wbQuelle.Worksheets.Count

                    If anzahlBlaetter <= 0 Then

                        LogEintrag _
                            wsLog, _
                            "TB__ WARNUNG", _
                            "", _
                            "Die Datei '" & _
                            dateiname & _
                            "' enthält keine Arbeitsblätter."

                        GoTo TBQuelleSchliessen

                    End If

                    ' Der Bereich wird erst nach erfolgreichem Öffnen
                    ' der Quelldatei und erfolgreicher Ermittlung der
                    ' Blattanzahl im Dashboard angelegt.
                    cLetzterNOK = LetzterNOKMarker(wsDash)

                    If cLetzterNOK <= 0 Then

                        LogEintrag _
                            wsLog, _
                            "TB__ WARNUNG", _
                            "", _
                            "Für den neuen Bereich '" & _
                            bereichsName & _
                            "' wurde kein vorhandener _NOK-Endmarker " & _
                            "als Vorgänger gefunden."

                        GoTo TBQuelleSchliessen

                    End If

                    cStart = cLetzterNOK + 1
                    cNOK = cStart + anzahlBlaetter

                    BereichNOKNameSicherstellen _
                        wsDash, _
                        cNOK, _
                        nokName

                    wsDash.Cells( _
                        zeileBezuege, _
                        cNOK).Value = bereichsName

                    wsDash.Range( _
                        wsDash.Cells(zeileSpieler1, cNOK), _
                        wsDash.Cells(letzteZeileDash, cNOK)).Value = 0

                    For blattIndex = 1 To anzahlBlaetter

                        Set wsQuelle = _
                            wbQuelle.Worksheets(blattIndex)

                        c = cStart + blattIndex - 1

                        sheetName = _
                            SichererText(wsQuelle.Name)

                        wsDash.Cells(1, c).Value = 0
                        wsDash.Cells(zeileSchwelle1, c).ClearContents
                        wsDash.Cells(zeileSchwelle2, c).ClearContents
                        wsDash.Cells(zeileBezuege, c).Value = sheetName

                        BereichDatenNameSicherstellen _
                            wsDash, _
                            c, _
                            ExcelNameBereinigen( _
                                excelBereichsName & _
                                "_" & _
                                sheetName)

                        schwelle1 = Empty
                        schwelle2 = Empty
                        dashboardWert = Empty

                        hatSchwelle1 = _
                            TBWertFuerDashboard( _
                                wsQuelle.Cells(1, 4).Value, _
                                dashboardWert) > 0

                        If hatSchwelle1 Then

                            schwelle1 = _
                                wsQuelle.Cells(1, 4).Value

                            wsDash.Cells( _
                                zeileSchwelle1, _
                                c).Value = dashboardWert

                        End If

                        dashboardWert = Empty

                        hatSchwelle2 = _
                            TBWertFuerDashboard( _
                                wsQuelle.Cells(1, 5).Value, _
                                dashboardWert) > 0

                        If hatSchwelle2 Then

                            schwelle2 = _
                                wsQuelle.Cells(1, 5).Value

                            wsDash.Cells( _
                                zeileSchwelle2, _
                                c).Value = dashboardWert

                        End If

                        Set spielerWerte = New Collection

                        letzteZeileQuelle = _
                            wsQuelle.Cells( _
                                wsQuelle.Rows.Count, _
                                3).End(xlUp).Row

                        If letzteZeileQuelle >= 3 Then

                            For i = 3 To letzteZeileQuelle

                                spieler = SichererText( _
                                    wsQuelle.Cells(i, 3).Value)

                                If Len(spieler) > 0 Then

                                    key = SpielerKey(spieler)

                                    If CollectionKeyExistiert( _
                                            spielerWerte, _
                                            key) Then

                                        LogEintrag _
                                            wsLog, _
                                            "DUPLIKAT TB__", _
                                            spieler, _
                                            "Spieler kommt mehrfach im Quellblatt '" & _
                                            sheetName & _
                                            "' der Datei '" & _
                                            dateiname & _
                                            "' vor. Der erste Wert wird verwendet."

                                    Else

                                        spielerWerte.Add _
                                            wsQuelle.Cells(i, 5).Value, _
                                            key

                                    End If

                                End If

                            Next i

                        End If

                        anzahlAktiveUnterSchwelle = 0

                        For i = zeileSpieler1 To letzteZeileDash

                            spieler = SichererText( _
                                wsDash.Cells(i, cSpieler).Value)

                            If Len(spieler) > 0 Then

                                key = SpielerKey(spieler)

                                wertGefunden = _
                                    CollectionKeyExistiert( _
                                        spielerWerte, _
                                        key)

                                If wertGefunden Then

                                    wert = _
                                        CollectionWert( _
                                            spielerWerte, _
                                            key)

                                    dashboardWert = Empty

                                    If TBWertFuerDashboard( _
                                            wert, _
                                            dashboardWert) > 0 Then

                                        wsDash.Cells(i, c).Value = _
                                            dashboardWert

                                    Else

                                        wsDash.Cells(i, c).ClearContents

                                    End If

                                    pruefen = False
                                    schwelleVerwenden = Empty

                                    If hatSchwelle1 And _
                                       hatSchwelle2 Then

                                        If IstEins( _
                                                wsDash.Cells( _
                                                    i, _
                                                    cT9).Value) Then

                                            schwelleVerwenden = schwelle1

                                        Else

                                            schwelleVerwenden = schwelle2

                                        End If

                                        pruefen = True

                                    ElseIf (Not hatSchwelle1) And _
                                           hatSchwelle2 Then

                                        schwelleVerwenden = schwelle2
                                        pruefen = True

                                    ElseIf hatSchwelle1 And _
                                           (Not hatSchwelle2) Then

                                        If IstEins( _
                                                wsDash.Cells( _
                                                    i, _
                                                    cT9).Value) Then

                                            schwelleVerwenden = schwelle1
                                            pruefen = True

                                        End If

                                    End If

                                    If pruefen Then

                                        If TBWertUnterSchwelle( _
                                                wert, _
                                                schwelleVerwenden, _
                                                istBewertbar) Then

                                            If IstAktiv( _
                                                    wsDash.Cells( _
                                                        i, _
                                                        cBasisdaten).Value) Then

                                                anzahlAktiveUnterSchwelle = _
                                                    anzahlAktiveUnterSchwelle + 1

                                            End If

                                            wsDash.Cells(i, c).Font.Color = _
                                                RGB(255, 0, 0)

                                            wsDash.Cells( _
                                                i, _
                                                cNOK).Value = _
                                                CLng( _
                                                    Val( _
                                                        wsDash.Cells( _
                                                            i, _
                                                            cNOK).Value)) + 1

                                        ElseIf istBewertbar Then

                                            wsDash.Cells(i, c).Font.ColorIndex = _
                                                xlAutomatic

                                        Else

                                            wsDash.Cells(i, c).Font.ColorIndex = _
                                                xlAutomatic

                                        End If

                                    Else

                                        wsDash.Cells(i, c).Font.ColorIndex = _
                                            xlAutomatic

                                    End If

                                Else

                                    wsDash.Cells(i, c).ClearContents

                                    wsDash.Cells(i, c).Font.ColorIndex = _
                                        xlAutomatic

                                End If

                            End If

                        Next i

                        wsDash.Cells(1, c).Value = _
                            anzahlAktiveUnterSchwelle

                    Next blattIndex

                    LogEintrag _
                        wsLog, _
                        "TB__ BEREICH ANGELEGT", _
                        "", _
                        "Bereich '" & _
                        bereichsName & _
                        "' wurde automatisch aus '" & _
                        dateiname & _
                        "' angelegt. Datenspalten: " & _
                        CStr(anzahlBlaetter) & _
                        ", _NOK-Marker: " & _
                        nokName

TBQuelleSchliessen:

                    If Not wbQuelle Is Nothing Then

                        If Not warBereitsOffen Then

                            On Error Resume Next

                            wbQuelle.Close SaveChanges:=False

                            On Error GoTo ScanFehler

                        End If

                    End If

                    Set wbQuelle = Nothing

                End If

            End If

        End If

NaechsteTBDatei:

        dateiname = Dir()

    Loop

    Exit Sub

DateiOeffnenFehler:

    LogEintrag _
        wsLog, _
        "TB__ WARNUNG", _
        "", _
        "Die Datei '" & _
        dateiname & _
        "' konnte nicht geöffnet werden. Fehler " & _
        CStr(Err.Number) & _
        ": " & _
        Err.Description

    On Error Resume Next

    If Not wbQuelle Is Nothing Then

        If Not warBereitsOffen Then
            wbQuelle.Close SaveChanges:=False
        End If

    End If

    Set wbQuelle = Nothing

    On Error GoTo ScanFehler

    Resume NaechsteTBDatei

ScanFehler:

    LogEintrag _
        wsLog, _
        "TB__ WARNUNG", _
        "", _
        "Fehler bei der automatischen TB__Bereichserkennung. " & _
        "Datei: '" & _
        dateiname & _
        "'. Fehler " & _
        CStr(Err.Number) & _
        ": " & _
        Err.Description

    On Error Resume Next

    If Not wbQuelle Is Nothing Then

        If Not warBereitsOffen Then
            wbQuelle.Close SaveChanges:=False
        End If

    End If

    On Error GoTo 0

End Sub

' ============================================================
' EXCEL-NAMEN AUS BEREICHS-/BLATTNAMEN ERZEUGEN
' ============================================================

Private Function ExcelNameBereinigen( _
    ByVal text As String) As String

    Dim i As Long
    Dim ch As String
    Dim ergebnis As String

    text = Trim$(text)

    For i = 1 To Len(text)

        ch = Mid$(text, i, 1)

        If ch Like "[A-Za-z0-9_]" Then

            ergebnis = ergebnis & ch

        Else

            ergebnis = ergebnis & "_"

        End If

    Next i

    Do While InStr(1, ergebnis, "__", vbBinaryCompare) > 0

        ergebnis = Replace( _
            ergebnis, _
            "__", _
            "_", _
            1, _
            -1, _
            vbBinaryCompare)

    Loop

    If Len(ergebnis) = 0 Then

        ergebnis = "_Bereich"

    End If

    If Mid$(ergebnis, 1, 1) Like "[0-9]" Then

        ergebnis = "_" & ergebnis

    End If

    If UCase$(ergebnis) = "R" Or _
       UCase$(ergebnis) = "C" Then

        ergebnis = "_" & ergebnis

    End If

    If Len(ergebnis) > 200 Then

        ergebnis = Left$(ergebnis, 200)

    End If

    ExcelNameBereinigen = ergebnis

End Function

' ============================================================
' GENERISCHEN _NOK-NAMEN SICHERSTELLEN
' ============================================================

Private Sub BereichNOKNameSicherstellen( _
    ByVal wsDash As Worksheet, _
    ByVal cNOK As Long, _
    ByVal nameText As String)

    Dim nm As name
    Dim rngZiel As Range
    Dim zielBezug As String

    If wsDash Is Nothing Then Exit Sub
    If cNOK <= 0 Then Exit Sub
    If Len(nameText) = 0 Then Exit Sub

    Set rngZiel = wsDash.Columns(cNOK)

    zielBezug = "=" & _
        wsDash.Name & "!" & _
        rngZiel.Address( _
            RowAbsolute:=True, _
            ColumnAbsolute:=True, _
            ReferenceStyle:=xlA1)

    Set nm = Nothing

    On Error Resume Next
    Set nm = ThisWorkbook.Names(nameText)
    On Error GoTo 0

    If nm Is Nothing Then

        Set nm = ThisWorkbook.Names.Add( _
            Name:=nameText, _
            RefersTo:=zielBezug, _
            Visible:=True)

    Else

        nm.Visible = True
        nm.RefersTo = zielBezug

    End If

End Sub

' ============================================================
' GENERISCHEN DATENSPALTEN-NAMEN SICHERSTELLEN
' ============================================================

Private Sub BereichDatenNameSicherstellen( _
    ByVal wsDash As Worksheet, _
    ByVal c As Long, _
    ByVal nameText As String)

    Dim nm As name
    Dim rngZiel As Range
    Dim zielBezug As String

    If wsDash Is Nothing Then Exit Sub
    If c <= 0 Then Exit Sub
    If Len(nameText) = 0 Then Exit Sub

    Set rngZiel = wsDash.Columns(c)

    zielBezug = "=" & _
        wsDash.Name & "!" & _
        rngZiel.Address( _
            RowAbsolute:=True, _
            ColumnAbsolute:=True, _
            ReferenceStyle:=xlA1)

    Set nm = Nothing

    On Error Resume Next
    Set nm = ThisWorkbook.Names(nameText)
    On Error GoTo 0

    If nm Is Nothing Then

        Set nm = ThisWorkbook.Names.Add( _
            Name:=nameText, _
            RefersTo:=zielBezug, _
            Visible:=True)

    Else

        nm.Visible = True
        nm.RefersTo = zielBezug

    End If

End Sub

' ============================================================
' DATENSPALTEN DER DYNAMISCHEN BEREICHE GRUPPIEREN
' ============================================================

Private Sub TBBereichVorhandenAktualisieren( _
    ByVal wsDash As Worksheet, _
    ByVal cBasisdaten As Long, _
    ByVal cNOK As Long, _
    ByVal zeileSummen As Long, _
    ByVal zeileSchwelle1 As Long, _
    ByVal zeileSchwelle2 As Long, _
    ByVal zeileBezuege As Long, _
    ByVal cSpieler As Long, _
    ByVal cT9 As Long, _
    ByVal zeileSpieler1 As Long, _
    ByVal letzteZeileDash As Long, _
    ByVal wsLog As Worksheet, _
    ByVal dateiname As String, _
    ByVal bereichsName As String)

    Dim wbQuelle As Workbook
    Dim wsQuelle As Worksheet
    Dim spielerWerte As Collection

    Dim cStart As Long
    Dim cEnde As Long
    Dim c As Long
    Dim i As Long
    Dim blattIndex As Long
    Dim anzahlBlaetter As Long
    Dim letzteZeileQuelle As Long

    Dim sheetName As String
    Dim datenName As String
    Dim spieler As String
    Dim key As String
    Dim wert As Variant
    Dim dashboardWert As Variant
    Dim schwelle1 As Variant
    Dim schwelle2 As Variant
    Dim schwelleVerwenden As Variant

    Dim hatSchwelle1 As Boolean
    Dim hatSchwelle2 As Boolean
    Dim pruefen As Boolean
    Dim istBewertbar As Boolean
    Dim wertGefunden As Boolean
    Dim anzahlAktiveUnterSchwelle As Long

    Dim warBereitsOffen As Boolean
    Dim dateipfad As String
    Dim cAnzahlDaten As Long

    If wsDash Is Nothing Then Exit Sub
    If cNOK <= 0 Then Exit Sub

    If Not BereichsGrenzenFuerNOKMarkerErmitteln( _
            wsDash, _
            DashboardSpalte(wsDash, LEG1_Basisdaten_Kopf), _
            cNOK, _
            cStart, _
            cEnde) Then

        LogEintrag _
            wsLog, _
            "TB__ WARNUNG", _
            "", _
            "Bereich '" & bereichsName & "' konnte nicht aktualisiert werden. " & _
            "Die dynamischen Bereichsgrenzen konnten nicht ermittelt werden."

        Exit Sub

    End If

    If cStart <= 0 Or cEnde < cStart Then Exit Sub

    If Len(ThisWorkbook.Path) = 0 Then Exit Sub

    dateipfad = _
        ThisWorkbook.Path & _
        Application.PathSeparator & _
        dateiname

    On Error Resume Next
    Set wbQuelle = Workbooks(dateiname)
    On Error GoTo OeffnenFehler

    If wbQuelle Is Nothing Then

        Set wbQuelle = Workbooks.Open( _
            FileName:=dateipfad, _
            UpdateLinks:=0, _
            ReadOnly:=True, _
            IgnoreReadOnlyRecommended:=True)

    Else

        warBereitsOffen = True

    End If

    If wbQuelle Is Nothing Then Exit Sub

    anzahlBlaetter = wbQuelle.Worksheets.Count

    If anzahlBlaetter <= 0 Then GoTo Aufraeumen

    cAnzahlDaten = cEnde - cStart + 1

    If anzahlBlaetter > cAnzahlDaten Then

        LogEintrag _
            wsLog, _
            "TB__ WARNUNG", _
            "", _
            "Bereich '" & bereichsName & "' hat " & _
            CStr(anzahlBlaetter) & " Quellblätter, aber nur " & _
            CStr(cAnzahlDaten) & " vorhandene Datenspalten. " & _
            "Neue Quellblätter wurden nicht automatisch ergänzt."

    End If

    ' Bestehende Ergebnisse des Bereichs vollständig zurücksetzen.
    wsDash.Range( _
        wsDash.Cells(zeileSpieler1, cNOK), _
        wsDash.Cells(letzteZeileDash, cNOK)).Value = 0

    wsDash.Range( _
        wsDash.Cells(zeileSpieler1, cNOK), _
        wsDash.Cells(letzteZeileDash, cNOK)).Font.ColorIndex = _
        xlAutomatic

    For c = cStart To cEnde

        wsDash.Cells(1, c).Value = 0
        wsDash.Cells(zeileSchwelle1, c).ClearContents
        wsDash.Cells(zeileSchwelle2, c).ClearContents

        If c <= cStart + anzahlBlaetter - 1 Then

            Set wsQuelle = wbQuelle.Worksheets(c - cStart + 1)

            sheetName = SichererText(wsQuelle.Name)

            wsDash.Cells(zeileBezuege, c).Value = sheetName

            datenName = ExcelNameBereinigen( _
                bereichsName & "_" & sheetName)

            BereichDatenNameSicherstellen _
                wsDash, _
                c, _
                datenName

            dashboardWert = Empty

            hatSchwelle1 = _
                TBWertFuerDashboard( _
                    wsQuelle.Cells(1, 4).Value, _
                    dashboardWert) > 0

            If hatSchwelle1 Then
                schwelle1 = wsQuelle.Cells(1, 4).Value
                wsDash.Cells(zeileSchwelle1, c).Value = dashboardWert
            Else
                schwelle1 = Empty
            End If

            dashboardWert = Empty

            hatSchwelle2 = _
                TBWertFuerDashboard( _
                    wsQuelle.Cells(1, 5).Value, _
                    dashboardWert) > 0

            If hatSchwelle2 Then
                schwelle2 = wsQuelle.Cells(1, 5).Value
                wsDash.Cells(zeileSchwelle2, c).Value = dashboardWert
            Else
                schwelle2 = Empty
            End If

            Set spielerWerte = New Collection

            letzteZeileQuelle = _
                wsQuelle.Cells( _
                    wsQuelle.Rows.Count, _
                    3).End(xlUp).Row

            If letzteZeileQuelle >= 3 Then

                For i = 3 To letzteZeileQuelle

                    spieler = SichererText( _
                        wsQuelle.Cells(i, 3).Value)

                    If Len(spieler) > 0 Then

                        key = SpielerKey(spieler)

                        If CollectionKeyExistiert( _
                                spielerWerte, _
                                key) Then

                            LogEintrag _
                                wsLog, _
                                "DUPLIKAT TB__", _
                                spieler, _
                                "Spieler kommt mehrfach im Quellblatt '" & _
                                sheetName & "' der Datei '" & _
                                dateiname & "' vor. Der erste Wert wird verwendet."

                        Else

                            spielerWerte.Add _
                                wsQuelle.Cells(i, 5).Value, _
                                key

                        End If

                    End If

                Next i

            End If

            anzahlAktiveUnterSchwelle = 0

            For i = zeileSpieler1 To letzteZeileDash

                spieler = SichererText( _
                    wsDash.Cells(i, cSpieler).Value)

                If Len(spieler) > 0 Then

                    key = SpielerKey(spieler)
                    wertGefunden = CollectionKeyExistiert( _
                        spielerWerte, _
                        key)

                    If wertGefunden Then

                        wert = CollectionWert( _
                            spielerWerte, _
                            key)

                        dashboardWert = Empty

                        If TBWertFuerDashboard( _
                                wert, _
                                dashboardWert) > 0 Then

                            wsDash.Cells(i, c).Value = dashboardWert

                        Else

                            wsDash.Cells(i, c).ClearContents

                        End If

                        pruefen = False
                        schwelleVerwenden = Empty
                        istBewertbar = False

                        If hatSchwelle1 And hatSchwelle2 Then

                            If IstEins( _
                                    wsDash.Cells(i, cT9).Value) Then

                                schwelleVerwenden = schwelle1

                            Else

                                schwelleVerwenden = schwelle2

                            End If

                            pruefen = True

                        ElseIf (Not hatSchwelle1) And hatSchwelle2 Then

                            schwelleVerwenden = schwelle2
                            pruefen = True

                        ElseIf hatSchwelle1 And (Not hatSchwelle2) Then

                            If IstEins( _
                                    wsDash.Cells(i, cT9).Value) Then

                                schwelleVerwenden = schwelle1
                                pruefen = True

                            End If

                        End If

                        If pruefen Then

                            If TBWertUnterSchwelle( _
                                    wert, _
                                    schwelleVerwenden, _
                                    istBewertbar) Then

                                If IstAktiv( _
                                        wsDash.Cells(i, cBasisdaten).Value) Then

                                    anzahlAktiveUnterSchwelle = _
                                        anzahlAktiveUnterSchwelle + 1

                                End If

                                wsDash.Cells(i, c).Font.Color = _
                                    RGB(255, 0, 0)

                                wsDash.Cells(i, cNOK).Value = _
                                    CLng( _
                                        Val( _
                                            wsDash.Cells(i, cNOK).Value)) + 1

                            Else

                                wsDash.Cells(i, c).Font.ColorIndex = _
                                    xlAutomatic

                            End If

                        Else

                            wsDash.Cells(i, c).Font.ColorIndex = _
                                xlAutomatic

                        End If

                    Else

                        wsDash.Cells(i, c).ClearContents
                        wsDash.Cells(i, c).Font.ColorIndex = xlAutomatic

                    End If

                End If

            Next i

            wsDash.Cells(1, c).Value = _
                anzahlAktiveUnterSchwelle

        Else

            ' Falls im Dashboard noch mehr Datenspalten vorhanden sind
            ' als Quellblätter, werden deren alte Werte entfernt.
            wsDash.Range( _
                wsDash.Cells(zeileSpieler1, c), _
                wsDash.Cells(letzteZeileDash, c)).ClearContents

            wsDash.Cells(zeileBezuege, c).ClearContents

        End If

    Next c

    wsDash.Cells(zeileBezuege, cNOK).Value = bereichsName

    LogEintrag _
        wsLog, _
        "TB__ BEREICH AKTUALISIERT", _
        "", _
        "Bereich '" & bereichsName & "' wurde aus '" & _
        dateiname & "' neu eingelesen. Datenspalten geprüft: " & _
        CStr(cAnzahlDaten) & "."

Aufraeumen:

    If Not wbQuelle Is Nothing Then

        If Not warBereitsOffen Then

            On Error Resume Next
            wbQuelle.Close SaveChanges:=False
            On Error GoTo 0

        End If

    End If

    Exit Sub

OeffnenFehler:

    LogEintrag _
        wsLog, _
        "TB__ WARNUNG", _
        "", _
        "Bereich '" & bereichsName & "' konnte aus '" & _
        dateiname & "' nicht aktualisiert werden. Fehler " & _
        CStr(Err.Number) & ": " & Err.Description

    Resume Aufraeumen

End Sub

Private Sub TB_BereichsformatierungSetzen( _
    ByVal ws As Worksheet, _
    ByVal cBasisdaten As Long, _
    ByVal zeileSchwelle1 As Long, _
    ByVal zeileSchwelle2 As Long, _
    ByVal zeileBezuege As Long)

    Dim markerSpalten() As Long
    Dim markerAnzahl As Long
    Dim cNOK As Long
    Dim cStart As Long
    Dim cEnde As Long
    Dim c As Long

    If cBasisdaten <= 0 Then Exit Sub
    If zeileSchwelle1 <= 0 Then Exit Sub
    If zeileSchwelle2 <= 0 Then Exit Sub
    If zeileBezuege <= 0 Then Exit Sub

    markerAnzahl = NOKMarkerSpaltenErmitteln( _
        ws, _
        markerSpalten)

    If markerAnzahl <= 0 Then Exit Sub

    For c = 1 To markerAnzahl

        cNOK = markerSpalten(c)

        If BereichsGrenzenFuerNOKMarkerErmitteln( _
                ws, _
                cBasisdaten, _
                cNOK, _
                cStart, _
                cEnde) Then

            If cEnde >= cStart Then

                ' ------------------------------------------------
                ' SCHWELLENZEILEN DER DATENSPALTEN
                ' 1000er-Trennzeichen und blaue Schrift.
                ' ------------------------------------------------

                With ws.Range( _
                        ws.Cells(zeileSchwelle1, cStart), _
                        ws.Cells(zeileSchwelle1, cEnde))

                    .NumberFormat = "#,##0"
                    .Font.Color = RGB(0, 0, 255)

                End With

                With ws.Range( _
                        ws.Cells(zeileSchwelle2, cStart), _
                        ws.Cells(zeileSchwelle2, cEnde))

                    .NumberFormat = "#,##0"
                    .Font.Color = RGB(0, 0, 255)

                End With

                ' Zeile 1 des _NOK-Markers enthält immer die Anzahl
                ' der Datenspalten des unmittelbar davorliegenden Bereichs.
                ws.Cells(1, cNOK).Value = cEnde - cStart + 1
                ws.Cells(1, cNOK).NumberFormat = "#,##0"

                ' ------------------------------------------------
                ' _NOK-MARKER
                ' Die Breite entspricht immer der Breite des
                ' unmittelbar davorliegenden _NOK-Markers.
                ' ------------------------------------------------

                ws.Columns(cNOK).Hidden = False

                If c > 1 Then
                    ws.Columns(cNOK).ColumnWidth = _
                        ws.Columns(markerSpalten(c - 1)).ColumnWidth
                End If

                ' ------------------------------------------------
                ' BEZUG DES _NOK-MARKERS
                ' 90 Grad nach links, rot, ohne Füllung.
                ' ------------------------------------------------

                With ws.Cells(zeileBezuege, cNOK)

                    .Orientation = xlUpward
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .Font.Color = RGB(255, 0, 0)
                    .Font.Size = 9
                    .Interior.Pattern = xlNone

                End With

            End If

        End If

    Next c

End Sub

Private Sub BereichsDatenspaltenGruppieren( _
    ByVal ws As Worksheet, _
    ByVal cBasisdaten As Long)

    Dim markerSpalten() As Long
    Dim markerAnzahl As Long
    Dim i As Long

    Dim cStart As Long
    Dim cEnde As Long
    Dim c As Long
    Dim bereitsGruppiert As Boolean

    Dim bereichWarOffen() As Boolean
    Dim bereichGueltig() As Boolean

    If ws Is Nothing Then Exit Sub
    If cBasisdaten <= 0 Then Exit Sub

    markerAnzahl = NOKMarkerSpaltenErmitteln( _
        ws, _
        markerSpalten)

    If markerAnzahl <= 0 Then Exit Sub

    ReDim bereichWarOffen(1 To markerAnzahl)
    ReDim bereichGueltig(1 To markerAnzahl)

    ' ------------------------------------------------------------
    ' AKTUELLEN ZUSTAND DER BEREICHE MERKEN
    '
    ' Ein Bereich gilt als offen, wenn alle seine Datenspalten
    ' sichtbar sind. Ist mindestens eine Datenspalte ausgeblendet,
    ' wird der Bereich als geschlossen behandelt.
    '
    ' Dieser Zustand wird vor dem Gruppieren gesichert, damit der
    ' Abgleich die vom Benutzer gewählte Ansicht nicht verändert.
    ' ------------------------------------------------------------

    For i = 1 To markerAnzahl

        cStart = 0
        cEnde = -1

        If BereichsGrenzenFuerNOKMarkerErmitteln( _
                ws, _
                cBasisdaten, _
                markerSpalten(i), _
                cStart, _
                cEnde) Then

            If cStart <= cEnde Then

                bereichGueltig(i) = True
                bereichWarOffen(i) = True

                For c = cStart To cEnde

                    If ws.Columns(c).Hidden Then

                        bereichWarOffen(i) = False
                        Exit For

                    End If

                Next c

            End If

        End If

    Next i

    ' ------------------------------------------------------------
    ' BEREICHE GRUPPIEREN
    ' ------------------------------------------------------------

    For i = 1 To markerAnzahl

        cStart = 0
        cEnde = -1

        If BereichsGrenzenFuerNOKMarkerErmitteln( _
                ws, _
                cBasisdaten, _
                markerSpalten(i), _
                cStart, _
                cEnde) Then

            If cStart <= cEnde Then

                bereitsGruppiert = True

                For c = cStart To cEnde

                    If ws.Columns(c).OutlineLevel <= 1 Then

                        bereitsGruppiert = False
                        Exit For

                    End If

                Next c

                If Not bereitsGruppiert Then

                    On Error Resume Next

                    ws.Range( _
                        ws.Columns(cStart), _
                        ws.Columns(cEnde)).Group

                    On Error GoTo 0

                End If

            End If

        End If

    Next i

    ' ------------------------------------------------------------
    ' URSPRÜNGLICHEN OFFEN/GESCHLOSSEN-ZUSTAND WIEDERHERSTELLEN
    '
    ' Die _NOK-Marker bleiben immer sichtbar.
    ' ------------------------------------------------------------

    For i = 1 To markerAnzahl

        cStart = 0
        cEnde = -1

        If bereichGueltig(i) Then

            If BereichsGrenzenFuerNOKMarkerErmitteln( _
                    ws, _
                    cBasisdaten, _
                    markerSpalten(i), _
                    cStart, _
                    cEnde) Then

                If cStart <= cEnde Then

                    ws.Range( _
                        ws.Columns(cStart), _
                        ws.Columns(cEnde)).EntireColumn.Hidden = _
                        Not bereichWarOffen(i)

                End If

            End If

        End If

        ws.Columns(markerSpalten(i)).Hidden = False

    Next i

End Sub

' ============================================================
' BEREICHSGRENZEN EINES _NOK-BEREICHS ERMITTELN
' ============================================================

Private Function BereichsGrenzenFuerNOKMarkerErmitteln( _
    ByVal ws As Worksheet, _
    ByVal cBasisdaten As Long, _
    ByVal cNOKMarker As Long, _
    ByRef cStart As Long, _
    ByRef cEnde As Long) As Boolean

    Dim markerSpalten() As Long
    Dim markerAnzahl As Long
    Dim cVorherigerMarker As Long

    BereichsGrenzenFuerNOKMarkerErmitteln = False

    cStart = 0
    cEnde = -1

    If cBasisdaten <= 0 Then Exit Function
    If cNOKMarker <= cBasisdaten Then Exit Function

    markerAnzahl = NOKMarkerSpaltenErmitteln( _
        ws, _
        markerSpalten)

    If Not IstErkannterNOKMarker( _
            markerSpalten, _
            markerAnzahl, _
            cNOKMarker) Then Exit Function

    cVorherigerMarker = VorherigerNOKMarker( _
        ws, _
        cNOKMarker, _
        cBasisdaten)

    BereichsGrenzenErmitteln _
        cVorherigerMarker, _
        cNOKMarker, _
        cStart, _
        cEnde

    If cStart <= 0 Then Exit Function
    If cStart > cEnde Then Exit Function

    BereichsGrenzenFuerNOKMarkerErmitteln = True

End Function

' ============================================================
' BEREICHSGRENZEN ERMITTELN
' ============================================================

Private Sub BereichsGrenzenErmitteln( _
    ByVal cStartMarker As Long, _
    ByVal cEndMarker As Long, _
    ByRef cStart As Long, _
    ByRef cEnde As Long)

    cStart = cStartMarker + 1
    cEnde = cEndMarker - 1

    If cEnde < cStart Then

        cStart = 0
        cEnde = -1

    End If

End Sub

' ============================================================
' LETZTE SPIELERZEILE
' ============================================================

Private Function LetzteSpielerZeile( _
    ByVal ws As Worksheet, _
    ByVal cSpieler As Long, _
    ByVal startZeile As Long) As Long

    Dim letzteZeile As Long

    If cSpieler <= 0 Then

        LetzteSpielerZeile = startZeile - 1
        Exit Function

    End If

    letzteZeile = _
        ws.Cells( _
            ws.Rows.Count, _
            cSpieler).End(xlUp).Row

    If letzteZeile < startZeile Then

        LetzteSpielerZeile = startZeile - 1

    Else

        LetzteSpielerZeile = letzteZeile

    End If

End Function

' ============================================================
' SPIELER-KEY
' ============================================================

Private Function SpielerKey( _
    ByVal spieler As String) As String

    SpielerKey = LCase$(Trim$(spieler))

End Function

' ============================================================
' AKTIV
' ============================================================

Private Function IstAktiv( _
    ByVal wert As Variant) As Boolean

    IstAktiv = False

    If IsError(wert) Then Exit Function
    If IsNull(wert) Then Exit Function
    If IsEmpty(wert) Then Exit Function

    If IsNumeric(wert) Then

        If CDbl(wert) = 1 Then
            IstAktiv = True
        End If

    Else

        If Trim$(CStr(wert)) = "1" Then
            IstAktiv = True
        End If

    End If

End Function

' ============================================================
' IST EINS
' ============================================================

Private Function IstEins( _
    ByVal wert As Variant) As Boolean

    IstEins = False

    If IsError(wert) Then Exit Function
    If IsNull(wert) Then Exit Function
    If IsEmpty(wert) Then Exit Function

    If IsNumeric(wert) Then

        If CDbl(wert) = 1 Then
            IstEins = True
        End If

    Else

        If Trim$(CStr(wert)) = "1" Then
            IstEins = True
        End If

    End If

End Function

' ============================================================
' NUMERISCHER WERT' ============================================================

Private Function NumerischerWert( _
    ByVal wert As Variant, _
    ByRef zahl As Double) As Boolean

    NumerischerWert = False
    zahl = 0

    If IsError(wert) Then Exit Function
    If IsNull(wert) Then Exit Function
    If IsEmpty(wert) Then Exit Function

    ' Boolean-Werte werden hier bewusst nicht als Zahlen behandelt.
    If VarType(wert) = vbBoolean Then Exit Function

    If IsNumeric(wert) Then

        zahl = CDbl(wert)
        NumerischerWert = True

    End If

End Function

' ============================================================
' WERTTYP FÜR VAULTS ERMITTELN
' ============================================================

Private Function TBWertTypErmitteln( _
    ByVal wert As Variant, _
    ByRef zahl As Double, _
    ByRef boolWert As Boolean) As Long

    Const VAULTS_WERT_UNGUELTIG As Long = 0
    Const VAULTS_WERT_ZAHL As Long = 1
    Const VAULTS_WERT_BOOLEAN As Long = 2

    Dim textWert As String

    TBWertTypErmitteln = VAULTS_WERT_UNGUELTIG
    zahl = 0
    boolWert = False

    If IsError(wert) Then Exit Function
    If IsNull(wert) Then Exit Function
    If IsEmpty(wert) Then Exit Function

    If VarType(wert) = vbBoolean Then

        boolWert = CBool(wert)
        TBWertTypErmitteln = VAULTS_WERT_BOOLEAN
        Exit Function

    End If

    textWert = Trim$(CStr(wert))

    If Len(textWert) = 0 Then Exit Function

    If IsNumeric(wert) Then

        zahl = CDbl(wert)
        TBWertTypErmitteln = VAULTS_WERT_ZAHL
        Exit Function

    End If

    Select Case UCase$(textWert)

        Case "WAHR", "TRUE"

            boolWert = True
            TBWertTypErmitteln = VAULTS_WERT_BOOLEAN

        Case "FALSCH", "FALSE"

            boolWert = False
            TBWertTypErmitteln = VAULTS_WERT_BOOLEAN

    End Select

End Function

' ============================================================
' VAULTS-WERT FÜR DAS DASHBOARD AUFBEREITEN
' ============================================================

Private Function TBWertFuerDashboard( _
    ByVal wert As Variant, _
    ByRef dashboardWert As Variant) As Long

    Dim zahl As Double
    Dim boolWert As Boolean

    TBWertFuerDashboard = _
        TBWertTypErmitteln( _
            wert, _
            zahl, _
            boolWert)

    Select Case TBWertFuerDashboard

        Case 1

            dashboardWert = zahl

        Case 2

            If boolWert Then
                dashboardWert = 1
            Else
                dashboardWert = 0
            End If

        Case Else

            dashboardWert = Empty

    End Select

End Function

' ============================================================
' VAULTS-WERT GEGEN SCHWELLENWERT BEWERTEN
' ============================================================

Private Function TBWertUnterSchwelle( _
    ByVal wert As Variant, _
    ByVal schwellenwert As Variant, _
    ByRef istBewertbar As Boolean) As Boolean

    Dim wertTyp As Long
    Dim schwellenTyp As Long

    Dim wertZahl As Double
    Dim schwellenZahl As Double

    Dim wertBoolean As Boolean
    Dim schwellenBoolean As Boolean

    Const VAULTS_WERT_ZAHL As Long = 1
    Const VAULTS_WERT_BOOLEAN As Long = 2

    istBewertbar = False
    TBWertUnterSchwelle = False

    wertTyp = TBWertTypErmitteln( _
        wert, _
        wertZahl, _
        wertBoolean)

    schwellenTyp = TBWertTypErmitteln( _
        schwellenwert, _
        schwellenZahl, _
        schwellenBoolean)

    If wertTyp = VAULTS_WERT_ZAHL And _
       schwellenTyp = VAULTS_WERT_ZAHL Then

        istBewertbar = True
        TBWertUnterSchwelle = _
            (wertZahl < schwellenZahl)

        Exit Function

    End If

    If wertTyp = VAULTS_WERT_BOOLEAN And _
       schwellenTyp = VAULTS_WERT_BOOLEAN Then

        istBewertbar = True

        ' False = 0, True = 1.
        ' Damit wird auch ein boolescher Schwellenwert
        ' eindeutig und ohne VBA-Sonderbehandlung bewertet.
        TBWertUnterSchwelle = _
            (Not wertBoolean And schwellenBoolean)

    End If

End Function

' ============================================================
' UPDATE-DATEN PRÜFEN
' ============================================================

Private Sub UpdateDatenPruefen( _
    ByVal wsUpdate As Worksheet, _
    ByVal letzteZeileUpdate As Long)

    If letzteZeileUpdate < 2 Then

        Err.Raise 1004, , _
            "Im Blatt 'update' wurden keine Spielerdaten gefunden."

    End If

End Sub

' ============================================================
' DASHBOARD-DUPLIKATE BEREINIGEN
' ============================================================

Private Sub DashboardDuplikateBereinigen( _
    ByVal wsDash As Worksheet, _
    ByVal wsLog As Worksheet, _
    ByVal cSpieler As Long, _
    ByVal cBasisdaten As Long, _
    ByRef letzteZeileDash As Long, _
    ByVal zeileSpieler1 As Long)

    Dim col As Collection
    Dim i As Long
    Dim spieler As String
    Dim key As String
    Dim ersteZeile As Long

    Set col = New Collection

    If letzteZeileDash < zeileSpieler1 Then Exit Sub

    For i = zeileSpieler1 To letzteZeileDash

        spieler = SichererText( _
            wsDash.Cells(i, cSpieler).Value)

        If Len(spieler) > 0 Then

            key = SpielerKey(spieler)

            If CollectionKeyExistiert(col, key) Then

                ersteZeile = CLng( _
                    CollectionWert( _
                        col, _
                        key))

                wsDash.Cells( _
                    i, _
                    cBasisdaten).Value = 0

                LogEintrag _
                    wsLog, _
                    "DUPLIKAT DASHBOARD", _
                    spieler, _
                    "Doppelte Zeile " & _
                    CStr(i) & _
                    " deaktiviert; erste Zeile: " & _
                    CStr(ersteZeile)

            Else

                col.Add i, key

            End If

        End If

    Next i

End Sub

' ============================================================
' DATENÄNDERUNGEN ERMITTELN
' ============================================================

Private Function DatenAenderungenErmitteln( _
    ByVal wsDash As Worksheet, _
    ByVal wsUpdate As Worksheet, _
    ByVal dashZeile As Long, _
    ByVal updateZeile As Long, _
    ByVal cSpieler As Long, _
    ByVal cRang As Long, _
    ByVal cH As Long, _
    ByVal cMacht As Long, _
    ByVal cG As Long, _
    ByVal cM As Long, _
    ByVal cS As Long, _
    ByVal cE As Long, _
    ByVal cJoined As Long, _
    ByVal cAliasse As Long) As Boolean

    DatenAenderungenErmitteln = False

    If WerteVerschieden( _
            wsDash.Cells(dashZeile, cSpieler).Value, _
            wsUpdate.Cells(updateZeile, UPDATE_SPIELER_SPALTE).Value) Then

        DatenAenderungenErmitteln = True
        Exit Function

    End If

    If WerteVerschieden( _
            wsDash.Cells(dashZeile, cRang).Value, _
            wsUpdate.Cells(updateZeile, UPDATE_RANG_SPALTE).Value) Then

        DatenAenderungenErmitteln = True
        Exit Function

    End If

    If WerteVerschieden( _
            wsDash.Cells(dashZeile, cH).Value, _
            wsUpdate.Cells(updateZeile, UPDATE_H_SPALTE).Value) Then

        DatenAenderungenErmitteln = True
        Exit Function

    End If

    If WerteVerschieden( _
            wsDash.Cells(dashZeile, cMacht).Value, _
            wsUpdate.Cells(updateZeile, UPDATE_MACHT_SPALTE).Value) Then

        DatenAenderungenErmitteln = True
        Exit Function

    End If

    If WerteVerschieden( _
            wsDash.Cells(dashZeile, cG).Value, _
            wsUpdate.Cells(updateZeile, UPDATE_G_SPALTE).Value) Then

        DatenAenderungenErmitteln = True
        Exit Function

    End If

    If WerteVerschieden( _
            wsDash.Cells(dashZeile, cM).Value, _
            wsUpdate.Cells(updateZeile, UPDATE_M_SPALTE).Value) Then

        DatenAenderungenErmitteln = True
        Exit Function

    End If

    If WerteVerschieden( _
            wsDash.Cells(dashZeile, cS).Value, _
            wsUpdate.Cells(updateZeile, UPDATE_S_SPALTE).Value) Then

        DatenAenderungenErmitteln = True
        Exit Function

    End If

    If WerteVerschieden( _
            wsDash.Cells(dashZeile, cE).Value, _
            wsUpdate.Cells(updateZeile, UPDATE_E_SPALTE).Value) Then

        DatenAenderungenErmitteln = True
        Exit Function

    End If

    If WerteVerschieden( _
            wsDash.Cells(dashZeile, cJoined).Value, _
            wsUpdate.Cells(updateZeile, UPDATE_JOINED_SPALTE).Value) Then

        DatenAenderungenErmitteln = True
        Exit Function

    End If

    If WerteVerschieden( _
            wsDash.Cells(dashZeile, cAliasse).Value, _
            wsUpdate.Cells(updateZeile, UPDATE_ALIASE_SPALTE).Value) Then

        DatenAenderungenErmitteln = True
        Exit Function

    End If

End Function

' ============================================================
' WERTE VERSCHIEDEN
' ============================================================

Private Function WerteVerschieden( _
    ByVal wert1 As Variant, _
    ByVal wert2 As Variant) As Boolean

    Dim text1 As String
    Dim text2 As String

    If IsError(wert1) Or IsError(wert2) Then

        If IsError(wert1) And IsError(wert2) Then
            WerteVerschieden = False
        Else
            WerteVerschieden = True
        End If

        Exit Function

    End If

    text1 = SichererText(wert1)
    text2 = SichererText(wert2)

    WerteVerschieden = (text1 <> text2)

End Function

' ============================================================
' DATEN ÜBERNEHMEN
' ============================================================

Private Sub DatenUebernehmen( _
    ByVal wsDash As Worksheet, _
    ByVal wsUpdate As Worksheet, _
    ByVal dashZeile As Long, _
    ByVal updateZeile As Long, _
    ByVal cSpieler As Long, _
    ByVal cRang As Long, _
    ByVal cH As Long, _
    ByVal cMacht As Long, _
    ByVal cG As Long, _
    ByVal cM As Long, _
    ByVal cS As Long, _
    ByVal cE As Long, _
    ByVal cJoined As Long, _
    ByVal cAliasse As Long)

    wsDash.Cells( _
        dashZeile, _
        cSpieler).Value = _
        wsUpdate.Cells( _
            updateZeile, _
            UPDATE_SPIELER_SPALTE).Value

    wsDash.Cells( _
        dashZeile, _
        cRang).Value = _
        wsUpdate.Cells( _
            updateZeile, _
            UPDATE_RANG_SPALTE).Value

    wsDash.Cells( _
        dashZeile, _
        cH).Value = _
        wsUpdate.Cells( _
            updateZeile, _
            UPDATE_H_SPALTE).Value

    wsDash.Cells( _
        dashZeile, _
        cMacht).Value = _
        wsUpdate.Cells( _
            updateZeile, _
            UPDATE_MACHT_SPALTE).Value

    wsDash.Cells( _
        dashZeile, _
        cG).Value = _
        wsUpdate.Cells( _
            updateZeile, _
            UPDATE_G_SPALTE).Value

    wsDash.Cells( _
        dashZeile, _
        cM).Value = _
        wsUpdate.Cells( _
            updateZeile, _
            UPDATE_M_SPALTE).Value

    wsDash.Cells( _
        dashZeile, _
        cS).Value = _
        wsUpdate.Cells( _
            updateZeile, _
            UPDATE_S_SPALTE).Value

    wsDash.Cells( _
        dashZeile, _
        cE).Value = _
        wsUpdate.Cells( _
            updateZeile, _
            UPDATE_E_SPALTE).Value

    wsDash.Cells( _
        dashZeile, _
        cJoined).Value = _
        wsUpdate.Cells( _
            updateZeile, _
            UPDATE_JOINED_SPALTE).Value

    wsDash.Cells( _
        dashZeile, _
        cAliasse).Value = _
        wsUpdate.Cells( _
            updateZeile, _
            UPDATE_ALIASE_SPALTE).Value

End Sub

' ============================================================
' T9 BERECHNEN
' ============================================================

Private Sub T9AlleSpielerNeuBerechnen( _
    ByVal wsDash As Worksheet, _
    ByVal cG As Long, _
    ByVal cS As Long, _
    ByVal cT9 As Long, _
    ByVal zeileSpieler1 As Long, _
    ByVal letzteZeile As Long)

    Dim i As Long
    Dim gWert As Double
    Dim sWert As Double

    If cG <= 0 Then Exit Sub
    If cS <= 0 Then Exit Sub
    If cT9 <= 0 Then Exit Sub

    If letzteZeile < zeileSpieler1 Then Exit Sub

    For i = zeileSpieler1 To letzteZeile

        gWert = 0
        sWert = 0

        If NumerischerWert( _
                wsDash.Cells(i, cG).Value, _
                gWert) And _
           NumerischerWert( _
                wsDash.Cells(i, cS).Value, _
                sWert) Then

            If gWert = 9 And sWert = 9 Then

                wsDash.Cells( _
                    i, _
                    cT9).Value = 1

            Else

                wsDash.Cells( _
                    i, _
                    cT9).Value = 0

            End If

        Else

            wsDash.Cells( _
                i, _
                cT9).Value = 0

        End If

    Next i

End Sub

' ============================================================
' T9 SEIT AKTUALISIEREN
' ============================================================

Private Sub T9SeitAktualisieren( _
    ByVal wsDash As Worksheet, _
    ByVal dashZeile As Long, _
    ByVal cT9 As Long, _
    ByVal cT9Seit As Long)

    Dim alterT9 As Variant
    Dim neuerT9 As Variant

    alterT9 = wsDash.Cells( _
        dashZeile, _
        cT9Seit).Value

    neuerT9 = wsDash.Cells( _
        dashZeile, _
        cT9).Value

    If IstEins(neuerT9) Then

        If IsError(alterT9) Then

            wsDash.Cells( _
                dashZeile, _
                cT9Seit).Value = GetUTCNow()

            wsDash.Cells( _
                dashZeile, _
                cT9Seit).NumberFormat = _
                "dd.mm.yyyy hh:mm"

        ElseIf Len(SichererText(alterT9)) = 0 Then

            wsDash.Cells( _
                dashZeile, _
                cT9Seit).Value = GetUTCNow()

            wsDash.Cells( _
                dashZeile, _
                cT9Seit).NumberFormat = _
                "dd.mm.yyyy hh:mm"

        End If
    Else

        wsDash.Cells( _
            dashZeile, _
            cT9Seit).ClearContents

    End If

End Sub

' ============================================================
' NEUER SPIELER - FARBE
' ============================================================

Private Sub NeueSpielerFarbeSetzen( _
    ByVal ws As Worksheet, _
    ByVal zeile As Long, _
    ByVal cSpieler As Long, _
    ByVal cLetzterNOK As Long)
    Dim cEnde As Long

    cEnde = cLetzterNOK

    If cSpieler <= 0 Then Exit Sub
    If cEnde < cSpieler Then Exit Sub

    ws.Range( _
        ws.Cells(zeile, cSpieler), _
        ws.Cells(zeile, cEnde)).Font.Color = _
        RGB(0, 0, 255)

End Sub

' ============================================================
' INAKTIVE ZEILE - FARBE
' ============================================================

Private Sub ZeileInaktivFarbe( _
    ByVal ws As Worksheet, _
    ByVal zeile As Long, _
    ByVal cSpieler As Long, _
    ByVal cEnde As Long)

    If cSpieler <= 0 Then Exit Sub
    If cEnde <= 0 Then Exit Sub
    If cEnde < cSpieler Then Exit Sub

    ws.Range( _
        ws.Cells(zeile, cSpieler), _
        ws.Cells(zeile, cEnde)).Font.Color = _
        RGB(150, 150, 150)

End Sub

' ============================================================
' AUTOMATISCHE ZEILENFARBE
' ============================================================

Private Sub ZeileAutomatischeFarbe( _
    ByVal ws As Worksheet, _
    ByVal zeile As Long, _
    ByVal cSpieler As Long, _
    ByVal cBasisdaten As Long, _
    ByVal cLetzterNOK As Long)

    Dim cEnde As Long

    cEnde = cLetzterNOK

    If cSpieler <= 0 Then Exit Sub
    If cEnde < cSpieler Then Exit Sub

    If IstAktiv( _
            ws.Cells( _
                zeile, _
                cBasisdaten).Value) Then

        ws.Range( _
            ws.Cells(zeile, cSpieler), _
            ws.Cells(zeile, cEnde)).Font.ColorIndex = _
            xlAutomatic

    Else

        ZeileInaktivFarbe _
            ws, _
            zeile, _
            cSpieler, _
            cEnde

    End If

End Sub

' ============================================================
' NEUE SPIELERFARBE ZURÜCKSETZEN
' ============================================================

Private Sub NeueSpielerFarbeZuruecksetzen( _
    ByVal ws As Worksheet, _
    ByVal cInternNeu As Long, _
    ByVal cSpieler As Long, _
    ByVal cLetzterNOK As Long, _
    ByVal letzteZeile As Long, _
    ByVal zeileSpieler1 As Long)

    Dim i As Long
    Dim zeitpunkt As Date
    Dim wert As Variant
    Dim cEnde As Long

    If letzteZeile < zeileSpieler1 Then Exit Sub
    If cInternNeu <= 0 Then Exit Sub
    If cSpieler <= 0 Then Exit Sub

    cEnde = cLetzterNOK

    If cEnde < cSpieler Then Exit Sub

    zeitpunkt = GetUTCNow()

    For i = zeileSpieler1 To letzteZeile

        wert = ws.Cells(i, cInternNeu).Value

        If Not IsError(wert) Then

            If Len(SichererText(wert)) > 0 Then

                If IsDate(wert) Then

                    If CDate(wert) < zeitpunkt Then

                        ws.Range( _
                            ws.Cells(i, cSpieler), _
                            ws.Cells(i, cEnde)).Font.ColorIndex = _
                            xlAutomatic

                        ws.Cells( _
                            i, _
                            cInternNeu).ClearContents

                    End If

                End If

            End If

        End If

    Next i

End Sub

' ============================================================
' ENDGÜLTIGE SCHRIFTFARBEN
' ============================================================

Private Sub EndgueltigeSchriftfarbenSetzen( _
    ByVal ws As Worksheet, _
    ByVal cBasisdaten As Long, _
    ByVal letzteZeile As Long, _
    ByVal neueZeilen As Collection, _
    ByVal cSpieler As Long, _
    ByVal cLetzterNOK As Long, _
    ByVal zeileSpieler1 As Long)

    Dim i As Long
    Dim cEnde As Long

    cEnde = cLetzterNOK

    If letzteZeile < zeileSpieler1 Then Exit Sub
    If cEnde < cSpieler Then Exit Sub

    For i = zeileSpieler1 To letzteZeile

        If IstAktiv( _
                ws.Cells( _
                    i, _
                    cBasisdaten).Value) Then

            If CollectionKeyExistiert( _
                    neueZeilen, _
                    CStr(i)) Then

                ' Neue Spielerfarbe bleibt zunächst bestehen.

            Else

                ' Nur die Basisdaten werden hier auf die automatische
                ' Schriftfarbe zurückgesetzt. Die nachfolgenden Bereiche
                ' (z. B. Carter/Chests) haben ihre eigenen Prüfungen und
                ' Markierungslogiken, die unmittelbar davor bereits die
                ' korrekten Einzelzellfarben gesetzt haben.
                ws.Range( _
                    ws.Cells(i, cSpieler), _
                    ws.Cells(i, cBasisdaten)).Font.ColorIndex = _
                    xlAutomatic

            End If

        Else

            ZeileInaktivFarbe _
                ws, _
                i, _
                cSpieler, _
                cLetzterNOK

        End If

    Next i

    ' T9 darf bei einem gerade neu angelegten Spieler
    ' die blaue Farbe NICHT überschreiben.
    T9SchriftfarbenSetzen _
        ws, _
        cBasisdaten, _
        cSpieler, _
        cLetzterNOK, _
        zeileSpieler1, _
        letzteZeile, _
        neueZeilen

    NOKFarbenSetzen _
        ws, _
        cBasisdaten, _
        cLetzterNOK, _
        zeileSpieler1, _
        letzteZeile

    ' Inaktive Spieler haben immer höchste Priorität:
    ' grau überdeckt auch T9- und NOK-Farben.
    For i = zeileSpieler1 To letzteZeile

        If Not IstAktiv( _
                ws.Cells( _
                    i, _
                    cBasisdaten).Value) Then

            ZeileInaktivFarbe _
                ws, _
                i, _
                cSpieler, _
                cEnde

        End If

    Next i

End Sub

' ============================================================
' T9 SCHRIFTFARBEN SETZEN
' ============================================================

Private Sub T9SchriftfarbenSetzen( _
    ByVal ws As Worksheet, _
    ByVal cBasisdaten As Long, _
    ByVal cSpieler As Long, _
    ByVal cLetzterNOK As Long, _
    ByVal zeileSpieler1 As Long, _
    ByVal letzteZeile As Long, _
    ByVal neueZeilen As Collection)

    Dim rngT9 As Range
    Dim cT9 As Long
    Dim i As Long

    If cBasisdaten <= 0 Then Exit Sub
    If cSpieler <= 0 Then Exit Sub
    If cLetzterNOK <= 0 Then Exit Sub

    If letzteZeile < zeileSpieler1 Then Exit Sub

    Set rngT9 = _
        NameManagerRangeErmitteln( _
            ws, _
            NAME_LEG1_T9)

    If rngT9 Is Nothing Then Exit Sub

    cT9 = rngT9.Column

    If cT9 <= cSpieler Then Exit Sub
    If cT9 >= cBasisdaten Then Exit Sub
    If cT9 >= cLetzterNOK Then Exit Sub

    For i = zeileSpieler1 To letzteZeile

        ' Ein neu angelegter Spieler bleibt auch in LEG1_T9 blau.
        If CollectionKeyExistiert( _
                neueZeilen, _
                CStr(i)) Then

            ' Nichts ändern.

        ElseIf IstEins( _
                ws.Cells(i, cT9).Value) Then

            ws.Cells( _
                i, _
                cT9).Font.Color = _
                RGB(255, 0, 0)

        Else

            ws.Cells( _
                i, _
                cT9).Font.ColorIndex = _
                xlAutomatic

        End If

    Next i

End Sub

' ============================================================
' SORTIERUNG ERFORDERLICH
' ============================================================

Private Function SortierungErforderlich( _
    ByVal ws As Worksheet, _
    ByVal cInternSortierung As Long, _
    ByVal zeileSpieler1 As Long, _
    ByVal letzteZeile As Long) As Boolean

    Dim i As Long

    SortierungErforderlich = False

    If cInternSortierung <= 0 Then Exit Function
    If letzteZeile < zeileSpieler1 Then Exit Function

    For i = zeileSpieler1 To letzteZeile

        If IstEins( _
                ws.Cells( _
                    i, _
                    cInternSortierung).Value) Then

            SortierungErforderlich = True
            Exit Function

        End If

    Next i

End Function

' ============================================================
' SORTIERKENNZEICHEN LÖSCHEN
' ============================================================

Private Sub SortierungskennzeichenLoeschen( _
    ByVal ws As Worksheet, _
    ByVal cInternSortierung As Long, _
    ByVal zeileSpieler1 As Long, _
    ByVal letzteZeile As Long)

    If cInternSortierung <= 0 Then Exit Sub
    If letzteZeile < zeileSpieler1 Then Exit Sub

    ws.Range( _
        ws.Cells(zeileSpieler1, cInternSortierung), _
        ws.Cells(letzteZeile, cInternSortierung)).ClearContents

End Sub

' ============================================================
' CARTER-BEREICH PRÜFEN
' ============================================================

Private Sub CarterBereichPruefen( _
    ByVal wsDash As Worksheet, _
    ByVal cBasisdaten As Long, _
    ByVal zeileSpieler1 As Long, _
    ByVal letzteZeile As Long)

    Dim cCarterNOK As Long
    Dim cBereichStart As Long
    Dim cBereichEnde As Long

    Dim zeileSchwelle2 As Long

    Dim i As Long
    Dim c As Long

    Dim schwelle As Double
    Dim wert As Double
    Dim anzahlUnterSchwelle As Long
    Dim schwelleGueltig As Boolean
    Dim spielerAktiv As Boolean
    If cBasisdaten <= 0 Then Exit Sub

    cCarterNOK = DashboardSpalte(wsDash, CARTER_NOK_Kopf)
    If cCarterNOK <= 0 Then Exit Sub

    If letzteZeile < zeileSpieler1 Then Exit Sub

    If Not BereichsGrenzenFuerNOKMarkerErmitteln( _
            wsDash, _
            cBasisdaten, _
            cCarterNOK, _
            cBereichStart, _
            cBereichEnde) Then

        Exit Sub

    End If

    zeileSchwelle2 = _
        DashboardZeileErmitteln( _
            wsDash, _
            NAME_DASHBOARD_SCHWELLE2)

    If zeileSchwelle2 <= 0 Then Exit Sub

    If cBereichStart <= 0 Or _
       cBereichStart > cBereichEnde Then

        For i = zeileSpieler1 To letzteZeile

            wsDash.Cells( _
                i, _
                cCarterNOK).Value = 0

            wsDash.Cells( _
                i, _
                cCarterNOK).Font.ColorIndex = _
                xlAutomatic

        Next i

        Exit Sub

    End If

    ' Alte rote Markierungen und alte Zeile-1-Zählungen entfernen.
    wsDash.Range( _
        wsDash.Cells(zeileSpieler1, cBereichStart), _
        wsDash.Cells(letzteZeile, cBereichEnde)).Font.ColorIndex = _
        xlAutomatic

    wsDash.Range( _
        wsDash.Cells(1, cBereichStart), _
        wsDash.Cells(1, cBereichEnde)).ClearContents

    ' Carter_NOK enthält pro Spieler die Anzahl der Carter-Werte,
    ' die den jeweiligen Schwellenwert nicht erreichen.
    ' In Zeile 1 wird je Carter-Spalte gezählt, wie viele Spieler
    ' in dieser Spalte NOK sind.
    For i = zeileSpieler1 To letzteZeile

        anzahlUnterSchwelle = 0
        spielerAktiv = IstAktiv( _
            wsDash.Cells(i, cBasisdaten).Value)

        For c = cBereichStart To cBereichEnde

            schwelleGueltig = _
                NumerischerWert( _
                    wsDash.Cells( _
                        zeileSchwelle2, _
                        c).Value, _
                    schwelle)

            If schwelleGueltig Then

                If NumerischerWert( _
                        wsDash.Cells( _
                            i, _
                            c).Value, _
                        wert) Then

                    If wert < schwelle Then

                        anzahlUnterSchwelle = _
                            anzahlUnterSchwelle + 1

                        If spielerAktiv Then
                            wsDash.Cells(1, c).Value = _
                                CLng(wsDash.Cells(1, c).Value) + 1
                        End If

                        wsDash.Cells( _
                            i, _
                            c).Font.Color = _
                            RGB(255, 0, 0)

                    Else

                        wsDash.Cells( _
                            i, _
                            c).Font.ColorIndex = _
                            xlAutomatic

                    End If

                Else

                    ' Nicht numerische / leere Werte erreichen
                    ' den Schwellenwert nicht und werden daher
                    ' als NOK gezählt und rot markiert.
                    anzahlUnterSchwelle = _
                        anzahlUnterSchwelle + 1

                    If spielerAktiv Then
                        wsDash.Cells(1, c).Value = _
                            CLng(wsDash.Cells(1, c).Value) + 1
                    End If

                    wsDash.Cells( _
                        i, _
                        c).Font.Color = _
                        RGB(255, 0, 0)

                End If

            Else

                ' Ohne gültigen Schwellenwert kann die Spalte
                ' nicht geprüft werden. Bestehende Markierung
                ' bleibt deshalb entfernt.
                wsDash.Cells( _
                    i, _
                    c).Font.ColorIndex = _
                    xlAutomatic

            End If

        Next c

        wsDash.Cells( _
            i, _
            cCarterNOK).Value = _
            anzahlUnterSchwelle

        If anzahlUnterSchwelle > 0 Then

            wsDash.Cells( _
                i, _
                cCarterNOK).Font.Color = _
                RGB(255, 0, 0)

        Else

            wsDash.Cells( _
                i, _
                cCarterNOK).Font.ColorIndex = _
                xlAutomatic

        End If

    Next i

End Sub

' ============================================================
' NOK-FARBEN GENERISCH SETZEN
' ============================================================

Private Sub NOKFarbenSetzen( _
    ByVal ws As Worksheet, _
    ByVal cBasisdaten As Long, _
    ByVal cLetzterNOK As Long, _
    ByVal zeileSpieler1 As Long, _
    ByVal letzteZeile As Long)

    Dim markerSpalten() As Long
    Dim markerAnzahl As Long
    Dim i As Long
    Dim j As Long
    Dim wert As Double
    Dim cNOK As Long

    If cBasisdaten <= 0 Then Exit Sub
    If cLetzterNOK <= 0 Then Exit Sub
    If letzteZeile < zeileSpieler1 Then Exit Sub

    markerAnzahl = NOKMarkerSpaltenErmitteln( _
        ws, _
        markerSpalten)

    If markerAnzahl <= 0 Then Exit Sub

    For j = 1 To markerAnzahl

        cNOK = markerSpalten(j)

        If cNOK <= cLetzterNOK Then

            For i = zeileSpieler1 To letzteZeile

                If NumerischerWert( _
                        ws.Cells(i, cNOK).Value, _
                        wert) Then

                    If wert > 0 Then
                        ws.Cells(i, cNOK).Font.Color = RGB(255, 0, 0)
                    Else
                        ws.Cells(i, cNOK).Font.ColorIndex = xlAutomatic
                    End If

                Else

                    ws.Cells(i, cNOK).Font.ColorIndex = xlAutomatic

                End If

            Next i

        End If

    Next j

End Sub

' ============================================================
' CHESTS-SPALTEN PRÜFEN
' ============================================================

Private Sub ChestsSpaltenPruefen( _
    ByVal wsDash As Worksheet, _
    ByVal cBasisdaten As Long, _
    ByVal zeileSummen As Long, _
    ByVal zeileSchwelle2 As Long, _
    ByVal zeileBezuege As Long, _
    ByVal cSpieler As Long, _
    ByVal zeileSpieler1 As Long, _
    ByVal letzteZeileDash As Long, _
    ByVal wsLog As Worksheet)

    Dim cChestsNOK As Long
    Dim cBereichStart As Long
    Dim cBereichEnde As Long
    Dim c As Long
    Dim cVorher As Long
    Dim cNachher As Long
    Dim cEinfügen As Long
    Dim cSuche As Long

    Dim kopf As String
    Dim datum As Date
    Dim datumText As String
    Dim nameText As String
    Dim quellSheetName As String
    Dim quellDatum As Date

    Dim wbChests As Workbook
    Dim wbWarBereitsOffen As Boolean
    Dim wsQuelle As Worksheet
    Dim datumGefunden As Boolean
    Dim vorhandeneChestsSpalten As Boolean

    If cBasisdaten <= 0 Then Exit Sub

    cChestsNOK = DashboardSpalte(wsDash, CHESTS_NOK_Kopf)
    If cChestsNOK <= 0 Then Exit Sub

    If zeileSummen <= 0 Then Exit Sub
    If zeileSchwelle2 <= 0 Then Exit Sub
    If zeileBezuege <= 0 Then Exit Sub

    Set wbChests = ChestsArbeitsmappeOeffnen( _
        wsLog, _
        wbWarBereitsOffen)

    If wbChests Is Nothing Then

        Err.Raise 1004, , _
            "Die Datei '" & _
            CHESTS_DATEI & _
            "' konnte nicht geöffnet werden."

    End If

    ' ------------------------------------------------------------
    ' CHESTS-BEREICH VOLLSTÄNDIG DYNAMISCH AUFBAUEN
    '
    ' Jedes gültige Datumsblatt der Quelldatei muss als eigene
    ' Dashboard-Spalte vorhanden sein.
    '
    ' Bereits vorhandene historische Spalten bleiben erhalten.
    ' Neue Quellblätter werden automatisch chronologisch vor
    ' Chests_NOK eingefügt.
    ' ------------------------------------------------------------

    For Each wsQuelle In wbChests.Worksheets

        quellSheetName = SichererText(wsQuelle.Name)

        If ChestsDatumGueltig( _
                quellSheetName, _
                quellDatum) Then

            cChestsNOK = DashboardSpalte( _
                wsDash, _
                CHESTS_NOK_Kopf)

            cBereichStart = cChestsNOK
            cBereichStart = cBereichStart - 1

            If cBereichStart > 0 Then

                datumGefunden = False

                For cSuche = 1 To cBereichStart

                    If ChestsDatumGueltig( _
                            SichererText( _
                                wsDash.Cells( _
                                    zeileBezuege, _
                                    cSuche).Value), _
                            datum) Then

                        If Format$( _
                                datum, _
                                "yyyymmdd") = _ 
                           quellSheetName Then

                            datumGefunden = True
                            Exit For

                        End If

                    End If

                Next cSuche

                If Not datumGefunden Then

                    cEinfügen = cChestsNOK

                    For cSuche = cBereichStart To 1 Step -1

                        If ChestsDatumGueltig( _
                                SichererText( _
                                    wsDash.Cells( _
                                        zeileBezuege, _
                                        cSuche).Value), _
                                datum) Then

                            If datum < quellDatum Then

                                cEinfügen = cSuche + 1
                                Exit For

                            End If

                        End If

                    Next cSuche

                    ' Wenn kein früheres Datum vorhanden ist, wird
                    ' vor der ersten vorhandenen Chests-Spalte
                    ' eingefügt.
                    If cEinfügen = cChestsNOK Then

                        For cSuche = 1 To cBereichStart

                            If ChestsDatumGueltig( _
                                    SichererText( _
                                        wsDash.Cells( _
                                            zeileBezuege, _
                                            cSuche).Value), _
                                    datum) Then

                                cEinfügen = cSuche
                                Exit For

                            End If

                        Next cSuche

                    End If

                    ' ------------------------------------------------
                    ' NEUE SPALTE EINFÜGEN
                    '
                    ' Die Nachbarspalte liefert Formatierung und
                    ' Schwellenwert. Spielerwerte und alte Zählung
                    ' werden anschließend entfernt.
                    ' ------------------------------------------------

                    cVorher = cEinfügen - 1
                    cNachher = cEinfügen

                    wsDash.Columns(cEinfügen).Insert _
                        Shift:=xlToRight, _
                        CopyOrigin:=xlFormatFromLeftOrAbove

                    If cVorher >= cBereichStart And _
                       cVorher < cEinfügen Then

                        wsDash.Columns(cVorher).Copy
                        wsDash.Columns(cEinfügen).PasteSpecial _
                            Paste:=xlPasteFormats

                        wsDash.Cells( _
                            zeileSchwelle2, _
                            cEinfügen).Value = _
                            wsDash.Cells( _
                                zeileSchwelle2, _
                                cVorher).Value

                    ElseIf cNachher < cChestsNOK Then

                        wsDash.Columns(cNachher).Copy
                        wsDash.Columns(cEinfügen).PasteSpecial _
                            Paste:=xlPasteFormats

                        wsDash.Cells( _
                            zeileSchwelle2, _
                            cEinfügen).Value = _
                            wsDash.Cells( _
                                zeileSchwelle2, _
                                cNachher).Value

                    End If

                    Application.CutCopyMode = False

                    wsDash.Cells( _
                        1, _
                        cEinfügen).ClearContents

                    wsDash.Cells( _
                        zeileSummen, _
                        cEinfügen).ClearContents

                    wsDash.Cells( _
                        zeileBezuege, _
                        cEinfügen).Value = _
                        quellSheetName

                    nameText = _
                        "Chests_" & _
                        quellSheetName

                    ChestsNameSicherstellen _
                        wsDash, _
                        cEinfügen, _
                        zeileSummen, _
                        nameText

                End If

            End If

        End If

    Next wsQuelle

    ' Nach allen Einfügungen Marker und Bereichsgrenzen erneut
    ' ausschließlich über den Namensmanager ermitteln.
    cChestsNOK = DashboardSpalte( _
        wsDash, _
        CHESTS_NOK_Kopf)

    If cChestsNOK <= 0 Then GoTo ChestsBeenden

    If Not BereichsGrenzenFuerNOKMarkerErmitteln( _
            wsDash, _
            cBasisdaten, _
            cChestsNOK, _
            cBereichStart, _
            cBereichEnde) Then

        GoTo ChestsBeenden

    End If

    ' Vor der vollständigen Chests-Prüfung alte Zählungen
    ' zurücksetzen.
    If letzteZeileDash >= zeileSpieler1 Then

        wsDash.Range( _
            wsDash.Cells(zeileSpieler1, cChestsNOK), _
            wsDash.Cells(letzteZeileDash, cChestsNOK)).Value = 0

        wsDash.Range( _
            wsDash.Cells(zeileSpieler1, cChestsNOK), _
            wsDash.Cells(letzteZeileDash, cChestsNOK)).Font.ColorIndex = _
            xlAutomatic

    End If

    If cBereichStart <= cBereichEnde Then

        For c = cBereichStart To cBereichEnde

            kopf = SichererText( _
                wsDash.Cells( _
                    zeileBezuege, _
                    c).Value)

            If ChestsDatumGueltig( _
                    kopf, _
                    datum) Then

                datumText = Format$( _
                    datum, _
                    "yyyymmdd")

                nameText = _
                    "Chests_" & _
                    datumText

                ChestsNameSicherstellen _
                    wsDash, _
                    c, _
                    zeileSummen, _
                    nameText

                ChestsSpalteImportierenUndPruefen _
                    wsDash, _
                    wbChests, _
                    c, _
                    datum, _
                    zeileSchwelle2, _
                    cSpieler, _
                    cChestsNOK, _
                    zeileSpieler1, _
                    letzteZeileDash, _
                    wsLog

            Else

                ChestsSpalteRotMarkierungLoeschen _
                    wsDash, _
                    c, _
                    zeileSpieler1, _
                    letzteZeileDash

                wsDash.Cells(1, c).ClearContents

            End If

        Next c

    End If

ChestsBeenden:

    If Not wbWarBereitsOffen Then

        On Error Resume Next

        wbChests.Close SaveChanges:=False

        On Error GoTo 0

    End If

End Sub

' ============================================================
' CHESTS-ARBEITSMAPPE ÖFFNEN
' ============================================================

Private Function ChestsArbeitsmappeOeffnen( _
    ByVal wsLog As Worksheet, _
    ByRef warBereitsOffen As Boolean) As Workbook

    Dim wb As Workbook
    Dim dateiPfad As String

    warBereitsOffen = False

    On Error Resume Next

    Set wb = Workbooks(CHESTS_DATEI)

    On Error GoTo 0

    If Not wb Is Nothing Then

        warBereitsOffen = True

        Set ChestsArbeitsmappeOeffnen = wb

        Exit Function

    End If

    If Len(ThisWorkbook.Path) = 0 Then

        LogEintrag _
            wsLog, _
            "CHESTS WARNUNG", _
            "", _
            "Die LEG1-Arbeitsmappe wurde noch nicht gespeichert. " & _
            "Der Pfad zu '" & _
            CHESTS_DATEI & _
            "' kann nicht ermittelt werden."

        Exit Function

    End If

    dateiPfad = _
        ThisWorkbook.Path & _
        Application.PathSeparator & _
        CHESTS_DATEI

    If Dir(dateiPfad) = "" Then

        LogEintrag _
            wsLog, _
            "CHESTS WARNUNG", _
            "", _
            "Die Datei '" & _
            CHESTS_DATEI & _
            "' wurde nicht gefunden: " & _
            dateiPfad

        Exit Function

    End If

    On Error GoTo OeffnenFehler

    Set wb = Workbooks.Open( _
        FileName:=dateiPfad, _
        UpdateLinks:=0, _
        ReadOnly:=True)

    Set ChestsArbeitsmappeOeffnen = wb

    Exit Function

OeffnenFehler:

    LogEintrag _
        wsLog, _
        "CHESTS WARNUNG", _
        "", _
        "Die Datei '" & _
        CHESTS_DATEI & _
        "' konnte nicht geöffnet werden. Fehler " & _
        CStr(Err.Number) & _
        ": " & _
        Err.Description

    Set ChestsArbeitsmappeOeffnen = Nothing

End Function

' ============================================================
' CHESTS-SPALTE IMPORTIEREN UND PRÜFEN
' ============================================================

Private Sub ChestsSpalteImportierenUndPruefen( _
    ByVal wsDash As Worksheet, _
    ByVal wbChests As Workbook, _
    ByVal cDash As Long, _
    ByVal datum As Date, _
    ByVal zeileSchwelle2 As Long, _
    ByVal cSpieler As Long, _
    ByVal cChestsNOK As Long, _
    ByVal zeileSpieler1 As Long, _
    ByVal letzteZeileDash As Long, _
    ByVal wsLog As Worksheet)

    Dim wsQuelle As Worksheet

    Dim sheetName As String
    Dim letzteZeileQuelle As Long
    Dim i As Long

    Dim spieler As String
    Dim key As String
    Dim wert As Variant
    Dim zahl As Double
    Dim schwelle As Double

    Dim spielerWerte As Collection

    Dim anzahlAktiveUnterSchwelle As Long
    Dim wertGefunden As Boolean

    sheetName = Format$( _
        datum, _
        "yyyymmdd")

    Set wsQuelle = Nothing

    On Error Resume Next

    Set wsQuelle = _
        wbChests.Worksheets(sheetName)

    On Error GoTo 0

    If wsQuelle Is Nothing Then

        Err.Raise 1004, , _
            "Das Chests-Quellblatt '" & _
            sheetName & _
            "' wurde in '" & _
            CHESTS_DATEI & _
            "' nicht gefunden."

    End If

    If Not NumerischerWert( _
            wsDash.Cells( _
                zeileSchwelle2, _
                cDash).Value, _
            schwelle) Then

        Err.Raise 1004, , _
            "Der Schwellenwert in " & _
            wsDash.Cells( _
                zeileSchwelle2, _
                cDash).Address(False, False) & _
            " ist kein numerischer Wert."

    End If

    ChestsSpalteRotMarkierungLoeschen _
        wsDash, _
        cDash, _
        zeileSpieler1, _
        letzteZeileDash

    Set spielerWerte = New Collection

    letzteZeileQuelle = _
        wsQuelle.Cells( _
            wsQuelle.Rows.Count, _
            CHESTS_SPIELER_SPALTE).End(xlUp).Row

    If letzteZeileQuelle >= CHESTS_DATENSTART Then

        For i = CHESTS_DATENSTART To letzteZeileQuelle

            spieler = SichererText( _
                wsQuelle.Cells( _
                    i, _
                    CHESTS_SPIELER_SPALTE).Value)

            If Len(spieler) > 0 Then

                key = SpielerKey(spieler)

                If CollectionKeyExistiert( _
                        spielerWerte, _
                        key) Then

                    LogEintrag _
                        wsLog, _
                        "DUPLIKAT CHESTS", _
                        spieler, _
                        "Spieler kommt mehrfach im Quellblatt '" & _
                        sheetName & _
                        "' vor. Der erste Wert wird verwendet."

                Else

                    wert = _
                        wsQuelle.Cells( _
                            i, _
                            CHESTS_WERT_SPALTE).Value

                    spielerWerte.Add _
                        wert, _
                        key

                End If
            End If

        Next i

    End If

    For i = zeileSpieler1 To letzteZeileDash

        spieler = SichererText( _
            wsDash.Cells( _
                i, _
                cSpieler).Value)

        If Len(spieler) > 0 Then

            key = SpielerKey(spieler)

            wertGefunden = _
                CollectionKeyExistiert( _
                    spielerWerte, _
                    key)

            If wertGefunden Then

                wert = CollectionWert( _
                    spielerWerte, _
                    key)

                ' Den Quellenwert immer in die entsprechende
                ' Dashboard-Zelle übernehmen.
                wsDash.Cells( _
                    i, _
                    cDash).Value = wert

                If NumerischerWert( _
                        wert, _
                        zahl) Then

                    If zahl < schwelle Then

                        If IstAktiv( _
                                wsDash.Cells( _
                                    i, _
                                    DashboardSpalte(wsDash, LEG1_Basisdaten_Kopf)).Value) Then

                            anzahlAktiveUnterSchwelle = _
                                anzahlAktiveUnterSchwelle + 1

                        End If

                        wsDash.Cells( _
                            i, _
                            cDash).Font.Color = _
                            RGB(255, 0, 0)

                    Else

                        wsDash.Cells( _
                            i, _
                            cDash).Font.ColorIndex = _
                            xlAutomatic

                    End If

                Else

                    wsDash.Cells( _
                        i, _
                        cDash).Font.ColorIndex = _
                        xlAutomatic

                End If

            Else

                ' Kein Quellwert: keine Wertübernahme und kein NOK.
                wsDash.Cells( _
                    i, _
                    cDash).ClearContents

            End If

        End If

    Next i

    wsDash.Cells(1, cDash).Value = _
        anzahlAktiveUnterSchwelle

    ' Chests_NOK enthält die Anzahl der Chests-Spalten unterhalb
    ' des Schwellenwerts für diesen Spieler.
    For i = zeileSpieler1 To letzteZeileDash

        spieler = SichererText( _
            wsDash.Cells(i, cSpieler).Value)

        If Len(spieler) > 0 Then

            key = SpielerKey(spieler)

            If CollectionKeyExistiert(spielerWerte, key) Then

                wert = CollectionWert(spielerWerte, key)

                If NumerischerWert(wert, zahl) Then

                    If zahl < schwelle Then
                        wsDash.Cells(i, cChestsNOK).Value = _
                            CLng(wsDash.Cells(i, cChestsNOK).Value) + 1
                    End If

                End If

            End If

        End If

    Next i

End Sub

' ============================================================
' CHESTS-ROT-MARKIERUNG EINER SPALTE LÖSCHEN
' ============================================================

Private Sub ChestsSpalteRotMarkierungLoeschen( _
    ByVal ws As Worksheet, _
    ByVal cDash As Long, _
    ByVal zeileSpieler1 As Long, _
    ByVal letzteZeileDash As Long)

    If cDash <= 0 Then Exit Sub
    If letzteZeileDash < zeileSpieler1 Then Exit Sub

    ws.Range( _
        ws.Cells(zeileSpieler1, cDash), _
        ws.Cells(letzteZeileDash, cDash)).Font.ColorIndex = _
        xlAutomatic

End Sub

' ============================================================
' CHESTS_NOK ZURÜCKSETZEN
' ============================================================

Private Sub ChestsNOKAlleSpielerZuruecksetzen( _
    ByVal wsDash As Worksheet, _
    ByVal cChestsNOK As Long, _
    ByVal zeileSpieler1 As Long, _
    ByVal letzteZeileDash As Long)

    If cChestsNOK <= 0 Then Exit Sub
    If letzteZeileDash < zeileSpieler1 Then Exit Sub

    wsDash.Range( _
        wsDash.Cells(zeileSpieler1, cChestsNOK), _
        wsDash.Cells(letzteZeileDash, cChestsNOK)).Value = 0

    wsDash.Range( _
        wsDash.Cells(zeileSpieler1, cChestsNOK), _
        wsDash.Cells(letzteZeileDash, cChestsNOK)).Font.ColorIndex = _
        xlAutomatic

End Sub

' ============================================================
' CHESTS DATUM GÜLTIG
' ============================================================

Private Function ChestsDatumGueltig( _
    ByVal text As String, _
    ByRef datum As Date) As Boolean

    ChestsDatumGueltig = False

    text = Trim$(text)

    If Len(text) <> 8 Then Exit Function
    If Not IsNumeric(text) Then Exit Function

    On Error GoTo Fehler

    datum = DateSerial( _
        CLng(Left$(text, 4)), _
        CLng(Mid$(text, 5, 2)), _
        CLng(Right$(text, 2)))

    If Format$( _
            datum, _
            "yyyymmdd") <> text Then

        ChestsDatumGueltig = False
        Exit Function