Attribute VB_Name = "m_digitalerEEB"
Option Private Module

' ------------------------------------------------------------
' Spalten Namen für Blatt Stärke
' ------------------------------------------------------------
Private Const FUEST As String = "B"
Private Const BEZEICHNUNG As String = "C"
Private Const organisation As String = "D"
Private Const HERKUNFT As String = "E"
Private Const ZUG As String = "F"
Private Const TRUPP_STAFFEL As String = "G"
Private Const GRUPPE As String = "H"
Private Const PERSON As String = "I"
Private Const GERAETE_FAHRZEUGE As String = "J"
Private Const AUFTRAEGE As String = "K"
Private Const ERREICHBARKEIT As String = "L"
Private Const VERFUEGBAR_BIS As String = "M"
Private Const ABLOESUNG_DATUM As String = "N"
Private Const ANFORDERUNGS_ID As String = "O"
Private Const ZUGESAGT_FUER As String = "P"
Private Const ZUGESAGT_VON As String = "Q"
Private Const VORGESEHENE_EINHEIT As String = "R"
Private Const VORGESEHENER_AUFTRAG As String = "S"
Private Const EINGETR_ZUGEW As String = "T"
Private Const einsatzende As String = "U"
Private Const RUECKFUEHRUNG As String = "V"
Private Const BEMERKUNGEN As String = "W"
Private Const status As String = "Z"
Private Const SCHICHT As String = "AA"
Private Const ID_EEB As String = "AB"
Private Const WEIBLICH As String = "AC"
Private Const DIVERS As String = "AD"
Private Const vegetarisch As String = "AE"
Private Const vegan As String = "AF"
Private Const UEBERNACHTUNG_M As String = "AG"
Private Const UEBERNACHTUNG_W As String = "AH"
Private Const UEBERNACHTUNG_D As String = "AI"
Private Const STAERKE_FUEHRER As String = "AJ"
Private Const STAERKE_UNTERFUEHRER As String = "AK"
Private Const STAERKE_HELFER As String = "AL"
Private Const gesamtStaerke As String = "AM"

' Bogen, der bereit zum Schreiben ist
Private mBereiterBogen As TEebBogen
Private mBogenBereit As Boolean
 

' Wird bei Enter im Eingabefeld aufgerufen.
Public Sub verarbeiteEingabefeldEnter()
    Dim inhalt As String

    inhalt = Trim$(ufr_digitalerEEBEingabe.txt_eingabefeld.text)
    ' CR/LF (Zeilenende) entfernen
    Do While Len(inhalt) > 0 And (Right$(inhalt, 1) = vbCr Or Right$(inhalt, 1) = vbLf)
        inhalt = Left$(inhalt, Len(inhalt) - 1)
    Loop

    ufr_digitalerEEBEingabe.txt_eingabefeld.text = ""          ' löst selbst wieder _Change aus
    Call verarbeiteScan(inhalt)
End Sub

Public Sub eingabefeldGeaendert()
    If Len(Trim$(ufr_digitalerEEBEingabe.txt_eingabefeld.text)) = 0 Then Exit Sub
    Call StatusAnzeigen("Mit Enter bestätigen.", vbBlack)
End Sub

' Verarbeitet die übergebene EEB Eingabe
Private Sub verarbeiteScan(ByVal scanText As String)
    If Len(scanText) = 0 Then Exit Sub
 
    If Left$(Trim$(scanText), 5) = "data:" Then
        Call StatusAnzeigen("Das ist ein Bild-Link, kein " & _
            "EEB-Code. Gebe einen EEB-Link ein.", vbRed)
        Exit Sub
    End If

    On Error GoTo Fehler
 
    If EebIstSegmentCode(scanText) Then ' Bogen besteht aus mehreren QR-Codes/Segmenten
        Dim ergebnis As TEebSammelErgebnis
        ergebnis = m_digitalerEEBParsing.EebSegmentSammeln(scanText)
 
        Select Case ergebnis.status
            Case eebSammelVollstaendig
                mBereiterBogen = m_digitalerEEBParsing.EebSegmenteZuBogen()
                mBogenBereit = True
                Call StatusAnzeigen("Alle " & ergebnis.anzahl & " Teile erfasst. Bitte bestätigen.", RGB(0, 128, 0))
                ufr_digitalerEEBEingabe.cmd_ok.Enabled = True
 
            Case eebSammelGesammelt
                mBogenBereit = False
                ufr_digitalerEEBEingabe.cmd_ok.Enabled = False
                Call StatusAnzeigen("Teil " & ergebnis.teilNr & "/" & ergebnis.anzahl & " erfasst (" & _
                    ergebnis.haben & " von " & ergebnis.anzahl & "). Noch benötigt: Teil " & _
                    m_digitalerEEBParsing.EebSegmentFehlendeTeile(), RGB(0, 102, 204))
 
            Case eebSammelDuplikat
                Call StatusAnzeigen("Teil " & ergebnis.teilNr & " war schon erfasst (" & ergebnis.haben & _
                    " von " & ergebnis.anzahl & " vorhanden). Noch benötigt: Teil " & _
                    m_digitalerEEBParsing.EebSegmentFehlendeTeile(), RGB(0, 102, 204))
 
            Case eebSammelFremd
                mBogenBereit = False
                ufr_digitalerEEBEingabe.cmd_ok.Enabled = False
                Call StatusAnzeigen("Dieser Code gehört zu einer ANDEREN Serie - die vorherige, " & _
                    "unvollständige Sammlung wurde verworfen. Teil " & ergebnis.teilNr & "/" & _
                    ergebnis.anzahl & " neu erfasst. Noch benötigt: Teil " & m_digitalerEEBParsing.EebSegmentFehlendeTeile(), vbRed)
        End Select
 
    Else
        ' Einzelner, vollständiger Code: eine evtl. noch offene Teile-Sammlung
        ' ist damit hinfällig und soll den nächsten Scan nicht als "fremd" stören.
        mBereiterBogen = EebBogenLesen(scanText)
        Call m_digitalerEEBParsing.EebSammlungVerwerfen
        mBogenBereit = True
        Call StatusAnzeigen("Code vollständig erfasst. Bitte bestätigen.", RGB(0, 128, 0))
        ufr_digitalerEEBEingabe.cmd_ok.Enabled = True
    End If
 
    Exit Sub
 
Fehler:
    mBogenBereit = False
    ufr_digitalerEEBEingabe.cmd_ok.Enabled = False
    Call StatusAnzeigen("Fehler: " & Err.Description, vbRed)
End Sub
 
' Erzeugt eine neue Zeile und schreibt die Daten des Bogens
Public Sub verarbeiteOkClick()
    Dim zeile As Range

    If Not mBogenBereit Then Exit Sub

    On Error GoTo Fehler

    Application.ScreenUpdating = False
    t_staerke.Unprotect (m_util.getPasswort())
    
    Set zeile = Selection.Rows(1).EntireRow
    zeile.Copy
    t_staerke.Rows(zeile.Row).Insert Shift:=xlDown
    Application.CutCopyMode = False ' Zwischenablage leeren, kann evtl. Konflikte mit Hintergrundprozessen verursachen

    Call m_digitalerEEB.eebBogenInZeileSchreiben(zeile.Row, mBereiterBogen)
    t_staerke.Range(FUEST & zeile.Row).Select
    ' Sammlung der Teile freigeben, damit der nächste mehrteilige Bogen nicht
    ' als "andere Serie" gemeldet wird.
    Call m_digitalerEEBParsing.EebSammlungVerwerfen
    mBogenBereit = False

    Call Unload(ufr_digitalerEEBEingabe)
Aufraeumen:
    t_staerke.Protect (m_util.getPasswort())
    Application.ScreenUpdating = True
    Exit Sub

Fehler:
    Application.CutCopyMode = False ' Zwischenablage leeren, kann evtl. Konflikte mit Hintergrundprozessen verursachen
    t_staerke.Protect (m_util.getPasswort())
    Application.ScreenUpdating = True
    Call MsgBox("EEB konnte nicht eingefügt werden: " & Err.Description, vbCritical)
End Sub


' ----- Hilfsfunktionen -----

Private Sub StatusAnzeigen(ByVal text As String, ByVal farbe As Long)
    With ufr_digitalerEEBEingabe
        .lbl_status.Caption = text
        .lbl_status.ForeColor = farbe
    End With
End Sub

Public Sub zustandZuruecksetzen()
    mBogenBereit = False
    Call m_digitalerEEBParsing.EebSammlungVerwerfen
    With ufr_digitalerEEBEingabe
        .cmd_ok.Enabled = False
        .txt_eingabefeld.text = ""
        .txt_eingabefeld.SetFocus
        .lbl_status.Caption = ""
    End With
End Sub

' Extrahiert Kontaktdaten der ersten ranghöchsten Führungskraft (staerkeRolle: 2=Führer, 1=Unterführer, 0=Mannschaft).
' Format: "Vorname Nachname: Kontaktdaten"
Private Function erreichbarkeitDesFuehrers(ByRef b As TEebBogen) As String
    Dim i As Long
    Dim maxGefundeneRolle As Long
    Dim fuehrerIndex As Long
    Dim p As TEebPerson
    
    erreichbarkeitDesFuehrers = ""
    If b.personalAnzahl = 0 Then Exit Function
    
    maxGefundeneRolle = -1
    fuehrerIndex = 0
    
    For i = 1 To b.personalAnzahl
        p = b.personal(i)
        If p.staerkeRolle > maxGefundeneRolle Then
            maxGefundeneRolle = p.staerkeRolle
            fuehrerIndex = i
        End If
    Next i
    
    ' Formatiere Ausgabe mit vollständigem Namen und Kontaktdaten
    If fuehrerIndex > 0 Then
        p = b.personal(fuehrerIndex)
        Dim fullName As String
        fullName = Trim$(p.vorname & " " & p.nachname)
        If Len(fullName) > 0 And Len(Trim$(p.kontakte)) > 0 Then
            erreichbarkeitDesFuehrers = fullName & ": " & Trim$(p.kontakte)
        End If
    End If
End Function

' Bildet die Nummern aus dem digitalen EEB-Bogen auf die Abkürzungen der Organisationen in Excel ab.
Private Function organisationAbkuerzung(ByVal orgCode As Long) As String
    Select Case orgCode
        Case 1: organisationAbkuerzung = "THW"
        Case 2: organisationAbkuerzung = "FW"
        Case 3: organisationAbkuerzung = "POL"
        Case 4: organisationAbkuerzung = "BPOL"
        Case 5: organisationAbkuerzung = "DRK"
        Case 6: organisationAbkuerzung = "JUH"
        Case 7: organisationAbkuerzung = "MHD"
        Case 8: organisationAbkuerzung = "ASB"
        Case 9: organisationAbkuerzung = "DLRG"
        Case 10: organisationAbkuerzung = "BW"
        Case 11: organisationAbkuerzung = "" ' Rettungsdienst existiert in Excel nicht
        ' Zivil und Wasserorganisationen existieren in EEB nicht
        Case Else: organisationAbkuerzung = ""
    End Select
End Function

' Schreibt einen digitalen EEBBogen in die übergeben Zeile in das Stärke-Blatt.
Private Sub eebBogenInZeileSchreiben(ByVal zielzeile As Long, b As TEebBogen)
    Dim i As Long
    Dim frauen As Long
    Dim diversZahl As Long
    Dim fahrzeugText As String
    Dim p As TEebPerson
    Dim f As TEebFahrzeug

    ' --- Personalzahlen ausrechnen ---
    frauen = 0
    diversZahl = 0
    For i = 1 To b.personalAnzahl
        p = b.personal(i)
        If p.geschlecht = 1 Then frauen = frauen + 1
        If p.geschlecht = 2 Then diversZahl = diversZahl + 1
    Next i

    ' --- Abweichende Dienststellen (ab Schema 10): "Name (Dienststelle)" ---
    Dim fremdeText As String
    fremdeText = ""
    For i = 1 To b.personalAnzahl
        p = b.personal(i)
        If Len(p.dienststelle) > 0 Then
            If Len(fremdeText) > 0 Then fremdeText = fremdeText & ", "
            fremdeText = fremdeText & Trim$(p.vorname & " " & p.nachname) & " (" & p.dienststelle & ")"
        End If
    Next i
    If Len(fremdeText) > 0 Then fremdeText = "Abweichende Dienststelle: " & fremdeText

    ' --- Fahrzeuge: Typen+Kennzeichen in einen Text kombinieren ---
    fahrzeugText = ""
    For i = 1 To b.fahrzeugAnzahl
        f = b.fahrzeuge(i)
        If Len(fahrzeugText) > 0 Then fahrzeugText = fahrzeugText & "; "
        fahrzeugText = fahrzeugText & Trim$(f.typ)
        If Len(Trim$(f.kennzeichen)) > 0 Then
            fahrzeugText = fahrzeugText & " " & Trim$(f.kennzeichen)
        End If
    Next i

    With t_staerke
        .Cells(zielzeile, BEZEICHNUNG).Value = b.einheitsTyp
        .Cells(zielzeile, organisation).Value = organisationAbkuerzung(b.organisation)
        .Cells(zielzeile, HERKUNFT).Value = b.einheitOrt
        .Cells(zielzeile, ZUG).Value = ""
        .Cells(zielzeile, TRUPP_STAFFEL).Value = ""
        .Cells(zielzeile, GRUPPE).Value = ""
        .Cells(zielzeile, PERSON).Value = ""
        .Cells(zielzeile, GERAETE_FAHRZEUGE).Value = fahrzeugText
        .Cells(zielzeile, AUFTRAEGE).Value = b.ortAuftrag
        .Cells(zielzeile, ERREICHBARKEIT).Value = erreichbarkeitDesFuehrers(b)
        .Cells(zielzeile, VERFUEGBAR_BIS).Value = IIf(b.zeitraumBis <> 0, Format$(b.zeitraumBis, "dd.mm.yyyy hh:nn"), "")
        .Cells(zielzeile, ABLOESUNG_DATUM).Value = ""
        .Cells(zielzeile, ANFORDERUNGS_ID).Value = ""
        .Cells(zielzeile, ZUGESAGT_FUER).Value = IIf(b.zeitraumVon <> 0, Format$(b.zeitraumVon, "dd.mm.yyyy hh:nn"), "")
        .Cells(zielzeile, ZUGESAGT_VON).Value = ""
        .Cells(zielzeile, VORGESEHENE_EINHEIT).Value = ""
        .Cells(zielzeile, VORGESEHENER_AUFTRAG).Value = b.ortAuftrag
        .Cells(zielzeile, EINGETR_ZUGEW).Value = IIf(b.einsatzbeginn <> 0, Format$(b.einsatzbeginn, "dd.mm.yyyy hh:nn"), "")
        .Cells(zielzeile, einsatzende).Value = IIf(b.einsatzende <> 0, Format$(b.einsatzende, "dd.mm.yyyy hh:nn"), "")
        .Cells(zielzeile, RUECKFUEHRUNG).Value = ""
        .Cells(zielzeile, BEMERKUNGEN).Value = b.sonstiges & _
            IIf(Len(fremdeText) > 0, IIf(Len(b.sonstiges) > 0, "; ", "") & fremdeText, "")

        .Cells(zielzeile, status).Value = ""
        .Cells(zielzeile, SCHICHT).Value = ""
        .Cells(zielzeile, ID_EEB).Value = ""
        .Cells(zielzeile, WEIBLICH).Value = frauen
        .Cells(zielzeile, DIVERS).Value = diversZahl
        .Cells(zielzeile, vegetarisch).Value = b.vegetarisch
        .Cells(zielzeile, vegan).Value = b.vegan
        .Cells(zielzeile, UEBERNACHTUNG_M).Value = b.unterbringungM
        .Cells(zielzeile, UEBERNACHTUNG_W).Value = b.unterbringungW
        .Cells(zielzeile, UEBERNACHTUNG_D).Value = b.unterbringungD
        .Cells(zielzeile, STAERKE_FUEHRER).Value = b.staerkeFuehrer
        .Cells(zielzeile, STAERKE_UNTERFUEHRER).Value = b.staerkeUnterfuehrer
        .Cells(zielzeile, STAERKE_HELFER).Value = b.staerkeMannschaft
        ' Gesamtstärke wird automatisch berechnet
    End With
End Sub
