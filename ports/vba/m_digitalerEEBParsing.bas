Attribute VB_Name = "m_digitalerEEBParsing"
' ===========================================================================
'  Dekoder für Daten aus dem digitalen Erfassugsbogen: https://erfassungsbogen.app/
'  ---------------------------------------------------------------------
'  Liest den Text, den ein QR-Scanner (Tastatur-Emulation) aus einem
'  Einheiten-Erfassungsbogen-Code in die Zelle schreibt, und macht daraus
'  lesbare Tabellenzeilen.
'
'  Ablauf wie in der App:
'     URL/Fragment  ->  "B." + Base41  ->  Bytes
'     Bytes         ->  Magic "EEB2" bzw. "EEB2C" (Signaturkette wird
'                       übersprungen, NICHT geprüft)
'     Rest          ->  Deflate (raw, RFC 1951) entpacken
'     Binärstrom    ->  Felder des Bogens (Schema 2..10)
'
'  Abwärtskompatibilität: alle Schema-Versionen 2 bis 10 werden gelesen,
'  Felder jüngerer Versionen bleiben bei älteren Codes leer/0 (wie
'  decodeBinaer in vendor/eeb-format/src/codec.ts). Neu je Version:
'     3  Ernährung je Person        7  Stand minutengenau
'     6  Übungs-Flag                8  weitere Fahrerlaubnisklassen
'     9  Sitzplätze je Fahrzeug    10  abweichende Dienststelle je Person
'
'  Einstiegspunkte für EINEN vollständigen Code:
'     EebBogenLesen         (Text -> TEebBogen, wirft bei Fehlern)
'     EebBogenAlsText       (Text -> Klartext für eine Zelle/MsgBox)
'
'  Einstiegspunkte für MEHRTEILIGE Codes ("EEBS.<teil>.<anzahl>.<id>.<chunk>",
'  ein Bogen über mehrere QR-Codes verteilt - Gegenstück zu segmentSammeln/
'  segmentePayload/segmenteZuBogen in vendor/eeb-format/src/codec.ts):
'     EebIstSegmentCode       (Text -> Boolean; VOR dem Einlesen prüfen, ob
'                              dieser Code zu einer Serie gehört statt
'                              EebBogenLesen)
'     EebSegmentSammeln       (Text -> TEebSammelErgebnis; einen Teil aufnehmen)
'     EebSegmenteZuBogen      (-> TEebBogen, wirft wenn noch unvollständig)
'     EebBogenAlsTextAusBogen (TEebBogen -> Klartext, für die Bestätigung
'                              nach EebSegmenteZuBogen)
'     EebSegmentFehlendeTeile / EebSegmentVorhandeneTeile /
'     EebSegmentGesammelteAnzahl / EebSegmentBenoetigteAnzahl /
'     EebSegmentIstVollstaendig / EebSammlungAktiv   (Statusabfragen für die UI)
'     EebSammlungVerwerfen    (laufende Sammlung verwerfen, z. B. Abbrechen)
'
'  Ablauf für die UI: pro Scan zuerst EebIstSegmentCode prüfen. Bei False wie
'  bisher EebBogenLesen. Bei True EebSegmentSammeln aufrufen und anhand von
'  ergebnis.status (eebSammelVollstaendig/eebSammelGesammelt/eebSammelDuplikat/
'  eebSammelFremd) sowie ergebnis.teilNr/ergebnis.anzahl anzeigen, z. B.
'  "Teil 1/4 erfasst" bzw. "weitere Teile benötigt: " & EebSegmentFehlendeTeile.
'  Bei eebSammelVollstaendig kann die UI mit
'  EebBogenAlsTextAusBogen(EebSegmenteZuBogen()) eine Bestätigung zeigen, bevor
'  sie mit EebBogenInZeileSchreiben (Modul m_digitalerEEB) schreibt und danach
'  EebSammlungVerwerfen aufruft.
'
'  Braucht das Modul m_digitalerEEBVokabular (EebVokabText) für die Klartexte
'  zu den Zahlencodes.
' ===========================================================================
Option Explicit

Private Const B41_ALPHABET As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ$*-/:"
Private Const EEB_EPOCHE As Date = #1/1/2020#
Private Const SCHEMA_MAX As Long = 10
' Obergrenze der Signaturstufen im Container (wie MAX_STUFEN in der App).
Private Const MAX_STUFEN As Long = 32
' Ein Beruf steht als Freitext mit diesem Präfix in den Zusatzqualifikationen
' (kein eigenes Feld im Transportformat, siehe src/app/beruf.ts in der App).
Private Const BERUF_PRAEFIX As String = "Beruf: "

' -------------------------------------------------- Datenstrukturen des Bogens

Public Type TEebPerson
    nachname As String
    vorname As String
    staerkeRolle As Long          ' 0=Mannschaft 1=Unterführer 2=Führer
    funktionen As String
    fahrerlaubnis As String
    geschlecht As Long            ' 0=M 1=W 2=D
    ernaehrung As Long            ' 0=Fleisch 1=Vegetarisch 2=Vegan
    kontakte As String
    qualifikationen As String      ' ohne die Berufe (siehe berufe)
    berufe As String               ' Einträge mit Präfix "Beruf: " aus den
                                   ' Zusatzqualifikationen, ohne Präfix
    dienststelle As String         ' abweichende Dienststelle, leer = wie die
                                   ' Einheit - ab Schema 10
End Type

Public Type TEebFahrzeug
    typ As String
    kennzeichen As String
    funkrufname As String
    stanKonform As String
    aenderungen As String
    sitzplaetze As Long           ' inkl. Fahrer/in; 0 = nicht angegeben (Richtwert
                                   ' des Fahrzeugtyps gilt) - ab Schema 9
End Type

Public Type TEebBogen
    schemaVersion As Long
    uebung As Boolean
    stand As Date
    organisation As Long
    organisationName As String
    einheitsTyp As String
    einheitOrt As String          ' unterste Hierarchie-Ebene
    hierarchie As String
    standortRef As Long           ' THW-OV-Nummer, 0 = keine
    ortAuftrag As String
    zeitraumVon As Date
    zeitraumBis As Date
    einsatzbeginn As Date         ' 0 = nicht angegeben
    einsatzende As Date
    personalErfassung As Long     ' 0=vollständig 1=nur Stärke
    staerkeFuehrer As Long
    staerkeUnterfuehrer As Long
    staerkeMannschaft As Long
    staerkeGesamt As Long
    staerkeManuell As Boolean
    unterbringungM As Long
    unterbringungW As Long
    unterbringungD As Long
    vegetarisch As Long
    vegan As Long
    hatSofortbedarf As Boolean
    verpflegungPersonen As Long
    dieselLiter As Long
    benzinLiter As Long
    gemischLiter As Long
    unterbringung As Boolean
    ruhezeit As Boolean
    sonstiges As String
    signaturstufen As Long        ' >0 = signierter Code (Signatur NICHT geprüft)
    personal() As TEebPerson
    personalAnzahl As Long
    fahrzeuge() As TEebFahrzeug
    fahrzeugAnzahl As Long
End Type

' ---------------------------------------------- Mehrteilige Codes (neu): Typen

' Status einer laufenden Sammlung mehrteiliger Codes, siehe EebSegmentSammeln.
Public Enum EebSammelStatus
    eebSammelGesammelt = 0      ' passt zur laufenden Serie, noch nicht vollständig
    eebSammelDuplikat = 1       ' dieser Teil liegt schon vor
    eebSammelFremd = 2          ' gehört zu einer ANDEREN Serie - alte Sammlung wurde verworfen
    eebSammelVollstaendig = 3   ' alle Teile beieinander
End Enum

' Ein gescannter Teil eines mehrteiligen Codes ("EEBS.<teil>.<anzahl>.<id>.<chunk>").
Public Type TEebSegmentTeil
    teilNr As Long          ' 1-basiert
    anzahl As Long          ' Gesamtzahl der Teile dieser Serie (>= 2)
    id As Double            ' Prüfsumme über den vollständigen Payload (32 Bit unsigned)
    chunk() As Long         ' Byte-Abschnitt des Payloads (bereits Base41/64-dekodiert)
    chunkLen As Long
End Type

' Rückgabe von EebSegmentSammeln - für die Anzeige in der UI.
Public Type TEebSammelErgebnis
    status As EebSammelStatus
    teilNr As Long           ' Teilnummer des soeben verarbeiteten Scans
    anzahl As Long           ' Gesamtzahl der Teile dieser Serie
    haben As Long            ' wie viele Teile jetzt insgesamt vorliegen
End Type

' --------------------------------------------------------------- Modulzustand

Private mBytes() As Long          ' aktuell gelesener Bytestrom
Private mPos As Long              ' Leseposition darin
Private mLen As Long
Private mBitPos As Long           ' Bitleser für Deflate
Private mBitBuf As Long
Private mBitCnt As Long

' Laufende Sammlung eines mehrteiligen Codes (EebSegmentSammeln/EebSegmenteZuBogen).
Private mSegTeile() As TEebSegmentTeil
Private mSegAnzahl As Long        ' Anzahl belegter Einträge in mSegTeile

' ===========================================================================
'  Öffentliche Einstiegspunkte
' ===========================================================================

' Dekodiert einen gescannten Code und gibt ihn als Klartext zurück
' (z. B. für eine Zelle oder MsgBox).
Public Function EebBogenAlsText(ByVal qrText As String) As String
    On Error GoTo Fehler
    EebBogenAlsText = BogenAlsTextFormatieren(EebBogenLesen(qrText))
    Exit Function
Fehler:
    EebBogenAlsText = "Fehler: " & Err.Description
End Function

' Wie EebBogenAlsText, aber für einen bereits fertigen TEebBogen - insbesondere
' für die Bestätigungsanzeige nach EebSegmenteZuBogen (mehrteiliger Code).
Public Function EebBogenAlsTextAusBogen(ByRef b As TEebBogen) As String
    On Error GoTo Fehler
    EebBogenAlsTextAusBogen = BogenAlsTextFormatieren(b)
    Exit Function
Fehler:
    EebBogenAlsTextAusBogen = "Fehler: " & Err.Description
End Function

Private Function BogenAlsTextFormatieren(ByRef b As TEebBogen) As String
    Dim s As String, i As Long
    s = "Stand: " & Format$(b.stand, "dd.mm.yyyy hh:nn") & vbLf
    If b.uebung Then s = s & "*** ÜBUNG ***" & vbLf
    s = s & "Einheit: " & EebOrgName(b.organisation, b.organisationName) & " " & _
        b.einheitsTyp & " " & b.einheitOrt & vbLf
    s = s & "Auftrag: " & b.ortAuftrag & vbLf
    s = s & "Zeitraum: " & Format$(b.zeitraumVon, "dd.mm.yyyy") & " - " & _
        Format$(b.zeitraumBis, "dd.mm.yyyy") & vbLf
    s = s & "Stärke: " & b.staerkeFuehrer & "/" & b.staerkeUnterfuehrer & "/" & _
        b.staerkeMannschaft & "//" & b.staerkeGesamt & vbLf
    For i = 1 To b.personalAnzahl
        s = s & "  " & b.personal(i).nachname & ", " & b.personal(i).vorname & _
            " (" & b.personal(i).funktionen & ")"
        If Len(b.personal(i).dienststelle) > 0 Then s = s & " [" & b.personal(i).dienststelle & "]"
        s = s & vbLf
    Next i
    For i = 1 To b.fahrzeugAnzahl
        s = s & "  " & b.fahrzeuge(i).typ & " " & b.fahrzeuge(i).kennzeichen & vbLf
    Next i
    BogenAlsTextFormatieren = s
End Function

' Kernfunktion für EINEN vollständigen Code: Scan-Text -> gefüllte Struktur.
' Wirft bei ungültigen Daten. Für mehrteilige Codes ("EEBS.") siehe
' EebIstSegmentCode/EebSegmentSammeln/EebSegmenteZuBogen weiter unten.
Public Function EebBogenLesen(ByVal qrText As String) As TEebBogen
    Dim t As String, raute As Long, roh() As Long, rohLen As Long
    t = Trim$(qrText)
    ' Volle App-URL: nur das Fragment hinter '#' trägt Daten.
    raute = InStr(t, "#")
    If raute > 0 Then t = Mid$(t, raute + 1)
    ' Vorlagen-Marker "V." (geteilte Vorlage) abtrennen.
    If Left$(t, 2) = "V." Then t = Mid$(t, 3)
    ' Mehrteiliger Code: Kopf "EEBS.<teil>.<anzahl>.<id>." vor dem Datenteil.
    ' Diese Funktion liest nur EINEN vollständigen Code - mehrteilige Codes
    ' laufen über EebSegmentSammeln/EebSegmenteZuBogen (die UI muss das VOR
    ' dem Aufruf hier mit EebIstSegmentCode abfangen).
    If Left$(t, 5) = "EEBS." Then
        Err.Raise vbObjectError + 3, , "Mehrteiliger Code (Teil " & SegmentTeilNr(t) & _
            "): bitte über EebIstSegmentCode/EebSegmentSammeln einlesen, nicht über EebBogenLesen."
    End If
    If Left$(t, 2) = "B." Then
        rohLen = Base41Dekodieren(Mid$(t, 3), roh)
    Else
        rohLen = Base64UrlDekodieren(t, roh)
    End If
    EebBogenLesen = BogenAusRohPayloadLesen(roh, rohLen)
End Function

' Gemeinsamer Rest von EebBogenLesen und EebSegmenteZuBogen: rohe Payload-Bytes
' (bereits Base41/64-dekodiert, bei mehrteiligen Codes aus den Chunks
' zusammengesetzt) -> Magic-Prüfung, optionaler EEB2C-Signaturkopf, Deflate,
' Bogen-Felder.
Private Function BogenAusRohPayloadLesen(ByRef roh() As Long, ByVal rohLen As Long) As TEebBogen
    Dim start As Long, stufen As Long, b As TEebBogen
    If rohLen < 5 Then Err.Raise vbObjectError + 1, , "Kein EEB2-Code (zu kurz)"
    If Not (roh(0) = &H45 And roh(1) = &H45 And roh(2) = &H42 And roh(3) = &H32) Then
        Err.Raise vbObjectError + 2, , "Kein EEB2-Code (falsche Kennung)"
    End If
    start = 4
    stufen = 0
    If roh(4) = &H43 Then                      ' "EEB2C" = signierter Container
        LeserSetzen roh, rohLen
        mPos = 5
        stufen = LiesVarint()
        If stufen < 1 Or stufen > MAX_STUFEN Then _
            Err.Raise vbObjectError + 4, , "EEB2C: unplausible Stufenzahl"
        Dim i As Long, kartenLen As Long
        For i = 1 To stufen
            mPos = mPos + 96                   ' 32 Byte Pubkey + 64 Byte Signatur
            kartenLen = LiesVarint()
            mPos = mPos + kartenLen            ' Absenderkarte überspringen
        Next i
        start = mPos
    End If
    Dim komprimiert() As Long, kLen As Long, klar() As Long, klarLen As Long
    kLen = TeilBytes(roh, start, rohLen - start, komprimiert)
    klarLen = InflateRoh(komprimiert, kLen, klar)
    LeserSetzen klar, klarLen
    b = BogenLesen()
    b.signaturstufen = stufen
    BogenAusRohPayloadLesen = b
End Function

' Teilnummer aus dem Kopf "EEBS.<teil>.<anzahl>.<id>." - nur für Meldungen.
Private Function SegmentTeilNr(ByVal t As String) As String
    Dim rest As String, punkt As Long
    rest = Mid$(t, 6)
    punkt = InStr(rest, ".")
    If punkt > 1 Then
        SegmentTeilNr = Left$(rest, punkt - 1)
    Else
        SegmentTeilNr = "?"
    End If
End Function

' ===========================================================================
'  Mehrteilige Codes (Segmentierung)
'  ---------------------------------------------------------------------
'  Gegenstück zu parseSegmentUrl/segmentSammeln/segmentePayload/segmenteZuBogen
'  in vendor/eeb-format/src/codec.ts. Kopf im URL-Fragment:
'     "EEBS.<teilNr>.<anzahl>.<id>.<chunk>"
'  teilNr (1-basiert), anzahl (Gesamtzahl, >= 2) und id (32-Bit-FNV-1a-Prüf-
'  summe über den vollständigen Payload) binden die Teile aneinander; chunk
'  ist ein Byte-Abschnitt des Payloads, wie ein Einzelcode Base41 (Marker
'  "B.") oder Base64url kodiert.
' ===========================================================================

' true, wenn der gescannte Text/Link zu einem mehrteiligen Code gehört
' (Fragment beginnt mit "EEBS."). VOR EebBogenLesen prüfen.
Public Function EebIstSegmentCode(ByVal qrText As String) As Boolean
    EebIstSegmentCode = (Left$(EebFragmentInhalt(qrText), 5) = "EEBS.")
End Function

' Extrahiert den Datenteil hinter '#' aus einer vollen App-URL, einem nackten
' '#...'-Fragment oder einem rohen Payload-Text (unverändert).
Private Function EebFragmentInhalt(ByVal qrText As String) As String
    Dim t As String, raute As Long
    t = Trim$(qrText)
    raute = InStr(t, "#")
    If raute > 0 Then
        EebFragmentInhalt = Mid$(t, raute + 1)
    Else
        EebFragmentInhalt = t
    End If
End Function

' Zerlegt einen einzelnen Segment-Code in seine Kopfdaten + dekodierten Chunk.
' Wirft bei ungültigem Kopf oder ungültigen Zeichen im Datenteil.
Public Function EebSegmentParsen(ByVal qrText As String) As TEebSegmentTeil
    Dim frag As String, rest As String, datenTeil As String
    Dim punkt1 As Long, punkt2 As Long, punkt3 As Long
    Dim teilNrText As String, anzahlText As String, idText As String
    Dim roh() As Long, rohLen As Long
    Dim t As TEebSegmentTeil

    frag = EebFragmentInhalt(qrText)
    If Left$(frag, 5) <> "EEBS." Then
        Err.Raise vbObjectError + 50, , "Kein EEB2-Segment-Code"
    End If
    rest = Mid$(frag, 6)

    punkt1 = InStr(rest, ".")
    If punkt1 <= 0 Then GoTo UngueltigerKopf
    punkt2 = InStr(punkt1 + 1, rest, ".")
    If punkt2 <= 0 Then GoTo UngueltigerKopf
    punkt3 = InStr(punkt2 + 1, rest, ".")
    If punkt3 <= 0 Then GoTo UngueltigerKopf

    teilNrText = Left$(rest, punkt1 - 1)
    anzahlText = Mid$(rest, punkt1 + 1, punkt2 - punkt1 - 1)
    idText = Mid$(rest, punkt2 + 1, punkt3 - punkt2 - 1)
    If Not (IstGanzzahl(teilNrText) And IstGanzzahl(anzahlText) And IstGanzzahl(idText)) Then _
        GoTo UngueltigerKopf

    t.teilNr = CLng(teilNrText)
    t.anzahl = CLng(anzahlText)
    t.id = CDbl(idText)
    If t.anzahl < 2 Or t.teilNr < 1 Or t.teilNr > t.anzahl Then GoTo UngueltigerKopf

    datenTeil = Mid$(rest, punkt3 + 1)
    If Left$(datenTeil, 2) = "B." Then
        rohLen = Base41Dekodieren(Mid$(datenTeil, 3), roh)
    Else
        rohLen = Base64UrlDekodieren(datenTeil, roh)
    End If
    t.chunk = roh
    t.chunkLen = rohLen

    EebSegmentParsen = t
    Exit Function

UngueltigerKopf:
    Err.Raise vbObjectError + 51, , "EEB2-Segment: ungültiger Kopf"
End Function

Private Function IstGanzzahl(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    IstGanzzahl = True
End Function

' Nimmt einen gescannten Segment-Code in die laufende Sammlung auf und gibt
' den neuen Stand zurück (für die Anzeige "Teil 1/4 erfasst" o. Ä.).
' Gehört der Code zu einer anderen Serie (andere id/anzahl) als die laufende
' Sammlung, wird diese verworfen und eine neue begonnen (status = eebSammelFremd).
Public Function EebSegmentSammeln(ByVal qrText As String) As TEebSammelErgebnis
    Dim neu As TEebSegmentTeil
    Dim ergebnis As TEebSammelErgebnis
    Dim passt As Boolean, fremd As Boolean
    Dim i As Long

    neu = EebSegmentParsen(qrText)

    passt = (mSegAnzahl > 0)
    If passt Then passt = (mSegTeile(1).id = neu.id) And (mSegTeile(1).anzahl = neu.anzahl)

    If passt Then
        For i = 1 To mSegAnzahl
            If mSegTeile(i).teilNr = neu.teilNr Then
                ergebnis.status = eebSammelDuplikat
                ergebnis.teilNr = neu.teilNr
                ergebnis.anzahl = neu.anzahl
                ergebnis.haben = mSegAnzahl
                EebSegmentSammeln = ergebnis
                Exit Function
            End If
        Next i
    End If

    fremd = (mSegAnzahl > 0) And Not passt

    If passt Then
        mSegAnzahl = mSegAnzahl + 1
        ReDim Preserve mSegTeile(1 To mSegAnzahl)
    Else
        mSegAnzahl = 1
        ReDim mSegTeile(1 To 1)
    End If
    mSegTeile(mSegAnzahl) = neu

    ergebnis.teilNr = neu.teilNr
    ergebnis.anzahl = neu.anzahl
    ergebnis.haben = mSegAnzahl
    If mSegAnzahl = neu.anzahl Then
        ergebnis.status = eebSammelVollstaendig
    ElseIf fremd Then
        ergebnis.status = eebSammelFremd
    Else
        ergebnis.status = eebSammelGesammelt
    End If

    EebSegmentSammeln = ergebnis
End Function

' Setzt die Teile der laufenden Sammlung zu einem vollständigen Bogen zusammen.
' Wirft, wenn noch nicht alle Teile vorliegen oder die Prüfsumme nicht passt
' (defekte/vermischte Teile). Die Sammlung selbst bleibt danach unverändert -
' die UI ruft nach dem Schreiben EebSammlungVerwerfen auf.
Public Function EebSegmenteZuBogen() As TEebBogen
    Dim anzahl As Long, i As Long, k As Long, gesamtLen As Long, pos As Long
    Dim sortiert() As TEebSegmentTeil
    Dim roh() As Long
    Dim pruefwert As Double

    If mSegAnzahl = 0 Then Err.Raise vbObjectError + 52, , "Keine Segmente gesammelt"
    anzahl = mSegTeile(1).anzahl
    If mSegAnzahl <> anzahl Then _
        Err.Raise vbObjectError + 53, , "Noch nicht vollständig: " & mSegAnzahl & " von " & anzahl & " Teilen"

    sortiert = mSegTeile
    SegmenteSortieren sortiert, mSegAnzahl

    For i = 1 To anzahl
        If sortiert(i).teilNr <> i Then _
            Err.Raise vbObjectError + 54, , "EEB2-Segmente: unvollständig oder vermischt (Teil " & i & "/" & anzahl & " fehlt)"
        gesamtLen = gesamtLen + sortiert(i).chunkLen
    Next i

    ReDim roh(0 To Application.Max(1, gesamtLen))
    For i = 1 To anzahl
        For k = 0 To sortiert(i).chunkLen - 1
            roh(pos) = sortiert(i).chunk(k)
            pos = pos + 1
        Next k
    Next i

    pruefwert = Fnv1aPruefsumme(roh, gesamtLen)
    If pruefwert <> mSegTeile(1).id Then _
        Err.Raise vbObjectError + 55, , "EEB2-Segmente: Prüfsumme falsch (defekte Teile)"

    EebSegmenteZuBogen = BogenAusRohPayloadLesen(roh, gesamtLen)
End Function

Private Sub SegmenteSortieren(ByRef teile() As TEebSegmentTeil, ByVal anzahl As Long)
    Dim i As Long, j As Long, tmp As TEebSegmentTeil
    For i = 1 To anzahl - 1
        For j = 1 To anzahl - i
            If teile(j).teilNr > teile(j + 1).teilNr Then
                tmp = teile(j)
                teile(j) = teile(j + 1)
                teile(j + 1) = tmp
            End If
        Next j
    Next i
End Sub

' 32-Bit-FNV-1a-Prüfsumme (Gegenstück zu pruefsumme() in codec.ts). Als Double
' gehalten, weil VBA Long nur 32-Bit VORZEICHENBEHAFTET ist (Bereich bis
' 2^31-1) und die Multiplikation sonst überliefe; Mult32Mod bildet die 32-Bit-
' Multiplikation modulo 2^32 exakt nach (Äquivalent zu Math.imul + ">>> 0").
Private Function Fnv1aPruefsumme(ByRef daten() As Long, ByVal anzahl As Long) As Double
    Dim h As Double
    Dim i As Long, lowByte As Long, neuByte As Long
    h = 2166136261#                              ' 0x811c9dc5
    For i = 0 To anzahl - 1
        lowByte = h - Int(h / 256) * 256         ' h Mod 256 (als Long-sicherer Wert)
        neuByte = lowByte Xor daten(i)
        h = h - lowByte + neuByte                ' nur das unterste Byte von h ersetzen
        h = Mult32Mod(h, 16777619)               ' * 0x01000193 mod 2^32
    Next i
    Fnv1aPruefsumme = h
End Function

' (h * p) Mod 2^32, mit h als "unsigned 32-Bit" in einem Double (0 .. 2^32-1)
' und p als kleiner Long-Konstante. Aufgeteilt in High-/Low-16-Bit-Teilpro-
' dukte, damit keine Zwischensumme die 2^53-Genauigkeitsgrenze von Double
' überschreitet.
Private Function Mult32Mod(ByVal h As Double, ByVal p As Long) As Double
    Const ZWEI_HOCH_32 As Double = 4294967296#   ' 2^32
    Dim hi As Double, lo As Double, term1 As Double, term1a As Double, term2 As Double, summe As Double
    hi = Int(h / 65536)
    lo = h - hi * 65536
    term1 = hi * p
    term1a = (term1 - Int(term1 / 65536) * 65536) * 65536   ' (term1 Mod 65536) * 65536
    term2 = lo * p
    summe = term1a + term2
    Mult32Mod = summe - Int(summe / ZWEI_HOCH_32) * ZWEI_HOCH_32
End Function

' Verwirft eine laufende Sammlung (z. B. Abbrechen-Button in der UI).
Public Sub EebSammlungVerwerfen()
    mSegAnzahl = 0
    Erase mSegTeile
End Sub

' true, wenn gerade Teile einer Serie gesammelt werden (noch nicht verworfen
' oder fertiggestellt).
Public Function EebSammlungAktiv() As Boolean
    EebSammlungAktiv = (mSegAnzahl > 0)
End Function

Public Function EebSegmentIstVollstaendig() As Boolean
    EebSegmentIstVollstaendig = (mSegAnzahl > 0) And (mSegAnzahl = mSegTeile(1).anzahl)
End Function

Public Function EebSegmentGesammelteAnzahl() As Long
    EebSegmentGesammelteAnzahl = mSegAnzahl
End Function

' 0, solange noch kein Teil gescannt wurde (die Gesamtzahl steht erst mit dem
' ersten Teil fest).
Public Function EebSegmentBenoetigteAnzahl() As Long
    If mSegAnzahl > 0 Then EebSegmentBenoetigteAnzahl = mSegTeile(1).anzahl
End Function

' Kommaseparierte Liste der noch fehlenden Teilnummern, z. B. "3, 5" - für
' einen Hinweis "noch benötigt: Teil 3, 5". Leerer String, wenn nichts fehlt
' oder keine Sammlung läuft.
Public Function EebSegmentFehlendeTeile() As String
    Dim anzahl As Long, i As Long, j As Long, gefunden As Boolean, s As String
    If mSegAnzahl = 0 Then Exit Function
    anzahl = mSegTeile(1).anzahl
    For i = 1 To anzahl
        gefunden = False
        For j = 1 To mSegAnzahl
            If mSegTeile(j).teilNr = i Then gefunden = True: Exit For
        Next j
        If Not gefunden Then
            If Len(s) > 0 Then s = s & ", "
            s = s & i
        End If
    Next i
    EebSegmentFehlendeTeile = s
End Function

' Kommaseparierte, aufsteigend sortierte Liste der bereits vorhandenen
' Teilnummern, z. B. "1, 2, 4" - Gegenstück zu EebSegmentFehlendeTeile.
Public Function EebSegmentVorhandeneTeile() As String
    Dim i As Long, j As Long, nrs() As Long, tmp As Long, s As String
    If mSegAnzahl = 0 Then Exit Function
    ReDim nrs(1 To mSegAnzahl)
    For i = 1 To mSegAnzahl
        nrs(i) = mSegTeile(i).teilNr
    Next i
    For i = 1 To mSegAnzahl - 1
        For j = 1 To mSegAnzahl - i
            If nrs(j) > nrs(j + 1) Then
                tmp = nrs(j): nrs(j) = nrs(j + 1): nrs(j + 1) = tmp
            End If
        Next j
    Next i
    For i = 1 To mSegAnzahl
        If Len(s) > 0 Then s = s & ", "
        s = s & nrs(i)
    Next i
    EebSegmentVorhandeneTeile = s
End Function

' ===========================================================================
'  Base41 / Base64url  (Gegenstück zu base41Dekodieren in src/codec.ts)
' ===========================================================================

Private Function Base41Dekodieren(ByVal s As String, ByRef aus() As Long) As Long
    Dim i As Long, j As Long, rest As Long, n As Long, v As Long
    Dim d(0 To 2) As Long, p As Long
    ReDim aus(0 To Application.Max(1, Len(s)))
    n = Len(s)
    p = 0
    For i = 1 To n Step 3
        rest = n - i + 1
        If rest = 1 Then Err.Raise vbObjectError + 10, , "Base41: unvollständige Gruppe"
        For j = 0 To IIf(rest < 3, rest - 1, 2)
            d(j) = InStr(1, B41_ALPHABET, Mid$(s, i + j, 1), vbBinaryCompare) - 1
            If d(j) < 0 Then Err.Raise vbObjectError + 11, , _
                "Base41: ungültiges Zeichen '" & Mid$(s, i + j, 1) & "'"
        Next j
        If rest >= 3 Then
            v = d(0) + d(1) * 41 + d(2) * 1681
            If v > 65535 Then Err.Raise vbObjectError + 12, , "Base41: Wert außerhalb 16 Bit"
            aus(p) = v \ 256: aus(p + 1) = v Mod 256
            p = p + 2
        Else
            v = d(0) + d(1) * 41
            If v > 255 Then Err.Raise vbObjectError + 13, , "Base41: Restbyte außerhalb 8 Bit"
            aus(p) = v
            p = p + 1
        End If
    Next i
    Base41Dekodieren = p
End Function

Private Function Base64UrlDekodieren(ByVal s As String, ByRef aus() As Long) As Long
    Const a As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_"
    Dim i As Long, wert As Long, puffer As Long, bits As Long, p As Long
    ReDim aus(0 To Application.Max(1, Len(s)))
    For i = 1 To Len(s)
        wert = InStr(1, a, Mid$(s, i, 1), vbBinaryCompare) - 1
        If wert < 0 Then Err.Raise vbObjectError + 14, , "Kein EEB2-Code (ungültige Zeichen)"
        puffer = puffer * 64 + wert
        bits = bits + 6
        If bits >= 8 Then
            bits = bits - 8
            aus(p) = (puffer \ (2 ^ bits)) And 255
            p = p + 1
            puffer = puffer And ((2 ^ bits) - 1)
        End If
    Next i
    Base64UrlDekodieren = p
End Function

' ===========================================================================
'  Deflate (raw, RFC 1951) - reine VBA-Umsetzung, keine Systembibliothek
' ===========================================================================

Private Function InflateRoh(ByRef ein() As Long, ByVal einLen As Long, ByRef aus() As Long) As Long
    Dim lenBase As Variant, lenExtra As Variant, distBase As Variant, distExtra As Variant, clcOrder As Variant
    lenBase = Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258)
    lenExtra = Array(0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0)
    distBase = Array(1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577)
    distExtra = Array(0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13)
    clcOrder = Array(16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15)

    Dim ausLen As Long, kapazitaet As Long
    kapazitaet = 4096
    ReDim aus(0 To kapazitaet - 1)
    ausLen = 0
    mBitPos = 0: mBitBuf = 0: mBitCnt = 0

    Dim letzter As Long, typ As Long
    Dim lenCount(0 To 15) As Long, lenSym() As Long
    Dim distCount(0 To 15) As Long, distSym() As Long

    Do
        letzter = LiesBits(ein, einLen, 1)
        typ = LiesBits(ein, einLen, 2)

        If typ = 0 Then                             ' unkomprimierter Block
            mBitBuf = 0: mBitCnt = 0
            If mBitPos + 4 > einLen Then Err.Raise vbObjectError + 20, , "Deflate: Block zu kurz"
            Dim roheLen As Long, k As Long
            roheLen = ein(mBitPos) + ein(mBitPos + 1) * 256
            mBitPos = mBitPos + 4
            For k = 1 To roheLen
                PufferSchreiben aus, kapazitaet, ausLen, ein(mBitPos)
                mBitPos = mBitPos + 1
            Next k

        ElseIf typ = 1 Or typ = 2 Then
            Dim laengen() As Long, i As Long

            If typ = 1 Then                         ' feste Huffman-Tabellen
                ReDim laengen(0 To 287)
                For i = 0 To 143: laengen(i) = 8: Next i
                For i = 144 To 255: laengen(i) = 9: Next i
                For i = 256 To 279: laengen(i) = 7: Next i
                For i = 280 To 287: laengen(i) = 8: Next i
                HuffmanBauen laengen, 288, lenCount, lenSym

                ReDim laengen(0 To 29)
                For i = 0 To 29: laengen(i) = 5: Next i
                HuffmanBauen laengen, 30, distCount, distSym
            Else                                     ' dynamische Tabellen
                Dim nLen As Long, nDist As Long, nCode As Long
                nLen = LiesBits(ein, einLen, 5) + 257
                nDist = LiesBits(ein, einLen, 5) + 1
                nCode = LiesBits(ein, einLen, 4) + 4

                Dim clc(0 To 18) As Long, clcCount(0 To 15) As Long, clcSym() As Long
                For i = 0 To 18: clc(i) = 0: Next i
                For i = 0 To nCode - 1
                    clc(clcOrder(i)) = LiesBits(ein, einLen, 3)
                Next i
                HuffmanBauen clc, 19, clcCount, clcSym

                ReDim laengen(0 To nLen + nDist - 1)
                Dim sym As Long, wieOft As Long, wert As Long
                i = 0
                Do While i < nLen + nDist
                    sym = HuffmanLesen(ein, einLen, clcCount, clcSym)
                    If sym < 16 Then
                        laengen(i) = sym
                        i = i + 1
                    Else
                        wert = 0
                        If sym = 16 Then
                            If i = 0 Then Err.Raise vbObjectError + 21, , "Deflate: Wiederholung ohne Vorgänger"
                            wert = laengen(i - 1)
                            wieOft = 3 + LiesBits(ein, einLen, 2)
                        ElseIf sym = 17 Then
                            wieOft = 3 + LiesBits(ein, einLen, 3)
                        Else
                            wieOft = 11 + LiesBits(ein, einLen, 7)
                        End If
                        Do While wieOft > 0
                            laengen(i) = wert
                            i = i + 1
                            wieOft = wieOft - 1
                        Loop
                    End If
                Loop

                Dim teilA() As Long, teilB() As Long
                ReDim teilA(0 To nLen - 1)
                For i = 0 To nLen - 1: teilA(i) = laengen(i): Next i
                ReDim teilB(0 To nDist - 1)
                For i = 0 To nDist - 1: teilB(i) = laengen(nLen + i): Next i
                HuffmanBauen teilA, nLen, lenCount, lenSym
                HuffmanBauen teilB, nDist, distCount, distSym
            End If

            Dim s As Long, laenge As Long, abstand As Long, ds As Long, j As Long
            Do
                s = HuffmanLesen(ein, einLen, lenCount, lenSym)
                If s < 256 Then
                    PufferSchreiben aus, kapazitaet, ausLen, s
                ElseIf s = 256 Then
                    Exit Do
                Else
                    s = s - 257
                    If s >= 29 Then Err.Raise vbObjectError + 22, , "Deflate: ungültiger Längencode"
                    laenge = lenBase(s) + LiesBits(ein, einLen, lenExtra(s))
                    ds = HuffmanLesen(ein, einLen, distCount, distSym)
                    abstand = distBase(ds) + LiesBits(ein, einLen, distExtra(ds))
                    If abstand > ausLen Then Err.Raise vbObjectError + 23, , "Deflate: Abstand zu groß"
                    For j = 1 To laenge
                        PufferSchreiben aus, kapazitaet, ausLen, aus(ausLen - abstand)
                    Next j
                End If
            Loop

        Else
            Err.Raise vbObjectError + 24, , "Deflate: reservierter Blocktyp"
        End If
    Loop While letzter = 0

    InflateRoh = ausLen
End Function

Private Sub PufferSchreiben(ByRef puffer() As Long, ByRef kapazitaet As Long, ByRef anzahl As Long, ByVal wert As Long)
    If anzahl >= kapazitaet Then
        kapazitaet = kapazitaet * 2
        ReDim Preserve puffer(0 To kapazitaet - 1)
    End If
    puffer(anzahl) = wert
    anzahl = anzahl + 1
End Sub

Private Function LiesBits(ByRef ein() As Long, ByVal einLen As Long, ByVal wieViele As Long) As Long
    Dim wert As Long
    If wieViele = 0 Then LiesBits = 0: Exit Function
    wert = mBitBuf
    Do While mBitCnt < wieViele
        If mBitPos >= einLen Then Err.Raise vbObjectError + 25, , "Deflate: Daten zu kurz"
        wert = wert + ein(mBitPos) * (2 ^ mBitCnt)
        mBitPos = mBitPos + 1
        mBitCnt = mBitCnt + 8
    Loop
    mBitBuf = wert \ (2 ^ wieViele)
    mBitCnt = mBitCnt - wieViele
    LiesBits = wert And ((2 ^ wieViele) - 1)
End Function

' Kanonische Huffman-Tabelle nach dem Verfahren aus puff.c:
' count(l) = Anzahl Codes der Länge l, sym() = Symbole in Code-Reihenfolge.
Private Sub HuffmanBauen(ByRef laengen() As Long, ByVal anzahl As Long, ByRef count() As Long, ByRef sym() As Long)
    Dim i As Long, l As Long
    Dim offs(0 To 16) As Long
    For l = 0 To 15: count(l) = 0: Next l
    For i = 0 To anzahl - 1: count(laengen(i)) = count(laengen(i)) + 1: Next i
    count(0) = 0
    offs(1) = 0
    For l = 1 To 15: offs(l + 1) = offs(l) + count(l): Next l
    ReDim sym(0 To anzahl)
    For i = 0 To anzahl - 1
        If laengen(i) <> 0 Then
            sym(offs(laengen(i))) = i
            offs(laengen(i)) = offs(laengen(i)) + 1
        End If
    Next i
End Sub

Private Function HuffmanLesen(ByRef ein() As Long, ByVal einLen As Long, ByRef count() As Long, ByRef sym() As Long) As Long
    Dim code As Long, erster As Long, index As Long, l As Long, anzahl As Long
    code = 0: erster = 0: index = 0
    For l = 1 To 15
        code = code Or LiesBits(ein, einLen, 1)
        anzahl = count(l)
        If code - anzahl < erster Then
            HuffmanLesen = sym(index + (code - erster))
            Exit Function
        End If
        index = index + anzahl
        erster = (erster + anzahl) * 2
        code = code * 2
    Next l
    Err.Raise vbObjectError + 26, , "Deflate: ungültiger Code"
End Function

Private Function TeilBytes(ByRef ein() As Long, ByVal ab As Long, ByVal anzahl As Long, ByRef aus() As Long) As Long
    Dim i As Long
    ReDim aus(0 To Application.Max(1, anzahl))
    For i = 0 To anzahl - 1: aus(i) = ein(ab + i): Next i
    TeilBytes = anzahl
End Function

' ===========================================================================
'  Leser für den entpackten Binärstrom
' ===========================================================================

Private Sub LeserSetzen(ByRef daten() As Long, ByVal anzahl As Long)
    mBytes = daten
    mLen = anzahl
    mPos = 0
End Sub

Private Function LiesU8() As Long
    If mPos >= mLen Then Err.Raise vbObjectError + 30, , "EEB2: Datenstrom zu kurz"
    LiesU8 = mBytes(mPos)
    mPos = mPos + 1
End Function

Private Function LiesU16() As Long
    Dim a As Long, b As Long
    a = LiesU8(): b = LiesU8()
    LiesU16 = a + b * 256
End Function

Private Function LiesU32() As Double
    Dim a As Long, b As Long, c As Long, d As Long
    a = LiesU8(): b = LiesU8(): c = LiesU8(): d = LiesU8()
    LiesU32 = a + b * 256# + c * 65536# + d * 16777216#
End Function

Private Function LiesVarint() As Long
    Dim wert As Double, faktor As Double, b As Long
    wert = 0: faktor = 1
    Do
        b = LiesU8()
        wert = wert + (b And &H7F) * faktor
        faktor = faktor * 128
        If (b And &H80) = 0 Then Exit Do
    Loop
    LiesVarint = CLng(wert)
End Function

' Länge (Varint) + UTF-8-Bytes -> VBA-String
Private Function LiesString() As String
    Dim n As Long, i As Long, b As Long, b2 As Long, b3 As Long, b4 As Long
    Dim cp As Long, s As String
    n = LiesVarint()
    i = 0
    Do While i < n
        b = LiesU8(): i = i + 1
        If b < &H80 Then
            cp = b
        ElseIf (b And &HE0) = &HC0 Then
            b2 = LiesU8(): i = i + 1
            cp = (b And &H1F) * 64 + (b2 And &H3F)
        ElseIf (b And &HF0) = &HE0 Then
            b2 = LiesU8(): b3 = LiesU8(): i = i + 2
            cp = (b And &HF) * 4096 + (b2 And &H3F) * 64 + (b3 And &H3F)
        Else
            b2 = LiesU8(): b3 = LiesU8(): b4 = LiesU8(): i = i + 3
            cp = (b And &H7) * 262144 + (b2 And &H3F) * 4096 + (b3 And &H3F) * 64 + (b4 And &H3F)
        End If
        If cp < &H10000 Then
            s = s & ChrW$(cp)
        Else                                  ' außerhalb der BMP: Ersatzpaar
            cp = cp - &H10000
            s = s & ChrW$(&HD800 + (cp \ 1024)) & ChrW$(&HDC00 + (cp Mod 1024))
        End If
    Loop
    LiesString = s
End Function

' Telefonnummer: Varint-Ziffernanzahl + BCD (2 Ziffern je Byte)
Private Function LiesBcd() As String
    Dim n As Long, i As Long, b As Long, s As String
    n = LiesVarint()
    For i = 1 To (n + 1) \ 2
        b = LiesU8()
        s = s & Chr$(48 + (b \ 16))
        If Len(s) < n Then s = s & Chr$(48 + (b And 15))
    Next i
    LiesBcd = s
End Function

' Vokabular-Wert: Varint-Code (0 = Freitext folgt).
' art: "E" Einheitstyp, "F" Funktion, "V" Fahrzeugtyp, "K" Kennwort,
'      "H<org>_" Hierarchie-Ebene. Tabellen gibt es nur für das THW und
'      die Ebenen; sonst erscheint der Code als "#12".
Private Function LiesVokabular(ByVal art As String, ByVal organisation As Long) As String
    Dim code As Long
    code = LiesVarint()
    If code = 0 Then
        LiesVokabular = LiesString()
    ElseIf art = "H" Then
        LiesVokabular = EebVokabText("H" & organisation & "_" & code, code)
    ElseIf art = "K" Then
        LiesVokabular = EebVokabText("K" & code, code)
    ElseIf organisation = 1 Then                  ' Vokabulartabellen bisher nur THW
        LiesVokabular = EebVokabText(art & code, code)
    Else
        LiesVokabular = "#" & code
    End If
End Function

' ===========================================================================
'  Bogen aus dem Binärstrom lesen (Reihenfolge wie decodeBinaer in codec.ts)
' ===========================================================================
Private Function BogenLesen() As TEebBogen
    Dim b As TEebBogen
    Dim v As Long, pflags As Long, anzahl As Long, i As Long

    v = LiesVarint()
    If v < 2 Or v > SCHEMA_MAX Then
        Err.Raise vbObjectError + 40, , "EEB2: nicht unterstützte Schema-Version " & v & _
            " (dieser Dekoder kennt 2 bis " & SCHEMA_MAX & ")"
    End If
    b.schemaVersion = v

    If v >= 7 Then
        b.stand = EEB_EPOCHE + LiesU32() / 1440#
    Else
        b.stand = EEB_EPOCHE + LiesU16()
    End If

    EinheitLesen b
    EinsatzLesen b

    pflags = LiesU8()
    b.personalErfassung = pflags And 1
    b.uebung = (pflags And 16) <> 0

    Dim unterbringungManuell As Boolean
    Dim verpflegungManuell As Boolean
    unterbringungManuell = (pflags And 4) <> 0
    verpflegungManuell = (pflags And 8) <> 0

    If (pflags And 2) <> 0 Then
        b.staerkeManuell = True
        b.staerkeFuehrer = LiesU8()
        b.staerkeUnterfuehrer = LiesU8()
        b.staerkeMannschaft = LiesU8()
    End If
    If unterbringungManuell Then
        b.unterbringungM = LiesU8(): b.unterbringungW = LiesU8(): b.unterbringungD = LiesU8()
    End If
    If verpflegungManuell Then
        b.vegetarisch = LiesU8(): b.vegan = LiesU8()
    End If

    anzahl = LiesVarint()
    b.personalAnzahl = anzahl
    ReDim b.personal(1 To Application.Max(1, anzahl))
    For i = 1 To anzahl
        b.personal(i) = PersonLesen(v, b.organisation)
    Next i

    anzahl = LiesVarint()
    b.fahrzeugAnzahl = anzahl
    ReDim b.fahrzeuge(1 To Application.Max(1, anzahl))
    For i = 1 To anzahl
        b.fahrzeuge(i) = FahrzeugLesen(b.organisation, b.einheitOrt)
    Next i

    If LiesU8() = 1 Then
        b.hatSofortbedarf = True
        b.verpflegungPersonen = LiesU8()
        Dim davonVeg As Long
        If v < 3 Then davonVeg = LiesU8()          ' v2: Aggregatzahl im Sofortbedarf
        b.dieselLiter = LiesVarint()
        b.benzinLiter = LiesVarint()
        b.gemischLiter = LiesVarint()
        Dim sflags As Long
        sflags = LiesU8()
        b.unterbringung = (sflags And 1) <> 0
        b.ruhezeit = (sflags And 2) <> 0
        If davonVeg > 0 And Not verpflegungManuell Then
            b.vegetarisch = davonVeg
            verpflegungManuell = True      ' wie im TS-Original: hebt die Aggregatzahl zur manuellen Angabe
        End If
    End If

    If LiesU8() = 1 Then b.sonstiges = LiesString()

    If mPos <> mLen Then Err.Raise vbObjectError + 41, , "EEB2: überschüssige Daten am Ende"

    ' Stärke: manuell gemeldet oder aus den Rollen des Personals abgeleitet.
    If Not b.staerkeManuell Then
        For i = 1 To b.personalAnzahl
            Select Case b.personal(i).staerkeRolle
                Case 2: b.staerkeFuehrer = b.staerkeFuehrer + 1
                Case 1: b.staerkeUnterfuehrer = b.staerkeUnterfuehrer + 1
                Case Else: b.staerkeMannschaft = b.staerkeMannschaft + 1
            End Select
        Next i
    End If
    b.staerkeGesamt = b.staerkeFuehrer + b.staerkeUnterfuehrer + b.staerkeMannschaft

    ' Vegetarisch/Vegan: manuell gemeldet (bzw. aus der alten Sofortbedarf-
    ' Aggregatzahl übernommen) oder aus der Ernährungsangabe des Personals abgeleitet.
    If Not verpflegungManuell Then
        For i = 1 To b.personalAnzahl
            Select Case b.personal(i).ernaehrung
                Case 1: b.vegetarisch = b.vegetarisch + 1
                Case 2: b.vegan = b.vegan + 1
            End Select
        Next i
    End If

    ' Unterbringung M/W/D: manuell gemeldet oder aus dem Geschlecht des Personals abgeleitet.
    If Not unterbringungManuell Then
        For i = 1 To b.personalAnzahl
            Select Case b.personal(i).geschlecht
                Case 0: b.unterbringungM = b.unterbringungM + 1
                Case 1: b.unterbringungW = b.unterbringungW + 1
                Case Else: b.unterbringungD = b.unterbringungD + 1
            End Select
        Next i
    End If

    BogenLesen = b
End Function

Private Sub EinheitLesen(ByRef b As TEebBogen)
    Dim flags As Long, orgName As String, anzahl As Long, i As Long
    Dim hflags As Long, bez As String, nam As String, zeile As String, altName As String

    b.organisation = LiesU8()
    flags = LiesU8()
    If (flags And 1) <> 0 Then orgName = LiesString()
    b.einheitsTyp = LiesVokabular("E", b.organisation)
    b.organisationName = orgName

    If (flags And 2) <> 0 Then
        ' THW: Hierarchie steckt als OV-Nummer im Verzeichnis der App und ist
        ' hier nicht auflösbar - die Nummer wird ausgewiesen.
        b.standortRef = LiesVarint()
        b.hierarchie = "OV-Nr. " & b.standortRef
        b.einheitOrt = ""
    Else
        altName = LiesString()                     ' früheres Freitextfeld
        anzahl = LiesVarint()
        For i = 1 To anzahl
            hflags = LiesU8()
            bez = LiesVokabular("H", b.organisation)
            nam = LiesString()
            If (hflags And 1) <> 0 Then Call LiesBcd          ' Telefon
            If (hflags And 2) <> 0 Then Call LiesString       ' E-Mail
            If (hflags And 4) <> 0 Then Call LiesString       ' Kurzform
            If i = 1 Then b.einheitOrt = Trim$(nam)
            zeile = bez & " " & nam
            b.hierarchie = b.hierarchie & IIf(Len(b.hierarchie) > 0, " / ", "") & Trim$(zeile)
        Next i
        If anzahl = 0 And Len(altName) > 0 Then
            b.einheitOrt = Trim$(altName)
            b.hierarchie = altName
        End If
    End If
End Sub

Private Sub EinsatzLesen(ByRef b As TEebBogen)
    Dim flags As Long
    flags = LiesU8()
    b.zeitraumVon = EEB_EPOCHE + LiesU16()
    b.zeitraumBis = EEB_EPOCHE + LiesU16()
    b.ortAuftrag = LiesString()
    If (flags And 1) <> 0 Then b.einsatzbeginn = EEB_EPOCHE + LiesU32() / 1440#
    If (flags And 2) <> 0 Then b.einsatzende = EEB_EPOCHE + LiesU32() / 1440#
End Sub

Private Function PersonLesen(ByVal v As Long, ByVal organisation As Long) As TEebPerson
    Dim p As TEebPerson, flags As Long, anzahl As Long, i As Long, s As String
    p.vorname = LiesString()
    p.nachname = LiesString()
    flags = LiesU8()
    p.fahrerlaubnis = EebFahrerlaubnisText(flags And 15)
    p.geschlecht = (flags \ 16) And 3
    p.staerkeRolle = (flags \ 64) And 3
    If v >= 3 Then p.ernaehrung = LiesU8() Else p.ernaehrung = 0

    If v >= 8 Then                                  ' weitere Fahrerlaubnisklassen
        anzahl = LiesVarint()
        For i = 1 To anzahl
            s = EebFahrerlaubnisText(LiesU8())
            If Len(s) > 0 And s <> "-" Then p.fahrerlaubnis = p.fahrerlaubnis & "+" & s
        Next i
    End If

    If v >= 10 Then p.dienststelle = Trim$(LiesString())   ' abweichende Dienststelle, ab Schema 10

    anzahl = LiesVarint()
    For i = 1 To anzahl
        s = LiesVokabular("F", organisation)
        p.funktionen = p.funktionen & IIf(Len(p.funktionen) > 0, ", ", "") & s
    Next i

    anzahl = LiesVarint()
    For i = 1 To anzahl
        s = KontaktLesen()
        p.kontakte = p.kontakte & IIf(Len(p.kontakte) > 0, ", ", "") & s
    Next i

    ' Zusatzqualifikationen; Freitexte mit Präfix "Beruf: " sind Berufe und
    ' landen (ohne Präfix) im eigenen Feld.
    anzahl = LiesVarint()
    For i = 1 To anzahl
        s = LiesVokabular("F", organisation)
        If Left$(s, Len(BERUF_PRAEFIX)) = BERUF_PRAEFIX Then
            s = Trim$(Mid$(s, Len(BERUF_PRAEFIX) + 1))
            p.berufe = p.berufe & IIf(Len(p.berufe) > 0, ", ", "") & s
        Else
            p.qualifikationen = p.qualifikationen & IIf(Len(p.qualifikationen) > 0, ", ", "") & s
        End If
    Next i

    PersonLesen = p
End Function

Private Function KontaktLesen() As String
    Dim flags As Long, art As Long, wert As String, artText As String
    flags = LiesU8()
    art = flags And 3
    If (flags And 8) <> 0 Then
        Call LiesU8                                   ' altes E-Mail-Template
        wert = "(Standardadresse)"
    ElseIf art = 2 Then
        wert = LiesString()
    Else
        wert = LiesBcd()
    End If
    Select Case art
        Case 0: artText = "Mobil"
        Case 1: artText = "Festnetz"
        Case Else: artText = "E-Mail"
    End Select
    KontaktLesen = artText & ": " & wert & " (" & IIf((flags And 4) <> 0, "D", "P") & ")"
End Function

Private Function FahrzeugLesen(ByVal organisation As Long, ByVal einheitOrt As String) As TEebFahrzeug
    Dim f As TEebFahrzeug, flags As Long
    flags = LiesU8()
    f.typ = LiesVokabular("V", organisation)
    If (flags And 16) <> 0 Then
        Call LiesVarint                              ' altes THW-Kennzeichen (Schema <= 3)
        f.kennzeichen = ""
    Else
        f.kennzeichen = LiesString()
    End If
    If (flags And 4) <> 0 Then f.funkrufname = FunkrufLesen(organisation, einheitOrt)
    If (flags And 1) <> 0 Then f.stanKonform = IIf((flags And 2) <> 0, "ja", "nein")
    If (flags And 8) <> 0 Then f.aenderungen = LiesString()
    If (flags And 32) <> 0 Then f.sitzplaetze = LiesU8()   ' abweichende Sitzplatzzahl, ab Schema 9
    FahrzeugLesen = f
End Function

Private Function FunkrufLesen(ByVal organisation As Long, ByVal einheitOrt As String) As String
    Dim eigen As Boolean, kennwort As String, ort As String, anzahl As Long, i As Long, teile As String
    eigen = (LiesU8() = 1)
    kennwort = LiesVokabular("K", organisation)
    If eigen Then ort = einheitOrt Else ort = LiesString()
    anzahl = LiesVarint()
    For i = 1 To anzahl
        teile = teile & IIf(Len(teile) > 0, "/", "") & LiesU8()
    Next i
    FunkrufLesen = Trim$(kennwort & " " & ort & " " & teile)
End Function

' ===========================================================================
'  Klartext für Enumerationen
' ===========================================================================

Public Function EebOrgName(ByVal org As Long, ByVal freitext As String) As String
    Dim s As String
    Select Case org
        Case 1: s = "THW"
        Case 2: s = "Feuerwehr"
        Case 3: s = "Polizei"
        Case 4: s = "Bundespolizei"
        Case 5: s = "DRK"
        Case 6: s = "Johanniter (JUH)"
        Case 7: s = "Malteser (MHD)"
        Case 8: s = "ASB"
        Case 9: s = "DLRG"
        Case 10: s = "Bundeswehr"
        Case 11: s = "Rettungsdienst"
        Case 255: s = "Sonstige"
        Case Else: s = "Organisation #" & org
    End Select
    If Len(freitext) > 0 Then s = s & " (" & freitext & ")"
    EebOrgName = s
End Function

Public Function EebFahrerlaubnisText(ByVal k As Long) As String
    Select Case k
        Case 0: EebFahrerlaubnisText = ""
        Case 1: EebFahrerlaubnisText = "AM"
        Case 2: EebFahrerlaubnisText = "A1"
        Case 3: EebFahrerlaubnisText = "A2"
        Case 4: EebFahrerlaubnisText = "A"
        Case 5: EebFahrerlaubnisText = "B"
        Case 6: EebFahrerlaubnisText = "BE"
        Case 7: EebFahrerlaubnisText = "C1"
        Case 8: EebFahrerlaubnisText = "C1E"
        Case 9: EebFahrerlaubnisText = "C"
        Case 10: EebFahrerlaubnisText = "CE"
        Case 11: EebFahrerlaubnisText = "D1"
        Case 12: EebFahrerlaubnisText = "D1E"
        Case 13: EebFahrerlaubnisText = "D"
        Case 14: EebFahrerlaubnisText = "DE"
        Case Else: EebFahrerlaubnisText = "#" & k
    End Select
End Function

Public Function EebRolleText(ByVal r As Long) As String
    Select Case r
        Case 2: EebRolleText = "Führer"
        Case 1: EebRolleText = "Unterführer"
        Case Else: EebRolleText = "Mannschaft"
    End Select
End Function

Public Function EebGeschlechtText(ByVal g As Long) As String
    Select Case g
        Case 1: EebGeschlechtText = "w"
        Case 2: EebGeschlechtText = "d"
        Case Else: EebGeschlechtText = "m"
    End Select
End Function

Public Function EebErnaehrungText(ByVal e As Long) As String
    Select Case e
        Case 1: EebErnaehrungText = "vegetarisch"
        Case 2: EebErnaehrungText = "vegan"
        Case Else: EebErnaehrungText = "Fleisch"
    End Select
End Function
