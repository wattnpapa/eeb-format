Attribute VB_Name = "m_digitalerEEBVokabular"
' ===========================================================================
'  Vokabular-Tabellen für den EEB-Dekoder (m_digitalerEEBParsing).
'
'  GENERIERT - nicht von Hand aendern. Quelle: erfassungsbogen.app,
'  npm run vokabular:excel-vba (scripts/excel-vokabular-vba.mts).
'  347 Einträge, Stand des Vokabulars: bos-vokabulare 2f94a46.
'
'  Die Codes im QR-Code sind append-only: eine ältere Tabelle liest neuere
'  Bögen weiter, unbekannte Codes erscheinen als "#123". Zum Nachziehen das
'  Modul neu erzeugen und in der Mappe ersetzen.
' ===========================================================================
Option Explicit

Private mVokab As Collection

' Klartext zu einem Vokabular-Code. Unbekannter Code -> "#<code>", damit im
' Blatt sichtbar bleibt, dass da ein Wert stand.
Public Function EebVokabText(ByVal schluessel As String, ByVal code As Long) As String
    Dim s As String
    If mVokab Is Nothing Then VokabularAufbauen
    On Error Resume Next
    s = mVokab.item(schluessel)
    On Error GoTo 0
    If Len(s) = 0 Then s = "#" & code
    EebVokabText = s
End Function

Private Sub VokabularAufbauen()
    Set mVokab = New Collection
    ' Einheitstypen (THW) (46)
    EinheitstypenFuellen1
    ' Funktionen und Zusatzqualifikationen (THW) (228)
    FunktionenFuellen1
    FunktionenFuellen2
    FunktionenFuellen3
    FunktionenFuellen4
    ' Fahrzeugtypen (THW) (43)
    FahrzeugeFuellen1
    ' Funkruf-Kennwörter (alle Organisationen) (7)
    KennwoerterFuellen1
    ' Hierarchie-Ebenen je Organisation (23)
    EbenenFuellen1
End Sub

' Ein Eintrag. Doppelte Schlüssel können nicht vorkommen (die Tabellen sind
' generiert), ein Add-Fehler soll aber nicht die ganze Mappe anhalten.
Private Sub v(ByVal schluessel As String, ByVal text As String)
    On Error Resume Next
    mVokab.Add text, schluessel
    On Error GoTo 0
End Sub

' --- Einheitstypen (THW)

Private Sub EinheitstypenFuellen1()
    v "E1", "Media Team"
    v "E2", "Virtual Operations Support Team"
    v "E3", "Zugtrupp Technischer Zug"
    v "E4", "Bergungsgruppe"
    v "E5", "Bergungsgruppe (Abstützsystem Holz)"
    v "E6", "Fachgruppe Räumen (A)"
    v "E7", "Fachgruppe Räumen (B)"
    v "E8", "Fachgruppe Räumen (C)"
    v "E9", "Fachgruppe Wassergefahren (A)"
    v "E10", "Fachgruppe Wassergefahren (B)"
    v "E11", "Fachgruppe Brückenbau"
    v "E12", "Fachgruppe Ortung (A)"
    v "E13", "Fachgruppe Ortung (B)"
    v "E14", "Fachgruppe Ortung (C)"
    v "E15", "Fachgruppe Sprengen"
    v "E16", "Fachgruppe Notversorgung und Notinstandsetzung"
    v "E17", "Schnelleinsatzeinheit Bergung Ausland"
    v "E18", "Fachgruppe Schwere Bergung (A)"
    v "E19", "Fachgruppe Schwere Bergung (B)"
    v "E20", "Trupp Einsatzstellensicherung"
    v "E21", "Trupp Mobiler Hochwasserpegel"
    v "E22", "Trupp Unbemannte Luftfahrtsysteme"
    v "E23", "Fachgruppe Bergungstauchen"
    v "E24", "Fachgruppe Infrastruktur"
    v "E25", "Fachgruppe Elektroversorgung"
    v "E26", "Fachgruppe Trinkwasserversorgung"
    v "E27", "Fachgruppe Wasserschaden/Pumpen (A)"
    v "E28", "Fachgruppe Wasserschaden/Pumpen (B)"
    v "E29", "Fachgruppe Wasserschaden/Pumpen (C)"
    v "E30", "Fachgruppe Ölschaden (A)"
    v "E31", "Fachgruppe Ölschaden (B)"
    v "E32", "Fachgruppe Ölschaden (C)"
    v "E33", "Schnelleinsatzeinheit Wasser Ausland"
    v "E34", "Einsatznachsorgeteam"
    v "E35", "Schnelleinsatzeinheit Logistikabwicklung Flughafen"
    v "E36", "Zugtrupp Fachzug Logistik"
    v "E37", "Fachgruppe Logistik-Materialwirtschaft"
    v "E38", "Fachgruppe Logistik-Verpflegung"
    v "E39", "Trupp Logistik-Schwerer Transport"
    v "E40", "System Bereitstellungsraum 500"
    v "E41", "Zugtrupp Fachzug Führung/Kommunikation"
    v "E42", "Fachgruppe Führungsunterstützung"
    v "E43", "Fachgruppe Kommunikation (A)"
    v "E44", "Fachgruppe Kommunikation (B)"
    v "E45", "Stabspersonal"
    v "E46", "Ortsverband (Stab)"
End Sub

' --- Funktionen und Zusatzqualifikationen (THW)

Private Sub FunktionenFuellen1()
    v "F1", "ZFü"
    v "F2", "ZTrFü"
    v "F3", "GrFü"
    v "F4", "TrFü"
    v "F5", "He"
    v "F6", "He (Res)"
    v "F7", "SGL"
    v "F8", "Ltr VOST"
    v "F9", "stv. Ltr VOST"
    v "F10", "GrLtr VOST"
    v "F11", "TrLtr VOST"
    v "F12", "PSF"
    v "F13", "Peer"
    v "F30", "Spr"
    v "F31", "SanHe"
    v "F32", "AGT"
    v "F33", "CBRN"
    v "F34", "FK IuK"
    v "F35", "FüGeh"
    v "F36", "BoFü"
    v "F37", "BoFü man."
    v "F38", "Bed. Motorsäge"
    v "F39", "Bed. Ladekran"
    v "F40", "Bed. Mobilkran"
    v "F41", "Bed. Bagger/Radlader"
    v "F42", "Bed. Teleskoplader"
    v "F43", "Bed. Schreitbagger"
    v "F44", "Bed. Hubarbeitsbühne"
    v "F45", "Bed. Arbeitsplattform"
    v "F46", "Bed. TOG"
    v "F47", "GabelSt"
    v "F48", "GelSt"
    v "F49", "Masch. NEA"
    v "F50", "Masch. SEA"
    v "F51", "Masch. Pumpen"
    v "F52", "Masch. TWAA"
    v "F53", "Masch. Richtfunk"
    v "F54", "Masch. Separation"
    v "F55", "Masch. Skimmer"
    v "F56", "RettHuFü"
    v "F57", "Spreng"
    v "F58", "SprengGeh"
    v "F59", "THW-Schw"
    v "F60", "PE/PVC-Schw"
    v "F61", "BFB"
    v "F62", "TB Ortung"
    v "F63", "TB Ölschaden"
    v "F64", "TB VOST"
    v "F65", "Abn. EGS"
    v "F66", "NwFü"
    v "F67", "LogFü"
    v "F68", "Ltr THW-FüSt"
    v "F69", "Ltr Stab THW-FüSt"
    v "F70", "Ltr FmBetrieb"
    v "F71", "FmFü"
    v "F72", "Ltr ENT"
    v "F73", "SGL S1"
    v "F74", "SGL S2"
    v "F75", "SGL S3"
    v "F76", "SGL S4"
End Sub

Private Sub FunktionenFuellen2()
    v "F77", "SGL S5"
    v "F78", "SGL S6"
    v "F79", "BergTaucher"
    v "F80", "TauchSanHe"
    v "F81", "TauchGerWart"
    v "F82", "LuFzFernFü"
    v "F83", "Ausw. Fernerkundung"
    v "F84", "TW-Laborant"
    v "F85", "Hygiene"
    v "F86", "Koch"
    v "F87", "GerWart St V"
    v "F88", "Bef. P. Logistik"
    v "F89", "Bef. P. Technik"
    v "F90", "Bef. P. Elektro"
    v "F91", "Hägglunds"
    v "F92", "Fotograf"
    v "F93", "Redakteur"
    v "F94", "Presse MT"
    v "F95", "SoMe"
    v "F96", "Videograf"
    v "F100", "Kf ADR Stückgut"
    v "F101", "Kf ADR Tank"
    v "F102", "Kf ADR Kl. 1"
    v "F200", "AE-He"
    v "F201", "AB"
    v "F202", "BÖ"
    v "F203", "Fü ÖGA"
    v "F204", "FaBe 1/2"
    v "F205", "FaBe1/2R"
    v "F206", "FaHe ESS"
    v "F207", "FaHe MHP"
    v "F208", "FernMel"
    v "F209", "GrFü Bel"
    v "F210", "GrFü B"
    v "F211", "GrFü B 2"
    v "F212", "GrFü BT"
    v "F213", "GrFü BrB"
    v "F214", "GrFü E"
    v "F215", "GrFü SEE"
    v "F216", "GrFü F"
    v "F217", "GrFü I"
    v "F218", "GrFü K"
    v "F219", "GrFüLogMW"
    v "F220", "GrFüLogV"
    v "F221", "GrFü MT"
    v "F222", "GrFü N"
    v "F223", "GrFü ÖGA"
    v "F224", "GrFü Öl"
    v "F225", "GrFü O"
    v "F226", "GrFü R"
    v "F227", "GrFü SB"
    v "F228", "GrFü Sp"
    v "F229", "GrFü TW"
    v "F230", "GrFü W"
    v "F231", "GrFü WP"
    v "F232", "HeAw"
    v "F233", "JuHe"
    v "F234", "Ltr FK"
    v "F235", "Ltr Log"
    v "F236", "LdF ÖGA"
End Sub

Private Sub FunktionenFuellen3()
    v "F237", "Koch"
    v "F238", "OB"
    v "F239", "OJB"
    v "F240", "PsFk (R)"
    v "F241", "S 6/FeFü"
    v "F242", "SM"
    v "F243", "SPLtr"
    v "F244", "Peer (R)"
    v "F245", "TrFü Bel"
    v "F246", "TrFü B"
    v "F247", "TrFü B S"
    v "F248", "TrFü BT"
    v "F249", "TrFü BrB"
    v "F250", "TrFü ESS"
    v "F251", "TrFü E"
    v "F252", "TrFü FmT"
    v "F253", "TrFü FuK"
    v "F254", "TrFü Fü"
    v "F255", "TrFü F"
    v "F256", "TrFü I"
    v "F257", "TrFü K"
    v "F258", "TrFü Log"
    v "F259", "TrFüLogMW"
    v "F260", "TrFüLogV"
    v "F261", "TrFü MT"
    v "F262", "TrFü MHP"
    v "F263", "TrFü N"
    v "F264", "TrFü ÖGA"
    v "F265", "TrFü Öl"
    v "F266", "TrFü O"
    v "F267", "TrFü O S"
    v "F268", "TrFü R"
    v "F269", "TrFü SB"
    v "F270", "TrFü TS"
    v "F271", "TrFü Sp"
    v "F272", "TrFü TW"
    v "F273", "TrFü UL"
    v "F274", "TrFü W"
    v "F275", "TrFü WP"
    v "F276", "TrFü WV"
    v "F277", "TrFü FGe"
    v "F278", "VwB"
    v "F279", "ZFü FZ FK"
    v "F280", "ZFüFZLog"
    v "F281", "ZTrFü FK"
    v "F282", "ZTrFüLog"
    v "F283", "stv. OB"
    v "F284", "stv. OJB"
    v "F285", "AS GerW"
    v "F286", "AusMotSä"
    v "F287", "ABO"
    v "F288", "AusbGA"
    v "F289", "BPLogBGR"
    v "F290", "BA AS"
    v "F291", "BA IT"
    v "F292", "BA Koch"
    v "F293", "BA Kf"
    v "F294", "BA SprFu"
    v "F295", "BezJB"
    v "F296", "BJB"
End Sub

Private Sub FunktionenFuellen4()
    v "F297", "BuJL"
    v "F298", "BJSp"
    v "F299", "BS"
    v "F300", "ELA"
    v "F301", "GerW"
    v "F302", "HeSpr"
    v "F303", "ITB"
    v "F304", "Koch Ausl"
    v "F305", "KB"
    v "F306", "LJB"
    v "F307", "LS"
    v "F308", "PGA"
    v "F309", "PRH"
    v "F310", "PLGA"
    v "F311", "SiBe"
    v "F312", "SiFa"
    v "F313", "SvP-OV"
    v "F314", "UmwAud"
    v "F315", "Ver DS"
    v "F316", "stvBezJB"
    v "F317", "stv. BJB"
    v "F318", "stv BuJL"
    v "F319", "stv BJSp"
    v "F320", "stv BS"
    v "F321", "stvHeSpr"
    v "F322", "stv. LJB"
    v "F323", "stv LS"
    v "F324", "ABC-FaHe"
    v "F325", "Bau ESS"
    v "F326", "BoFüSee"
    v "F327", "E-Schw."
    v "F328", "FaBe 3/4"
    v "F329", "G-Schw."
    v "F330", "G-S. 1BG"
    v "F331", "GerWLogV"
    v "F332", "HygFa"
    v "F333", "Kf Häg"
    v "F334", "KüHe"
    v "F335", "LehrTau"
    v "F336", "LtrStabE"
    v "F337", "LtrFüStE"
    v "F338", "Masch DL"
    v "F339", "OrtHe"
    v "F340", "RechFü"
    v "F341", "S1-4"
    v "F342", "SM(SBA)"
    v "F343", "TeBe ABC"
    v "F344", "TeBeDe"
End Sub

' --- Fahrzeugtypen (THW)

Private Sub FahrzeugeFuellen1()
    v "V1", "Fernmeldekraftwagen"
    v "V2", "Führungskraftwagen"
    v "V3", "Führungs- und Kommunikationskraftwagen"
    v "V4", "Gerätekraftwagen"
    v "V5", "Mannschaftstransportwagen Fachgruppe"
    v "V6", "Mannschaftslastwagen IV"
    v "V7", "Lastkraftwagen Kipper"
    v "V8", "Lastkraftwagen mit Ladebordwand"
    v "V9", "Lastkraftwagen mit Ladekran"
    v "V10", "Lastkraftwagen geländegängig mit Ladekran"
    v "V11", "Lastkraftwagen Wechsellader"
    v "V12", "LKW WLF Abrollsystem Tank"
    v "V13", "Sattelzugmaschine"
    v "V14", "Auflieger Sattelzug"
    v "V15", "Personenkraftwagen geländegängig"
    v "V16", "Baumaschine Bagger"
    v "V17", "Baumaschine Radlader"
    v "V18", "Baumaschine Teleskoplader"
    v "V19", "Schreitbagger"
    v "V20", "Mobilkran"
    v "V21", "Gabelstapler"
    v "V22", "Mehrzweckarbeitsboot"
    v "V23", "Mannschaftstransportwagen Technischer Zug"
    v "V24", "Mannschaftstransportwagen geländegängig"
    v "V25", "Mehrzweckkraftwagen"
    v "V26", "Mehrzweckgerätewagen"
    v "V27", "Mannschaftstransportwagen Ortsverband"
    v "V40", "Anhänger (2 t Nutzlast)"
    v "V41", "Anhänger mit Spezialaufbau für FGr K (1 t Nutzlast)"
    v "V42", "Anhänger Führung und Lage"
    v "V43", "Anhänger Plane/Spriegel"
    v "V44", "Anhänger Plattform"
    v "V45", "Anhänger Tieflader"
    v "V46", "Anhänger Drucklufterzeuger"
    v "V47", "Anhänger mit Netzersatzanlage (mittel)"
    v "V48", "Anhänger mit Netzersatzanlage (groß)"
    v "V49", "Anhänger mit Netzersatzanlage (sehr groß)"
    v "V50", "Anhänger mit Schmutzwasserpumpe (klein)"
    v "V51", "Anhänger mit Schmutzwasserpumpe (mittel)"
    v "V52", "Anhänger mit Schmutzwasserpumpe (groß)"
    v "V53", "Anhänger TWAA"
    v "V54", "Anhänger BDF-Lafette"
    v "V55", "Anhänger mit Abstützsystem Holz"
End Sub

' --- Funkruf-Kennwörter (alle Organisationen)

Private Sub KennwoerterFuellen1()
    v "K1", "Heros"
    v "K2", "Florian"
    v "K3", "Rotkreuz"
    v "K4", "Akkon"
    v "K5", "Johannes"
    v "K6", "Sama"
    v "K7", "Pelikan"
End Sub

' --- Hierarchie-Ebenen je Organisation

Private Sub EbenenFuellen1()
    v "H1_1", "OV"
    v "H1_2", "RB"
    v "H1_3", "LV"
    v "H2_1", "Gemeinde"
    v "H2_2", "LK"
    v "H2_3", "Bezirk"
    v "H2_4", "Land"
    v "H5_1", "OV"
    v "H5_2", "KV"
    v "H5_3", "LV"
    v "H6_1", "OV"
    v "H6_2", "RV"
    v "H6_3", "LV"
    v "H7_1", "OG"
    v "H7_2", "DG"
    v "H7_3", "LG"
    v "H8_1", "OV"
    v "H8_2", "KV"
    v "H8_3", "RV"
    v "H8_4", "LV"
    v "H9_1", "OG"
    v "H9_2", "Bezirk"
    v "H9_3", "LV"
End Sub

