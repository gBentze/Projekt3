# Projekt3
Option Compare Database

'Tabelle Abteilung, Mitkreis, typeFzgGrp
Const dateipfad = "S:\VDG-VFD\Laure\LaureDokumente\Praktikum\MeineAufgaben\Abschlussarbeitsphase\ProjektAccessDashboard\Daten"
Const xlBlattName As String = "Tabelle1"

 Option Compare Database
'Tabelle Abteilung, Mitkreis, typeFzgGrp
Const dateipfad = "D:\DantenbankAccess\Abschlussphase"
Const xlBlattName As String = "Tabelle1"

Private Type FrzStruc
    stammNr As Long
    mitarbeiterName As String
    mitarbeitervorname As String '
    fin As String
    modellBzg As String
    herstName As String
    grundPreis As Double
    modellName As String
    bestellNr As String
    bestelldat As Date
    kfzbrief As String
    rechnungDat As Date
    vertragsende As Date
    dLLeasingRate As Double
    leasingNr As String
    kennzeichen As String
    vertragsdauer As Integer
    kiloStand As Long
    datEz As Date
    'gisUebernahme As Boolean
    rueckgabeDat As Date
    netLesingrate As Double
    'kaufinteresse As Boolean
    verkaufDat As Date
    anlieferungDat As Date
    schaeden As Double
    abmeldeDat As Date
    austattungWert As Double
    verkPreis As Double
    fzgStatus As String
    kraftstoff As String
    dienstwagenNz As String ' Variable sert a extraire le nom et le prenom
End Type

Private FzgXLSXdata() As FrzStruc
'Anzahl Zeilen in Excel Tabelle
Private XLSXmax As Integer
Private Sub test()
Call ImportDaten
Call WriteXLSXDaten
Call CloseXLSXApp(True)
End Sub
Private Sub ImportDaten()
Dim xlpfad As String
xlpfad = dateipfad & "\FzgAll.xlsx"
Dim vFin As Variant, vModellBzg, vHerstName, vGrundPreis, vModellName, vBestellNr, vBestelldat, vKfzBrief, vRechnungDat, vVertragsende
Dim vDLLeasingRate As Variant, vLeasingNr, vKennzeichen, vVertragsdauer, vKiloStand, vDatEz
Dim vRueckgabeDat, vNetLesingrate, vVerkaufDat, vAnlieferungDat, vSchaeden, vAbmeldeDat
Dim vAustattungWert As Variant, vVerkPreis, vFzgStatus, vKraftstoff, vStammNr, vDienstwagenNz
'vKaufinteresse, vGisUebernahme
'Dim dienstwagenNz As String ' Variable sert a extraire le nom et le prenom

Dim i As Long
Dim iRowS As Integer
Dim iRowL As Long
Dim iCol As Integer
Dim sCol As String
' Verweis auf Excel-Bibliothek muss gesetzt sein
Dim xlsApp As Excel.Application
Dim Blatt As Excel.Worksheet
Dim MsgAntw As Integer
' Konstante: Name des einzulesenden Arbeitsblattes
' Excel vorbereiten
On Error Resume Next
Set xlsApp = GetObject(, "Excel.Application")
If xlsApp Is Nothing Then
    Set xlsApp = CreateObject("Excel.Application")
End If
On Error GoTo 0
' Exceldatei readonly öffnen
xlsApp.Workbooks.Open xlpfad, , True
    ' Erste Zeile wird statisch angegeben
    iRowS = 2
    ' Letzte Zeile auf Tabellenblatt wird dynamisch ermittelt
    iRowL = xlsApp.Worksheets(xlBlattName).Cells(xlsApp.Rows.Count, 1).End(xlUp).Row
 
  
    'xlBlattName excel Datei FzgAllMitKunde.xslx
    'Stammnummer
    sCol = "A"
    vStammNr = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Dienstwagennutzer
    sCol = "B"
    vDienstwagenNz = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Fahrzeug-Identifikationsnummer
    sCol = "C"
    vFin = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Modell/Typ
    sCol = "D"
    vModellName = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'HERSTNAME
    sCol = "E"
    vHerstName = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Grundpreis
    sCol = "F"
    vGrundPreis = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    ' ModellNr
    sCol = "G"
    vModellBzg = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Bestellnummer
    sCol = "H"
    vBestellNr = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Bestelldatum
    sCol = "I"
    vBestelldat = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Kraftfahrzeugbriefnummer
    sCol = "j"
    vKfzBrief = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Rechnungsdatum
    sCol = "K"
    vRechnungDat = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Status
    sCol = "L"
    vFzgStatus = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Vertragsende
    sCol = "M"
    vVertragsende = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'DL Leasing Rate
    sCol = "N"
    vDLLeasingRate = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Leasing Vertragsnummer
    sCol = "O"
    vLeasingNr = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'amtl. KFZ Kennzeichen
    sCol = "P"
    vKennzeichen = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Vertragsdauer
    sCol = "Q"
    vVertragsdauer = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Kilometer-Stand
    sCol = "R"
    vKiloStand = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Datum EZ
    sCol = "S"
    vDatEz = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Fahrzeug ins GIS übernommen
'    sCol = "T"
'    vUebernahme = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Rueckgabedatum
    sCol = "U"
    vRueckgabeDat = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Netto-Leasingrate
    sCol = "V"
    vNetLesingrate = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Kaufinteresse
'    sCol = "W"
'    vKaufinteresse = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Verkaufsdatum
    sCol = "X"
    vVerkaufDat = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Anlieferungsdatum
    sCol = "Y"
    vAnlieferungDat = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Schäden
    sCol = "Z"
    vSchaeden = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Abmeldatum
    sCol = "AA"
    vAbmeldeDat = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Wert der Ausstattung
    sCol = "AB"
    vAustattungWert = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Verkaufpreis
    sCol = "AC"
    vVerkPreis = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Motor
    sCol = "AD"
    vKraftstoff = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
  

    '--------------------
    ' Werte in globaler Variable speichern
    XLSXmax = iRowL - 1
    ReDim FzgXLSXdata(1 To XLSXmax)
  
   
'Set xlsApp = Nothing

    With xlsApp.Worksheets(xlBlattName)
        For i = 1 To XLSXmax
                  
            FzgXLSXdata(i).abmeldeDat = vAbmeldeDat(i, 1)
            FzgXLSXdata(i).anlieferungDat = vAnlieferungDat(i, 1)
            FzgXLSXdata(i).austattungWert = vAustattungWert(i, 1)
            FzgXLSXdata(i).bestelldat = vBestelldat(i, 1)
            FzgXLSXdata(i).bestellNr = vBestellNr(i, 1)
            FzgXLSXdata(i).datEz = vDatEz(i, 1)
            FzgXLSXdata(i).dLLeasingRate = vDLLeasingRate(i, 1)
            FzgXLSXdata(i).fin = vFin(i, 1)
            'FzgXLSXdata(i).gisUebernahme = vGisUebernahme(i, 1)'Boolean?
            FzgXLSXdata(i).grundPreis = vGrundPreis(i, 1)
            FzgXLSXdata(i).herstName = vHerstName(i, 1)
            'FzgXLSXdata(i).kaufinteresse = vKaufinteresse(i, 1)'Boolean?
            FzgXLSXdata(i).kennzeichen = vKennzeichen(i, 1)
            FzgXLSXdata(i).kfzbrief = vKfzBrief(i, 1)
            FzgXLSXdata(i).kiloStand = vKiloStand(i, 1)
            FzgXLSXdata(i).leasingNr = vLeasingNr(i, 1)
            FzgXLSXdata(i).modellBzg = vModellBzg(i, 1)
            FzgXLSXdata(i).modellName = vModellName(i, 1)
            FzgXLSXdata(i).netLesingrate = vNetLesingrate(i, 1)
            FzgXLSXdata(i).rechnungDat = vRechnungDat(i, 1)
            FzgXLSXdata(i).rueckgabeDat = vRueckgabeDat(i, 1)
            FzgXLSXdata(i).schaeden = vSchaeden(i, 1)
            FzgXLSXdata(i).verkaufDat = vVerkaufDat(i, 1)
            FzgXLSXdata(i).verkPreis = vVerkPreis(i, 1)
            FzgXLSXdata(i).vertragsdauer = vVertragsdauer(i, 1)
            FzgXLSXdata(i).vertragsende = vVertragsende(i, 1)
            FzgXLSXdata(i).fzgStatus = vFzgStatus(i, 1)
            FzgXLSXdata(i).kraftstoff = vKraftstoff(i, 1)
            FzgXLSXdata(i).stammNr = vStammNr(i, 1)
            '#################### extraire le nom et le prenom
            dienstwagenNz = vDienstwagenNz(i, 1)
            FzgXLSXdata(i).mitarbeiterName = Mid(dienstwagenNz, 1, InStr(1, dienstwagenNz, ",", vbTextCompare) - 1)
            FzgXLSXdata(i).mitarbeitervorname = Mid(dienstwagenNz, InStr(1, dienstwagenNz, ",", vbTextCompare) + 1, Len(dienstwagenNz))

        Next i
    End With
Set xlsApp = Nothing

End Sub
Private Sub WriteXLSXDaten()
' Daten aus Variable in Tabelle übertragen
Dim i As Integer
'Dim lngMitKreisID As Long
Dim sSQL As String
Dim bestellID As Long, statusID As Long, herstID As Long, kraftID As Long, fzgModellID As Long, fzgID As Long, austattungID As Long, dwNzrID As Long, stammNrID As Long
Dim MsgAntw As Integer, mitID As Long
' Schleife über alle Datensätze in der Variablen
For i = 1 To XLSXmax
    ' SQL-String erstellen und Daten schreiben
  
    bestellID = Nz(DLookup("BestellID", "tblBestellung", "BestellNr = '" & FzgXLSXdata(i).bestellNr & "'"), 0)
    statusID = Nz(DLookup("StatusID", "tblFzgStatus", "Status = '" & FzgXLSXdata(i).fzgStatus & "'"), 0)
    herstID = Nz(DLookup("HerstID", "tblHersteller", "HerstName = '" & FzgXLSXdata(i).herstName & "'"), 0)
    kraftID = Nz(DLookup("KraftID", "tblKraftstoff", "Kraftstoffart= '" & FzgXLSXdata(i).kraftstoff & "'"), 0)
    mitID = Nz(DLookup("MitID", "tblMitarbeiter", "StammNr= " & FzgXLSXdata(i).stammNr & ""), 0)
  
    fzgModellID = Nz(DLookup("FzgModellID", "tblFzgModell", "ModellNr= '" & FzgXLSXdata(i).modellBzg & "'"), 0)
    fzgID = Nz(DLookup("FzgID", "tblFzg", "FahrgestellNr= '" & FzgXLSXdata(i).fin & "'"), 0)
    DoCmd.SetWarnings False
'1.######################################## tblBestellung#################################################
    If bestellID = 0 Then
        sSQL = "INSERT INTO tblBestellung  (BestellNr,BestellDat , Kraftfahrzeugbriefnummer, AnlieferungDat, Rechnungsdatum )" & _
        "VALUES ('" & FzgXLSXdata(i).bestellNr & "', '" & FzgXLSXdata(i).bestelldat & "', '" & FzgXLSXdata(i).kfzbrief & "','" & _
        FzgXLSXdata(i).anlieferungDat & "','" & FzgXLSXdata(i).rechnungDat & "');"
        DoCmd.RunSQL sSQL
        bestellID = Nz(DLookup("BestellID", "tblBestellung", "BestellNr= '" & FzgXLSXdata(i).bestellNr & "'"))
'    Else
'        sSQL = "UPDATE tblBestellung SET (" & _
'            " BestellNr = '" & FzgXLSXdata(i).bestellNr & "'" & _
'            ",BestellDat='" & FzgXLSXdata(i).bestelldat & "'" & _
'            ",Kraftfahrzeugbriefnummer='" & FzgXLSXdata(i).kfzbrief & "'" & _
'            ",AnlieferungDat = '" & FzgXLSXdata(i).anlieferungDat & "'" & _
'            ",Rechnungsdatum = '" & FzgXLSXdata(i).rechnungDat & "') WHERE BestellID = " & bestellID & ";"
'        DoCmd.RunSQL sSQL
    End If
'2.######################################## tblStatus #################################################
    If statusID = 0 Then
        sSQL = "INSERT INTO tblFzgStatus (Status) VALUES ('" & FzgXLSXdata(i).fzgStatus & "');"
        DoCmd.RunSQL sSQL
        statusID = Nz(DLookup("StatusID", "tblFzgStatus", "Status= '" & FzgXLSXdata(i).fzgStatus & "'"))
      
    End If
  
'#########################################tblDienstwagennutzer##########################################

    If mitID = 0 Then
        sSQL = "INSERT INTO tblMitarbeiter(Nachname, Vorname, StammNr)VALUES('" & _
         FzgXLSXdata(i).mitarbeiterName & "','" & FzgXLSXdata(i).mitarbeitervorname & "'," & FzgXLSXdata(i).stammNr & ");"
            DoCmd.RunSQL sSQL
      mitID = Nz(DLookup("MitID", "tblMitarbeiter", "StammNr= " & FzgXLSXdata(i).stammNr & ""))
    End If

  
'3.######################################## tblHersteller #################################################
    If herstID = 0 Then
     sSQL = "INSERT INTO tblHersteller (HerstName ) VALUES ('" & FzgXLSXdata(i).herstName & "');"
        DoCmd.RunSQL sSQL
     herstID = Nz(DLookup("HerstID", "tblHersteller", "HerstName= '" & FzgXLSXdata(i).herstName & "'"))
    End If

'4.######################################## tblKraftstoff #################################################
    If kraftID = 0 Then
     sSQL = "INSERT INTO tblKraftstoff (Kraftstoffart ) VALUES ('" & FzgXLSXdata(i).kraftstoff & "');"
        DoCmd.RunSQL sSQL
     kraftID = Nz(DLookup("KraftID", "tblKraftstoff", "Kraftstoffart= '" & FzgXLSXdata(i).kraftstoff & "'"))
    End If

''5.######################################## tblAustattung #################################################
'    If austattungID = 0 Then
'     sSQL = "INSERT INTO tblAustattung (Austattung ) VALUES (" & Str(FzgXLSXdata(i).austattungWert) & ");"
'        DoCmd.RunSQL sSQL
'     austattungID = Nz(DLookup("AustattungID", "tblAustattung", "Austattung= " & Str(FzgXLSXdata(i).austattungWert) & ""))
'     Debug.Print sSQL
'    End If
'6.######################################## tblFzgModell #################################################
     ' Abteilungskuerzel mit den Schlüsselwerten in Variablen ersetzen
    If fzgModellID = 0 Then
         sSQL = "INSERT INTO tblFzgModell (KraftID, HerstID, ModellTyp, ModellNr) " & _
         "VALUES(" & kraftID & "," & herstID & ",'" & FzgXLSXdata(i).modellName & "', '" & _
         FzgXLSXdata(i).modellBzg & "');"

        Debug.Print sSQL
        DoCmd.RunSQL sSQL
        fzgModellID = Nz(DLookup("FzgModellID", "tblFzgModell", "ModellNr= '" & FzgXLSXdata(i).modellBzg & "'"))
        'in diesem Fall werden alle Fzg mit dem gleichen ModellBzg nicht gespeichert. Eine Möglichkeit, um das zu ändern und einfach alles Daten einfügen?
    End If
'7######################################## tblFzg #################################################
    If fzgID = 0 Then
         sSQL = "INSERT INTO tblFzg (StatusID, AustattungID, FzgModellID, FahrgestellNr, GrundPreis) " & _
         "VALUES(" & statusID & ", " & austattungID & "," & fzgModellID & ",'" & FzgXLSXdata(i).fin & "', " & _
         Str(FzgXLSXdata(i).grundPreis) & ");"

        Debug.Print sSQL
        DoCmd.RunSQL sSQL
        fzgID = Nz(DLookup("FzgID", "tblFzg", "FahrgestellNr= '" & FzgXLSXdata(i).fin & "'")) ', " & Str(Nz(FzgXLSXdata(i).grundPreis)) & ""))
        ' in diesem Fall werden alle Fzg mit dem gleichen ModellBzg nicht gespeichert.
    End If

''8.######################################## tblVerkFzgAus #################################################
'
'    If IsNull(DLookup("AustattungID", "tblVerknuepftFzg_Aust", "AustattungID=" & austattungID & " AND FzgID= " & fzgID)) Then
'        sSQLMit = "INSERT INTO tblVerknuepftFzg_Aust (AustattungID,fzgID) VALUES (" _
'        & austattungID & "," & fzgID & ");"
'        Debug.Print sSQLMit
'        DoCmd.RunSQL sSQLMit
'    End If
''7######################################## tblModellTyp_Kraftstoff #################################################
'
'    If modKrafID = 0 Then
'        sSQL = "INSERT INTO tblModellTyp_Kraftstoff (ModBz, KraftNr) " & _
'        "VALUES (" & FzgXLSXdata(i).modellBzg & ", " & FzgXLSXdata(i).kraftstoffNr & ");"
'        Debug.Print sSQL
'        DoCmd.RunSQL sSQL
'    End If
DoCmd.SetWarnings True
Next i
'###############################Feststellungen und Vorschläge ################################################
  
End Sub
Private Sub CloseXLSXApp(bShowInfo As Boolean)
''' Excel-Instanz beenden
' Verweis auf Excel-Bibliothek muss gesetzt sein
Dim xlsApp As Excel.Application
Dim MsgAntw As Integer
' Excel-Instanz suchen
  On Error Resume Next
    Set xlsApp = GetObject(, "Excel.Aplication")
    If xlsApp Is Nothing Then
        ' keine Excel-Instanz vorhanden
        ' Meldung
        If bShowInfo Then
            MsgAntw = MsgBox("Es wurde keine Excel-Instanz gefunden.", vbInformation, "Excel-Instanz beenden")
        End If
        ' Ende
        Exit Sub
    End If
On Error GoTo 0
    ' Excel schließen und resetten
    xlsApp.Quit
    Set xlsApp = Nothing
    ' Meldung
    If bShowInfo Then
        MsgAntw = MsgBox("Die Excel-Instanz wurde beendet.", vbInformation, "Excel-Instanz beenden")
    End If
End Sub








































































































































































































































































































