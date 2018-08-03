# Projekt3
Option Compare Database

'Tabelle Abteilung, Mitkreis, typeFzgGrp
Const dateipfad = "S:\VDG-VFD\Laure\LaureDokumente\Praktikum\MeineAufgaben\Abschlussarbeitsphase\ProjektAccessDashboard\Daten"
Const xlBlattName As String = "Tabelle1"

Private Type FrzStruc
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
Dim xlpfad1 As String
xlpfad = dateipfad & "\FzgAll.xlsx"
Dim vFin As Variant, vModellBzg, vHerstName, vGrundPreis, vModellName, vBestellNr, vBestelldat, vKfzBrief, vRechnungDat, vVertragsende
Dim vDLLeasingRate As Variant, vLeasingNr, vKennzeichen, vVertragsdauer, vKiloStand, vDatEz
Dim vRueckgabeDat, vNetLesingrate, vVerkaufDat, vAnlieferungDat, vSchaeden, vAbmeldeDat
Dim vAustattungWert As Variant, vVerkPreis, vFzgStatus, vKraftstoff As Variant
'vKaufinteresse, vGisUebernahme
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
  
    
    'xlBlattName excel Datei FzgAll
    sCol = "A"
    vFin = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Modell/Typ
    sCol = "B"
    vModellName = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'HerstName
    sCol = "C"
    vHerstName = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Grundpreis
    sCol = "D"
    vGrundPreis = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'ModellNr
    sCol = "E"
    vModellBzg = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Bestellnummer
    sCol = "F"
    vBestellNr = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Bestelldatum
    sCol = "G"
    vBestelldat = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Kraftfahrzeugbriefnummer
    sCol = "H"
    vKfzBrief = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Rechnungsdatum
    sCol = "I"
    vRechnungDat = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Status
    sCol = "j"
    vFzgStatus = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Vertragsende
    sCol = "K"
    vVertragsende = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'DL Leasing Rate
    sCol = "L"
    vDLLeasingRate = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Leasing Vertragsnummer
    sCol = "M"
    vLeasingNr = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'amtl. KFZ Kennzeichen
    sCol = "N"
    vKennzeichen = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Vertragsdauer
    sCol = "O"
    vVertragsdauer = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Kilometer-Stand
    sCol = "P"
    vKiloStand = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Datum EZ
    sCol = "Q"
    vDatEz = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Fahrzeug ins GIS übernommen
'    sCol = "R"
'    vGisUebernahme = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Rueckgabedatum
    sCol = "S"
    vRueckgabeDat = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Netto-Leasingrate
    sCol = "K"
    vVertragsende = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Netto-LeasingRate
    sCol = "T"
    vNetLesingrate = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Kaufinteresse
'    sCol = "U"
'    vKaufinteresse = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Verkaufsdatum
    sCol = "V"
    vVerkaufDat = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Anlieferungsdatum
    sCol = "W"
    vAnlieferungDat = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Schäden
    sCol = "X"
    vSchaeden = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Abmeldedatum
    sCol = "Y"
    vAbmeldeDat = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Wert der Ausstattung
    sCol = "Z"
    vAustattungWert = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Verkaufspreis
    sCol = "AA"
    vVerkPreis = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)

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
'            FzgXLSXdata(i).kraftstoff = vKraftstoff(i, 1)
        
            

        Next i
    End With
Set xlsApp = Nothing

End Sub
Private Sub WriteXLSXDaten()
' Daten aus Variable in Tabelle übertragen
Dim i As Integer
'Dim lngMitKreisID As Long
Dim sSQL As String
Dim bestellID As Long, statusID As Long, herstID As Long, kraftID As Long, fzgModellID As Long, fzgID As Long, austattungID As Long
Dim MsgAntw As Integer
' Schleife über alle Datensätze in der Variablen
For i = 1 To XLSXmax
    ' SQL-String erstellen und Daten schreiben
    
    bestellID = Nz(DLookup("BestellID", "tblBestellung", "BestellNr = '" & FzgXLSXdata(i).bestellNr & "'"), 0)
    statusID = Nz(DLookup("StatusID", "tblFzgStatus", "Status = '" & FzgXLSXdata(i).fzgStatus & "'"), 0)
    herstID = Nz(DLookup("HerstID", "tblHersteller", "HerstName = '" & FzgXLSXdata(i).herstName & "'"), 0)
    kraftID = Nz(DLookup("KraftID", "tblKraftstoffArt", "Kraftstoff= '" & FzgXLSXdata(i).kraftstoff & "'"), 0)
    'austattungID = Nz(DLookup("AustattungID", "tblAustattung", "Austattung= " & FzgXLSXdata(i).austattungWert & ""), 0)
    ' Double macht fehler
    
    fzgModellID = Nz(DLookup("FzgModellID", "tblFzgModell", "ModellBzg= '" & FzgXLSXdata(i).modellBzg & "'"), 0)
    fzgID = Nz(DLookup("FzgID", "tblFzg", "FahrgestellNr= '" & FzgXLSXdata(i).fin & "'"), 0)
    DoCmd.SetWarnings False
'1.######################################## tblBestellung#################################################
    If bestellID = 0 Then
        sSQL = "INSERT INTO tblBestellung (BestellNr,BestellDat, Kraftfahrzeugbriefnummer, AnlieferungDat, Rechnungsdatum )" & _
        "VALUES ('" & FzgXLSXdata(i).bestellNr & "', '" & FzgXLSXdata(i).bestelldat & "', '" & FzgXLSXdata(i).kfzbrief & "','" & _
        FzgXLSXdata(i).anlieferungDat & "','" & FzgXLSXdata(i).rechnungDat & "');"
        
        DoCmd.RunSQL sSQL
        bestellID = Nz(DLookup("BestellID", "tblBestellung", "BestellNr= '" & FzgXLSXdata(i).bestellNr & "'"))
    End If
'2.######################################## tblStatus #################################################
    If statusID = 0 Then
        sSQL = "INSERT INTO tblFzgStatus (Status) VALUES ('" & FzgXLSXdata(i).fzgStatus & "');"
        DoCmd.RunSQL sSQL
        statusID = Nz(DLookup("StatusID", "tblFzgStatus", "Status= '" & FzgXLSXdata(i).fzgStatus & "'"))
        
    End If
'3.######################################## tblHersteller #################################################
    If herstID = 0 Then
     sSQL = "INSERT INTO tblHersteller (HerstName ) VALUES ('" & FzgXLSXdata(i).herstName & "');"
        DoCmd.RunSQL sSQL
     herstID = Nz(DLookup("HerstID", "tblHersteller", "HerstName= '" & FzgXLSXdata(i).herstName & "'"))
    End If
    
'4.######################################## tblKraftstoff #################################################
    If kraftID = 0 Then
'     sSQL = "INSERT INTO tblKraftstoffArt (Kraftstoff ) VALUES ('" & FzgXLSXdata(i).kraftstoff & "');"
'        DoCmd.RunSQL sSQL
     kraftID = Nz(DLookup("KraftID", "tblKraftstoffArt", "Kraftstoff= '" & FzgXLSXdata(i).kraftstoff & "'"))
    End If
    
'5.######################################## tblAustattung #################################################
    If austattungID = 0 Then
     sSQL = "INSERT INTO tblAustattung (Austattung ) VALUES (" & Str(FzgXLSXdata(i).austattungWert) & ");"
        DoCmd.RunSQL sSQL
     austattungID = Nz(DLookup("AustattungID", "tblAustattung", "Austattung= " & Str(FzgXLSXdata(i).austattungWert) & ""))
     Debug.Print sSQL
    End If
'6.######################################## tblFzgModell #################################################
     ' Abteilungskuerzel mit den Schlüsselwerten in Variablen ersetzen
    If fzgModellID = 0 Then
         sSQL = "INSERT INTO tblFzgModell (KraftID, HerstID, Modellname, ModellBzg) " & _
         "VALUES(" & kraftID & "," & herstID & ",'" & FzgXLSXdata(i).modellName & "', '" & _
         FzgXLSXdata(i).modellBzg & "');"
         
        Debug.Print sSQL
        DoCmd.RunSQL sSQL
        fzgModellID = Nz(DLookup("FzgModellID", "tblFzgModell", "ModellBzg= '" & FzgXLSXdata(i).modellBzg & "'"))
        ' in diesem Fall werden alle Fzg mit dem gleichen ModellBzg nicht gespeichert.
    End If
'7######################################## tblFzg #################################################
    If fzgID = 0 Then
         sSQL = "INSERT INTO tblFzg (StatusID, AustattungID, FzgModellID, FahrgestellNr, GrundPreis) " & _
         "VALUES(" & statusID & ", " & austattungID & "," & fzgModellID & ",'" & FzgXLSXdata(i).fin & "', " & _
         Str(FzgXLSXdata(i).grundPreis) & ");"
         
        Debug.Print sSQL
        DoCmd.RunSQL sSQL
        'fzgID = Nz(DLookup("FzgID", "tblFzg", "FahrgestellNr= '" & FzgXLSXdata(i).fin & "', " & Str(Nz(FzgXLSXdata(i).grundPreis)) & ""))
        ' in diesem Fall werden alle Fzg mit dem gleichen ModellBzg nicht gespeichert.
    End If
'8.######################################## tblVerkFzgAus #################################################
    
    If IsNull(DLookup("AustattungID", "tblVerknuepftFzg_Aust", "AustattungID=" & austattungID & " AND FzgID= " & fzgID)) Then
        sSQLMit = "INSERT INTO tblVerknuepftFzg_Aust (AustattungID,fzgID) VALUES (" _
        & austattungID & "," & fzgID & ");"
        Debug.Print sSQLMit
        DoCmd.RunSQL sSQLMit
    End If
'7######################################## tblModellTyp_Kraftstoff #################################################
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
    'eine Kostenstelle kann mehrere Abteilungen haben und eine Abteilung kann mehrere Kostenstellen haben.
    'also m:n Beziehung
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























































































































































































































































































































