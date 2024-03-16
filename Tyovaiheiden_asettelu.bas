
''''''''''''''''''''''''''''''''''''''''
' VÄLILEHDEN OLEMASSAOLON TESTAUS (FUNKTIO)
' Tarkistetaan, että välilehteä olemassa
' Palauttaa True tai False
''''''''''''''''''''''''''''''''''''''''
Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

''''''''''''''''''''''''''''''''''''''''
' STRING in ARRAY testaaminen (FUNKTIO)
' Tarkistetaan, että jokin teksti on arrayssa.
' Palauttaa True tai False
''''''''''''''''''''''''''''''''''''''''

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = Not IsError(Application.Match(stringToBeFound, arr, 0))
End Function

''''''''''''''''''''''''''''''''''''''''
' LISTA EXCELIIN (FUNKTIO)
' "Printataan" dataa Excelin taulukkoon
''''''''''''''''''''''''''''''''''''''''

Sub PrintArray(Data As Variant, Cl As Range)
    Cl.Resize(1, UBound(Data)) = Data
End Sub

''''''''''''''''''''''''''''''''''''''''
' UUDEN DATAN LUKEMINEN EXPLORER IKKUNAN KAUTTA (tyhjä lähtötaulukko) (FUNKTIO)
' Kun halutaan lukea valmiista tiedostosta suoraan data, ei tarvita siis copy-pastea
''''''''''''''''''''''''''''''''''''''''

Function Avaa_ja_lue_tiedosto()

Dim MasterFile As Workbook
Set MasterFile = ThisWorkbook
Dim MasterFileName As String
MasterFileName = ThisWorkbook.Name

Dim MyFileName As String
Dim MyFile As Workbook



Set MyFile = Workbooks.Open(Application.GetOpenFilename())
MyFile.Activate
Sheets("Taul1").Copy Before:=Workbooks(MasterFileName).Sheets(1)
If WorksheetExists("Original Data") Then ' Poistetaan jo oleva Original data välilehti
    Application.DisplayAlerts = False
    Sheets("Original Data").Delete
    Application.DisplayAlerts = True
End If
ActiveSheet.Name = "Original Data"
Workbooks(2).Close SaveChanges:=False
'Application.DisplayAlerts = False
'Worksheets("Taul1").Delete
'On Error Resume Next
'Worksheets("Taul2").Delete
'On Error Resume Next
'Worksheets("Taul3").Delete
'On Error Resume Next
'Application.DisplayAlerts = True

End Function

''''''''''''''''''''''''''''''''''''''''
' TYÖVAIHEIDEN JAKAMINEN ERILLISIIN SARAKKEISIIN (FUNKTIO)
' Funktio, joka sijoittaa työnvaiheen oikeaan sarakkeeseen
' MIKÄLI työnnimeä ei löydy (Case kohdat), päätyy työnvaihe "MUUT VAIHEET" sarakkeeseen
''''''''''''''''''''''''''''''''''''''''

Function Valitaan_tyon_paikka(paikka As Integer, vanhapaikka As Integer, siirtaja As Integer)
    Dim src As Worksheet
    Dim trg As Worksheet ' Funktiossa nämä (Dim ja Set rivit) määritellään uudelleen -- laiskuuttani.
    Set src = ThisWorkbook.Worksheets("Original Data")
    Set trg = ThisWorkbook.Worksheets("Processed")
    Dim vertailukohde As String
    Dim raja As Integer
    raja = 0 ' Tarvitaan jokin mitta, kuinka monta riviä luetaan kerralla
    Do While siirtaja >= raja
        vertailukohde = src.Range("L" & (vanhapaikka + raja)).Value ' Luetaan originaalista datasta työnvaihe
        Select Case vertailukohde
            Case "TURVALLISTAMINEN TUOTANTO"
                src.Range(("L" & vanhapaikka + raja)).Copy trg.Range("L" & paikka) ' Sijoitetaan työ uuteen välilehteen
            Case "TURVALLISTAMINEN AUTOMAATIO"
                src.Range(("L" & vanhapaikka + raja)).Copy trg.Range("M" & paikka) ' HUOM!
            Case "TURVALLISTAMINEN PNEUMATIIKKA"
                src.Range(("L" & vanhapaikka + raja)).Copy trg.Range("N" & paikka) ' Näissä kaikissa on sarakkeen NIMI (kirjain)
            Case "TURVALLISTAMINEN SÄHKÖ"
                src.Range(("L" & vanhapaikka + raja)).Copy trg.Range("O" & paikka) ' Eli tässä ei ole "automaattisuutta"
            Case "TURVALLISTAMINEN MEKAANINEN"
                src.Range(("L" & vanhapaikka + raja)).Copy trg.Range("P" & paikka) ' Uuden työn lisääminen vaatii manuaalista
            Case "MEKAANINEN TAAKKA"
                src.Range(("L" & vanhapaikka + raja)).Copy trg.Range("Q" & paikka) ' Siirtelyä
            Case "TELINETARVE"
                src.Range(("L" & vanhapaikka + raja)).Copy trg.Range("R" & paikka)
            Case "TULITYÖLUPA"
                src.Range(("L" & vanhapaikka + raja)).Copy trg.Range("S" & paikka)
            Case "PROSESSITYÖLUPA"
                src.Range(("L" & vanhapaikka + raja)).Copy trg.Range("T" & paikka)
            Case "KORKEALLA TYÖSKENTELY"
                src.Range(("L" & vanhapaikka + raja)).Copy trg.Range("U" & paikka)
            Case "TESTAUSTARVE"
                src.Range(("L" & vanhapaikka + raja)).Copy trg.Range("Z" & paikka)
            Case Else
                If Not IsEmpty(trg.Range("AB" & paikka)) Then                       ' JOS ei löydy saraketta, tieto menee viimeiseen sarakkeeseen, pilkulla erottaen
                    trg.Range("AB" & paikka) = trg.Range("AB" & paikka).Value & ", " & src.Range(("L" & vanhapaikka + raja)).Value
                Else
                    src.Range(("L" & vanhapaikka + raja)).Copy trg.Range("AB" & paikka)
                End If
        End Select
        raja = raja + 1
    Loop
End Function

Sub CombineOrders() ' PÄÄOHJELMA

If IsEmpty(Range("A1")) Then
    Avaa_ja_lue_tiedosto
End If
    

''''''''''''''''''''''''''''''''''''''''
' ALKUMUUTTUJAT
' Määritellään joitakin muuttujia
''''''''''''''''''''''''''''''''''''''''
ActiveSheet.Name = "Original Data"
Dim MasterFile As Workbook ' Työskentelytiedosto
Set MasterFile = ThisWorkbook
Dim MasterFileName As String
MasterFileName = ThisWorkbook.Name

If WorksheetExists("Processed") Then    ' JOS meillä on jo välilehti, niin ei luoda uutta
Else
Sheets.Add(After:=Sheets("Original Data")).Name = "Processed"
End If


Dim src As Worksheet
Dim trg As Worksheet

Set src = ThisWorkbook.Worksheets("Original Data")  ' Source eli välilehti, johon data luettiin
Set trg = ThisWorkbook.Worksheets("Processed")  ' Target eli välilehti, jonne pyöritelty aineisto viedään

''''''''''''''''''''''''''''''''''''''''
' UUDEN DATAN PÄÄLLEKIRJOITUS
' JOS Prosessoidun datan välilehdessä on jo dataa (solussa A1), kysytään päällekirjoituksesta.
''''''''''''''''''''''''''''''''''''''''

If Not IsEmpty(trg.Range("A1")) Then
    Dim answer As Integer
    
    answer = MsgBox("Lopputulosten välilehdessä on jo dataa. Haluatko varmasti prosessoida Original Data välilehden aineiston?", vbQuestion + vbYesNo + vbDefaultButton2, "Prosessoitua aineistoa on jo olemassa")
    
    If answer = vbYes Then
        trg.Cells.Clear ' Tyhjennä lopputulosten välilehti
    Else
        Exit Sub
    End If
End If


''''''''''''''''''''''''''''''''''''''''
' HEADEREIDEN ASETTAMINEN LOPPUTULOSTEN VÄLILEHTEEN
' 10 tietoa (mm. sijainti, työpiste jne.) ovat tietyn työtehtävän alla aina samat
' Lisäksi eri työvaiheet (11 - 28) saavat omat sarakkeet
''''''''''''''''''''''''''''''''''''''''

Dim trgRowNames As Variant
ReDim trgRowNames(1 To 28)
trgRowNames(1) = "SIJAINTI"
trgRowNames(2) = "PRIORITEETTI"
trgRowNames(3) = "TYÖTILAUKSEN NUMERO"
trgRowNames(4) = "TOIMINTAPAIKKA"
trgRowNames(5) = "TOIMINTAPAIKAN NIMI"
trgRowNames(6) = "TYÖN NIMI"
trgRowNames(7) = "LAJITTELUKENTTÄ"
trgRowNames(8) = "METSÄN VASTUUHENKILÖ"
trgRowNames(9) = "SUUNNITTELURYHMÄ"
trgRowNames(10) = "TYÖPISTE"
trgRowNames(11) = "TURVALLISTAMISLISTAN NUMERO"
trgRowNames(12) = "TURVALLISTAMINEN TUOTANTO"
trgRowNames(13) = "TURVALLISTAMINEN AUTOMAATIO"
trgRowNames(14) = "TURVALLISTAMINEN PNEUMATIKKA"
trgRowNames(15) = "TURVALLISTAMINEN SÄHKÖ"
trgRowNames(16) = "TURVALLISTAMINEN MEKAANINEN"
trgRowNames(17) = "MEKAANINEN TAAKKA"
trgRowNames(18) = "TELINETARVE"
trgRowNames(19) = "TULITYÖLUPA"
trgRowNames(20) = "PROSESSITYÖLUPA"
trgRowNames(21) = "KORKEALLA TYÖSKENTELY"
trgRowNames(22) = "TURVALLISTETTU"
trgRowNames(23) = "TYÖ ALOITETTU"
trgRowNames(24) = "TYÖ PÄÄTETTY"
trgRowNames(25) = "TURVALLISTAMINEN PURETTU"
trgRowNames(26) = "TESTAUSTARVE"
trgRowNames(27) = "TESTAUS VALMIS"
trgRowNames(28) = "MUUT VAIHEET"

PrintArray trgRowNames, trg.[A1] ' Kutsutaan funktiota -- Transpoosia ja asettamista varten pääosin

''''''''''''''''''''''''''''''''''''''''
' TILAUSTEN RYHMITTELYN VALMISTELU (MUUTTUVAT SARAKKEET)
' Määritellään joitakin muuttujia valmiiksi
''''''''''''''''''''''''''''''''''''''''
Dim Lastrow As Integer
Lastrow = src.Cells(Rows.Count, 1).End(xlUp).Row ' Rivien määrä alkuperäisessä taulukossa
Dim i As Integer
Dim j As Integer
Dim z As Integer
Dim n As Integer
Dim toisto As Integer

''''''''''''''''''''''''''''''''''''''''
' TILAUSTEN RYHMITTELY JA KOPIOINTI
' Selvitetään jokaisen tilauksen kaikki rivit
' Siirretään tilauksen yhteiset tiedot suoraan toiseen välilehteen
' Siirretään jokaiselle työlle ominaiset vaiheet omiin sarakkeisiin
''''''''''''''''''''''''''''''''''''''''

j = 0 ' Alustetaan luku, joka kertyy, kun sama työluku toistuu
toisto = 0 ' KUN tiettyä työlukua ei enään tule, päätetään looppi
For i = 2 To Lastrow ' Hypätään Header rivi yli
    If j <> 0 Then  ' Käsitellään rivit, joissa toistuu sama työluku
        j = j - 1 ' Hypitään yli ne rivit, missä on sama työkoodi ollut
        GoTo NextIteration ' Otetaan seuraava rivi
    End If
    Do While toisto < 1
        If src.Cells(i, 3) <> src.Cells(i + j + 1, 3) Then  ' Astutaan rivi eteenpäin, ja katsotaan, onko se sama kuin edeltävä
            toisto = 1  ' Luvut eroavat, looppi päättyy, eli yhden työluvun kaikki vaiheet on löydetty
        Else
            j = j + 1   ' ON SAMA, astutaan seuraavalle riville (eli uudelleen looppiin)
        End If
    Loop
    toisto = 0
    
    
    
    
    src.Range(("A" & i), ("C" & i)).Copy trg.Range("A" & Rows.Count).End(xlUp).Offset(1, 0) ' Siirretään "vakiorivit". HUOM! Hypätään D (Vaihe) sarake yli
    src.Range(("E" & i), ("K" & i)).Copy trg.Range("D" & Rows.Count).End(xlUp).Offset(1, 0)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LOGIIKKA TIETTYYN TYÖKOODIIN KUULUVIEN TYÖVAIHEIDEN SIIRTELYYN LÄHTEE TÄSTÄ
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    n = 0
    z = 2 ' Lähdetään rivistä 2; ensimmäinen rivihän on Header
    Do While n = 0
        If src.Cells(i, 3) <> trg.Cells(z, 3) Then ' Etsitään lähtö-välilehdestä työnkoodi, ja sen vastaava rivi lopputulosten välilehdestä
            z = z + 1   ' Tämä siksi, että saadaan oikealle riville työnvaihe lopputulosten taulukossa
        Else
            n = 1
        End If
    Loop
    n = 0
    Valitaan_tyon_paikka z, i, j ' z on rivi, i viittaa lähtötaulukon riviin, j kertoo rivien määrästä, mitkä kuuluvat samaan työkoodiin
NextIteration:
Next i

'''''''''''''''''''''''''''''''''''''''''''''''
' VÄRJÄÄMINEN JA BOLDAUS
' Ylimmän header rivin ulkoasun muokkaaminen
'''''''''''''''''''''''''''''''''''''''''''''''

LastCol = Split(trg.Cells(1, Columns.Count).End(xlToLeft).Address, "$")(1) ' Selvitetään viimeinen sarakkeen kirjain

trg.Range(("A1"), (LastCol & 1)).Interior.ColorIndex = 4 ' Vihreä pohja
trg.Range(("K1")).Interior.ColorIndex = 27 ' Keltainen pohja
trg.Range(("V1"), ("Y1")).Interior.ColorIndex = 27 ' Keltainen pohja
trg.Range(("AA1")).Interior.ColorIndex = 27 ' Keltainen pohja

trg.Range(("A1"), (LastCol & 1)).Columns.AutoFit ' Venytetään solut
trg.Range(("A1"), (LastCol & 1)).Font.Bold = True ' Boldataan teksti



'''''''''''''''''''''''''''''''''''''''''
' VANHAA KOODIA; EI KÄYTÖSSÄ!
' LIITTYI TEHTÄVIEN VAIHEIDEN LUOKITTELUUN
' AUTOMATISOITU VERSIO -- VIRALLISESSA VERSIOSSA KIINTEÄT VAIHEET
'''''''''''''''''''''''''''''''''''''''''

'ReDim tyo_kuvaukset(1 To Lastrow) As Variant    ' Erilaisten työvaiheiden taulukko
'Dim cell As Range

'i = LBound(tyo_kuvaukset)
'Dim x As String
'x = src.Cells(2, 12)
'MsgBox x
'For Each cell In src.Range("L2:L" & Lastrow)   ' Lähdetään käymään läpi erilaisia työnvaiheita
'If IsInArray(cell.Value, tyo_kuvaukset) Then    ' JOS työnimike on jo listassa, hyppää yli
'    GoTo NextOne
'Else
'    tyo_kuvaukset(i) = cell.Value   ' JOS ei ole jo listalla, lisää se
'End If
'i = i + 1
'On Error Resume Next
'NextOne:
'Next cell
' HOX!!! JOS POHJATAULUKKO MUUTTUU, TUOSSA ALLA OLEVA K1 pitää muuttaa
'PrintArray tyo_kuvaukset, trg.[K1]  ' Lopuksi transpoosi, ja laitetaan "11" (eli K1) sarakkeeseen
'MsgBox "Täällä"


End Sub
