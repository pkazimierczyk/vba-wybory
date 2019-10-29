Attribute VB_Name = "Module3"
Option Explicit

Sub Przycisk_start()
okno_startowe.Show
ActiveWorkbook.Sheets.Add(after:=Sheets(Sheets.Count)).Name = "Dane wej�ciowe"
Sheets("Dane wej�ciowe").Cells(1, 1) = "Liczba mandat�w do zdobycia"
Sheets("Dane wej�ciowe").Cells(1, 2) = liczba_mandatow
Sheets("Dane wej�ciowe").Cells(2, 1) = "Pr�g wyborczy (%)"
Sheets("Dane wej�ciowe").Cells(2, 2) = prog
If prog_KKW <> prog Then
    Sheets("Dane wej�ciowe").Cells(3, 1) = "Pr�g wyborczy dla KKW (%)"
    Sheets("Dane wej�ciowe").Cells(3, 2) = prog_KKW
    If prog_mniejszosc <> prog Then
        Sheets("Dane wej�ciowe").Cells(4, 1) = "Pr�g wyborczy dla komitet�w mniejszo�ci (%)"
        Sheets("Dane wej�ciowe").Cells(4, 2) = prog_mniejszosc
    End If
ElseIf prog_mniejszosc <> prog Then
        Sheets("Dane wej�ciowe").Cells(3, 1) = "Pr�g wyborczy dla komitet�w mniejszo�ci (%)"
        Sheets("Dane wej�ciowe").Cells(3, 2) = prog_mniejszosc
End If

Sheets("Dane wej�ciowe").Cells(1, 4) = "nr okr�gu"
Sheets("Dane wej�ciowe").Cells(1, 5) = "liczba uprawnionych do g�osowania"
Sheets("Dane wej�ciowe").Cells(1, 5).WrapText = True

Sheets("Dane wej�ciowe").Cells(1, 8) = "nr listy"
Dim i As Integer
For i = 1 To liczba_okregow
    Sheets("Dane wej�ciowe").Cells(1 + i, 4) = i
    Next i
For i = 1 To liczba_list
    Sheets("Dane wej�ciowe").Cells(1 + i, 8) = i
    Next i
    
'wyszarzenie list
Sheets("Dane wej�ciowe").Columns(8).Font.Color = RGB(225, 225, 225)
    
'formatowanie
Sheets("Dane wej�ciowe").Columns.AutoFit
Sheets("Dane wej�ciowe").Columns(2).Font.Bold = True
Sheets("Dane wej�ciowe").Columns(5).ColumnWidth = 19
Sheets("Dane wej�ciowe").Columns(6).ColumnWidth = 18
Sheets("Dane wej�ciowe").Rows.AutoFit
uprawnieni_podswietlenie
Sheets("Dane wej�ciowe").Activate
MsgBox "Uzupe�nij liczby uprawnionych do g�osowania do wyliczenia liczby mandat�w w okr�gach"

'zablokowanie arkusza
Range(Sheets("Dane wej�ciowe").Cells(2, 5), Sheets("Dane wej�ciowe").Cells(1 + liczba_okregow, 5)).Locked = False
Sheets("Dane wej�ciowe").Protect

End Sub
