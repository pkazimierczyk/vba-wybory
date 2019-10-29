Attribute VB_Name = "Module3"
Option Explicit

Sub Przycisk_start()
okno_startowe.Show
ActiveWorkbook.Sheets.Add(after:=Sheets(Sheets.Count)).Name = "Dane wejœciowe"
Sheets("Dane wejœciowe").Cells(1, 1) = "Liczba mandatów do zdobycia"
Sheets("Dane wejœciowe").Cells(1, 2) = liczba_mandatow
Sheets("Dane wejœciowe").Cells(2, 1) = "Próg wyborczy (%)"
Sheets("Dane wejœciowe").Cells(2, 2) = prog
If prog_KKW <> prog Then
    Sheets("Dane wejœciowe").Cells(3, 1) = "Próg wyborczy dla KKW (%)"
    Sheets("Dane wejœciowe").Cells(3, 2) = prog_KKW
    If prog_mniejszosc <> prog Then
        Sheets("Dane wejœciowe").Cells(4, 1) = "Próg wyborczy dla komitetów mniejszoœci (%)"
        Sheets("Dane wejœciowe").Cells(4, 2) = prog_mniejszosc
    End If
ElseIf prog_mniejszosc <> prog Then
        Sheets("Dane wejœciowe").Cells(3, 1) = "Próg wyborczy dla komitetów mniejszoœci (%)"
        Sheets("Dane wejœciowe").Cells(3, 2) = prog_mniejszosc
End If

Sheets("Dane wejœciowe").Cells(1, 4) = "nr okrêgu"
Sheets("Dane wejœciowe").Cells(1, 5) = "liczba uprawnionych do g³osowania"
Sheets("Dane wejœciowe").Cells(1, 5).WrapText = True

Sheets("Dane wejœciowe").Cells(1, 8) = "nr listy"
Dim i As Integer
For i = 1 To liczba_okregow
    Sheets("Dane wejœciowe").Cells(1 + i, 4) = i
    Next i
For i = 1 To liczba_list
    Sheets("Dane wejœciowe").Cells(1 + i, 8) = i
    Next i
    
'wyszarzenie list
Sheets("Dane wejœciowe").Columns(8).Font.Color = RGB(225, 225, 225)
    
'formatowanie
Sheets("Dane wejœciowe").Columns.AutoFit
Sheets("Dane wejœciowe").Columns(2).Font.Bold = True
Sheets("Dane wejœciowe").Columns(5).ColumnWidth = 19
Sheets("Dane wejœciowe").Columns(6).ColumnWidth = 18
Sheets("Dane wejœciowe").Rows.AutoFit
uprawnieni_podswietlenie
Sheets("Dane wejœciowe").Activate
MsgBox "Uzupe³nij liczby uprawnionych do g³osowania do wyliczenia liczby mandatów w okrêgach"

'zablokowanie arkusza
Range(Sheets("Dane wejœciowe").Cells(2, 5), Sheets("Dane wejœciowe").Cells(1 + liczba_okregow, 5)).Locked = False
Sheets("Dane wejœciowe").Protect

End Sub
