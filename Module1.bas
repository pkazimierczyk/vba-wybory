Attribute VB_Name = "Module1"
Option Explicit
Public liczba_mandatow As Integer
Public liczba_okregow As Integer
Public liczba_list As Integer
Public prog As Integer
Public prog_KKW As Integer
Public prog_mniejszosc As Integer



Sub uprawnieni_podswietlenie()

Dim warunek As FormatCondition
Set warunek = Range(Sheets("Dane wej�ciowe").Cells(2, 5), Sheets("Dane wej�ciowe").Cells(1 + liczba_okregow, 5)) _
.FormatConditions.Add(xlCellValue, xlEqual, 0)

With warunek.Borders
.LineStyle = xlContinuous
.Color = RGB(250, 150, 150)
End With

'tworzenie przycisku
Dim przycisk_okreslenie_liczby_mandatow As Button
Set przycisk_okreslenie_liczby_mandatow = Sheets("Dane wej�ciowe").Buttons.Add( _
Width:=120, Height:=40, Left:=Cells(3 + liczba_okregow, 4).Left, Top:=Cells(3 + liczba_okregow, 4).Top)
With przycisk_okreslenie_liczby_mandatow
    .Text = "Ustal liczb� mandat�w dla ka�dego okr�gu"
    .OnAction = "okreslenie_liczby_mandatow_na_okreg"
End With

End Sub



Sub okreslenie_liczby_mandatow_na_okreg()

Sheets("Dane wej�ciowe").Unprotect
liczba_okregow = Sheets("Dane wej�ciowe").Application.WorksheetFunction.Max(Columns(4))
liczba_list = Sheets("Dane wej�ciowe").Application.WorksheetFunction.Max(Columns(8))
liczba_mandatow = Sheets("Dane wej�ciowe").Cells(1, 2)

' Sprawdzenie czy s� dane
If Application.WorksheetFunction.CountBlank(Sheets("Dane wej�ciowe").Range(Cells(2, 5), Cells(1 + liczba_okregow, 5))) > 0 Then
    MsgBox ("Uzupe�nij dane dot. liczby uprawnionych do g�osowania w okr�gach")
    Exit Sub
End If

' Sprawdzenie poprawno�ci danych
Dim j As Integer
For j = 1 To liczba_okregow
    If Not IsNumeric(Sheets("Dane wej�ciowe").Cells(1 + j, 5)) Then
        MsgBox ("W kolumnie 'liczba uprawnionych do g�osowania' musz� by� wy��cznie liczby ca�kowite dodatnie!")
        Exit Sub
    ElseIf (Sheets("Dane wej�ciowe").Cells(1 + j, 5)) < 1 _
    Or (Sheets("Dane wej�ciowe").Cells(1 + j, 5)) - Int(Sheets("Dane wej�ciowe").Cells(1 + j, 5)) <> 0 Then
        MsgBox ("W kolumnie 'liczba uprawnionych do g�osowania' musz� by� wy��cznie liczby ca�kowite dodatnie!")
        Exit Sub
    End If
Next j

   
' Wyznaczenie jednolitej normy przedstawicielskiej
Dim JNP As Single
JNP = Application.WorksheetFunction.Sum(Sheets("Dane wej�ciowe").Range("e:e")) / liczba_mandatow

' Ustalenie liczby mandat�w - krok 1
Dim i1 As Integer
Dim teoretycznaLM() As Single
ReDim teoretycznaLM(1 To liczba_okregow)
For i1 = 1 To liczba_okregow
    teoretycznaLM(i1) = Sheets("Dane wej�ciowe").Cells(1 + i1, 5) / JNP
    Next i1

' Ustalanie liczby mandat�w - krok 2
Dim i2 As Integer
Dim teoretycznaLM2() As Single
ReDim teoretycznaLM2(1 To liczba_okregow)
For i2 = 1 To liczba_okregow
    teoretycznaLM2(i2) = Application.WorksheetFunction.Round(teoretycznaLM(i2), 0)
'    ActiveCell.Offset(1, i2) = teoretycznaLM2(i2)
    Next i2

' Ustalanie liczby mandat�w - krok 3
Dim maks_onm As Single
Dim nr_listy_maks_onm As Integer
Dim min_onm As Single
Dim nr_listy_min_onm As Integer
Dim i3 As Integer
Dim i4 As Integer
Dim i5 As Integer
Dim i6 As Integer
' je�li si� uda�o
If liczba_mandatow = Application.WorksheetFunction.Sum(teoretycznaLM2) Then

Else
    Dim obywatele_na_mandat() As Single
    ReDim obywatele_na_mandat(1 To liczba_okregow)
    For i5 = 1 To liczba_okregow
        obywatele_na_mandat(i5) = Sheets("Dane wej�ciowe").Cells(1 + i5, 5) / teoretycznaLM2(i5)
        Next i5
    ' je�li rozdano za ma�o mandat�w
    If liczba_mandatow > Application.WorksheetFunction.Sum(teoretycznaLM2) Then
        Do While liczba_mandatow > Application.WorksheetFunction.Sum(teoretycznaLM2)
            maks_onm = obywatele_na_mandat(1)
            nr_listy_maks_onm = 1
            For i6 = 2 To liczba_okregow
                If maks_onm < obywatele_na_mandat(i6) Then
                    maks_onm = obywatele_na_mandat(i6)
                    nr_listy_maks_onm = i6
                End If
                Next i6
            teoretycznaLM2(nr_listy_maks_onm) = teoretycznaLM2(nr_listy_maks_onm) + 1
        Loop
    ' je�li rozdano za du�o mandat�w
    Else
        Do While liczba_mandatow < Application.WorksheetFunction.Sum(teoretycznaLM2)
            min_onm = obywatele_na_mandat(1)
            nr_listy_min_onm = 1
            For i6 = 2 To liczba_okregow
                If min_onm > obywatele_na_mandat(i6) Then
                    min_onm = obywatele_na_mandat(i6)
                    nr_listy_min_onm = i6
                End If
                Next i6
            teoretycznaLM2(nr_listy_min_onm) = teoretycznaLM2(nr_listy_min_onm) - 1
        Loop
    End If
        
End If

'wype�nienie tabeli liczb� mandat�w
For i3 = 1 To liczba_okregow
    Sheets("Dane wej�ciowe").Cells(1 + i3, 6) = teoretycznaLM2(i3)
    Next i3

'wypisanie okr�g�w w arkuszu dane
Sheets("Dane wej�ciowe").Range(Cells(1, 11), Cells(1, 11).End(xlToRight).End(xlDown)).Clear
Dim i7 As Integer

For i7 = 1 To liczba_okregow
    Cells(1, 10 + i7) = "Okr�g nr " & i7
Next i7
Range(Cells(1, 9), Cells(1, 10 + i7)).Orientation = xlUpward


'stworzenie list rozwijanych mo�liwych liczb kandydat�w
Dim i8 As Integer
Dim i9 As Integer
Dim mandaty As Integer
Dim mozliwe As Variant
For i8 = 1 To liczba_okregow
    mandaty = Cells(i8 + 1, 6)
    mozliwe = Array(0, mandaty)
    ReDim Preserve mozliwe(UBound(mozliwe) + mandaty + 1)
    For i9 = 1 To mandaty
        mozliwe(i9 + 2) = mandaty + i9
    Next i9
    
    With Range(Cells(2, 10 + i8), Cells(1 + liczba_list, 10 + i8)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(mozliwe, ",")
    End With
    
    With Range(Cells(2, 9), Cells(1 + liczba_list, 10)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(Array("�", "tak"), ",")
    End With
Next i8

'formatowanie
Sheets("Dane wej�ciowe").Cells(1, 6) = "liczba mandat�w w okr�gu"
Sheets("Dane wej�ciowe").Cells(1, 6).WrapText = True
Sheets("Dane wej�ciowe").Columns(6).ColumnWidth = 16
Sheets("Dane wej�ciowe").Columns(8).Font.Color = RGB(0, 0, 0)
Sheets("Dane wej�ciowe").Cells(1, 9) = "KKW"
Sheets("Dane wej�ciowe").Cells(1, 10) = "mniejszo�� nar."
Sheets("Dane wej�ciowe").Columns.AutoFit
Sheets("Dane wej�ciowe").Rows.AutoFit
Sheets("Dane wej�ciowe").Range(Cells(2, 6), Cells(1 + liczba_okregow, 6)).Font.Bold = True

'tworzenie przycisku
Dim tworzenie_arkuszy_okregow As Button
Set tworzenie_arkuszy_okregow = Sheets("Dane wej�ciowe").Buttons.Add _
(Width:=120, Height:=40, Left:=Cells(3 + liczba_list, 8).Left, Top:=Cells(3 + liczba_list, 8).Top)
With tworzenie_arkuszy_okregow
    .Text = "Utw�rz arkusze okr�g�w wyborczych"
    .OnAction = "tworzenie_arkuszy_okregow"
End With

'dezaktywacja kolumn KKW i mniejszosci, je�li nie maj� osobnych prog�w
Dim czy_KKW As Boolean
Dim czy_mniejszosc As Boolean
If "Pr�g wyborczy dla KKW (%)" = Sheets("Dane Wej�ciowe").Cells(3, 1).Value Then
    czy_KKW = True
End If
If "Pr�g wyborczy dla komitet�w mniejszo�ci (%)" = Sheets("Dane Wej�ciowe").Cells(3, 1).Value Or "Pr�g wyborczy dla komitet�w mniejszo�ci (%)" = Sheets("Dane Wej�ciowe").Cells(4, 1).Value Then
    czy_mniejszosc = True
End If
If czy_KKW = False Then
    With Range(Sheets("Dane wej�ciowe").Cells(2, 9), Sheets("Dane wej�ciowe").Cells(1 + liczba_list, 9))
        .Locked = True
        .ColumnWidth = 0
    End With
Else
    Range(Sheets("Dane wej�ciowe").Cells(2, 9), Sheets("Dane wej�ciowe").Cells(1 + liczba_list, 9)).Locked = False
End If
If czy_mniejszosc = False Then
    With Range(Sheets("Dane wej�ciowe").Cells(2, 10), Sheets("Dane wej�ciowe").Cells(1 + liczba_list, 10))
        .Locked = True
        .ColumnWidth = 0
    End With
Else
    Range(Sheets("Dane wej�ciowe").Cells(2, 10), Sheets("Dane wej�ciowe").Cells(1 + liczba_list, 10)).Locked = False
End If

Range(Sheets("Dane wej�ciowe").Cells(2, 11), Sheets("Dane wej�ciowe").Cells(1 + liczba_list, 8 + liczba_okregow + 2)).Locked = False
Sheets("Dane wej�ciowe").Protect

'komunikat
Dim c1 As String, c2 As String, c3 As String
If czy_KKW = True Then
    c1 = "koalicyjne komitety wyborcze"
        If czy_mniejszosc = True Then
            c2 = ", komitety mniejszo�ci narodowych oraz "
        Else
            c2 = " oraz "
        End If
ElseIf czy_mniejszosc = True Then
    c1 = "komitety mniejszo�ci narodowych oraz "
End If

MsgBox ("Wska� " & c1 & c2 & c3 & "liczby kandydat�w na poszczeg�lnych listach w okr�gach")

End Sub

Sub tworzenie_arkuszy_okregow()

liczba_okregow = Sheets("Dane wej�ciowe").Application.WorksheetFunction.Max(Columns(4))
liczba_list = Sheets("Dane wej�ciowe").Application.WorksheetFunction.Max(Columns(8))
liczba_mandatow = Sheets("Dane wej�ciowe").Cells(1, 2)

' Sprawdzenie, czy s� dane i s� ok
Dim l As Integer
For l = 1 To liczba_list
    If Sheets("Dane wej�ciowe").Cells(l + 1, 9) = Sheets("Dane wej�ciowe").Cells(l + 1, 10) And Sheets("Dane wej�ciowe").Cells(l + 1, 9) = "tak" Then
        MsgBox ("�aden komitet nie mo�e by� jednocze�nie KKW i komitetem mniejszo�ci")
        Exit Sub
    End If
Next l
If Application.WorksheetFunction.CountBlank(Sheets("Dane wej�ciowe").Range(Cells(2, 11), Cells(1 + liczba_list, 10 + liczba_okregow))) > 0 Then
    MsgBox ("Uzupe�nij dane o liczbach kandydat�w na listach")
    Exit Sub
End If


'pytanie pomocnicze dla u�ytkownika
Dim pytanie As Integer
pytanie = MsgBox("Na pewno oznaczy�e� wszystkie komitety koalicyjne i mniejszo�ci narodowych?", vbYesNo)
If pytanie = 7 Then
    Exit Sub
End If

'tworzenie arkuszy
Dim i As Integer
Dim i1 As Integer
Dim i2 As Integer
Dim i21 As Integer
Dim tabela_list As Range
Set tabela_list = Range(Sheets("Dane wej�ciowe").Cells(2, 11), Sheets("Dane wej�ciowe").Cells(2, 11).End(xlToRight).End(xlDown))
For i = 1 To liczba_okregow
    ActiveWorkbook.Sheets.Add(after:=Sheets(Sheets.Count)).Name = "Okr�g nr " & i
    Sheets("Okr�g nr " & i).Cells.Interior.ColorIndex = 16
    Sheets("Okr�g nr " & i).Cells.HorizontalAlignment = xlCenter
    Sheets("Okr�g nr " & i).Cells(2, 1) = "Miejsce na li�cie"
    Sheets("Okr�g nr " & i).Cells(1, 2) = "Numer listy"
    Sheets("Okr�g nr " & i).Cells(1, 1) = i
    Range(Sheets("Okr�g nr " & i).Cells(1, 2), Sheets("Okr�g nr " & i).Cells(1, 1 + liczba_list)).Merge
    
    Cells(Sheets("Dane wej�ciowe").Cells(1 + i, 6) * 2 + 3, 1) = "Suma g�os�w"
    Cells(Sheets("Dane wej�ciowe").Cells(1 + i, 6) * 2 + 4, 1) = "Procent g�os�w"
    Cells(Sheets("Dane wej�ciowe").Cells(1 + i, 6) * 2 + 5, 1) = "Frekwencja"
    
    Sheets("Okr�g nr " & i).Cells(1, 2 + liczba_list) = "G�osy niewa�ne"

    With Sheets("Okr�g nr " & i).Cells(1, 3 + liczba_list)
        .Locked = False
        .Borders.LineStyle = xlContinuous
        .Interior.ColorIndex = 0
    End With
    
    'wpisanie miejsc na listach
    For i1 = 1 To Sheets("Dane wej�ciowe").Cells(1 + i, 6) * 2
        Sheets("Okr�g nr " & i).Cells(2 + i1, 1) = i1
    Next i1

    ActiveSheet.Range(Cells(2, 2), Cells(2, 1 + liczba_list)).Locked = False
    
    'wpisanie numer�w list i wybielenie kom�rek edytowalnych
    For i2 = 1 To liczba_list
        Sheets("Okr�g nr " & i).Cells(2, 1 + i2) = i2
        For i21 = 1 To tabela_list(i2, i)
            Sheets("Okr�g nr " & i).Cells(2 + i21, 1 + i2).Locked = False
            Sheets("Okr�g nr " & i).Cells(2 + i21, 1 + i2).Interior.ColorIndex = 0
            Sheets("Okr�g nr " & i).Cells(2 + i21, 1 + i2).Borders.LineStyle = xlContinuous
        Next i21
    Next i2
    
    'formatowanie
    With Union(Cells(1, 1), Cells(1, 2), Cells(2, 1), Range(Cells(3, 1), Cells(3, 1).End(xlDown)), _
    Range(Cells(2, 2), Cells(2, 2).End(xlToRight)), Cells(1, 2 + liczba_list), _
    Cells(Sheets("Dane wej�ciowe").Cells(1 + i, 6) * 2 + 5, 1), Cells(Sheets("Dane wej�ciowe").Cells(1 + i, 6) * 2 + 5, 2), _
    Range(Cells(Sheets("Dane wej�ciowe").Cells(1 + i, 6) * 2 + 3, 2), Cells(Sheets("Dane wej�ciowe").Cells(1 + i, 6) * 2 + 4, 1 + liczba_list)))
        .Interior.ColorIndex = 0
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With
    Cells(1, 1).Interior.ColorIndex = 16
    
    ActiveSheet.Columns(1).AutoFit

    With Range(Cells(Sheets("Dane wej�ciowe").Cells(1 + i, 6) * 2 + 3, 1), Cells(Sheets("Dane wej�ciowe").Cells(1 + i, 6) * 2 + 3, 1 + liczba_list)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With


    Dim nr_listy As Integer
    For nr_listy = 1 To liczba_list
        Columns(1 + nr_listy).ColumnWidth = 8
    Next nr_listy
    
    ActiveSheet.Columns(liczba_list + 2).AutoFit
    Columns(liczba_list + 3).ColumnWidth = 8
    
    'tworzenie przycisku "Podsumuj okr�g"
    Dim podsumuj_okreg As Button
    Set podsumuj_okreg = ActiveSheet.Buttons.Add _
    (Width:=120, Height:=40, Left:=Cells(Application.WorksheetFunction.Max(ActiveSheet.Columns(1)) + 7, 1).Left, _
    Top:=Cells(Application.WorksheetFunction.Max(ActiveSheet.Columns(1)) + 7, 1).Top)
    With podsumuj_okreg
        .Text = "Podsumuj okr�g"
        .OnAction = "suma_w_okregu"
    End With
    
    'blokowanie
    ActiveSheet.Range(Cells(2, 2), Cells(2, 1 + liczba_list)).Locked = True
    ActiveSheet.Protect
    
    
Next i

tworzenia_arkusza_wyniki_zbiorcze

Sheets("Dane wej�ciowe").Protect
Sheets("Okr�g nr 1").Activate
End Sub

Sub suma_w_okregu()
Dim nr_okregu As Integer
nr_okregu = Mid(ActiveSheet.Name, 10)

liczba_okregow = Application.WorksheetFunction.Max(Sheets("Dane wej�ciowe").Columns(4))
liczba_list = Application.WorksheetFunction.Max(Sheets("Dane wej�ciowe").Columns(8))
liczba_mandatow = Sheets("Dane wej�ciowe").Cells(1, 2)

ActiveSheet.Unprotect
Dim maks_miejsc As Integer
Dim i As Integer
maks_miejsc = Application.WorksheetFunction.Max(ActiveSheet.Columns(1))

'sprawdzenie czy s� dane i czy s� ok
Dim j As Integer
Dim tabela_list As Range
Set tabela_list = Range(Sheets("Dane wej�ciowe").Cells(2, 10 + ActiveSheet.Cells(1, 1)), Sheets("Dane wej�ciowe").Cells(2, 10 + ActiveSheet.Cells(1, 1)).End(xlDown))
For i = 1 To liczba_list
    For j = 1 To tabela_list(i)
        If IsEmpty(Cells(j + 2, i + 1)) Then
            MsgBox ("Uzupe�nij dane")
            ActiveSheet.Range(Cells(2, 2), Cells(2, 1 + liczba_list)).Locked = True
            Range(Cells(maks_miejsc + 3, 2), Cells(maks_miejsc + 5, liczba_list + 1)).Value = ""
            ActiveSheet.Protect
            Exit Sub
        End If
        Call czy_liczba(Cells(j + 2, i + 1), "Wszystkie dane musz� by� liczbami nieujemnymi")
        If Cells(j + 2, i + 1) = "" Then
            Range(Cells(maks_miejsc + 3, 2), Cells(maks_miejsc + 5, liczba_list + 1)).Value = ""
            ActiveSheet.Protect
            Exit Sub
        End If
    Next j
Next i

If IsEmpty(Cells(1, liczba_list + 3)) Then
    MsgBox ("Uzupe�nij dane. Pami�taj o g�osach niewa�nych.")
    ActiveSheet.Range(Cells(2, 2), Cells(2, 1 + liczba_list)).Locked = True
    Range(Cells(maks_miejsc + 3, 2), Cells(maks_miejsc + 5, liczba_list + 1)).Value = ""
    ActiveSheet.Protect
    Exit Sub
End If
Call czy_liczba(Cells(1, liczba_list + 3), "Wszystkie dane musz� by� liczbami nieujemnymi")
If Cells(1, liczba_list + 3) = "" Then
    Range(Cells(maks_miejsc + 3, 2), Cells(maks_miejsc + 5, liczba_list + 1)).Value = ""
    ActiveSheet.Protect
    Exit Sub
End If
                
'sprawdzenie czy frekwencja <= 100%
If Application.WorksheetFunction.Sum(Range(Cells(3, 2), Cells(2 + maks_miejsc, 1 + liczba_list))) > Sheets("Dane wej�ciowe").Cells(1 + nr_okregu, 5) Then
    MsgBox ("Frekwencja ponad 100%. Popraw wyniki.")
    Exit Sub
End If

'suma g�os�w i procent
For i = 1 To liczba_list
    Cells(maks_miejsc + 3, 1 + i) = Application.WorksheetFunction.Sum(Range(Cells(3, 1 + i), Cells(2 + maks_miejsc, 1 + i)))
    Cells(maks_miejsc + 4, 1 + i) = Cells(maks_miejsc + 3, 1 + i) / _
    Application.WorksheetFunction.Sum(Range(Cells(3, 2), Cells(2 + maks_miejsc, 1 + liczba_list)))
Next i

'frekwencja
Cells(maks_miejsc + 5, 2) = Application.WorksheetFunction.Sum(Cells(1, liczba_list + 3), Range(Cells(maks_miejsc + 3, 2), Cells(maks_miejsc + 3, 1 + liczba_list))) _
/ Sheets("Dane wej�ciowe").Cells(1 + nr_okregu, 5)
Range(Cells(maks_miejsc + 4, 2), Cells(maks_miejsc + 5, 1 + i)).NumberFormat = "0.00%"
ActiveSheet.Protect
End Sub

Sub suma_wszystkie_okregi()

liczba_okregow = Application.WorksheetFunction.Max(Sheets("Dane wej�ciowe").Columns(4))
liczba_list = Application.WorksheetFunction.Max(Sheets("Dane wej�ciowe").Columns(8))
liczba_mandatow = Sheets("Dane wej�ciowe").Cells(1, 2)

Dim i As Integer
For i = 1 To liczba_okregow
    Sheets("Okr�g nr " & i).Activate
    suma_w_okregu
Next i
Sheets("Wyniki zbiorcze").Activate
End Sub

Sub tworzenia_arkusza_wyniki_zbiorcze()

ActiveWorkbook.Sheets.Add(after:=Sheets(Sheets.Count)).Name = "Wyniki zbiorcze"
Sheets("Wyniki zbiorcze").Cells.HorizontalAlignment = xlCenter
Sheets("Wyniki zbiorcze").Cells.Interior.ColorIndex = 16
liczba_okregow = Application.WorksheetFunction.Max(Sheets("Dane wej�ciowe").Columns(4))
liczba_list = Application.WorksheetFunction.Max(Sheets("Dane wej�ciowe").Columns(8))
liczba_mandatow = Sheets("Dane wej�ciowe").Cells(1, 2)

Dim i1 As Integer
Dim i2 As Integer

Sheets("Wyniki zbiorcze").Cells(2, 1) = "Numer okr�gu"
Sheets("Wyniki zbiorcze").Cells(1, 2) = "Numer listy"
Range(Sheets("Wyniki zbiorcze").Cells(1, 2), Sheets("Wyniki zbiorcze").Cells(1, 1 + liczba_list)).Merge
Sheets("Wyniki zbiorcze").Cells(1, 2 + liczba_list) = "G�osy niewa�ne"

'wpisanie numer�w okr�g�w
For i1 = 1 To liczba_okregow
    Sheets("Wyniki zbiorcze").Cells(2 + i1, 1) = i1
Next i1

'wpisanie numer�w list
For i2 = 1 To liczba_list
    Sheets("Wyniki zbiorcze").Cells(2, 1 + i2) = i2
Next i2

Cells(liczba_okregow + 3, 1) = "Suma g�os�w"
Cells(liczba_okregow + 4, 1) = "Procent g�os�w"
Cells(liczba_okregow + 5, 1) = "Frekwencja"


'tworzenie przycisku "Przelicz wyniki w okr�gach i pobierz dane"
Dim przelicz_pobierz As Button
Set przelicz_pobierz = ActiveSheet.Buttons.Add _
(Width:=120, Height:=40, Left:=Cells(3, liczba_list + 3).Left, _
Top:=Cells(3, liczba_list + 3).Top)
With przelicz_pobierz
    .Text = "Przelicz wyniki w okr�gach i pobierz dane"
    .OnAction = "suma_wszystkie_okregi"
    .OnAction = "pobieranie_wynikow"
End With

'formatowanie
With Union(Range(Cells(1, 2), Cells(1, 2).End(xlToRight).Offset(0, 1)), _
Range(Cells(2, 1), Cells(liczba_okregow + 4, liczba_list + 1)), Range(Cells(liczba_okregow + 5, 1), Cells(liczba_okregow + 5, 2)))
    .Interior.ColorIndex = 0
    .Borders.LineStyle = xlContinuous
End With
With Union(Range(Cells(1, 2), Cells(1, 2).End(xlToRight).Offset(0, 1)), Range(Cells(2, 1), Cells(liczba_okregow + 5, 1)), _
Range(Cells(2, 2), Cells(2, 1 + liczba_list)), Range(Cells(liczba_okregow + 3, 2), Cells(liczba_okregow + 5, liczba_list + 1)))
    .Font.Bold = True
End With
With Range(Cells(liczba_okregow + 3, 1), Cells(liczba_okregow + 3, liczba_list + 1)).Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlMedium
End With

ActiveSheet.Columns(1).ColumnWidth = 17
Dim nr_listy As Integer
For nr_listy = 1 To liczba_list
    Columns(1 + nr_listy).ColumnWidth = 8
Next nr_listy
ActiveSheet.Columns(liczba_list + 2).AutoFit
Columns(liczba_list + 3).ColumnWidth = 8

Sheets("Wyniki zbiorcze").Protect

End Sub

Sub pobieranie_wynikow()

Sheets("Wyniki zbiorcze").Unprotect

liczba_okregow = Application.WorksheetFunction.Max(Sheets("Dane wej�ciowe").Columns(4))
liczba_list = Application.WorksheetFunction.Max(Sheets("Dane wej�ciowe").Columns(8))
liczba_mandatow = Sheets("Dane wej�ciowe").Cells(1, 2)
prog = Sheets("Dane wej�ciowe").Cells(2, 2)
If "Pr�g wyborczy dla KKW (%)" = Sheets("Dane Wej�ciowe").Cells(3, 1).Value Then
    prog_KKW = Sheets("Dane wej�ciowe").Cells(3, 2)
ElseIf "Pr�g wyborczy dla komitet�w mniejszo�ci (%)" = Sheets("Dane Wej�ciowe").Cells(3, 1).Value Then
    prog_mniejszosc = Sheets("Dane wej�ciowe").Cells(3, 2)
Else
    prog_mniejszosc = Sheets("Dane wej�ciowe").Cells(4, 2)
End If

'czyszczenie starych danych
Sheets("Wyniki zbiorcze").Range(Cells(3, 2), Cells(2 + liczba_okregow + 3, 1 + liczba_list)).Value = ""
Sheets("Wyniki zbiorcze").Cells(1, liczba_list + 3).Value = ""

'pobieranie
Dim i As Integer
Dim j As Integer
For i = 1 To liczba_okregow
    Sheets("Wyniki zbiorcze").Cells(1, liczba_list + 3) = _
    Sheets("Wyniki zbiorcze").Cells(1, liczba_list + 3) + Sheets("Okr�g nr " & i).Cells(1, liczba_list + 3)
    For j = 1 To liczba_list
        Sheets("Wyniki zbiorcze").Cells(2 + i, 1 + j) = _
        Sheets("Okr�g nr " & i).Cells(Application.WorksheetFunction.Max(Sheets("Okr�g nr " & i).Columns(1)) + 3, 1 + j)
    Next j
Next i

'stwierdzenie wej�cia listy
Dim k As Integer
For k = 1 To liczba_list
    Cells(liczba_okregow + 3, 1 + k) = Application.WorksheetFunction.Sum(Range(Cells(3, 1 + k), Cells(2 + liczba_okregow, 1 + k)))
    Cells(liczba_okregow + 4, 1 + k) = Cells(liczba_okregow + 3, 1 + k) / Application.WorksheetFunction.Sum(Range(Cells(3, 2), Cells(2 + liczba_okregow, 1 + liczba_list)))
    If Sheets("Dane wej�ciowe").Cells(1 + k, 10) = "tak" Then
        If Cells(liczba_okregow + 4, 1 + k) >= prog_mniejszosc / 100 Then
            Cells(liczba_okregow + 4, 1 + k).Font.Bold = True
            Cells(liczba_okregow + 4, 1 + k).Interior.Color = vbGreen
        End If
    ElseIf Sheets("Dane wej�ciowe").Cells(1 + k, 9) = "tak" Then
        If Cells(liczba_okregow + 4, 1 + k) >= prog_KKW / 100 Then
            Cells(liczba_okregow + 4, 1 + k).Font.Bold = True
            Cells(liczba_okregow + 4, 1 + k).Interior.Color = vbGreen
        End If
    Else
        If Cells(liczba_okregow + 4, 1 + k) >= prog / 100 Then
           Cells(liczba_okregow + 4, 1 + k).Font.Bold = True
           Cells(liczba_okregow + 4, 1 + k).Interior.Color = vbGreen
        End If
    End If
Next k

'frekwencja
Cells(liczba_okregow + 5, 2) = (Application.WorksheetFunction.Sum(Range(Cells(liczba_okregow + 3, 2), Cells(liczba_okregow + 3, liczba_list + 1))) + Cells(1, liczba_list + 3)) _
/ Application.WorksheetFunction.Sum(Range(Sheets("Dane wej�ciowe").Cells(2, 5), Sheets("Dane wej�ciowe").Cells(1 + liczba_okregow, 5)))
Range(Cells(liczba_okregow + 4, 2), Cells(liczba_okregow + 5, 1 + liczba_list)).NumberFormat = "0.00%"
Sheets("Wyniki zbiorcze").Cells.HorizontalAlignment = xlCenter


'tworzenie przycisk�w wyboru metody
Dim mandaty_dHondt As Button
Set mandaty_dHondt = ActiveSheet.Buttons.Add _
(Width:=120, Height:=40, Left:=Cells(liczba_okregow + 9, liczba_list + 3).Left, _
Top:=Cells(liczba_okregow + 9, liczba_list + 3).Top)
With mandaty_dHondt
    .Text = "Ustal liczb� mandat�w metod� d'Hondta"
    .OnAction = "mandaty_dHondt"
End With

Dim mandaty_Sainte_Lague As Button
Set mandaty_Sainte_Lague = ActiveSheet.Buttons.Add _
(Width:=120, Height:=40, Left:=Cells(liczba_okregow + 9, liczba_list + 6).Left, _
Top:=Cells(liczba_okregow + 9, liczba_list + 6).Top)
With mandaty_Sainte_Lague
    .Text = "Ustal liczb� mandat�w metod� Sainte-Lague"
    .OnAction = "mandaty_Sainte_Lague"
End With

Sheets("Wyniki zbiorcze").Protect

End Sub

Sub mandaty_dHondt()
mandaty ("d'Hondt")
End Sub

Sub mandaty_Sainte_Lague()
mandaty ("Sainte-Lague")
End Sub

Sub mandaty(metoda)
Sheets("Wyniki zbiorcze").Unprotect
Dim x As Integer
Dim y As Integer
If metoda = "d'Hondt" Then
    x = 1
    y = 0
ElseIf metoda = "Sainte-Lague" Then
    x = 2
    y = 1
End If
    
Dim maks_mandatow As Integer
liczba_okregow = Application.WorksheetFunction.Max(Sheets("Dane wej�ciowe").Columns(4))
liczba_list = Application.WorksheetFunction.Max(Sheets("Dane wej�ciowe").Columns(8))
liczba_mandatow = Sheets("Dane wej�ciowe").Cells(1, 2)
maks_mandatow = Application.WorksheetFunction.Max(Sheets("Dane wej�ciowe").Columns(6))

Cells(liczba_okregow + 9, 1) = metoda

Dim mandaty_w_okregach() As Variant
mandaty_w_okregach = Application.Transpose(Range(Sheets("Dane wej�ciowe").Cells(2, 6), Sheets("Dane wej�ciowe").Cells(1 + liczba_okregow, 6)))

Dim i, j, k As Integer
Dim tabela_wspolczynnikow() As Variant
ReDim tabela_wspolczynnikow(liczba_okregow, liczba_list, maks_mandatow)
For i = 1 To liczba_okregow
    For j = 1 To liczba_list
        For k = 1 To mandaty_w_okregach(i)
            tabela_wspolczynnikow(i, j, k) = Cells(2 + i, 1 + j) / (x * k - y)
        Next k
    Next j
Next i

Dim listy As Integer
Dim i2, i3 As Integer
For listy = 1 To liczba_list
    If Not Cells(liczba_okregow + 4, 1 + listy).Interior.Color = vbGreen Then
        For i2 = 1 To liczba_okregow
            For i3 = 1 To maks_mandatow
                tabela_wspolczynnikow(i2, listy, i3) = 0
            Next i3
        Next i2
    End If
Next listy

Dim i4, i5, i6 As Integer
Sheets("Wyniki zbiorcze").Cells(liczba_okregow + 8 + 2, 1) = "Numer okr�gu"
Sheets("Wyniki zbiorcze").Cells(liczba_okregow * 2 + 8 + 3, 1) = "Liczba mandat�w"
Sheets("Wyniki zbiorcze").Cells(liczba_okregow * 2 + 8 + 4, 1) = "Procent mandat�w"
Sheets("Wyniki zbiorcze").Cells(liczba_okregow + 8 + 1, 2) = "Numer listy"


Range(Sheets("Wyniki zbiorcze").Cells(liczba_okregow + 8 + 1, 2), Sheets("Wyniki zbiorcze").Cells(liczba_okregow + 8 + 1, 1 + liczba_list)).Merge
For i4 = 1 To liczba_okregow
    Cells(liczba_okregow + 8 + 2 + i4, 1) = Cells(liczba_okregow + 8 + 2 + i4, 1).Offset(-8 - liczba_okregow, 0).Value
Next i4
For i5 = 1 To liczba_list
    Cells(liczba_okregow + 8 + 2, 1 + i5) = Cells(liczba_okregow + 8 + 2, 1 + i5).Offset(-8 - liczba_okregow, 0).Value
Next i5

Range(Cells(11 + liczba_okregow, 2), Cells(10 + 2 * liczba_okregow, 1 + liczba_list)) = 0
        
Dim okregi, mandaty As Integer
Dim mandat_dla_listy As Integer
Dim najw_wsp As Long
For okregi = 1 To liczba_okregow
    For mandaty = 1 To mandaty_w_okregach(okregi)
    najw_wsp = tabela_wspolczynnikow(okregi, 1, 1)
    mandat_dla_listy = 1
        For listy = 1 To liczba_list
            If tabela_wspolczynnikow(okregi, listy, 1) > najw_wsp Then
                najw_wsp = tabela_wspolczynnikow(okregi, listy, 1)
                mandat_dla_listy = listy
            End If
        Next listy
    
        Cells(10 + liczba_okregow + okregi, 1 + mandat_dla_listy) = Cells(10 + liczba_okregow + okregi, 1 + mandat_dla_listy) + 1
        'Cells(10 + liczba_okregow + okregi, 1 + mandat_dla_listy).Font.Bold = True
        For i6 = 1 To maks_mandatow - 1
            tabela_wspolczynnikow(okregi, mandat_dla_listy, i6) = tabela_wspolczynnikow(okregi, mandat_dla_listy, i6 + 1)
        Next i6
    Next mandaty
Next okregi

Dim i7 As Integer
For i7 = 1 To liczba_list
    Cells(liczba_okregow * 2 + 8 + 3, 1 + i7) = Application.WorksheetFunction.Sum(Range(Cells(liczba_okregow + 8 + 3, 1 + i7), Cells(liczba_okregow + 8 + 2 + liczba_okregow, 1 + i7)))
    Cells(liczba_okregow * 2 + 8 + 4, 1 + i7) = Cells(liczba_okregow * 2 + 8 + 4 - 1, 1 + i7) / liczba_mandatow
Next i7
Range(Cells(liczba_okregow * 2 + 8 + 4, 1 + 1), Cells(liczba_okregow * 2 + 8 + 4, 1 + liczba_list)).NumberFormat = "0.0%"

'formatowanie
With Union(Range(Cells(liczba_okregow + 9, 2), Cells(liczba_okregow + 9, 1 + liczba_list)), Range(Cells(liczba_okregow + 10, 1), Cells(liczba_okregow + 10, 1).End(xlToRight).End(xlDown)))
    .Interior.ColorIndex = 0
    .Borders.LineStyle = xlContinuous
End With
With Range(Cells(liczba_okregow * 2 + 11, 1), Cells(liczba_okregow * 2 + 11, liczba_list + 1)).Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlMedium
End With
With Union(Cells(liczba_okregow + 9, 2), Range(Cells(liczba_okregow + 10, 1), Cells(liczba_okregow * 2 + 12, 1)), _
Range(Cells(liczba_okregow + 10, 2), Cells(liczba_okregow + 10, liczba_list + 1)), Range(Cells(liczba_okregow * 2 + 11, 2), Cells(liczba_okregow * 2 + 12, 1 + liczba_list)))
    .Font.Bold = True
End With
Dim m1 As Integer
Dim m2 As Integer
For m1 = 1 To liczba_list
    For m2 = 1 To liczba_okregow + 1
        If Cells(liczba_okregow + 10 + m2, 1 + m1).Value = 0 Then
            Cells(liczba_okregow + 10 + m2, 1 + m1).Font.Color = RGB(225, 225, 225)
        Else
            Cells(liczba_okregow + 10 + m2, 1 + m1).Font.Color = RGB(0, 0, 0)
        End If
    Next m2
Next m1
    


Sheets("Wyniki zbiorcze").Protect

End Sub


