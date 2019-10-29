VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} okno_startowe 
   Caption         =   "Witaj szanowny u¿ytkowniku"
   ClientHeight    =   4200
   ClientLeft      =   -240
   ClientTop       =   -852
   ClientWidth     =   3900
   OleObjectBlob   =   "okno_startowe.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "okno_startowe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub mandaty_AfterUpdate()
Call czy_liczba(okno_startowe.mandaty, "Wpisz nieujemn¹ liczbê ca³kowit¹ proszê")
End Sub

Private Sub okregi_AfterUpdate()
Call czy_liczba(okno_startowe.okregi, "Wpisz nieujemn¹ liczbê ca³kowit¹ proszê")
End Sub

Private Sub komitety_AfterUpdate()
Call czy_liczba(okno_startowe.komitety, "Wpisz nieujemn¹ liczbê ca³kowit¹ proszê")
End Sub

Private Sub prog_AfterUpdate()
Call czy_liczba(okno_startowe.prog, "Wpisz nieujemn¹ liczbê ca³kowit¹ proszê")

If czy_prog_KKW.Value = False Then
    prog_KKW.Value = prog
End If

If czy_prog_mniejszosc.Value = False Then
    prog_mniejszosc.Value = prog
End If

End Sub

Private Sub czy_prog_KKW_Click()

If czy_prog_KKW.Value = True Then
    prog_KKW.Locked = False
    prog_KKW.BackColor = vbWhite
End If

If czy_prog_KKW.Value = False Then
    prog_KKW.Value = prog
    prog_KKW.BackColor = RGB(192, 192, 192)
    prog_KKW.Locked = True
End If

End Sub

Private Sub prog_KKW_AfterUpdate()
Call czy_liczba(okno_startowe.prog_KKW, "Wpisz nieujemn¹ liczbê ca³kowit¹ proszê")
End Sub

Private Sub czy_prog_mniejszosc_Click()

If czy_prog_mniejszosc.Value = True Then
    prog_mniejszosc.Locked = False
    prog_mniejszosc.BackColor = vbWhite
End If

If czy_prog_mniejszosc.Value = False Then
    prog_mniejszosc.Value = prog
    prog_mniejszosc.BackColor = RGB(192, 192, 192)
    prog_mniejszosc.Locked = True
End If

End Sub

Private Sub prog_mniejszosc_AfterUpdate()
Call czy_liczba(okno_startowe.prog_mniejszosc, "Wpisz nieujemn¹ liczbê ca³kowit¹ proszê")
End Sub





Public Sub zatwierdz_Click()
If okno_startowe.okregi = "" Or okno_startowe.mandaty = "" Or okno_startowe.komitety = "" Or _
okno_startowe.prog = "" Or okno_startowe.prog_KKW = "" Or okno_startowe.prog_mniejszosc = "" Then
    MsgBox ("Uzupe³nij dane")
Else
    Module1.liczba_mandatow = Int(okno_startowe.mandaty)
    Module1.liczba_okregow = Int(okno_startowe.okregi)
    Module1.liczba_list = Int(okno_startowe.komitety)
    Module1.prog = Int(okno_startowe.prog)
    Module1.prog_KKW = Int(okno_startowe.prog_KKW)
    Module1.prog_mniejszosc = Int(okno_startowe.prog_mniejszosc)
    okno_startowe.Hide
    okno_startowe.mandaty = ""
    okno_startowe.okregi = ""
    okno_startowe.komitety = ""
End If

    
End Sub
