# Generowanie raportów PDF za pomocą VBA

## Krok 1: Ustawianie zakresu eksportu

Aby określić, który obszar arkusza ma zostać zapisany jako PDF:

1. Wybierz zakres komórek, które mają zostać zapisane.
2. W VBA przypisz zakres za pomocą właściwości `PrintArea`:

   ```vba
   Sub UstawZakresEksportu()
       ActiveSheet.PageSetup.PrintArea = "$A$1:$D$20"
   End Sub
   ```

## Krok 2: Generowanie pliku PDF

Aby zapisać zakres w pliku PDF, użyj metody `ExportAsFixedFormat`:

```vba
Sub GenerujPDF()
    Dim Sciezka As String
    Sciezka = ThisWorkbook.Path & "\Raport.pdf"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=Sciezka
    MsgBox "Raport został zapisany jako PDF!", vbInformation
End Sub
```

## Krok 3: Automatyzacja nazwy pliku

Możesz dynamicznie tworzyć nazwy plików w zależności od daty lub innych zmiennych:

```vba
Sub GenerujPDFZNazwą()
    Dim Sciezka As String
    Sciezka = ThisWorkbook.Path & "\Raport_" & Format(Date, "YYYY-MM-DD") & ".pdf"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=Sciezka
    MsgBox "Raport zapisany jako: " & Sciezka, vbInformation
End Sub
```

## Krok 4: Eksport wielu arkuszy

Aby zapisać wszystkie arkusze w jednym pliku PDF:

```vba
Sub EksportWszystkichArkuszy()
    Dim Sciezka As String
    Sciezka = ThisWorkbook.Path & "\Raport_Zbiorczy.pdf"
    ThisWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=Sciezka
    MsgBox "Wszystkie arkusze zapisane jako PDF!", vbInformation
End Sub
```

## Krok 5: Obsługa błędów

Aby uniknąć problemów, np. braku dostępu do folderu zapisu:

```vba
Sub GenerujZObslugaBledow()
    On Error GoTo Blad
    Dim Sciezka As String
    Sciezka = ThisWorkbook.Path & "\Raport.pdf"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=Sciezka
    MsgBox "Raport zapisany!", vbInformation
    Exit Sub
Blad:
    MsgBox "Nie udało się zapisać pliku. Sprawdź lokalizację i spróbuj ponownie.", vbExclamation
End Sub
```

## Krok 6: Dodanie automatycznego układu strony

Aby dopasować eksport do formatu A4 i ustawić orientację:

```vba
Sub UstawieniaStrony()
    With ActiveSheet.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
End Sub
```

## Podsumowanie

Generowanie raportów PDF za pomocą VBA pozwala na szybkie tworzenie profesjonalnych dokumentów gotowych do udostępnienia. Ta automatyzacja oszczędza czas i eliminuje błędy związane z ręcznym przygotowywaniem raportów.
