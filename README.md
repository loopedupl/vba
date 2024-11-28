# Dynamiczne aktualizowanie wykresów

## Krok 1: Wykorzystanie dynamicznych zakresów danych

Dynamiczne zakresy umożliwiają automatyczne dostosowywanie źródła danych wykresu do zmieniających się wartości. Oto przykład definiowania zakresu dynamicznego przy użyciu funkcji `OFFSET`:

1. Przejdź do zakładki **Formuły** > **Menadżer nazw**.
2. Utwórz nową nazwę, np. `DynamiczneDane`, a jako formułę wpisz:

   ```excel
   =OFFSET(Arkusz1!$A$1,0,0,COUNTA(Arkusz1!$A:$A),1)
   ```

3. Użyj tej nazwy jako źródła danych wykresu.

## Krok 2: Aktualizacja wykresu za pomocą VBA

Aby dynamicznie odświeżyć wykres w VBA:

```vba
Sub AktualizujWykres()
    Dim Wykres As ChartObject
    Set Wykres = ActiveSheet.ChartObjects(1)
    Wykres.Chart.SetSourceData Source:=Range("DynamiczneDane")
End Sub
```

## Krok 3: Dodanie przycisku do aktualizacji

Możesz dodać przycisk, który uruchomi makro aktualizujące wykres:

1. Wstaw **Przycisk Formantu** z zakładki **Deweloper**.
2. Przypisz do niego makro `AktualizujWykres`.

## Krok 4: Automatyczne odświeżanie przy zmianie danych

Aby wykres aktualizował się automatycznie, gdy zmienią się dane, użyj zdarzenia arkusza:

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("A:A")) Is Nothing Then
        Call AktualizujWykres
    End If
End Sub
```

## Krok 5: Dostosowanie wyglądu wykresu

Podczas dynamicznego odświeżania możesz także dostosować wygląd wykresu, np. zmieniając kolory serii:

```vba
Sub ZmienKolorSerii()
    Dim Wykres As ChartObject
    Set Wykres = ActiveSheet.ChartObjects(1)
    Wykres.Chart.SeriesCollection(1).Interior.Color = RGB(255, 0, 0) ' Czerwony
End Sub
```

## Krok 6: Dodawanie komunikatów przy błędach danych

Aby upewnić się, że wykres nie zostanie zepsuty przez brak danych:

```vba
Sub SprawdzDane()
    Dim Wykres As ChartObject
    Set Wykres = ActiveSheet.ChartObjects(1)
    If WorksheetFunction.CountA(Range("A:A")) = 0 Then
        MsgBox "Brak danych do wykresu!", vbExclamation
    Else
        Call AktualizujWykres
    End If
End Sub
```

## Podsumowanie

Dynamiczne aktualizowanie wykresów eliminuje potrzebę ręcznego dostosowywania danych. Dzięki temu możesz zautomatyzować swoje raporty i zyskać więcej czasu na analizę wyników.
