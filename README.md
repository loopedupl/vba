# Tworzenie wykresów za pomocą VBA

## Krok 1: Tworzenie prostego wykresu

Aby utworzyć wykres liniowy z danych znajdujących się w zakresie `A1:B10`:

```vba
Sub TworzWykresLiniowy()
    Dim Wykres As ChartObject
    Set Wykres = ActiveSheet.ChartObjects.Add(Left:=100, Width:=400, Top:=50, Height:=300)
    Wykres.Chart.SetSourceData Source:=Range("A1:B10")
    Wykres.Chart.ChartType = xlLine
End Sub
```

## Krok 2: Tworzenie wykresu kolumnowego

Aby stworzyć wykres kolumnowy z danych:

```vba
Sub TworzWykresKolumnowy()
    Dim Wykres As ChartObject
    Set Wykres = ActiveSheet.ChartObjects.Add(Left:=100, Width:=400, Top:=50, Height:=300)
    Wykres.Chart.SetSourceData Source:=Range("A1:B10")
    Wykres.Chart.ChartType = xlColumnClustered
End Sub
```

## Krok 3: Dodawanie tytułu do wykresu

Możesz dodać tytuł do wykresu, aby lepiej opisać jego zawartość:

```vba
Sub DodajTytul()
    Dim Wykres As ChartObject
    Set Wykres = ActiveSheet.ChartObjects(1)
    Wykres.Chart.HasTitle = True
    Wykres.Chart.ChartTitle.Text = "Mój Wykres"
End Sub
```

## Krok 4: Formatowanie osi wykresu

Aby dostosować osie wykresu:

```vba
Sub FormatowanieOsi()
    Dim Wykres As ChartObject
    Set Wykres = ActiveSheet.ChartObjects(1)
    With Wykres.Chart.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "Kategoria"
    End With
    With Wykres.Chart.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "Wartość"
    End With
End Sub
```

## Krok 5: Dynamiczne zmienianie danych wykresu

Aby dynamicznie zmieniać dane wykresu:

```vba
Sub ZmienDaneWykresu()
    Dim Wykres As ChartObject
    Set Wykres = ActiveSheet.ChartObjects(1)
    Wykres.Chart.SetSourceData Source:=Range("C1:D10")
End Sub
```

## Krok 6: Tworzenie wykresu w nowym arkuszu

Aby utworzyć wykres na nowym arkuszu:

```vba
Sub WykresNaNowymArkuszu()
    Dim Wykres As Chart
    Set Wykres = Charts.Add
    Wykres.SetSourceData Source:=Range("A1:B10")
    Wykres.ChartType = xlPie
    Wykres.Location Where:=xlLocationAsNewSheet
End Sub
```

## Krok 7: Usuwanie wykresu

Aby usunąć istniejący wykres:

```vba
Sub UsunWykres()
    Dim Wykres As ChartObject
    For Each Wykres In ActiveSheet.ChartObjects
        Wykres.Delete
    Next Wykres
End Sub
```

## Podsumowanie

Tworzenie wykresów za pomocą VBA pozwala na automatyzację prezentacji danych w atrakcyjnej wizualnie formie. Opanowanie tej umiejętności przyspieszy Twoją pracę z raportami i umożliwi dynamiczne przedstawianie wyników analiz.
