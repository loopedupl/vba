# Sortowanie i filtrowanie danych

## Krok 1: Sortowanie danych

VBA pozwala na sortowanie danych w arkuszu według określonego kryterium. Aby posortować dane w kolumnie A w kolejności rosnącej:

```vba
Sub SortowanieRosnace()
    Range("A1:A20").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
End Sub
```

Dla kolejności malejącej:

```vba
Sub SortowanieMalejace()
    Range("A1:A20").Sort Key1:=Range("A1"), Order1:=xlDescending, Header:=xlYes
End Sub
```

W tym kodzie `Key1` odnosi się do kolumny, według której sortujemy, a `Order1` do kolejności sortowania.

## Krok 2: Sortowanie według wielu kryteriów

Jeśli chcesz posortować dane według więcej niż jednego kryterium, możesz to zrobić za pomocą metody `Sort`.

```vba
Sub SortowanieWielokryterialne()
    Range("A1:C20").Sort Key1:=Range("A1"), Order1:=xlAscending, _
                        Key2:=Range("B1"), Order2:=xlDescending, Header:=xlYes
End Sub
```

W tym przykładzie dane są najpierw sortowane według kolumny A w kolejności rosnącej, a następnie według kolumny B w kolejności malejącej.

## Krok 3: Filtrowanie danych

Filtrowanie umożliwia wyświetlanie tylko tych danych, które spełniają określone kryteria. Aby włączyć filtr na wybranym zakresie:

```vba
Sub WlaczFiltr()
    Range("A1:C20").AutoFilter
End Sub
```

Aby zastosować filtr na określonej kolumnie:

```vba
Sub FiltrowanieDanych()
    Range("A1:C20").AutoFilter Field:=2, Criteria1:=">100"
End Sub
```

W tym przykładzie wyświetlane są tylko wiersze, w których wartość w kolumnie B (`Field:=2`) jest większa niż 100.

## Krok 4: Wyłączenie filtra

Aby wyłączyć filtr w VBA:

```vba
Sub WylaczFiltr()
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False
End Sub
```

## Krok 5: Filtrowanie zaawansowane

Filtrowanie zaawansowane pozwala na wykorzystanie złożonych warunków. Przykład zastosowania filtra zaawansowanego:

```vba
Sub FiltrowanieZaawansowane()
    Range("A1:C20").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=Range("E1:F2")
End Sub
```

W tym przykładzie zakres `E1:F2` zawiera kryteria filtrowania.

## Krok 6: Kopiowanie wyników filtra

Możesz skopiować wyniki filtrowania do innego miejsca w arkuszu:

```vba
Sub KopiowanieWynikowFiltra()
    Range("A1:C20").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range("E1:F2"), _
                                    CopyToRange:=Range("H1:J1")
End Sub
```

## Podsumowanie

Sortowanie i filtrowanie danych to kluczowe umiejętności, które pozwalają efektywnie organizować i analizować dane w Excelu. Opanowanie tych technik pozwala na tworzenie dynamicznych raportów i automatyzację codziennych zadań analitycznych.
