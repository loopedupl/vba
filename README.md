# Pętle – Automatyzacja na wyższym poziomie

## Krok 1: Pętla For

Pętla `For` pozwala na wykonanie określonego bloku kodu określoną liczbę razy. Oto przykład:

```vba
Sub PrzykładFor()
    Dim i As Integer
    For i = 1 To 5
        MsgBox "Wartość i wynosi " & i
    Next i
End Sub
```

W tym przypadku pętla wykona się pięć razy, a zmienna `i` będzie przyjmować wartości od 1 do 5.

## Krok 2: Pętla For Each

Pętla `For Each` jest szczególnie przydatna, gdy chcesz iterować po kolekcjach, takich jak tablice czy obiekty.

```vba
Sub PrzykładForEach()
    Dim cell As Range
    For Each cell In Range("A1:A5")
        MsgBox "Wartość komórki to " & cell.Value
    Next cell
End Sub
```

W tym przypadku pętla przechodzi przez każdą komórkę w zakresie A1:A5, wykonując określoną akcję.

## Krok 3: Pętla Do While

Pętla `Do While` będzie kontynuować działanie, dopóki określony warunek jest prawdziwy.

```vba
Sub PrzykładDoWhile()
    Dim i As Integer
    i = 1
    Do While i <= 5
        MsgBox "Wartość i wynosi " & i
        i = i + 1
    Loop
End Sub
```

Pętla będzie wykonywać się, dopóki zmienna `i` nie przekroczy wartości 5.

## Krok 4: Pętla Do Until

Pętla `Do Until` działa w odwrotności do pętli `Do While`. Pętla będzie wykonywać się, dopóki warunek nie stanie się prawdziwy.

```vba
Sub PrzykładDoUntil()
    Dim i As Integer
    i = 1
    Do Until i > 5
        MsgBox "Wartość i wynosi " & i
        i = i + 1
    Loop
End Sub
```

W tej pętli proces będzie kontynuowany, dopóki zmienna `i` nie przekroczy 5.

## Krok 5: Przerwanie pętli za pomocą Exit

Czasami chcemy przerwać pętlę przed jej zakończeniem. Możemy to zrobić za pomocą instrukcji `Exit`.

```vba
Sub PrzykładExit()
    Dim i As Integer
    For i = 1 To 10
        If i = 5 Then Exit For
        MsgBox "Wartość i wynosi " & i
    Next i
End Sub
```

Tutaj pętla `For` zostanie przerwana, gdy `i` osiągnie wartość 5.

```vba
Sub PrzykładExitDo()
    Dim i As Integer
    i = 1
    Do
        MsgBox "Wartość i wynosi " & i
        i = i + 1
        If i = 4 Then Exit Do ' Zakończenie pętli po osiągnięciu wartości 4
    Loop
End Sub
```

W tym przykładzie pętla wykona się tylko trzy razy, a gdy zmienna `i` osiągnie wartość 4, pętla zostanie przerwana.

## Podsumowanie

Pętle to podstawowe narzędzie w VBA, które pozwala na automatyzację powtarzalnych zadań. Dzięki pętlom `For`, `For Each`, `Do While` i `Do Until` będziesz w stanie szybko i efektywnie przechodzić przez dane. Zastosowanie instrukcji `Exit` pozwala na pełną kontrolę nad przebiegiem pętli. Pętle to must-have w automatyzacji, dlatego warto je opanować, by zwiększyć swoją produktywność.
