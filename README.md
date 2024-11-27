# Instrukcje warunkowe – Decyzje w kodzie

## Krok 1: Podstawowa instrukcja If

Instrukcja `If` pozwala na wykonanie kodu tylko wtedy, gdy określony warunek jest prawdziwy. Oto przykład użycia:

```vba
Sub PrzykładIf()
    Dim x As Integer
    x = 5
    If x > 3 Then
        MsgBox "x jest większe od 3"
    End If
End Sub
```

W tym przypadku, jeśli zmienna `x` jest większa niż 3, wyświetli się komunikat.

## Krok 2: Instrukcja ElseIf i Else

Czasami chcemy sprawdzić więcej niż jeden warunek. W takim przypadku używamy instrukcji `ElseIf` i `Else`.

```vba
Sub PrzykładElseIf()
    Dim x As Integer
    x = 5
    If x > 10 Then
        MsgBox "x jest większe od 10"
    ElseIf x > 5 Then
        MsgBox "x jest większe od 5, ale mniejsze lub równe 10"
    Else
        MsgBox "x jest mniejsze lub równe 5"
    End If
End Sub
```

W tym przykładzie kod sprawdza, czy `x` jest większe niż 10, większe niż 5, czy mniejsze lub równe 5, wykonując odpowiednią akcję w każdym przypadku.

## Krok 3: Instrukcja Select Case

Jeśli masz do sprawdzenia wiele różnych warunków, zamiast używać wielu instrukcji `If`, lepiej jest skorzystać z instrukcji `Select Case`. Przykład:

```vba
Sub PrzykładSelectCase()
    Dim x As Integer
    x = 3
    Select Case x
        Case 1
            MsgBox "x wynosi 1"
        Case 2
            MsgBox "x wynosi 2"
        Case 3
            MsgBox "x wynosi 3"
        Case Else
            MsgBox "x ma inną wartość"
    End Select
End Sub
```

Instrukcja `Select Case` jest bardziej przejrzysta, gdy mamy wiele różnych wartości, które chcemy sprawdzić.

## Krok 4: Operatory logiczne

Aby tworzyć bardziej złożone warunki, możemy używać operatorów logicznych, takich jak `And`, `Or`, czy `Not`.

```vba
Sub PrzykładOperatoryLogiczne()
    Dim x As Integer
    Dim y As Integer
    x = 5
    y = 10

    If x > 3 And y < 15 Then
        MsgBox "Warunek jest spełniony"
    Else
        MsgBox "Warunek nie jest spełniony"
    End If
End Sub
```

Tutaj warunek sprawdza, czy obie zmienne `x` i `y` spełniają określone kryteria.

## Podsumowanie

Instrukcje warunkowe w VBA pozwalają na tworzenie bardziej dynamicznego i elastycznego kodu, który reaguje na zmienne i dane wejściowe. Dzięki tej lekcji opanujesz umiejętność podejmowania decyzji w swoim kodzie, co jest kluczowe w programowaniu VBA. Używaj instrukcji `If`, `Select Case` oraz operatorów logicznych, aby dostosować działanie swojego programu do różnych scenariuszy.
