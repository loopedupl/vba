# Nagrywanie i edytowanie makr

## Krok 1: Nagrywanie makra

Excel umożliwia nagrywanie makr, które zapisują Twoje działania w postaci kodu VBA. Aby nagrać makro:

1. Przejdź do zakładki `Deweloper` na pasku narzędzi.
2. Kliknij przycisk `Nagrywanie makra`.
3. Podaj nazwę dla makra oraz, jeśli chcesz, przypisz skrót klawiszowy.
4. Wybierz, gdzie zapisać makro: w tym skoroszycie lub w osobnym pliku.
5. Kliknij `OK` i zacznij wykonywać czynności, które chcesz nagrać.
6. Po zakończeniu nagrywania kliknij `Zatrzymaj nagrywanie`.

Makro zapisze Twoje działania jako kod VBA, który będziesz mógł później edytować.

## Krok 2: Edytowanie nagranego makra

Po nagraniu makra, możesz je edytować, aby dostosować kod do własnych potrzeb. Aby to zrobić:

1. Przejdź do edytora VBA (naciśnij `Alt + F11`).
2. W edytorze VBA znajdź swój moduł, w którym zapisane zostało makro.
3. Zobaczysz kod wygenerowany przez Excel. Możesz go edytować, np. zmieniając argumenty funkcji, dodając nowe instrukcje lub pętle.

Oto przykład kodu nagranego makra:

```vba
Sub NagranieMakra()
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Witaj w VBA!"
    Range("A2").Select
End Sub
```

W tym przykładzie makro zapisuje tekst "Witaj w VBA!" w komórce A1.

## Krok 3: Dostosowanie kodu VBA

Chociaż Excel nagrywa Twoje działania, kod może być bardzo podstawowy. Dlatego warto go edytować, aby zwiększyć jego efektywność. Możesz np.:

- Zastąpić konkretne wartości zmiennymi.
- Dodać pętle, aby wykonywać operacje na wielu komórkach.
- Zastosować warunki, aby makro działało tylko w określonych przypadkach.

Przykład edytowanego makra z pętlą:

```vba
Sub EdytowaneMakro()
    Dim i As Integer
    For i = 1 To 10
        Cells(i, 1).Value = "Wiersz " & i
    Next i
End Sub
```

To makro wpisuje "Wiersz 1", "Wiersz 2", ..., "Wiersz 10" do pierwszej kolumny.

## Krok 4: Uruchamianie i testowanie makra

Po edytowaniu makra możesz je uruchomić, klikając `F5` w edytorze VBA lub używając skrótu, jeśli został przypisany podczas nagrywania. Przetestuj makro na różnych danych, aby upewnić się, że działa zgodnie z oczekiwaniami.

## Podsumowanie

Nagrywanie i edytowanie makr to świetny sposób na automatyzację codziennych zadań w Excelu. Dzięki tej lekcji nauczyłeś się, jak rejestrować swoje działania, a także jak dostosować zapisany kod VBA, aby lepiej pasował do Twoich potrzeb. Teraz możesz tworzyć bardziej zaawansowane makra, które zaoszczędzą Ci czas i poprawią efektywność pracy.
