# Pierwsze kroki w edytorze VBA

## Krok 1: Otwieranie edytora VBA

Aby otworzyć edytor VBA w Excelu, należy nacisnąć `Alt + F11`. Zostaniesz przeniesiony do okna edytora, gdzie możesz pisać i uruchamiać swoje makra.

## Krok 2: Tworzenie modułów

Moduł w VBA to kontener na kod. Aby stworzyć nowy moduł:

1. Przejdź do `Wstaw` > `Moduł`.
2. Pojawi się pusty moduł, w którym możesz zacząć pisać swoje makra.

## Krok 3: Pisanie pierwszego makra

Makra w VBA to nic innego jak zestaw poleceń, które będą wykonywane automatycznie. Oto przykład prostego makra:

```vba
Sub Powitanie()
    MsgBox "Witaj w kursie VBA!"
End Sub
```

To makro wyświetli okno z powitaniem. Możesz je uruchomić, klikając przycisk `F5` lub wybierając `Uruchom` w górnym menu edytora.

## Krok 4: Uruchamianie makr

Aby uruchomić swoje makro:

1. Zapisz kod w edytorze.
2. Przejdź do Excela i naciśnij `Alt + F8`.
3. Wybierz swoje makro z listy i kliknij `Uruchom`.

## Krok 5: Edytowanie makr

Możesz edytować swoje makra w edytorze VBA, zmieniając kod. W przypadku zmiany funkcji makra, po zapisaniu zmian w edytorze, nowe polecenie będzie miało zastosowanie w przyszłych uruchomieniach.

## Krok 6: Zasady organizacji kodu

- Używaj modułów do grupowania funkcji, które są ze sobą powiązane.
- Nazwij moduły zgodnie z ich przeznaczeniem, np. `ModulMakra`, `ObslugaBledow`.
- Używaj odpowiednich komentarzy, aby łatwiej było odnaleźć poszczególne fragmenty kodu w przyszłości.

## Podsumowanie

Edytor VBA to narzędzie, które umożliwia tworzenie, edytowanie i uruchamianie makr. Dzięki tej lekcji nauczyłeś się, jak wstawić moduł, napisać swoje pierwsze makro i uruchomić je w Excelu. Teraz możesz rozwijać swoje umiejętności programowania w VBA, tworząc bardziej złożone makra i organizując swój kod w modułach.
