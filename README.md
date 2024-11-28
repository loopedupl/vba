# Interfejsy w VBA

## Co to są interfejsy?

Interfejs to abstrakcyjna definicja zestawu metod i właściwości, które klasy implementujące ten interfejs muszą zaimplementować. W VBA interfejsy są definiowane jako klasy z pustymi metodami i właściwościami.

---

## Dlaczego warto używać interfejsów?

- **Elastyczność**: Możesz tworzyć klasy implementujące te same metody, ale z różnym zachowaniem.
- **Łatwiejsza konserwacja kodu**: Zmiana implementacji interfejsu w jednej klasie nie wpływa na inne klasy.
- **Skalowalność**: Możesz dodawać nowe implementacje bez modyfikowania istniejącego kodu.

---

## Tworzenie interfejsu

### Przykład interfejsu

```vba
' Interfejs IObliczenia
Public Function Oblicz() As Double
End Function
```

### Implementacja interfejsu w klasie

1. Utwórz nową klasę.
2. Dodaj słowo kluczowe `Implements` oraz nazwę interfejsu.
3. Zaimplementuj wszystkie metody i właściwości zdefiniowane w interfejsie.

```vba
' Klasa implementująca interfejs IObliczenia
Implements IObliczenia

Private Function IObliczenia_Oblicz() As Double
    IObliczenia_Oblicz = 3.14 ' Przykładowa implementacja
End Function
```

---

## Korzystanie z interfejsów

### 1. **Tworzenie instancji klasy implementującej interfejs**

```vba
Dim obj As IObliczenia
Set obj = New KlasaObliczenia
MsgBox obj.Oblicz
```

### 2. **Wielokrotne implementacje**

Różne klasy mogą implementować ten sam interfejs, zapewniając różne zachowanie:

```vba
Dim obj1 As IObliczenia
Dim obj2 As IObliczenia

Set obj1 = New KlasaObliczeniaA
Set obj2 = New KlasaObliczeniaB

MsgBox obj1.Oblicz ' Wynik A
MsgBox obj2.Oblicz ' Wynik B
```

---

## Praktyczne zastosowanie interfejsów

- **Kalkulatory**: Tworzenie kalkulatorów różnych typów (np. matematycznego, finansowego).
- **Przetwarzanie danych**: Implementacja różnych metod przetwarzania danych w zależności od formatu wejściowego.
- **Strategie raportowania**: Różne style generowania raportów oparte na wspólnym interfejsie.

---

## Podsumowanie

Interfejsy to potężne narzędzie, które zwiększa elastyczność i modularność kodu w VBA. Dzięki nim możesz tworzyć aplikacje łatwiejsze do rozbudowy i utrzymania, co jest kluczowe w pracy z większymi projektami.
