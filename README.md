# Tworzenie i wykorzystanie klas w VBA

## Wprowadzenie do klas i obiektów

Klasa w VBA to szablon, na podstawie którego można tworzyć obiekty. Obiekty to instancje klas, które posiadają własne właściwości i metody. Dzięki klasom można lepiej organizować kod, grupując logicznie powiązane funkcjonalności.

---

## Tworzenie klasy w VBA

### 1. **Dodawanie klasy**

Aby utworzyć klasę:

1. Otwórz edytor VBA (Alt + F11).
2. Wybierz **Wstaw → Klasa modułu**.
3. Nadaj nazwę klasie (np. `Klient`).

---

### 2. **Definiowanie właściwości klasy**

Właściwości definiuje się za pomocą zmiennych prywatnych i procedur `Property Let`, `Property Get` oraz `Property Set`.

Przykład:

```vba
Private m_Imie As String

' Ustawianie właściwości
Public Property Let Imie(Value As String)
    m_Imie = Value
End Property

' Pobieranie właściwości
Public Property Get Imie() As String
    Imie = m_Imie
End Property
```

---

### 3. **Dodawanie metod do klasy**

Metody to procedury lub funkcje zdefiniowane w klasie.

Przykład:

```vba
Public Function Przywitaj() As String
    Przywitaj = "Witaj, " & m_Imie & "!"
End Function
```

---

## Wykorzystanie klasy w kodzie

### Tworzenie instancji klasy

Aby użyć klasy w VBA, należy utworzyć jej instancję:

```vba
Dim nowyKlient As Klient
Set nowyKlient = New Klient
```

### Ustawianie właściwości i wywoływanie metod

Po utworzeniu obiektu można korzystać z jego właściwości i metod:

```vba
nowyKlient.Imie = "Jan"
MsgBox nowyKlient.Przywitaj() ' Wyświetli: "Witaj, Jan!"
```

---

## Zastosowania klas w VBA

1. **Reprezentowanie danych**  
   Klasy mogą reprezentować bardziej złożone dane, np. produkty, klientów, zamówienia.

2. **Modularność kodu**  
   Kod staje się łatwiejszy w zarządzaniu i rozbudowie, gdy jest zorganizowany wokół klas.

3. **Wielokrotne użycie**  
   Raz napisane klasy można wykorzystywać w wielu projektach.

---

## Zalety programowania z użyciem klas

- **Lepsza organizacja kodu** – klasy grupują powiązane funkcjonalności.
- **Reużywalność** – klasy mogą być łatwo używane w różnych częściach programu.
- **Łatwiejsza konserwacja** – zmiany w jednej klasie automatycznie wpływają na wszystkie jej instancje.

---

## Przykład pełnej klasy

```vba
' Klasa: Produkt
Private m_Nazwa As String
Private m_Cena As Double

' Właściwość Nazwa
Public Property Let Nazwa(Value As String)
    m_Nazwa = Value
End Property

Public Property Get Nazwa() As String
    Nazwa = m_Nazwa
End Property

' Właściwość Cena
Public Property Let Cena(Value As Double)
    m_Cena = Value
End Property

Public Property Get Cena() As Double
    Cena = m_Cena
End Property

' Metoda obliczająca podatek VAT
Public Function CenaZVAT(VAT As Double) As Double
    CenaZVAT = m_Cena * (1 + VAT / 100)
End Function
```

---

## Podsumowanie

Tworzenie i używanie klas w VBA pozwala pisać bardziej czytelny, modularny i efektywny kod. Dzięki klasom można łatwiej zarządzać złożonymi projektami i wprowadzać zmiany w jednym miejscu, które automatycznie odzwierciedlają się w całym programie.
