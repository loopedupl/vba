# Obsługa Konstruktorów i Destruktorów

## Wstęp

Konstruktorzy i destruktory są specjalnymi mechanizmami w programowaniu obiektowym, które pozwalają na:

1. **Konstruktor**: Przygotowanie obiektu do użycia (inicjalizacja danych, konfiguracja początkowa).
2. **Destruktor**: Sprzątanie i zwalnianie zasobów (np. pamięci, plików, połączeń), gdy obiekt jest usuwany.

W VBA obsługujemy te mechanizmy za pomocą specjalnych metod.

---

## 1. **Konstruktor w VBA**

W VBA nie ma dedykowanego słowa kluczowego `constructor`. Zamiast tego używamy metody `Class_Initialize`, która jest automatycznie wywoływana podczas tworzenia nowej instancji klasy.

### Przykład: Inicjalizacja obiektu za pomocą konstruktora

```vba
Private pName As String

' Konstruktor - automatycznie wywoływany podczas inicjalizacji obiektu
Private Sub Class_Initialize()
    pName = "Domyślne imię"
    Debug.Print "Obiekt został zainicjalizowany."
End Sub

' Właściwość do zarządzania nazwą
Public Property Let Name(Value As String)
    pName = Value
End Property

Public Property Get Name() As String
    Name = pName
End Property
```

#### Użycie w module:

```vba
Dim person As New Person
Debug.Print person.Name ' Wyświetli "Domyślne imię"
```

---

## 2. **Destruktor w VBA**

Podobnie jak w przypadku konstruktora, VBA używa metody `Class_Terminate`, która jest automatycznie wywoływana, gdy obiekt jest usuwany lub kończy się jego zakres.

### Przykład: Zwalnianie zasobów za pomocą destruktora

```vba
Private Sub Class_Terminate()
    Debug.Print "Obiekt został usunięty z pamięci."
End Sub
```

#### Użycie w module:

```vba
Dim person As Person
Set person = New Person
Set person = Nothing ' Wywołuje destruktor
```

---

## 3. **Praktyczne zastosowania**

- **Konstruktor**:

  - Automatyczne wypełnianie właściwości wartościami domyślnymi.
  - Tworzenie połączenia z zewnętrzną bazą danych.
  - Inicjalizacja parametrów aplikacji.

- **Destruktor**:
  - Zamykanie plików i połączeń z bazami danych.
  - Zwolnienie pamięci zajmowanej przez duże struktury danych.

### Przykład: Praca z plikami w konstruktorze i destruktorze

```vba
Private filePath As String
Private fileHandle As Object

Private Sub Class_Initialize()
    filePath = "C:\dane.txt"
    Set fileHandle = CreateObject("Scripting.FileSystemObject").OpenTextFile(filePath, 1)
    Debug.Print "Plik otwarty."
End Sub

Private Sub Class_Terminate()
    If Not fileHandle Is Nothing Then fileHandle.Close
    Debug.Print "Plik zamknięty."
End Sub
```

---

## 4. **Zarządzanie pamięcią w VBA**

VBA automatycznie usuwa obiekty poza zakresem, ale możesz ręcznie wywołać destruktor za pomocą przypisania obiektu do `Nothing`.

### Przykład: Wywołanie destruktora

```vba
Dim obj As New MyClass
Set obj = Nothing ' Klasa_Terminate zostaje wywołana
```

---

## 5. **Podsumowanie**

Konstruktorzy i destruktory to kluczowe mechanizmy, które pozwalają na kontrolę procesu tworzenia i usuwania obiektów. Dzięki ich wykorzystaniu Twój kod będzie bardziej wydajny, czytelny i łatwiejszy w utrzymaniu.
