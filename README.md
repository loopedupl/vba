# Właściwości i Metody Klasy

## Wstęp

Klasy w VBA umożliwiają organizację kodu w logiczne jednostki, które mają swoje właściwości (atrybuty) i metody (działania). W tej lekcji nauczysz się, jak definiować właściwości i metody w klasach oraz jak kontrolować ich widoczność.

---

## 1. **Definiowanie właściwości w klasach**

Właściwości to atrybuty obiektu. Możesz je tworzyć za pomocą trzech specjalnych mechanizmów:

- **Property Get** – do odczytu wartości.
- **Property Let** – do ustawiania wartości typu prostego.
- **Property Set** – do ustawiania wartości typu obiektowego.

### Przykład: Definiowanie właściwości w klasie

```vba
' Klasa o nazwie "Person"
Private pName As String

' Właściwość do odczytu imienia
Public Property Get Name() As String
    Name = pName
End Property

' Właściwość do ustawiania imienia
Public Property Let Name(Value As String)
    pName = Value
End Property
```

#### Użycie w module:

```vba
Dim person As New Person
person.Name = "Jan Kowalski"
Debug.Print person.Name ' Wyświetli "Jan Kowalski"
```

---

## 2. **Tworzenie i wywoływanie metod**

Metody to działania, które obiekt może wykonywać. Możesz je definiować w klasach jako `Public Sub` lub `Function`.

### Przykład: Tworzenie metody w klasie

```vba
Public Sub Introduce()
    Debug.Print "Nazywam się " & pName
End Sub
```

#### Użycie w module:

```vba
person.Introduce ' Wyświetli "Nazywam się Jan Kowalski"
```

---

## 3. **Zarządzanie widocznością w klasach**

W VBA możesz kontrolować dostęp do elementów klasy, korzystając z modyfikatorów:

- **Public** – element widoczny dla całego projektu.
- **Private** – element dostępny tylko w obrębie danej klasy.

### Przykład: Widoczność właściwości i metod

```vba
Private pAge As Integer

' Właściwość prywatna
Private Property Get Age() As Integer
    Age = pAge
End Property

' Metoda publiczna
Public Sub DisplayAge()
    Debug.Print "Wiek: " & pAge
End Sub
```

#### Użycie w module:

```vba
Dim person As New Person
' Nie można użyć person.Age bezpośrednio (błąd kompilacji)
person.DisplayAge ' Wyświetli wiek
```

---

## 4. **Praktyczne zastosowania**

- Tworzenie klas reprezentujących dane biznesowe, np. produkty, zamówienia, pracowników.
- Budowanie obiektowych struktur dla złożonych raportów i analiz.
- Dynamiczne zarządzanie danymi za pomocą właściwości i metod.

---

## 5. **Podsumowanie**

Właściwości i metody to podstawowe elementy klas, które umożliwiają elastyczne zarządzanie danymi i logiką w VBA. Dzięki odpowiedniemu zarządzaniu widocznością możesz tworzyć bezpieczne i łatwe w utrzymaniu aplikacje.
