# Zaawansowane formuły i analiza danych w VBA

## Wstęp

Excel oferuje szerokie możliwości analizy danych, a VBA pozwala je jeszcze bardziej rozszerzyć. W tej lekcji nauczysz się tworzyć niestandardowe funkcje użytkownika (UDFs), zautomatyzować złożone kalkulacje finansowe i statystyczne oraz analizować trendy i prognozować przyszłe wartości.

---

## 1. **Tworzenie niestandardowych funkcji użytkownika (UDFs)**

Funkcje użytkownika pozwalają na tworzenie własnych formuł, które można wykorzystywać tak samo jak standardowe funkcje Excela.

### Przykład: Tworzenie funkcji obliczającej procentową zmianę wartości

```vba
Function PercentChange(oldValue As Double, newValue As Double) As Double
    If oldValue = 0 Then
        PercentChange = CVErr(xlErrDiv0)
    Else
        PercentChange = (newValue - oldValue) / oldValue
    End If
End Function
```

W arkuszu Excel wpisz w komórce:  
`=PercentChange(A1, B1)` – obliczy procentową zmianę między wartością w komórkach A1 i B1.

---

## 2. **Automatyzacja złożonych kalkulacji finansowych i statystycznych**

Za pomocą VBA możesz automatyzować złożone analizy finansowe i statystyczne.

### Przykład: Obliczanie wartości przyszłej inwestycji (Future Value)

```vba
Function FutureValue(rate As Double, periods As Integer, payment As Double, presentValue As Double) As Double
    FutureValue = presentValue * (1 + rate) ^ periods + payment * ((1 + rate) ^ periods - 1) / rate
End Function
```

W arkuszu Excel użyj:  
`=FutureValue(0.05, 10, 100, 1000)` – obliczy wartość przyszłą inwestycji z oprocentowaniem 5%, na 10 okresów, przy wpłatach 100 i wartości początkowej 1000.

---

## 3. **Analiza trendów i prognozowanie z wykorzystaniem Excela i VBA**

Analiza trendów i prognozowanie pozwala przewidywać przyszłe wartości na podstawie danych historycznych.

### Przykład: Prognozowanie przy użyciu średniej ruchomej

```vba
Function MovingAverage(dataRange As Range, period As Integer) As Double
    Dim total As Double
    Dim i As Integer

    total = 0
    For i = 1 To period
        total = total + dataRange.Cells(i, 1).Value
    Next i

    MovingAverage = total / period
End Function
```

W arkuszu Excel:  
`=MovingAverage(A1:A10, 5)` – obliczy średnią ruchomą dla ostatnich 5 okresów w zakresie A1:A10.

---

## 4. **Praktyczne zastosowania**

- Tworzenie kalkulatora rat kredytowych za pomocą VBA.
- Automatyzacja analizy wyników finansowych firmy.
- Prognozowanie sprzedaży na podstawie trendów historycznych.

---

## 5. **Podsumowanie**

Tworzenie niestandardowych funkcji i analiza danych w VBA otwiera szerokie możliwości automatyzacji i zaawansowanej analizy. Dzięki tej wiedzy możesz efektywnie przetwarzać duże zbiory danych, realizować kalkulacje finansowe oraz podejmować świadome decyzje na podstawie prognoz.
