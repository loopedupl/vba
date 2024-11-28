# Importowanie i eksportowanie danych (CSV, txt)

## Wstęp

Dane w formatach takich jak CSV (Comma-Separated Values) czy pliki tekstowe (txt) są powszechnie wykorzystywane do wymiany informacji między systemami. W tej lekcji dowiesz się, jak za pomocą VBA łatwo importować dane do arkusza Excel oraz eksportować je do tych formatów.

---

## 1. **Importowanie danych z plików CSV**

CSV to format, w którym dane są zapisane w postaci tekstowej, gdzie poszczególne wartości są oddzielone przecinkami (lub innymi separatorami).

### Przykład: Importowanie pliku CSV

Aby zaimportować dane z pliku CSV:

```vba
Sub ImportCSV()
    Dim ws As Worksheet
    Dim filePath As String

    Set ws = ThisWorkbook.Sheets("Dane")
    filePath = "C:\ścieżka\do\pliku.csv"

    With ws.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=ws.Range("A1"))
        .TextFileCommaDelimiter = True  ' Ustawienie przecinka jako separatora
        .TextFilePlatform = 65001  ' UTF-8
        .Refresh
    End With
End Sub
```

---

## 2. **Eksportowanie danych do pliku CSV**

Eksportowanie danych do pliku CSV pozwala na ich zapis w formacie, który można łatwo zaimportować do innych systemów.

### Przykład: Eksportowanie danych do CSV

```vba
Sub ExportCSV()
    Dim ws As Worksheet
    Dim filePath As String
    Dim dataRange As Range
    Dim outputFile As Integer
    Dim row As Range
    Dim cell As Range

    Set ws = ThisWorkbook.Sheets("Dane")
    filePath = "C:\ścieżka\do\eksport.csv"
    Set dataRange = ws.UsedRange

    outputFile = FreeFile
    Open filePath For Output As outputFile

    For Each row In dataRange.Rows
        Dim rowData As String
        rowData = ""
        For Each cell In row.Cells
            rowData = rowData & cell.Value & ","  ' Dodaj dane z komórki do wiersza
        Next cell
        rowData = Left(rowData, Len(rowData) - 1)  ' Usuń ostatni przecinek
        Print #outputFile, rowData
    Next row

    Close outputFile
End Sub
```

---

## 3. **Importowanie danych z plików tekstowych (txt)**

Pliki tekstowe często zawierają dane rozdzielone spacjami, tabulatorami lub innymi separatorami.

### Przykład: Importowanie pliku tekstowego

```vba
Sub ImportTXT()
    Dim filePath As String
    Dim fileContent As String
    Dim fileLine As String
    Dim rowIndex As Long

    filePath = "C:\ścieżka\do\pliku.txt"
    Open filePath For Input As #1

    rowIndex = 1
    Do Until EOF(1)
        Line Input #1, fileLine
        Cells(rowIndex, 1).Value = fileLine
        rowIndex = rowIndex + 1
    Loop

    Close #1
End Sub
```

---

## 4. **Eksportowanie danych do pliku tekstowego (txt)**

Eksportowanie do pliku tekstowego jest podobne do eksportowania do CSV, ale separator danych może być inny.

### Przykład: Eksportowanie danych do TXT

```vba
Sub ExportTXT()
    Dim ws As Worksheet
    Dim filePath As String
    Dim dataRange As Range
    Dim outputFile As Integer
    Dim row As Range
    Dim cell As Range

    Set ws = ThisWorkbook.Sheets("Dane")
    filePath = "C:\ścieżka\do\eksport.txt"
    Set dataRange = ws.UsedRange

    outputFile = FreeFile
    Open filePath For Output As outputFile

    For Each row In dataRange.Rows
        Dim rowData As String
        rowData = ""
        For Each cell In row.Cells
            rowData = rowData & cell.Value & vbTab  ' Separator tabulator
        Next cell
        rowData = Left(rowData, Len(rowData) - 1)  ' Usuń ostatni separator
        Print #outputFile, rowData
    Next row

    Close outputFile
End Sub
```

---

## 5. **Podsumowanie**

Dzięki VBA możesz łatwo automatyzować importowanie i eksportowanie danych w formatach CSV i tekstowych. To przydatna umiejętność, która pozwala na szybką wymianę danych między różnymi systemami i aplikacjami.
