# Integracja raportów z zewnętrznymi źródłami danych

## Krok 1: Tworzenie połączenia z bazą danych

Aby pobrać dane z bazy danych, należy użyć obiektu `ADODB.Connection`. W tym przykładzie łączymy się z bazą danych Access:

```vba
Sub PolaczenieZBazaDanych()
    Dim Conn As Object
    Dim Rs As Object
    Dim Query As String

    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Sciezka\Do\Bazy.accdb"

    Query = "SELECT * FROM Produkty"

    Set Rs = CreateObject("ADODB.Recordset")
    Rs.Open Query, Conn

    ' Przenosimy dane do arkusza
    Sheets("Raport").Range("A2").CopyFromRecordset Rs

    Rs.Close
    Conn.Close
End Sub
```

## Krok 2: Pobieranie danych z pliku CSV

Excel pozwala na importowanie danych z plików CSV bezpośrednio do arkusza za pomocą VBA. Oto przykład:

```vba
Sub ImportZPlikuCSV()
    Dim Sciezka As String
    Sciezka = "C:\Sciezka\Do\Pliku.csv"

    ' Importowanie danych z pliku CSV do aktywnego arkusza
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & Sciezka, Destination:=Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileSemicolonDelimiter = True
        .Refresh BackgroundQuery:=False
    End With
End Sub
```

## Krok 3: Pobieranie danych z API za pomocą VBA

Dzięki VBA możesz pobierać dane z zewnętrznych API. Poniżej znajduje się przykład pobierania danych z API przy użyciu `MSXML2.XMLHTTP`:

```vba
Sub PobieranieDanychZAPI()
    Dim Http As Object
    Dim URL As String
    Dim Odpowiedz As String

    Set Http = CreateObject("MSXML2.XMLHTTP")
    URL = "https://api.exemplo.com/dane"

    Http.Open "GET", URL, False
    Http.Send

    Odpowiedz = Http.responseText

    ' Przenosimy odpowiedź do komórki
    Sheets("Raport").Range("A1").Value = Odpowiedz
End Sub
```

## Krok 4: Aktualizowanie raportu na podstawie pobranych danych

Po zaimportowaniu danych z zewnętrznych źródeł, możesz wykorzystać je do uaktualnienia raportu w Excelu. Poniższy kod dodaje dane do tabeli w arkuszu:

```vba
Sub AktualizacjaRaportuZObcegoZrodla()
    Dim Arkusz As Worksheet
    Set Arkusz = ThisWorkbook.Sheets("Raport")

    ' Zakładając, że dane zostały już zaimportowane do komórek A2:B10
    ' Możemy teraz dodać je do tabeli w arkuszu

    Arkusz.Range("A2:B10").Sort Key1:=Arkusz.Range("A2"), Order1:=xlAscending, Header:=xlYes
End Sub
```

## Krok 5: Użycie Power Query do integracji z danymi zewnętrznymi

Możesz także używać Power Query w połączeniu z VBA, aby pobierać i przetwarzać dane z zewnętrznych źródeł. Poniżej znajdziesz przykład uruchomienia Power Query za pomocą VBA:

```vba
Sub UruchomPowerQuery()
    Dim PQ As WorkbookQuery

    ' Uruchomienie Power Query o nazwie "DaneZAPI"
    Set PQ = ThisWorkbook.Queries("DaneZAPI")
    PQ.Refresh BackgroundQuery:=False
End Sub
```

## Krok 6: Automatyczne generowanie raportu po integracji z danymi

Po zintegrowaniu zewnętrznych źródeł danych, możesz zautomatyzować proces generowania raportów, które będą zawsze aktualne. Oto przykład:

```vba
Sub GenerowanieRaportuPoIntegracji()
    ' Pobieranie danych z API
    Call PobieranieDanychZAPI

    ' Przetwarzanie danych w arkuszu
    Call AktualizacjaRaportuZObcegoZrodla

    ' Zapisanie raportu do pliku PDF
    Call ZapiszRaportPDF
End Sub
```

---

## Podsumowanie

Integracja raportów z zewnętrznymi źródłami danych za pomocą VBA to technika, która pozwala na automatyzację zbierania i przetwarzania danych z różnych systemów. Dzięki temu możesz tworzyć raporty, które zawsze będą zawierały aktualne informacje, bez potrzeby ręcznego wprowadzania danych.
