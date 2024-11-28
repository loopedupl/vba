# Praca z tabelami przestawnymi w raportach

## Krok 1: Tworzenie tabeli przestawnej z danych w arkuszu

Za pomocą VBA możemy stworzyć tabelę przestawną na podstawie danych z arkusza Excel. Oto jak wygląda kod, który tworzy prostą tabelę przestawną:

```vba
Sub TworzenieTabeliPrzestawnej()
    Dim ZakresDanych As Range
    Dim TabelaPrzestawna As PivotTable
    Dim Arkusz As Worksheet
    Set Arkusz = ThisWorkbook.Sheets("Dane")

    Set ZakresDanych = Arkusz.Range("A1:D100") ' Zakres danych

    ' Tworzenie tabeli przestawnej
    Set TabelaPrzestawna = ThisWorkbook.PivotTableWizard(SourceType:=xlDatabase, SourceData:=ZakresDanych)
    TabelaPrzestawna.RefreshTable
End Sub
```

## Krok 2: Modyfikacja tabeli przestawnej

Możesz również manipulować tabelą przestawną, dodając nowe elementy do wierszy, kolumn, wartości oraz filtrów. Przykład:

```vba
Sub ModyfikacjaTabeliPrzestawnej()
    Dim TabelaPrzestawna As PivotTable
    Set TabelaPrzestawna = ThisWorkbook.Sheets("Raport").PivotTables("Tabela1")

    ' Dodanie wiersza
    TabelaPrzestawna.PivotFields("Kategoria").Orientation = xlRowField
    TabelaPrzestawna.PivotFields("Kategoria").Position = 1

    ' Dodanie wartości
    TabelaPrzestawna.PivotFields("Sprzedaż").Orientation = xlDataField
    TabelaPrzestawna.PivotFields("Sprzedaż").Function = xlSum
End Sub
```

## Krok 3: Aktualizacja tabeli przestawnej

Po dokonaniu zmian w danych źródłowych należy odświeżyć tabelę przestawną, aby wyświetlała najnowsze dane. Oto przykład:

```vba
Sub AktualizacjaTabeliPrzestawnej()
    Dim TabelaPrzestawna As PivotTable
    Set TabelaPrzestawna = ThisWorkbook.Sheets("Raport").PivotTables("Tabela1")
    TabelaPrzestawna.RefreshTable
End Sub
```

## Krok 4: Formatowanie tabeli przestawnej

Możesz również dostosować wygląd tabeli przestawnej, np. ustawić format liczb, zmienić style, dodać nagłówki. Przykład formatowania:

```vba
Sub FormatowanieTabeliPrzestawnej()
    Dim TabelaPrzestawna As PivotTable
    Set TabelaPrzestawna = ThisWorkbook.Sheets("Raport").PivotTables("Tabela1")

    ' Zmiana formatu liczb
    TabelaPrzestawna.PivotFields("Sprzedaż").NumberFormat = "#,##0.00"

    ' Zastosowanie stylu
    TabelaPrzestawna.TableStyle2 = "PivotStyleLight16"
End Sub
```

## Krok 5: Tworzenie tabeli przestawnej na podstawie filtrów

Tabele przestawne mogą być również wykorzystywane do filtrowania danych. Dzięki VBA możemy dynamicznie ustawić filtry na tabeli:

```vba
Sub DodanieFiltraDoTabeliPrzestawnej()
    Dim TabelaPrzestawna As PivotTable
    Set TabelaPrzestawna = ThisWorkbook.Sheets("Raport").PivotTables("Tabela1")

    ' Dodanie filtra
    TabelaPrzestawna.PivotFields("Data").Orientation = xlPageField
    TabelaPrzestawna.PivotFields("Data").Position = 1
    TabelaPrzestawna.PivotFields("Data").CurrentPage = "2023"
End Sub
```

## Krok 6: Zapisywanie raportu z tabelą przestawną do PDF

Po utworzeniu tabeli przestawnej, możemy zapisać cały raport do pliku PDF, co ułatwia dystrybucję:

```vba
Sub ZapiszRaportPDF()
    Dim Sciezka As String
    Sciezka = ThisWorkbook.Path & "\Raport_Tabela_Przeswtna_" & Format(Date, "YYYY-MM-DD") & ".pdf"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=Sciezka
End Sub
```

## Podsumowanie

Tabela przestawna to jeden z najbardziej użytecznych elementów w Excelu, który pozwala na szybkie i dynamiczne podsumowanie danych. Dzięki VBA możesz w pełni zautomatyzować proces tworzenia, aktualizowania oraz formatowania tabel przestawnych w swoich raportach.
