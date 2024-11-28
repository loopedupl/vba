# Automatyczne generowanie raportów

## Krok 1: Tworzenie szablonu raportu

Zanim zaczniesz generować raporty, dobrze jest mieć ustalony szablon, który będzie uaktualniany. Przykład szablonu:

```vba
Sub SzablonRaport()
    Range("A1").Value = "Data"
    Range("B1").Value = "Sprzedaż"
    Range("A2").Value = Date
    Range("B2").Value = Application.WorksheetFunction.Sum(Range("C2:C10"))
End Sub
```

## Krok 2: Automatyczne zbieranie danych

Aby raport był zawsze aktualny, można zbierać dane z różnych zakresów. Przykład:

```vba
Sub PobierzDaneZInnegoArkusza()
    Dim Zrodlo As Worksheet
    Set Zrodlo = Worksheets("Dane")

    Range("A2").Value = Zrodlo.Range("B2").Value
    Range("B2").Value = Zrodlo.Range("C2").Value
End Sub
```

## Krok 3: Automatyczne zapisywanie raportu

Zapisz raport do pliku PDF lub Excel. Możesz ustawić dynamiczną nazwę pliku na podstawie daty:

```vba
Sub ZapiszRaport()
    Dim Sciezka As String
    Sciezka = ThisWorkbook.Path & "\Raport_" & Format(Date, "YYYY-MM-DD") & ".xlsx"
    ActiveWorkbook.SaveAs Filename:=Sciezka
End Sub
```

## Krok 4: Planowanie generowania raportów

Za pomocą VBA możemy również ustawić automatyczne generowanie raportów o określonych porach. W tym celu używamy obiektu `Application.OnTime`:

```vba
Sub GenerowanieRaportowNaZywo()
    Application.OnTime Now + TimeValue("01:00:00"), "ZapiszRaport"
End Sub
```

## Krok 5: Generowanie raportu na podstawie wyników filtrów

Raport może być generowany tylko z danych, które spełniają określone kryteria. Możesz użyć funkcji filtrujących w VBA, aby zebrać odpowiednie dane do raportu:

```vba
Sub GenerowanieRaportuwZaawansowanymFiltrem()
    ActiveSheet.Range("A1:B10").AutoFilter Field:=1, Criteria1:=">=1000"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:="Raport_Filtr.pdf"
End Sub
```

## Krok 6: Automatyczne generowanie raportu i wysyłanie mailem

Możesz użyć VBA do automatycznego generowania raportu i wysyłania go mailem:

```vba
Sub WyslijRaportMail()
    Dim OutlookApp As Object
    Dim OutlookMail As Object

    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)

    OutlookMail.Subject = "Automatyczny Raport"
    OutlookMail.Body = "W załączniku znajduje się automatycznie wygenerowany raport."
    OutlookMail.Attachments.Add "C:\Raport.pdf"
    OutlookMail.Send
End Sub
```

## Podsumowanie

Automatyczne generowanie raportów w Excelu z użyciem VBA pozwala na efektywne tworzenie cyklicznych raportów, które mogą być automatycznie eksportowane do różnych formatów i przesyłane na określone adresy e-mail. To oszczędza czas i eliminuje błędy związane z manualnym tworzeniem raportów.
