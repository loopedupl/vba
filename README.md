# Manipulacja arkuszami i plikami Excel

## Krok 1: Tworzenie i usuwanie arkuszy

W VBA możemy łatwo tworzyć nowe arkusze lub usuwać istniejące. Aby utworzyć nowy arkusz, używamy metody `Sheets.Add`. Aby usunąć arkusz, używamy metody `Sheets("NazwaArkusza").Delete`.

```vba
Sub TworzenieArkusza()
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Nowy Arkusz"
End Sub
```

W tym przykładzie dodajemy nowy arkusz na końcu arkuszy w skoroszycie.

Aby usunąć arkusz:

```vba
Sub UsuwanieArkusza()
    Sheets("Nowy Arkusz").Delete
End Sub
```

## Krok 2: Kopiowanie i przenoszenie arkuszy

Możesz skopiować lub przenieść arkusz z jednego skoroszytu do drugiego lub w obrębie tego samego skoroszytu.

Aby skopiować arkusz:

```vba
Sub KopiowanieArkusza()
    Sheets("Arkusz1").Copy After:=Sheets(Sheets.Count)
End Sub
```

Aby przenieść arkusz:

```vba
Sub PrzenoszenieArkusza()
    Sheets("Arkusz1").Move Before:=Sheets(1)
End Sub
```

## Krok 3: Praca z plikami Excel

Z poziomu VBA możesz otwierać, zapisywać i zamykać pliki Excel. Poniżej przedstawiamy sposób otwierania pliku:

```vba
Sub OtwieraniePliku()
    Workbooks.Open Filename:="C:\Ścieżka\Do\Pliku.xlsx"
End Sub
```

Aby zapisać plik:

```vba
Sub ZapisywaniePliku()
    ActiveWorkbook.SaveAs Filename:="C:\Ścieżka\Do\NowegoPliku.xlsx"
End Sub
```

Aby zamknąć plik:

```vba
Sub ZamknijPlik()
    ActiveWorkbook.Close
End Sub
```

## Krok 4: Zabezpieczanie arkuszy

Zabezpieczanie arkuszy pozwala kontrolować dostęp do danych w skoroszycie. Aby chronić arkusz przed edycją, używamy metody `Protect`.

```vba
Sub ZabezpieczArkusz()
    Sheets("Arkusz1").Protect Password:="haslo"
End Sub
```

Aby usunąć ochronę:

```vba
Sub UsuwanieOchrony()
    Sheets("Arkusz1").Unprotect Password:="haslo"
End Sub
```

## Krok 5: Automatyczne zamykanie i zapisywanie pliku

Możemy ustawić automatyczne zapisanie i zamknięcie pliku po wykonaniu określonych działań.

```vba
Sub ZapiszIZamknijPlik()
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub
```

## Krok 6: Praca z formatami plików

VBA pozwala na zapisywanie plików w różnych formatach, takich jak `.xlsx`, `.xlsm`, czy `.csv`. Oto jak zapisać plik w formacie `.xlsm`:

```vba
Sub ZapiszJakoXLSM()
    ActiveWorkbook.SaveAs Filename:="C:\Ścieżka\Do\Pliku.xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
End Sub
```

## Podsumowanie

Manipulacja arkuszami i plikami Excel to ważny element pracy z VBA. Dzięki tej lekcji nauczyłeś się, jak tworzyć i usuwać arkusze, kopiować je, zarządzać plikami Excel, zabezpieczać arkusze przed edycją oraz automatycznie zapisywać i zamykać pliki. Umiejętności te stanowią podstawę do bardziej zaawansowanych zadań automatyzujących procesy w Excelu.
