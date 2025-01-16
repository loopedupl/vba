# Szukanie wydarzeń na podstawie słowa kluczowego

## Cel lekcji:

Celem tej lekcji jest nauczenie się, jak za pomocą VBA w Outlooku wyszukiwać wydarzenia (spotkania) w kalendarzu na podstawie podanego słowa kluczowego. Dzięki tej automatyzacji możesz szybko znaleźć interesujące Cię spotkania i wydarzenia w Outlooku, oszczędzając czas na ręcznym przeszukiwaniu kalendarza.

## Krok 1: Przygotowanie środowiska

Zanim rozpoczniesz, upewnij się, że masz dostęp do edytora VBA w Outlooku. Możesz to zrobić, przechodząc do zakładki "Developer" (jeśli nie jest widoczna, należy ją włączyć w ustawieniach), a następnie klikając "Visual Basic".

## Krok 2: Tworzenie makra do wyszukiwania wydarzeń

W edytorze VBA stwórz nowe makro, które będzie przeszukiwać kalendarz Outlooka pod kątem wydarzeń zawierających określone słowo kluczowe. Poniżej znajduje się przykładowy kod VBA:

```vba
Sub FindEvents()
    Dim oNamespace As Outlook.NameSpace
    Dim oCalendar As Outlook.MAPIFolder
    Dim oItems As Outlook.Items
    Dim oItem As Object
    Dim sSearchKeyword As String
    Dim bEventFound As Boolean

    Set oNamespace = Application.GetNamespace("MAPI")
    Set oCalendar = oNamespace.GetDefaultFolder(olFolderCalendar)
    Set oItems = oCalendar.Items

    sSearchKeyword = InputBox("Wprowadź słowo kluczowe do wyszukania w wydarzeniach:", "Szukaj wydarzeń")

    If sSearchKeyword = "" Then
        MsgBox "Nie wprowadzono słowa kluczowego. Operacja anulowana.", vbExclamation
        Exit Sub
    End If

    bEventFound = False
    For Each oItem In oItems
        With oItem
            If InStr(LCase(.Subject), LCase(sSearchKeyword)) > 0 Or InStr(LCase(.Body), LCase(sSearchKeyword)) > 0 Then
                .Display
                bEventFound = True
            End If
        End With
    Next oItem

    If Not bEventFound Then
        MsgBox "Nie znaleziono wydarzeń zawierających słowo kluczowe '" & sSearchKeyword & "'.", vbInformation
    End If

End Sub
```

## Krok 3: Uruchomienie makra

Po zapisaniu makra, wróć do Outlooka i uruchom makro. Zostaniesz poproszony o wprowadzenie słowa kluczowego, które chcesz wyszukiwać w wydarzeniach kalendarza. Program przeszuka wydarzenia w kalendarzu w zadanym zakresie dat (np. w bieżącym miesiącu) i wyświetli wszystkie, które zawierają to słowo kluczowe w temacie lub treści.

## Podsumowanie:

Dzięki temu makro możesz szybko i łatwo przeszukiwać kalendarz Outlooka pod kątem określonych wydarzeń, oszczędzając czas na ręcznym wyszukiwaniu. Jest to świetne narzędzie do organizacji i efektywnego zarządzania spotkaniami w Outlooku.
