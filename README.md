# Oznaczanie wiadomości jako ważna w Outlook VBA

## Wprowadzenie

W tej lekcji nauczysz się, jak automatycznie oznaczać wiadomości e-mail jako ważne w Outlooku przy użyciu VBA. Dzięki temu łatwo wyróżnisz kluczowe wiadomości i szybciej na nie zareagujesz.

---

## Tworzenie Makra

### 1. Otwórz edytor VBA w Outlooku

1. W Outlooku naciśnij **Alt + F11**, aby otworzyć edytor VBA.
2. Wybierz **ThisOutlookSession** z listy modułów.

### 2. Dodaj poniższy kod do modułu

```vba
Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)

    Dim asEntryIds() As String
    Dim avEmails As Variant
    Dim i As Integer, j As Integer
    Dim oNamespace As Outlook.NameSpace
    Dim oItem As Object

    asEntryIds = Split(EntryIDCollection, ",")

    avEmails = Array("adress@example.com")

    Set oNamespace = Application.GetNamespace("MAPI")

    For i = LBound(asEntryIds) To UBound(asEntryIds)
        Set oItem = oNamespace.GetItemFromID(asEntryIds(i))
        For j = LBound(avEmails) To UBound(avEmails)
            If oItem.SenderEmailAddress = avEmails(j) Then
                oItem.Importance = olImportanceHigh
                oItem.Save
                Exit For
            End If
        Next j
    Next i

End Sub

```

### 3. Zapisz zmiany

1. Naciśnij **Ctrl + S**, aby zapisać kod.
2. Zamknij edytor VBA.

---

## Testowanie Makra

1. Wyślij do siebie wiadomość e-mail z frazami \"pilne\" lub \"ważne\" w temacie.
2. Sprawdź, czy wiadomość została oznaczona jako ważna w skrzynce odbiorczej.

---

## Podsumowanie

To makro pozwala automatycznie wyróżniać kluczowe wiadomości w Twojej skrzynce odbiorczej. Dzięki temu możesz lepiej zarządzać priorytetami i szybciej reagować na istotne sprawy.
