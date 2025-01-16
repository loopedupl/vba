# Tworzenie kontaktów na podstawie wiadomości

## Cel lekcji:

Celem tej lekcji jest nauczenie się, jak automatycznie tworzyć kontakty w Outlooku na podstawie treści wiadomości e-mail, wykorzystując VBA. Dzięki tej automatyzacji, zaoszczędzisz czas na ręcznym dodawaniu nowych kontaktów, szczególnie gdy wiadomości zawierają dane kontaktowe, takie jak adresy e-mail, numery telefonów czy imiona i nazwiska.

## Krok 1: Przygotowanie środowiska

Zanim zaczniesz pisać kod VBA, upewnij się, że masz dostęp do edytora VBA w Outlooku. Aby otworzyć edytor VBA, przejdź do zakładki "Developer" (jeśli nie jest widoczna, należy ją włączyć w ustawieniach), a następnie kliknij "Visual Basic".

## Krok 2: Tworzenie makra do tworzenia kontaktów

W edytorze VBA stwórz nowe makro, które będzie odczytywać treść wybranej wiadomości e-mail i na jej podstawie tworzyć nowy kontakt w Outlooku. Poniżej znajduje się przykładowy kod VBA:

```vba
Sub AddSendersToContacts()
    Dim oNamespace As Outlook.NameSpace
    Dim oInbox As Outlook.MAPIFolder
    Dim oItem As Object
    Dim oContact As Outlook.ContactItem

    Set oNamespace = Application.GetNamespace("MAPI")
    Set oInbox = oNamespace.GetDefaultFolder(olFolderInbox)

    For Each oItem In oInbox.Items
        Set oContact = Application.CreateItem(olContactItem)
        With oContact
            .FullName = oItem.SenderName
            .Email1Address = oItem.SenderEmailAddress
            .Save
        End With
    Next oItem
End Sub
```

## Krok 3: Uruchomienie makra

Po zapisaniu makra, wróć do Outlooka i uruchom makro. Wybierz wiadomość e-mail, z której chcesz stworzyć kontakt, a następnie uruchom makro. Nowy kontakt zostanie automatycznie utworzony w Twojej książce adresowej w Outlooku na podstawie danych z wiadomości.

## Podsumowanie:

Dzięki temu makro możesz zaoszczędzić czas, automatycznie tworząc kontakty w Outlooku na podstawie wiadomości e-mail. Zastosowanie VBA w Outlooku pozwala na łatwą organizację danych kontaktowych i szybkie dodawanie nowych osób do książki adresowej.
