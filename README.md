# Pobieranie kontaktów do pliku CSV

## Cel lekcji:

Celem tej lekcji jest nauczenie się, jak eksportować kontakty z Outlooka do pliku CSV za pomocą VBA. Dzięki tej automatyzacji będziesz mógł łatwo pobrać wszystkie kontakty z książki adresowej Outlooka i zapisać je w formacie CSV, który jest powszechnie używany w wielu aplikacjach i systemach.

## Krok 1: Przygotowanie środowiska

Zanim rozpoczniesz, upewnij się, że masz dostęp do edytora VBA w Outlooku. Możesz to zrobić, przechodząc do zakładki "Developer" (jeśli nie jest widoczna, należy ją włączyć w ustawieniach), a następnie klikając "Visual Basic".

## Krok 2: Tworzenie makra do eksportowania kontaktów

W edytorze VBA stwórz nowe makro, które będzie eksportować wszystkie kontakty z książki adresowej Outlooka do pliku CSV. Poniżej znajduje się przykładowy kod VBA:

```vba
Sub ExportContactsToCSV()
    Const FILE_PATH = "wpisz ścieżkę do pliku"

    Dim oFSO As Object
    Dim oFile As Object
    Dim avHeaders As Variant, avContact As Variant
    Dim oNamespace As Outlook.NameSpace
    Dim oContacts As Outlook.MAPIFolder
    Dim oContact As Object

    Set oFSO = CreateObject("Scripting.FileSystemObject")

    Set oFile = oFSO.CreateTextFile(FILE_PATH, True)

    avHeaders = Array("File name", "Email address")
    oFile.WriteLine Join(avHeaders, ",")

    Set oNamespace = Application.GetNamespace("MAPI")

    Set oContacts = oNamespace.GetDefaultFolder(olFolderContacts)

    For Each oContact In oContacts.Items
        avContact = Array(oContact.FullName, oContact.Email1Address)
        oFile.WriteLine Join(avContact, ",")
    Next oContact

    oFile.Close

End Sub
```

## Krok 3: Uruchomienie makra

Po zapisaniu makra, wróć do Outlooka i uruchom makro. Kontakty z Twojej książki adresowej zostaną zapisane w pliku CSV w wybranej lokalizacji. Plik CSV będzie zawierał podstawowe informacje, takie jak pełne imię i nazwisko, adres e-mail oraz numer telefonu kontaktu.

## Krok 4: Otwieranie pliku CSV

Po zakończeniu procesu eksportu, otwórz plik CSV za pomocą aplikacji takich jak Microsoft Excel, Google Sheets lub dowolnego edytora tekstu, aby zobaczyć zapisane dane kontaktowe.

## Podsumowanie:

Dzięki temu makro możesz łatwo eksportować wszystkie kontakty z Outlooka do pliku CSV, co ułatwia ich dalszą obróbkę lub import do innych aplikacji. Jest to świetne narzędzie do tworzenia kopii zapasowych danych kontaktowych lub migracji do innych systemów.
