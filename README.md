# Pobieranie wiadomości do pliku CSV

## Cel lekcji:

Celem tej lekcji jest nauczenie się, jak za pomocą VBA w Outlooku pobierać wiadomości e-mail i zapisywać je w pliku CSV. Dzięki tej automatyzacji będziesz mógł łatwo eksportować wiadomości z Outlooka, co ułatwi ich dalszą obróbkę lub archiwizowanie.

## Krok 1: Przygotowanie środowiska

Zanim rozpoczniesz, upewnij się, że masz dostęp do edytora VBA w Outlooku. Możesz to zrobić, przechodząc do zakładki "Developer" (jeśli nie jest widoczna, należy ją włączyć w ustawieniach), a następnie klikając "Visual Basic".

## Krok 2: Tworzenie makra do pobierania wiadomości

W edytorze VBA stwórz nowe makro, które będzie eksportować wybrane wiadomości z Outlooka do pliku CSV. Poniżej znajduje się przykładowy kod VBA:

```vba
Sub ExportMailsToCSV()
    Const FILE_PATH = "wpisz ścieżkę do pliku"

    Dim oFSO As Object
    Dim oFile As Object
    Dim avHeaders As Variant, avMail As Variant
    Dim oNamespace As Outlook.NameSpace
    Dim oInbox As Outlook.MAPIFolder
    Dim oItem As Object

    Set oFSO = CreateObject("Scripting.FileSystemObject")

    Set oFile = oFSO.CreateTextFile(FILE_PATH, True)

    avHeaders = Array("Sender Name", "Sender Email address", "Subject", "Recieved Date")
    oFile.WriteLine Join(avHeaders, ",")

    Set oNamespace = Application.GetNamespace("MAPI")

    Set oInbox = oNamespace.GetDefaultFolder(olFolderInbox)

    For Each oItem In oInbox.Items
        avMail = Array(oItem.SenderName, oItem.SenderEmailAddress, oItem.Subject, oItem.ReceivedTime)
        oFile.WriteLine Join(avMail, ",")
    Next oContact

    oFile.Close

End Sub
```

## Krok 3: Uruchomienie makra

Po zapisaniu makra, wróć do Outlooka i uruchom makro. Makro przejdzie przez wszystkie wiadomości w skrzynce odbiorczej i zapisze wybrane informacje (temat, nadawcę, datę otrzymania oraz treść wiadomości) do pliku CSV w wybranej lokalizacji.

## Krok 4: Sprawdzenie pliku CSV

Po zakończeniu procesu eksportu, przejdź do folderu, który ustawiłeś w makrze (np. `C:\Users\YourUsername\Documents\emails.csv`). Otwórz plik CSV za pomocą aplikacji takich jak Microsoft Excel, Google Sheets lub dowolnego edytora tekstu, aby zobaczyć zapisane dane.

## Podsumowanie:

Dzięki temu makro możesz łatwo eksportować wiadomości z Outlooka do pliku CSV, co pozwala na ich dalszą obróbkę, archiwizowanie lub import do innych aplikacji. Automatyzacja tego procesu pozwala zaoszczędzić czas i zwiększa efektywność zarządzania wiadomościami e-mail.
