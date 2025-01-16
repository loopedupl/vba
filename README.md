# Tworzenie folderów i przenoszenie wiadomości

## Cel lekcji:

Celem tej lekcji jest nauczenie się, jak za pomocą VBA w Outlooku tworzyć foldery i przenosić wiadomości e-mail do wybranych folderów. Dzięki tej automatyzacji będziesz mógł uporządkować swoją skrzynkę odbiorczą, organizując wiadomości w odpowiednich folderach.

## Krok 1: Przygotowanie środowiska

Zanim rozpoczniesz, upewnij się, że masz dostęp do edytora VBA w Outlooku. Możesz to zrobić, przechodząc do zakładki "Developer" (jeśli nie jest widoczna, należy ją włączyć w ustawieniach), a następnie klikając "Visual Basic".

## Krok 2: Tworzenie folderu i przenoszenie wiadomości

W edytorze VBA stwórz nowe makro, które będzie tworzyć foldery w Outlooku. Poniżej znajduje się przykładowy kod VBA, który tworzy folder w skrzynce odbiorczej:

```vba
Sub OrganizeBySender()
    Dim oNamespace As Outlook.NameSpace
    Dim oInbox As Outlook.MAPIFolder
    Dim oItem As Object
    Dim oSenderFolder As Outlook.MAPIFolder
    Dim sSenderName As String
    Dim i As Long

    Set oNamespace = Application.GetNamespace("MAPI")
    Set oInbox = oNamespace.GetDefaultFolder(olFolderInbox)

    For i = oInbox.Items.Count To 1 Step -1
        Set oItem = oInbox.Items(i)

        sSenderName = oItem.SenderName

        On Error Resume Next
        Set oSenderFolder = oInbox.Folders(sSenderName)
        On Error GoTo 0

        If oSenderFolder Is Nothing Then
            Set oSenderFolder = oInbox.Folders.Add(sSenderName)
        End If
        oItem.Move oSenderFolder

        Set oSenderFolder = Nothing
    Next i
End Sub
```

## Krok 3: Uruchomienie makra

Po zapisaniu makra, wróć do Outlooka i uruchom makro. Aby przenieść wiadomość, wybierz ją w skrzynce odbiorczej, a następnie uruchom makro. Wybrana wiadomość zostanie przeniesiona do folderu, który wcześniej utworzyłeś.

## Podsumowanie:

Dzięki tym makrom możesz łatwo tworzyć foldery w Outlooku i przenosić wiadomości do odpowiednich folderów, co pomoże Ci w lepszej organizacji skrzynki odbiorczej. Automatyzacja tych procesów pozwala zaoszczędzić czas i poprawić efektywność zarządzania wiadomościami.
