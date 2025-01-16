# Pobieranie załączników

## Cel lekcji:

Celem tej lekcji jest nauczenie się, jak za pomocą VBA w Outlooku pobierać załączniki z wiadomości e-mail i zapisywać je na dysku. Dzięki tej automatyzacji będziesz mógł szybko i efektywnie pobierać pliki z wiadomości e-mail bez konieczności ręcznego ich pobierania.

## Krok 1: Przygotowanie środowiska

Zanim rozpoczniesz, upewnij się, że masz dostęp do edytora VBA w Outlooku. Możesz to zrobić, przechodząc do zakładki "Developer" (jeśli nie jest widoczna, należy ją włączyć w ustawieniach), a następnie klikając "Visual Basic".

## Krok 2: Tworzenie makra do pobierania załączników

W edytorze VBA stwórz nowe makro, które będzie pobierać załączniki z wiadomości e-mail i zapisywać je na dysku. Poniżej znajduje się przykładowy kod VBA:

```vba
Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)

    Const FOLDER_PATH As String = "ścieżka do folderu"

    Dim avEntryIds() As String
    Dim oNamespace As Outlook.NameSpace
    Dim i As Integer
    Dim oItem As Object
    Dim oMail As Outlook.MailItem
    Dim oAttachments As Outlook.Attachments
    Dim oAttachment As Outlook.attachment

    avEntryIds = Split(EntryIDCollection, ",")

    Set oNamespace = Application.GetNamespace("MAPI")

    For i = LBound(avEntryIds) To UBound(avEntryIds)
        Set oItem = oNamespace.GetItemFromID(avEntryIds(i))

        If TypeOf oItem Is Outlook.MailItem Then
            Set oMail = oItem
            Set oAttachments = oMail.Attachments
            If oAttachments.Count > 0 Then
                For Each oAttachment In oAttachments
                    oAttachment.SaveAsFile FOLDER_PATH & oAttachment.FileName
                Next oAttachment
            End If
        End If
    Next i

End Sub
```

## Krok 3: Uruchomienie makra

Po zapisaniu makra, wróć do Outlooka i uruchom makro. Wybierz wiadomość e-mail, która zawiera załączniki, a następnie uruchom makro. Załączniki zostaną zapisane w wybranym folderze na Twoim dysku.

## Krok 4: Sprawdzenie zapisanych załączników

Po zakończeniu procesu pobierania załączników, przejdź do folderu, który ustawiłeś w makrze (np. `C:\Users\YourUsername\Documents\Attachments\`). Tam znajdziesz zapisane pliki.

## Podsumowanie:

Dzięki temu makro możesz łatwo pobierać załączniki z wiadomości e-mail w Outlooku i zapisywać je na dysku. Automatyzacja tego procesu pozwala zaoszczędzić czas i poprawić efektywność pracy z załącznikami.
