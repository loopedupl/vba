# Stwórz zaproszenie na spotkanie na podstawie wiadomości

## Cel lekcji:

Celem tej lekcji jest nauczenie się, jak stworzyć zaproszenie na spotkanie w Outlooku na podstawie treści wiadomości e-mail, wykorzystując VBA. Dzięki tej automatyzacji zaoszczędzisz czas, tworząc spotkania bez konieczności ręcznego kopiowania informacji z wiadomości.

## Krok 1: Przygotowanie środowiska

Zanim zaczniesz pisać kod VBA, upewnij się, że masz dostęp do edytora VBA w Outlook. Aby otworzyć edytor VBA, przejdź do zakładki "Developer" (jeśli nie jest widoczna, należy ją włączyć w ustawieniach), a następnie kliknij "Visual Basic".

## Krok 2: Tworzenie makra

W edytorze VBA stwórz nowe makro, które będzie odczytywać treść wybranej wiadomości e-mail i na jej podstawie tworzyć zaproszenie na spotkanie. Poniżej znajduje się przykładowy kod VBA:

```vba
Sub ConvertMailToInvite()

    Dim oSelection As Outlook.Selection
    Dim oSelectionItem As Object

    Set oSelection = Application.ActiveExplorer.Selection
    If oSelection Is Nothing Then
        MsgBox "Nie zaznaczono żadnego elementu.", vbExclamation, "Brak zaznaczenia"
        Exit Sub
    End If

    For Each oSelectionItem In oSelection
        If TypeOf oSelectionItem Is Outlook.MailItem Then
            MailToInvite oSelectionItem
        End If
    Next oSelectionItem

End Sub
```

```vba
Private Sub MailToInvite(p_oMail As Outlook.MailItem)

    Dim sMe As String, sSender As String, sTo As String, sCc As String
    Dim sRequired As String, sOptional As String
    Dim oInvite As Outlook.AppointmentItem

    sMe = Application.Session.CurrentUser.AddressEntry
    sSender = p_oMail.SenderEmailAddress
    sTo = p_oMail.To
    sCc = p_oMail.CC

    If Not sSender = sMe Then
        sRequired = sTo & ";" & sSender & ";"
    Else
        sRequired = sTo & ";"
    End If
    sOptional = sCc & ";"

    sRequired = Replace(sRequired, sMe & ";", "")
    sOptional = Replace(sOptional, sMe & ";", "")

    Set oInvite = Application.CreateItem(olAppointmentItem)

    With oInvite
        .Subject = p_oMail.Subject
        .RequiredAttendees = sRequired
        .OptionalAttendees = sOptional
        .Body = p_oMail.Body
        .Duration = 60
        .Start = Now() + 1
        .MeetingStatus = olMeeting
        .Display
        '.Send
    End With

End Sub
```

## Krok 3: Uruchomienie makra

Po zapisaniu makra, wróć do Outlooka i uruchom makro, aby stworzyć zaproszenie na spotkanie. Wystarczy, że wybierzesz wiadomość e-mail, a następnie uruchomisz makro. Zaproszenie na spotkanie zostanie automatycznie wygenerowane z danymi z wiadomości.

## Podsumowanie:

Dzięki temu makro możesz automatycznie tworzyć zaproszenia na spotkania w Outlooku na podstawie treści wiadomości e-mail. Zastosowanie VBA w Outlooku pozwala na znaczną automatyzację codziennych zadań, co oszczędza czas i zwiększa efektywność pracy.
