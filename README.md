# Masowe odpowiedzi na wiadomości w Outlook VBA

## Wprowadzenie

W tej lekcji dowiesz się, jak stworzyć makro w VBA, które umożliwia wysyłanie masowych odpowiedzi na wybrane wiadomości w Outlooku. To rozwiązanie jest szczególnie przydatne, gdy musisz szybko odpowiedzieć na wiele podobnych e-maili.

---

## Tworzenie Makra

### 1. Otwórz edytor VBA w Outlooku

1. W Outlooku naciśnij **Alt + F11**, aby otworzyć edytor VBA.
2. Wybierz **ThisOutlookSession** z listy modułów.

### 2. Dodaj poniższy kod do modułu

```vba
Sub BulkReply()

    Const REPLY_TEXT As String = "Dziękuję za odpowiedź." & vbCrLf & "Odpowiem wkrótce." & vbCrLf & vbCrLf & "Pozdrawiam, Twój Zespół"

    Dim oSelection As Outlook.Selection
    Dim oSelectionItem As Object
    Dim oMail As Outlook.MailItem
    Dim oReply As Outlook.MailItem

    Set oSelection = Application.ActiveExplorer.Selection

    If oSelection Is Nothing Then
        MsgBox "Nie zaznaczono żadnego elementu", vbExclamation
        Exit Sub
    End If

    For Each oSelectionItem In oSelection
        If TypeOf oSelectionItem Is Outlook.MailItem Then
            Set oMail = oSelectionItem
            Set oReply = oMail.Reply
            With oReply
                .Body = REPLY_TEXT & vbCrLf & vbCrLf & oMail.Body
                .Display
                '.Send
            End With
        End If
    Next oSelectionItem

End Sub
```

### 3. Zapisz zmiany

1. Naciśnij **Ctrl + S**, aby zapisać kod.
2. Zamknij edytor VBA.

---

## Testowanie Makra

1. W Outlooku wybierz kilka wiadomości w swojej skrzynce odbiorczej.
2. Otwórz **Developer > Makra** lub naciśnij **Alt + F8**, wybierz `BulkReply` i kliknij **Uruchom**.
3. Sprawdź, czy odpowiedzi zostały wysłane do wybranych odbiorców.

---

## Podsumowanie

Makro do masowych odpowiedzi pozwala zaoszczędzić czas i zwiększyć efektywność w komunikacji e-mailowej. Dzięki temu rozwiązaniu możesz szybko reagować na wiele wiadomości jednocześnie, zachowując profesjonalizm i spójność odpowiedzi.
