# Przypomnienie o braku załącznika w Outlook VBA

## Wprowadzenie

W tej lekcji dowiesz się, jak stworzyć makro w VBA, które automatycznie przypomina o dodaniu załącznika, jeśli w treści wiadomości znajduje się wzmianka o załączniku, ale faktycznie go brakuje. Dzięki temu unikniesz sytuacji, w których zapominasz dołączyć ważne pliki do e-maila.

---

## Tworzenie Makra

### 1. Otwórz edytor VBA w Outlooku

1. W Outlooku naciśnij **Alt + F11**, aby otworzyć edytor VBA.
2. Wybierz **ThisOutlookSession** z listy modułów.

### 2. Dodaj poniższy kod do modułu

```vba
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    Dim oMail As Outlook.MailItem
    Dim avKeywords As Variant
    Dim bHasAttachment As Boolean, bFound As Boolean
    Dim i As Integer
    Dim sPrompt As String

    ' List of keywords
    avKeywords = Array("załącznik", "załączony", "dołączony", "plik", "dołączam", "wysyłam", "załączam", "plików", "przesyłam", "załączeniu")

    If Not TypeOf Item Is Outlook.MailItem Then Exit Sub

    Set oMail = Item
    bHasAttachment = oMail.Attachments.Count > 0

    If bHasAttachment Then Exit Sub

    ' Check the email body for keywords
    For i = LBound(avKeywords) To UBound(avKeywords)
        If InStr(LCase(oMail.Body), avKeywords(i)) > 0 Or InStr(LCase(oMail.Subject), avKeywords(i)) > 0 Then
            bFound = True
            Exit For
        End If
    Next i

    If bFound Then
        sPrompt = "Wydaje się, że wspomniałeś o załączniku, ale nie dodałeś żadnych plików. " & _
         "Czy nadal chcesz wysłać ten e-mail?"
        Cancel = (MsgBox(sPrompt, vbYesNo + vbExclamation, "Przypomnienie o załączniku") = vbNo)
    End If

End Sub

```

### 3. Zapisz zmiany

1. Naciśnij **Ctrl + S**, aby zapisać kod.
2. Zamknij edytor VBA.

---

## Testowanie Makra

1. Utwórz nową wiadomość w Outlooku.
2. W treści wiadomości wpisz słowo \"załącznik\" lub \"w załączeniu\".
3. Spróbuj wysłać wiadomość bez dodawania załącznika. Powinno pojawić się okno z przypomnieniem.

---

## Podsumowanie

To makro pomoże Ci uniknąć nieprofesjonalnych sytuacji związanych z wysyłaniem e-maili bez załączników, gdy wspominasz o nich w treści wiadomości. Automatyzując ten proces, oszczędzisz czas i zwiększysz swoją produktywność.
