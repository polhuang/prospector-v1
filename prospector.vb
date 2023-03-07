Option Explicit


Sub ProspectNow()
Dim Outlook As Object
Set Outlook = CreateObject("Outlook.Application")
Dim OutlookMessage As Object

Dim DataRange As Range
Set DataRange = Range("D2", Range("d2").End(xlDown))
Dim Entry As Range
For Each Entry In Datarange
    Set OutlookMessage = Outlook.CreateItem(0)
        With OutlookMessage
            .To = Entry.Offset(0, 3) 
            'First name is at [0], Last name at [1], Title at [2], and e-mail at [3]. Template below only uses [0] and [3].
            .Subject = "Hi from Yadda Yadda" 
            .CC = "yadda@yadda.com" 
            .HTMLBody = "Hi " & Entry & ",<html> <br><br>" _
            & "I wanted to introduce myself as your yadda yadda. " & R.Offset(0, -1) & " yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda-yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda." _ 
            & "Yadda yadda yadda yadda yadda yadda yadda yadda yadda, yadda yadda yadda yadda yadda yadda, yadda yadda yadda yadda yadda, yadda yadda." _
            & "<br>Yadda yadda yadda, yadda yadda. Yadda yadda yadda yaada. Yadda, yadda yadda yadda yadda yadda yadda, yadda yadda yadda yadda yadda: Yadda, Yadda, Yadda<br>" _
            & "<br><b>Yadda Yadda /b><br>- Yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda. Yadda yadda yadda." _
            & "<br>- Yadda yadda yadda yadda Yadda, Yadda, Yadda Yadda, Yadda, yadda Yadda-Yadda Yadda<br><br><b>Yadda Yadda</b>" _
            & "<br>- Yadda yadda, yadda yadda-yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda/yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda/yadda" _
            & "<br>- Yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda<br>- Yadda yadda yadda-yadda yadda<br><br><b>Yadda-yadda Yadda</b><br>- Yadda yadda yadda yadda yadda yadda yadda yadda yadda yadda" _
            & "<br>- Yadda-yadda (yadda yadda Yadda)<br>- Yadda yadda yadda yadda yadda<br>- Yadda yadda yadda-yadda yadda </html>
            .Display
End With
Next Entry
Set Outlook = Nothing
Set Entry = Nothing
End Sub
