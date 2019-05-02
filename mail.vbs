Const MSG_RECIPIENT_LIST = "test@zxc.com"
Const MSG_SUBJECT = "IN"
Const MSG_BODY = ""
 
Dim olkApp, olkSes, olkMsg
Set olkApp = CreateObject("Outlook.Application")
Set olkSes = olkApp.GetNamespace("MAPI")
olkSes.Logon olkApp.DefaultProfileName
Set olkMsg = olkApp.CreateItem(0)
With olkMsg
    .To = MSG_RECIPIENT_LIST
    .Subject = MSG_SUBJECT
    .HTMLBody = MSG_BODY
    .Send
End With
olkSes.Logoff
Set olkMsg = Nothing
Set olkSes = Nothing
Set olkApp = Nothings
