# VB.net OAuth Connection to Outlook
Follow [Microsoft's documentation](https://docs.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth) to get the necessary parameters 

## Example Usage

> Dim o365 = New O365EasyEmail
>
> o365.InitToken(Client_ID/App_ID, client_secret,
>    tenant_ID, emailToAccess, Password)
> 
>    Console.WriteLine(o365.GetEmails())
> 
>    Console.WriteLine(o365.DeleteEmail("AAMkAGZkNjJhMjlkLTd"))
