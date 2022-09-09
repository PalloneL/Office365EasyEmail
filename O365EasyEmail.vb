Imports System
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Web

Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Net.Http

Public Class O365EasyEmail

    Public token As String
    Public tenant As String

    Sub InitToken(app_id As String, client_secret As String, tenant_id As String, email As String, password As String)

        Dim token_data As New TokenData(app_id, client_secret, email, password)
        tenant = tenant_id
        Dim token_url = "https://login.microsoftonline.com/" + tenant + "/oauth2/token"

        Dim result = TokenPost(token_url, token_data)
        token = result.GetValue("access_token").ToString
        'Console.WriteLine(token)
        'Console.WriteLine(GetEmails())

    End Sub
    Function HttpRequest(URL As String) As String
        Dim request As WebRequest = WebRequest.Create(URL)
        Dim dataStream As Stream = request.GetResponse.GetResponseStream()
        Dim sr As New StreamReader(dataStream)
        Return sr.ReadToEnd
    End Function
    Function TokenPost(URL As String, token_data As TokenData) As JObject

        Dim outgoingQueryString = HttpUtility.ParseQueryString(String.Empty)
        outgoingQueryString.Add("grant_type", token_data.grant_type)
        outgoingQueryString.Add("client_id", token_data.client_id)
        outgoingQueryString.Add("client_secret", token_data.client_secret)
        outgoingQueryString.Add("resource", token_data.resource)
        outgoingQueryString.Add("scope", token_data.scope)
        outgoingQueryString.Add("username", token_data.username)
        outgoingQueryString.Add("password", token_data.password)


        Dim postdata = outgoingQueryString.ToString()
        'Console.WriteLine(postdata)

        ' JSON Data
        Dim byteArray As Byte() = Encoding.UTF8.GetBytes(postdata)

        ' HTTP Request
        Dim request As WebRequest = WebRequest.Create(URL)
        request.Method = "POST"

        ' Requesting
        Dim dataStream As Stream = request.GetRequestStream()
        dataStream.Write(byteArray, 0, byteArray.Length)
        dataStream.Close()

        ' Recieving a Response
        dataStream = request.GetResponse.GetResponseStream()
        Dim reader As New StreamReader(dataStream)
        Dim responseFromServer As String = reader.ReadToEnd()
        dataStream.Close()
        reader.Close()

        Return JObject.Parse(responseFromServer)

    End Function

    Public Function GetEmails(Optional ByVal id As String = "") As JObject
        Dim URL = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages"
        If id.Length > 0 Then
            URL += id
        End If
        Try
            Dim request As WebRequest = WebRequest.Create(URL)
            request.Method = "GET"
            request.Headers.Add("Authorization", "Bearer " + token)
            request.Headers.Add("Prefer", "outlook.body-content-type=html")
            Dim dataStream = request.GetResponse.GetResponseStream()
            Dim reader As New StreamReader(dataStream)
            Dim responseFromServer As String = reader.ReadToEnd()
            dataStream.Close()
            reader.Close()

            Return JObject.Parse(responseFromServer)
        Catch
            Console.WriteLine("ERROR Getting email, ID: " + id)
            Return JObject.Parse("{}")
        End Try
    End Function

    Public Function DeleteEmail(id As String) As String
        Dim URL = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/" + id
        Dim request As WebRequest = WebRequest.Create(URL)
        request.Method = "Delete"
        request.Headers.Add("Authorization", "Bearer " + token)
        Try
            Dim dataStream = request.GetResponse.GetResponseStream()
            Dim reader As New StreamReader(dataStream)
            Dim responseFromServer As String = reader.ReadToEnd()
            dataStream.Close()
            reader.Close()

            Return ("Deleted Email with ID " + id)

        Catch
            Return ("ERROR Deleting Email " + id + ", please double check the ID")
        End Try
        'Return URL
    End Function

End Class

Public Class TokenData

    Public grant_type As String = "password"
    Public client_id, client_secret As String
    Public resource As String = "https://graph.microsoft.com"
    Public scope As String = "https://graph.microsoft.com"
    Public username, password As String

    Public Sub New(app As String, cs As String, email As String, pass As String)
        client_id = app
        client_secret = cs
        username = email
        password = pass
    End Sub


End Class
