# MailLib

A simple wrapper around MailKit that can be used to send email with modern authentication from Classic ASP apps.

## Installing

Copy the contents of bin/ to a place on your server, the use a command prompt to run:

```
C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe .\MailLib.dll /codebase /tlb
```

### Sample ASP Script

```VBScript
<html>
<head>
    <title>Test Email</title>
</head>
<body>
    <h1>Test Email</h1>
    <form action="TestEmailNet.asp?process=1" method="post">
        <p>
            Host
            <br />
                <input type="text" name="host" value="" />
        </p>

        <p>
            Port
            <br />
                <input type="text" name="port" value="587" />
        </p>

        <p>
            Username
            <br />
            <input type="text" name="username" value="" />
        </p>

        <p>
            Password
            <br />
            <input type="password" name="password" value="" />
        </p>

        <p>
            Sender Address
            <br />
            <input type="text" name="sender" value="" />
        </p>

        <p>
            Reply-To Address
            <br />
            <input type="text" name="replyto" />
        </p>

        <p>
            From Address
            <br />
            <input type="text" name="from" value="" />
        </p>

        <p>
            To
            <br />
            <input type="text" name="to" value="" />
        </p>

        <p>
            Subject
            <br />
            <input type="text" name="subject" value="Test Email" />
        </p>

        <p>
            Body (HTML allowed)
            <br />
            <textarea name="bodyhtml" rows="10" cols="50">&lt;b&gt;Test Email!&lt;/b&gt;</textarea>
        </p>

        <p>
            <input type="submit" value="Submit" />
        </p>
    </form>
    <%
    If Request.QueryString("process") <> "1" Then
        Response.End
    End If

    If Request.Form("to") = "" Then
        Response.Write "<strong>A To: address is required.</strong>"
        Response.End
    End If


    Dim EmailSender
    Set EmailSender = Server.CreateObject("MailLib.EmailSender")

    EmailSender.Host = Request.Form("host")
    EmailSender.Port = CInt(Request.Form("port"))
    EmailSender.UserName = Request.Form("username")
    EmailSender.Password = Request.Form("password")

    EmailSender.AddFromAddress Request.Form("from")
    EmailSender.AddToAddress Request.Form("to")

    EmailSender.Subject = Request.Form("subject")
    EmailSender.AppendToHtmlBody(Request.Form("bodyhtml"))

    EmailSender.Send

    If Err <> 0 Then ' error occurred
        Response.Write "<strong>" & Err.Description & "</strong>"
    else
        Response.Write "<strong>No error, check recipient's inbox.</strong>"
    End If
    %>
</body>
</html>
```
