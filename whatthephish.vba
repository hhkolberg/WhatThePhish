' Outlook VBA Macro for Phishing Email Detection
' Author: KaliNetz
' Date: 2024-09-14

Private WithEvents inboxItems As Outlook.Items
Private phishingThreshold As Integer

Private Sub Application_Startup()
    Dim outlookNamespace As Outlook.NameSpace
    Set outlookNamespace = Application.GetNamespace("MAPI")
    Set inboxItems = outlookNamespace.GetDefaultFolder(olFolderInbox).Items
    
    ' Set the risk score threshold for flagging an email as phishing
    phishingThreshold = 5  ' Adjust based on desired sensitivity
End Sub

Private Sub inboxItems_ItemAdd(ByVal Item As Object)
    On Error Resume Next
    If TypeOf Item Is MailItem Then
        Dim mail As MailItem
        Set mail = Item
        Dim riskScore As Integer
        riskScore = EvaluateEmail(mail)
        If riskScore >= phishingThreshold Then
            ' Action to take on detection
            MsgBox "Warning: Potential phishing email detected!" & vbCrLf & _
                   "Subject: " & mail.Subject & vbCrLf & _
                   "Risk Score: " & riskScore, vbExclamation, "Phishing Alert"
            ' Optional: Move email to Junk folder
            ' mail.Move Session.GetDefaultFolder(olFolderJunk)
            ' Optional: Forward to IT security
            ' mail.Forward.Recipients.Add "itsecurity@example.com"
        End If
    End If
End Sub

Function EvaluateEmail(mail As MailItem) As Integer
    Dim totalRiskScore As Integer
    totalRiskScore = 0
    
    ' Perform various analyses and accumulate the risk score
    totalRiskScore = totalRiskScore + KeywordAnalysis(mail)
    totalRiskScore = totalRiskScore + HeaderAnalysis(mail)
    totalRiskScore = totalRiskScore + LinkAnalysis(mail)
    totalRiskScore = totalRiskScore + AttachmentAnalysis(mail)
    totalRiskScore = totalRiskScore + SenderReputation(mail)
    totalRiskScore = totalRiskScore + ReplyToMismatch(mail)
    totalRiskScore = totalRiskScore + ContainsRTLOCharacter(mail)
    totalRiskScore = totalRiskScore + HomographAttackCheck(mail)
    
    ' Return the total risk score
    EvaluateEmail = totalRiskScore
End Function

' 1. Advanced Keyword Analysis with Weighted Scoring
Function KeywordAnalysis(mail As MailItem) As Integer
    Dim suspiciousWords As Variant
    Dim word As Variant
    Dim emailContent As String
    Dim score As Integer
    score = 0
    
    ' List of keywords with assigned weights
    suspiciousWords = Array( _
        Array("account verification", 2), _
        Array("urgent action required", 3), _
        Array("confirm your password", 2), _
        Array("update your information", 2), _
        Array("security alert", 2), _
        Array("unusual activity", 3), _
        Array("click here", 1), _
        Array("log in to your account", 2), _
        Array("suspended", 1), _
        Array("unauthorized access", 3), _
        Array("verify your account", 2))
    
    ' Combine subject and body for scanning
    emailContent = LCase(mail.Subject & " " & mail.Body)
    
    For Each word In suspiciousWords
        If InStr(emailContent, LCase(word(0))) > 0 Then
            score = score + word(1)
        End If
    Next word
    
    KeywordAnalysis = score
End Function

' 2. Email Header Analysis
Function HeaderAnalysis(mail As MailItem) As Integer
    Dim score As Integer
    score = 0
    ' Access email headers
    Dim headers As String
    headers = mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
    headers = LCase(headers)
    
    ' Check for common signs of spoofing
    If InStr(headers, "spf=fail") > 0 Or InStr(headers, "dmarc=fail") > 0 Then
        score = score + 3
    ElseIf InStr(headers, "spf=softfail") > 0 Or InStr(headers, "dmarc=quarantine") > 0 Then
        score = score + 2
    End If
    
    HeaderAnalysis = score
End Function

' 3. URL and Link Analysis
Function LinkAnalysis(mail As MailItem) As Integer
    Dim score As Integer
    score = 0
    Dim bodyContent As String
    bodyContent = mail.HTMLBody
    
    ' Extract URLs using regex
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "(href\s*=\s*[""']?)([^'"" >]+)"
    regex.Global = True
    regex.IgnoreCase = True
    
    Dim matches As Object
    Set matches = regex.Execute(bodyContent)
    
    Dim match As Object
    For Each match In matches
        Dim url As String
        url = match.SubMatches(1)
        If IsSuspiciousURL(url) Then
            score = score + 2
        End If
    Next match
    
    LinkAnalysis = score
End Function

Function IsSuspiciousURL(url As String) As Boolean
    ' Implement URL checking logic here
    ' For example, check for IP addresses, excessive subdomains, or mismatched display text
    Dim suspiciousPatterns As Variant
    suspiciousPatterns = Array( _
        "^\d{1,3}(\.\d{1,3}){3}", _                 ' IP addresses
        "([a-z0-9]+\.){3,}[a-z]{2,}", _             ' Excessive subdomains
        "[^\s]+@[^\.]+\.[a-z]{2,}", _               ' Emails used as URLs
        "[^\s]+//[^\s]+"                            ' Obfuscated URLs
    )
    
    Dim pattern As Variant
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    
    For Each pattern In suspiciousPatterns
        regex.Pattern = pattern
        If regex.Test(url) Then
            IsSuspiciousURL = True
            Exit Function
        End If
    Next pattern
    
    IsSuspiciousURL = False
End Function

' 4. Attachment Scanning
Function AttachmentAnalysis(mail As MailItem) As Integer
    Dim score As Integer
    score = 0
    Dim att As Attachment
    For Each att In mail.Attachments
        Dim fileExt As String
        fileExt = LCase(Right(att.FileName, Len(att.FileName) - InStrRev(att.FileName, ".")))
        Select Case fileExt
            Case "exe", "scr", "js", "vbs", "bat", "cmd", "ps1", "jar", "msi", "reg"
                score = score + 3
            Case "docm", "xlsm", "pptm"
                score = score + 2
            Case "zip", "rar", "7z"
                score = score + 1
        End Select
    Next att
    AttachmentAnalysis = score
End Function

' 5. Sender Reputation Check (Basic Implementation)
Function SenderReputation(mail As MailItem) As Integer
    Dim score As Integer
    score = 0
    Dim senderDomain As String
    senderDomain = GetDomainFromEmail(mail.SenderEmailAddress)
    
    ' Simple blacklist check (you can integrate with an external API or database)
    Dim blacklistedDomains As Variant
    blacklistedDomains = Array("malicious.com", "phishing.net", "bad-domain.org")
    
    Dim domain As Variant
    For Each domain In blacklistedDomains
        If LCase(senderDomain) = LCase(domain) Then
            score = score + 5
            Exit For
        End If
    Next domain
    
    SenderReputation = score
End Function

Function GetDomainFromEmail(emailAddress As String) As String
    Dim atPos As Integer
    atPos = InStr(emailAddress, "@")
    If atPos > 0 Then
        GetDomainFromEmail = Mid(emailAddress, atPos + 1)
    Else
        GetDomainFromEmail = ""
    End If
End Function

' 6. Reply-To and From Address Mismatch
Function ReplyToMismatch(mail As MailItem) As Integer
    Dim score As Integer
    score = 0
    If mail.ReplyRecipientNames <> "" And mail.SenderEmailAddress <> "" Then
        If LCase(mail.ReplyRecipientNames) <> LCase(mail.SenderEmailAddress) Then
            score = score + 2
        End If
    End If
    ReplyToMismatch = score
End Function

' 7. Right-to-Left Override Character Detection
Function ContainsRTLOCharacter(mail As MailItem) As Integer
    Dim score As Integer
    score = 0
    Const RTLO_CHAR As String = ChrW(&H202E)
    If InStr(mail.Subject, RTLO_CHAR) > 0 Or InStr(mail.Body, RTLO_CHAR) > 0 Then
        score = score + 3
    End If
    ContainsRTLOCharacter = score
End Function

' 8. Homograph Attack Detection
Function HomographAttackCheck(mail As MailItem) As Integer
    Dim score As Integer
    score = 0
    Dim displayName As String
    Dim emailAddress As String
    displayName = mail.SenderName
    emailAddress = mail.SenderEmailAddress
    
    If ContainsHomographCharacters(displayName) Or ContainsHomographCharacters(emailAddress) Then
        score = score + 2
    End If
    HomographAttackCheck = score
End Function

Function ContainsHomographCharacters(text As String) As Boolean
    ' Basic implementation: Check for characters outside ASCII range
    Dim i As Integer
    For i = 1 To Len(text)
        If AscW(Mid(text, i, 1)) > 127 Then
            ContainsHomographCharacters = True
            Exit Function
        End If
    Next i
    ContainsHomographCharacters = False
End Function

' Additional utility functions and features can be added here...
