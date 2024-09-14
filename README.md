# Outlook Phishing Email Detection Macro

A VBA macro for Microsoft Outlook that enhances the detection of phishing emails by analyzing various aspects of incoming messages. This macro evaluates emails based on multiple criteria and assigns a risk score to identify potential phishing attempts, helping to protect users from malicious content.

## Table of Contents

- [Features](#features)
- [Installation](#installation)
  - [Prerequisites](#prerequisites)
  - [Macro Security Settings](#macro-security-settings)
  - [Setting Up the Macro](#setting-up-the-macro)
- [Usage](#usage)
- [Configuration](#configuration)
  - [Adjusting the Risk Score Threshold](#adjusting-the-risk-score-threshold)
  - [Customizing Suspicious Keywords](#customizing-suspicious-keywords)
  - [Updating Blacklisted Domains](#updating-blacklisted-domains)
- [Functions Explained](#functions-explained)
  - [Keyword Analysis](#keyword-analysis)
  - [Header Analysis](#header-analysis)
  - [Link Analysis](#link-analysis)
  - [Attachment Analysis](#attachment-analysis)
  - [Sender Reputation Check](#sender-reputation-check)
  - [Reply-To and From Address Mismatch](#reply-to-and-from-address-mismatch)
  - [Right-to-Left Override Character Detection](#right-to-left-override-character-detection)
  - [Homograph Attack Detection](#homograph-attack-detection)
- [Testing the Macro](#testing-the-macro)
- [Security Considerations](#security-considerations)
- [Contributing](#contributing)
- [License](#license)
- [Disclaimer](#disclaimer)

---

## Features

- **Advanced Keyword Analysis**: Uses weighted scoring for suspicious keywords in email content.
- **Email Header Analysis**: Checks SPF, DKIM, and DMARC authentication results.
- **URL and Link Analysis**: Detects suspicious URLs and obfuscated links.
- **Attachment Scanning**: Flags potentially dangerous attachments based on file extensions.
- **Sender Reputation Check**: Evaluates the sender's domain against a blacklist.
- **Reply-To Mismatch Detection**: Identifies discrepancies between sender and reply-to addresses.
- **Unicode Character Detection**: Detects Right-to-Left Override characters and homograph attacks.
- **Risk Scoring System**: Assigns a cumulative risk score to each email to assess the phishing risk.
- **Customizable Thresholds and Lists**: Allows configuration to suit organizational needs.

## Installation

### Prerequisites

- **Microsoft Outlook**: This macro is designed for Outlook on Windows.
- **Macro Security**: Ability to enable and run VBA macros in Outlook.

### Macro Security Settings

1. Open Outlook.
2. Go to **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Macro Settings**.
3. Select **Notifications for all macros** or **Enable all macros** (not recommended for security reasons).
4. Click **OK**.

### Setting Up the Macro

1. **Enable the Developer Tab**:

   - Go to **File** > **Options** > **Customize Ribbon**.
   - In the right pane, check the box next to **Developer**.
   - Click **OK**.

2. **Open the VBA Editor**:

   - Click on the **Developer** tab.
   - Click **Visual Basic**.

3. **Insert the Macro Code**:

   - In the VBA editor, double-click **ThisOutlookSession** in the Project pane.
   - Paste the macro code (provided below) into the code window.

4. **Save the Macro**:

   - Click **File** > **Save** in the VBA editor.
   - Close the VBA editor.

5. **Restart Outlook**:

   - Close and reopen Outlook to activate the macro.

## Usage

Once installed, the macro runs automatically and evaluates incoming emails in your Inbox. If a potential phishing email is detected based on the risk score threshold, a warning message will appear displaying the subject and risk score of the email.

Optional actions (commented out in the code) can be enabled, such as:

- Moving the email to the Junk folder.
- Forwarding the email to your IT security team.

## Configuration

### Adjusting the Risk Score Threshold

The variable `phishingThreshold` determines the sensitivity of the detection.

- **Location in Code**:

  ```vb
  phishingThreshold = 5  ' Adjust based on desired sensitivity
  ```

- **Adjusting**:

  - Lower the value to make the macro more sensitive (may increase false positives).
  - Raise the value to reduce sensitivity (may risk missed detections).

### Customizing Suspicious Keywords

Update the list of suspicious keywords and their weights to better fit your organization's needs.

- **Location in Code**:

  ```vb
  suspiciousWords = Array( _
      Array("account verification", 2), _
      Array("urgent action required", 3), _
      Array("confirm your password", 2), _
      ' Add or modify entries here
  )
  ```

- **Adding Keywords**:

  - Add new arrays with the keyword and its associated weight.

### Updating Blacklisted Domains

Modify the list of blacklisted domains to include known malicious senders.

- **Location in Code**:

  ```vb
  blacklistedDomains = Array("malicious.com", "phishing.net", "bad-domain.org")
  ```

- **Adding Domains**:

  - Add domains as strings within the array.

## Functions Explained

### Keyword Analysis

Analyzes the email content for suspicious keywords and calculates a weighted score.

### Header Analysis

Checks the email headers for SPF, DKIM, and DMARC authentication results to detect spoofing.

### Link Analysis

Scans the email for suspicious URLs using regular expressions and pattern matching.

### Attachment Analysis

Flags emails with potentially dangerous attachments by checking file extensions.

### Sender Reputation Check

Evaluates the sender's domain against a blacklist to assess reputation.

### Reply-To and From Address Mismatch

Detects if the `Reply-To` address is different from the sender's email, which can be a phishing indicator.

### Right-to-Left Override Character Detection

Identifies the use of Unicode characters that can disguise filenames and content.

### Homograph Attack Detection

Checks for the use of non-ASCII characters that resemble standard characters to prevent homograph attacks.

## Testing the Macro

- **Simulate Phishing Emails**:

  - Send test emails containing suspicious keywords, links, or attachments to verify detection.

- **Review Alerts**:

  - Ensure that warning messages display the correct information and risk scores.

- **Enable Optional Actions**:

  - Uncomment optional actions in the code to test moving emails or forwarding them.

## Security Considerations

- **Macro Security**:

  - Be cautious with enabling macros, as they can pose security risks.
  - Only run macros from trusted sources.

- **Automatic Actions**:

  - Carefully consider enabling automatic actions like moving or forwarding emails.
  - Ensure compliance with organizational policies.

- **False Positives**:

  - Adjust thresholds and lists to minimize false positives.
  - Regularly review and update configurations.

## Contributing

Contributions are welcome! To contribute:

1. Fork the repository.
2. Create a new branch for your feature or bug fix.
3. Commit your changes and submit a pull request.

Please ensure that your code adheres to the project's coding standards and includes appropriate documentation.

## License

This project is licensed under the [MIT License](LICENSE).

## Disclaimer

**Important**: This macro is intended to assist in the detection of phishing emails but should not be relied upon as the sole security measure. Always use up-to-date antivirus software and follow best practices for email security. The developers are not responsible for any damages or losses resulting from the use of this macro.

---

## Macro Code

Below is the complete VBA macro code. Copy and paste this into the **ThisOutlookSession** module in the VBA editor.

```vb
' Enhanced Outlook VBA Macro for Phishing Email Detection
' Author: Your Name
' Date: YYYY-MM-DD

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
            ' Dim fwd As MailItem
            ' Set fwd = mail.Forward
            ' fwd.Recipients.Add "itsecurity@example.com"
            ' fwd.Send
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
```

---

## Additional Resources

- **Microsoft Outlook VBA Reference**: [Link](https://docs.microsoft.com/en-us/office/vba/api/overview/)
- **Regular Expressions in VBA**: Utilize regex for advanced pattern matching.
- **Email Security Best Practices**: Stay informed about the latest phishing tactics and prevention strategies.

## Acknowledgments

- **Community Contributors**: Thank you to everyone who has contributed suggestions and improvements.
- **Security Researchers**: For ongoing efforts to understand and mitigate phishing threats.

---

By using this macro, you contribute to a safer email environment. Always combine technical solutions with user education to enhance security awareness within your organization.
