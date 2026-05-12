Attribute VB_Name = "SaveUtilities"
Option Explicit
'======================================================================================='
Public Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application
        Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select
    Set objApp = Nothing
End Function

'======================================================================================='
Public Function UrlEncodeUtf8NoBom(ByVal sText As String) As String
    Dim oStream As Object
    Dim byteArray() As Byte
    Dim i As Long
    Dim sEncoded As String
    Dim startIndex As Long

    Set oStream = CreateObject("ADODB.Stream")
    oStream.Type = 2
    oStream.Mode = 3
    oStream.Charset = "UTF-8"
    oStream.Open
    oStream.WriteText sText
    oStream.Position = 0
    oStream.Type = 1
    byteArray = oStream.Read
    oStream.Close
    Set oStream = Nothing

    startIndex = LBound(byteArray)
    If (UBound(byteArray) - LBound(byteArray) >= 2) Then
        If byteArray(0) = &HEF And byteArray(1) = &HBB And byteArray(2) = &HBF Then
            startIndex = 3
        End If
    End If

    For i = startIndex To UBound(byteArray)
        sEncoded = sEncoded & "%" & Right("0" & Hex(byteArray(i)), 2)
    Next i

    UrlEncodeUtf8NoBom = sEncoded
End Function

'======================================================================================='
Public Sub ReplaceCharsForFileName(temporarySubjectLineString As String, sChr As String)
    temporarySubjectLineString = Replace(temporarySubjectLineString, "/", sChr)
    temporarySubjectLineString = Replace(temporarySubjectLineString, "\", sChr)
    temporarySubjectLineString = Replace(temporarySubjectLineString, ":", sChr)
    temporarySubjectLineString = Replace(temporarySubjectLineString, "?", sChr)
    temporarySubjectLineString = Replace(temporarySubjectLineString, Chr(34), sChr)
    temporarySubjectLineString = Replace(temporarySubjectLineString, "<", sChr)
    temporarySubjectLineString = Replace(temporarySubjectLineString, ">", sChr)
    temporarySubjectLineString = Replace(temporarySubjectLineString, "|", sChr)
    temporarySubjectLineString = Replace(temporarySubjectLineString, "[", sChr)
    temporarySubjectLineString = Replace(temporarySubjectLineString, "]", sChr)
End Sub

'======================================================================================='
Public Function formatName(str As String, personNameStartChar As String) As String
    Dim typeOfNameToClean As Integer

    Dim regexJustFirstNameAndLastName As Object
        Set regexJustFirstNameAndLastName = New RegExp
        regexJustFirstNameAndLastName.Pattern = "^\w+\s\w+$"
    If regexJustFirstNameAndLastName.Test(str) = True Then typeOfNameToClean = 1
    Set regexJustFirstNameAndLastName = Nothing

    Dim regexFirstNameLastNameAndFullDomain As Object
        Set regexFirstNameLastNameAndFullDomain = New RegExp
        regexFirstNameLastNameAndFullDomain.Pattern = "^\w+,\s\w+@\w+(\.\w+)+"
    If regexFirstNameLastNameAndFullDomain.Test(str) = True Then typeOfNameToClean = 2
    Set regexFirstNameLastNameAndFullDomain = Nothing

    Dim regexPlainEmailAddress As Object
        Set regexPlainEmailAddress = New RegExp
        regexPlainEmailAddress.Pattern = "^\w+@\w+\.\w+"
    If regexPlainEmailAddress.Test(str) = True Then typeOfNameToClean = 3
    Set regexPlainEmailAddress = Nothing

    Dim regexLastNameFirstNameAndAgency As Object
        Set regexLastNameFirstNameAndAgency = New RegExp
        regexLastNameFirstNameAndAgency.Pattern = "^[a-zA-Z_\-]+,\s[a-zA-Z_\-]+\s\(\w+\)$"
    If regexLastNameFirstNameAndAgency.Test(str) = True Then typeOfNameToClean = 4
    Set regexLastNameFirstNameAndAgency = Nothing

    Dim regexSingleNameAndDomain As Object
        Set regexSingleNameAndDomain = New RegExp
        regexSingleNameAndDomain.Pattern = "^([a-zA-Z_\-]+\s)+\([a-zA-Z_\-]+\)"
    If regexSingleNameAndDomain.Test(str) = True Then typeOfNameToClean = 5
    Set regexSingleNameAndDomain = Nothing

    Select Case typeOfNameToClean
        Case 1 ' John Doe
            formatName = "[[" & personNameStartChar & str & "]]"
        Case 2 ' Doe, John@domain.or.gov
            Dim fName As String, lname As String
            fName = Mid(str, InStr(str, ", ") + 2, InStr(str, "@") - (InStr(str, ", ") + 2))
            lname = Mid(str, 1, InStr(str, ",") - 1)
            formatName = "[[" & personNameStartChar & fName & " " & lname & "]]"
        Case 3 ' JohnDoe@gmail.com
            formatName = "[[" & personNameStartChar & Left(str, InStr(str, "@") - 1) & "]]"
        Case 4 ' Doe, John (Agency)
            Dim fname1 As String, lname1 As String
            fname1 = Mid(str, InStr(str, ", ") + 2, InStr(str, " (") - (InStr(str, ", ") + 2))
            lname1 = Mid(str, 1, InStr(str, ",") - 1)
            formatName = "[[" & personNameStartChar & fname1 & " " & lname1 & "]]"
        Case 5 ' Payroll (Agency)
            formatName = "[[" & Left(str, InStr(str, " (") - 1) & "]]"
        Case Else
            formatName = "[[" & str & "]]"
    End Select
End Function

'======================================================================================='
Public Sub SaveAsUTF8(filePath As String, content As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Mode = 3
    stm.Charset = "UTF-8"
    stm.Open
    stm.WriteText content
    stm.SaveToFile filePath, 2
    stm.Close
    Set stm = Nothing
End Sub

'======================================================================================='
Public Function ReadFileContent(ByVal filePath As String) As String
    Dim stage As String
    On Error GoTo ReadErr

    stage = "FSO check"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        ReadFileContent = ""
        Exit Function
    End If
    Set fso = Nothing

    stage = "Stream open"
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open

    stage = "LoadFromFile"
    stm.LoadFromFile filePath

    stage = "ReadText"
    stm.Position = 0
    stm.Type = 2
    ' Auto-detect encoding (UTF-8 with BOM, UTF-16, or system default).
    ' Outlook may save HTML in different encodings depending on content.
    stm.Charset = "_autodetect_all"
    ReadFileContent = stm.ReadText
    stm.Close
    Set stm = Nothing
    Exit Function

ReadErr:
    Err.Raise Err.Number, "ReadFileContent", _
        "[" & stage & "] " & Err.Description & " | path=" & filePath
End Function

'======================================================================================='
' HTML → Markdown conversion
' Properly handles: source whitespace, headings, bold/italic, images, tables, entities
Public Function ConvertHTMLToMarkdown(ByVal html As String, ByVal mName As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True

    ' Step 1: Remove noise blocks
    regEx.Pattern = "<head[\s\S]*?</head>"
    html = regEx.Replace(html, "")
    regEx.Pattern = "<style[\s\S]*?</style>"
    html = regEx.Replace(html, "")
    regEx.Pattern = "<script[\s\S]*?</script>"
    html = regEx.Replace(html, "")
    regEx.Pattern = "<!--[\s\S]*?-->"     ' removes conditional comments and VML comments
    html = regEx.Replace(html, "")
    regEx.Pattern = "</?v:[^>]*>"          ' VML tags
    html = regEx.Replace(html, "")

    ' Step 2: Extract tables → placeholders (before whitespace collapse)
    Dim tableRegEx As Object
    Set tableRegEx = CreateObject("VBScript.RegExp")
    tableRegEx.Pattern = "<table[\s\S]*?</table>"
    tableRegEx.Global = True
    tableRegEx.IgnoreCase = True

    Dim tableMatches As Object, tblMatch As Object
    Dim placeholders() As String, tableMarkdown() As String
    Dim idx As Long
    idx = 0
    Set tableMatches = tableRegEx.Execute(html)
    For Each tblMatch In tableMatches
        ReDim Preserve placeholders(idx)
        ReDim Preserve tableMarkdown(idx)
        placeholders(idx) = "%%TABLE_" & idx & "%%"
        tableMarkdown(idx) = ConvertHTMLTableToMarkdown(tblMatch.Value)
        html = Replace(html, tblMatch.Value, placeholders(idx))
        idx = idx + 1
    Next tblMatch

    ' Step 3: Collapse HTML source line breaks to spaces
    ' (In HTML, raw newlines are just whitespace - only <br>/<p> create visual breaks)
    regEx.Pattern = "(\r\n|\r|\n)"
    html = regEx.Replace(html, " ")

    ' Step 4: Replace img tags with Obsidian wikilinks
    html = ReplaceImageTagsWithMarkdownLinks(html, mName)

    ' Step 4b: Normalize non-breaking spaces to regular spaces early so
    ' bold/whitespace cleanup works correctly. Outlook reply headers use NBSP
    ' for alignment — without this, they survive the space-collapse step and
    ' produce wide gaps like "보낸 사람:        신석철".
    html = Replace(html, "&nbsp;", " ")
    html = Replace(html, "&#160;", " ")
    html = Replace(html, "&#xa0;", " ")
    html = Replace(html, "&#xA0;", " ")
    html = Replace(html, ChrW(160), " ")
    html = Replace(html, vbTab, " ")

    ' Step 5: Headings
    regEx.Pattern = "<h[12][^>]*>"
    html = regEx.Replace(html, vbCrLf & "## ")
    regEx.Pattern = "<h[34][^>]*>"
    html = regEx.Replace(html, vbCrLf & "### ")
    regEx.Pattern = "<h[56][^>]*>"
    html = regEx.Replace(html, vbCrLf & "#### ")
    regEx.Pattern = "</h[1-6]>"
    html = regEx.Replace(html, vbCrLf)

    ' Step 6: Strip bold/italic tags
    ' Outlook emails have deeply nested <b> tags in headers/signatures that produce
    ' mismatched ** markers. Plain text is cleaner and more readable for email content.
    regEx.Pattern = "</?(?:b|strong|i|em)[^>]*>"
    html = regEx.Replace(html, "")

    ' Step 7: Line breaks and block elements → newlines
    ' Only closing tags add newlines to avoid double-spacing from open+close pairs
    regEx.Pattern = "<br\s*/?>"
    html = regEx.Replace(html, vbCrLf)
    regEx.Pattern = "<(p|div|li|blockquote|tr)[^>]*>"
    html = regEx.Replace(html, " ")
    regEx.Pattern = "</(p|div|li|blockquote|tr)>"
    html = regEx.Replace(html, vbCrLf)

    ' Step 8: Decode remaining HTML entities (&nbsp; already handled in Step 4b)
    html = Replace(html, "&amp;", "&")
    html = Replace(html, "&lt;", "<")
    html = Replace(html, "&gt;", ">")
    html = Replace(html, "&quot;", Chr(34))
    html = Replace(html, "&#39;", "'")
    html = Replace(html, "&apos;", "'")

    ' Step 8b: Collapse multiple consecutive spaces to single space
    Do While InStr(html, "  ") > 0
        html = Replace(html, "  ", " ")
    Loop

    ' Step 9: Strip all remaining HTML tags
    regEx.Pattern = "<[^>]+>"
    html = regEx.Replace(html, "")

    ' Step 9b: Trim leading/trailing whitespace on each line (removes source indentation)
    Dim lines() As String
    Dim j As Long
    If InStr(html, vbCrLf) > 0 Then
        lines = Split(html, vbCrLf)
    Else
        lines = Split(html, vbLf)
    End If
    For j = 0 To UBound(lines)
        lines(j) = Trim(lines(j))
    Next j
    html = Join(lines, vbCrLf)

    ' Step 9c: Split reply chain header fields onto separate lines
    ' ChrW values: 보(48372) 낸(45240) 날(45216) 짜(51676) 받(48155) 는(45716) 사(49324) 람(46988) 참(52280) 조(51312) 제(51228) 목(47785)
    Dim replyFields As Variant
    replyFields = Array(ChrW(48372) & ChrW(45240) & " " & ChrW(49324) & ChrW(46988) & ":", _
                        ChrW(48372) & ChrW(45240) & " " & ChrW(45216) & ChrW(51676) & ":", _
                        ChrW(48155) & ChrW(45716) & " " & ChrW(49324) & ChrW(46988) & ":", _
                        ChrW(52280) & ChrW(51312) & ":", _
                        ChrW(51228) & ChrW(47785) & ":")
    Dim f As Long
    For f = 0 To UBound(replyFields)
        html = Replace(html, " " & replyFields(f), vbCrLf & replyFields(f))
    Next f

    ' Step 9d: Insert --- before signature block (image wikilink after non-image content)
    Dim sigLines() As String
    If InStr(html, vbCrLf) > 0 Then
        sigLines = Split(html, vbCrLf)
    Else
        sigLines = Split(html, vbLf)
    End If
    Dim sigInserted As Boolean
    Dim trimLine As String
    sigInserted = False
    Dim k As Long
    Dim prevNonEmpty As String
    For k = 1 To UBound(sigLines)
        If Not sigInserted Then
            trimLine = Trim(sigLines(k))
            If Left(trimLine, 3) = "![[" Then
                ' Scan backward to find the last non-empty line
                Dim m As Long
                prevNonEmpty = ""
                For m = k - 1 To 0 Step -1
                    If Trim(sigLines(m)) <> "" Then
                        prevNonEmpty = Trim(sigLines(m))
                        Exit For
                    End If
                Next m
                ' Insert --- only if preceded by actual content (not another image)
                If prevNonEmpty <> "" And Left(prevNonEmpty, 3) <> "![[" Then
                    sigLines(k) = "---" & vbCrLf & sigLines(k)
                    sigInserted = True
                End If
            End If
        End If
    Next k
    html = Join(sigLines, vbCrLf)

    ' Step 10: Normalize whitespace — collapse 3+ consecutive newlines to 2
    Do While InStr(html, vbCrLf & vbCrLf & vbCrLf) > 0
        html = Replace(html, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
    Loop
    Do While InStr(html, vbLf & vbLf & vbLf) > 0
        html = Replace(html, vbLf & vbLf & vbLf, vbLf & vbLf)
    Loop

    ' Step 11: Restore table placeholders
    Dim i As Long
    For i = 0 To idx - 1
        html = Replace(html, placeholders(i), vbCrLf & tableMarkdown(i) & vbCrLf)
    Next i

    ConvertHTMLToMarkdown = Trim(html)
End Function

'======================================================================================='
' Replace <img> tags with Obsidian wikilinks pointing to saved .files folder
Function ReplaceImageTagsWithMarkdownLinks(ByVal html As String, ByVal mName As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    regEx.Pattern = "<img[^>]*src\s*=\s*[""']([^""']+)[""'][^>]*>"

    Dim matches As Object, m As Object
    Dim result As String, imgFileName As String
    result = html
    Set matches = regEx.Execute(html)

    Dim i As Long
    For i = matches.Count - 1 To 0 Step -1
        Set m = matches(i)
        imgFileName = GetFileNameFromPath(m.SubMatches(0))
        Dim newSrc As String
        newSrc = Replace(baseFolder & mName & "\" & mName & ".files\" & imgFileName, "\", "/")
        Dim mdLink As String
        mdLink = "![[" & newSrc & "]]"
        result = Left(result, m.FirstIndex) & mdLink & Mid(result, m.FirstIndex + m.Length + 1)
    Next i

    ReplaceImageTagsWithMarkdownLinks = result
End Function

'======================================================================================='
Function GetFileNameFromPath(ByVal path As String) As String
    Dim pos As Long
    pos = InStrRev(path, "/")
    If pos = 0 Then pos = InStrRev(path, "\")
    If pos > 0 Then
        GetFileNameFromPath = Mid(path, pos + 1)
    Else
        GetFileNameFromPath = path
    End If
End Function

'======================================================================================='
' Convert an HTML <table> to Markdown table syntax
Public Function ConvertHTMLTableToMarkdown(tblHtml As String) As String
    On Error GoTo ErrorHandler
    Dim md As String
    Dim htmlDoc As Object, tbl As Object, rows As Object
    Dim r As Long, c As Long, colCount As Long
    Dim cellText As String

    Set htmlDoc = CreateObject("htmlfile")
    htmlDoc.Open
    htmlDoc.Write tblHtml
    htmlDoc.Close
    Set tbl = htmlDoc.getElementsByTagName("table")(0)
    Set rows = tbl.getElementsByTagName("tr")

    If rows.Length = 0 Then
        ConvertHTMLTableToMarkdown = ""
        Exit Function
    End If

    Dim headerCells As Object
    Set headerCells = rows(0).getElementsByTagName("th")
    If headerCells.Length = 0 Then Set headerCells = rows(0).getElementsByTagName("td")
    colCount = headerCells.Length
    If colCount = 0 Then
        ConvertHTMLTableToMarkdown = ""
        Exit Function
    End If

    md = "|"
    For c = 0 To colCount - 1
        md = md & " " & CleanCellText(headerCells(c).innerText) & " |"
    Next c
    md = md & vbCrLf & "|"
    For c = 0 To colCount - 1
        md = md & " --- |"
    Next c
    md = md & vbCrLf

    Dim currentCells As Object
    For r = 1 To rows.Length - 1
        Set currentCells = rows(r).getElementsByTagName("td")
        If currentCells.Length = 0 Then Set currentCells = rows(r).getElementsByTagName("th")
        If currentCells.Length > 0 Then
            md = md & "|"
            For c = 0 To currentCells.Length - 1
                md = md & " " & CleanCellText(currentCells(c).innerText) & " |"
            Next c
            md = md & vbCrLf
        End If
    Next r

    ConvertHTMLTableToMarkdown = md
    Exit Function
ErrorHandler:
    ConvertHTMLTableToMarkdown = ""
End Function

Public Function CleanCellText(ByVal text As String) As String
    CleanCellText = Trim(Replace(Replace(text, vbCrLf, " "), vbLf, " "))
End Function
