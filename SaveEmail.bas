Attribute VB_Name = "SaveEmail"

Option Explicit
'======================================================================================='

' Declare ShellExecute API for opening Obsidian links
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' ==================================================================================================
Sub ExtractEmail_MarkDown()

    Dim vaultPathToSaveFileTo As String
    Dim emailFileNameStartChr As String
    Dim emailTypeLink As String
    Dim personNameStartChar As String

    config vaultPathToSaveFileTo, personNameStartChar, emailFileNameStartChr, emailTypeLink

    Dim obj As Object
    Dim oMail As Outlook.mailItem
    On Error GoTo ErrHandler:

    Dim fileName As String, mName As String
    Dim temporarySubjectLineString As String
    Dim currentExplorer As Explorer
        Set currentExplorer = Application.ActiveExplorer
    Dim Selection As Selection
        Set Selection = currentExplorer.Selection

    For Each obj In Selection
        Set oMail = obj
        If oMail.Class <> 43 Then
            MsgBox "This code only works with Emails."
            GoTo EndClean:
        End If

        temporarySubjectLineString = oMail.subject
        ReplaceCharsForFileName temporarySubjectLineString, ""

        Dim dtDate As Date
            dtDate = oMail.ReceivedTime
        mName = Format(dtDate, "yyyymmdd", vbUseSystemDayOfWeek, vbUseSystem) & " " & temporarySubjectLineString

        ' Create per-email subfolder inside attachments (부모 폴더도 없으면 함께 생성)
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim mailFolder As String
        mailFolder = vaultPathToSaveFileTo & mName & "\"
        If Not fso.FolderExists(vaultPathToSaveFileTo) Then
            fso.CreateFolder vaultPathToSaveFileTo
        End If
        If Not fso.FolderExists(mailFolder) Then
            fso.CreateFolder mailFolder
        End If
        Set fso = Nothing

        ' (1) Save HTML (includes inline images in .files subfolder)
        Dim objItem As mailItem, htmlpath As String
        Set objItem = Application.ActiveExplorer.Selection(1)
        htmlpath = mailFolder & mName & ".html"
        objItem.SaveAs htmlpath, 5

        ' (2) Save attachments
        Dim attachments As Outlook.attachments
        Dim Attachment As Outlook.Attachment
        Set attachments = objItem.attachments
        Dim i As Long
        For i = 1 To attachments.Count
            Set Attachment = attachments(i)
            Attachment.SaveAsFile mailFolder & Attachment.fileName
        Next i

        ' (3) Build YAML frontmatter
        Dim sender As String
            sender = formatName(oMail.sender, personNameStartChar)
        Dim resultString As String

        resultString = "---" & vbCrLf
        resultString = resultString & "tags:" & vbCrLf
        resultString = resultString & "  - SOURCE/MAIL" & vbCrLf
        resultString = resultString & "title: """ & temporarySubjectLineString & """" & vbCrLf
        resultString = resultString & "date: " & Format(oMail.SentOn, "yyyy-MM-dd") & vbCrLf
        resultString = resultString & "from: """ & sender & """" & vbCrLf
        resultString = resultString & "ITSM: """""  & vbCrLf
        resultString = resultString & "ITSM_URL: """""  & vbCrLf
        resultString = resultString & "---" & vbCrLf & vbCrLf

        ' (4) HTML embed — right below YAML for quick preview
        resultString = resultString & Replace("![[" & baseFolder & mName & "/" & mName & ".html]]", "\", "/") & vbCrLf & vbCrLf

        ' (5) Visible attachment links
        Dim propertyAccessor As Outlook.propertyAccessor
        Dim propHidden As String
        Dim isHidden As Variant
        propHidden = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
        For i = 1 To attachments.Count
            Set Attachment = attachments(i)
            Set propertyAccessor = Attachment.propertyAccessor
            isHidden = propertyAccessor.GetProperty(propHidden)
            If Not isHidden Then
                resultString = resultString & Replace("![[" & baseFolder & mName & "/" & Attachment.fileName & "]]", "\", "/") & vbCrLf
            End If
        Next i

        ' (6) Convert HTML to Markdown for inline viewing
        Dim htmlContent As String
        htmlContent = ReadFileContent(htmlpath)

        Dim mdBody As String
        mdBody = ConvertHTMLToMarkdown(htmlContent, mName)

        resultString = resultString & vbCrLf & "---" & vbCrLf & vbCrLf
        resultString = resultString & mdBody & vbCrLf

        ' Save .md file
        fileName = mName & ".md"
        SaveAsUTF8 ObsidianFolder & fileName, resultString

        ' Open in Obsidian
        Dim obsidianURI As String
        obsidianURI = "obsidian://open?path=" & UrlEncodeUtf8NoBom(ObsidianFolder & fileName)
        ShellExecute 0, "open", obsidianURI, vbNullString, vbNullString, 1

    Next
    GoTo EndClean
ErrHandler:
    MsgBox "오류 발생 (줄 " & Erl & ")" & vbCrLf & _
           "오류 번호: " & Err.Number & vbCrLf & _
           "내용: " & Err.Description, vbCritical, "Outlook2Obsidian"
EndClean:
    Set obj = Nothing
    Set oMail = Nothing
End Sub

' ==================================================================================================
Sub ExtractEmail_html()
    ExtractEmail_MarkDown
End Sub
