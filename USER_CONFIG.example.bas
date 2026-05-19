Attribute VB_Name = "USER_CONFIG"
Option Explicit

Public ObsidianFolder
Public baseFolder
Public vaultPathToSaveFileTo

Public Sub config(vaultPathToSaveFileTo As String, personNameStartChar As String, Optional emailFileNameStartChr As String, Optional emailTypeLink As String, Optional meetingFileNameStartChr As String, Optional meetingTypeLink As String, Optional trainingFileNameStartChr As String, Optional trainingTypeLink As String, Optional contactFileNameStartChr As String, Optional contactTypeLink As String)

    '================================================'
    '====DECLARE=YOUR=FILE=PATH=TO=SAVE=FILES=TO====='
    '================================================'
    ' 1) Set INSTALL_FOLDER below to the directory where you placed the .bas
    '    files and vault-path.txt (ASCII path — never contains Korean).
    ' 2) Edit vault-path.txt (UTF-8) with Notepad. Put the absolute path to
    '    your Obsidian mail folder on a single line, e.g.:
    '        D:\GoogleDrive\Obsidian\Notes\Mail\
    '    Korean characters in the path are fully supported here.
    Const INSTALL_FOLDER As String = "C:\Outlook2Obsidian\"

    ObsidianFolder = ReadVaultPath(INSTALL_FOLDER & "vault-path.txt")
    baseFolder = "_attachments\"
    vaultPathToSaveFileTo = ObsidianFolder & baseFolder

    '================================================'
    ' 사람 이름 앞에 붙는 접두사 (wikilink 용)
    ' ex: "@" → [[@홍길동]]
    personNameStartChar = "@"
    '================================================'

    emailFileNameStartChr = "Email_"
    meetingFileNameStartChr = "M "
    trainingFileNameStartChr = "T "
    contactFileNameStartChr = "@"

    emailTypeLink = "[[+]]"
    meetingTypeLink = "[[&]]"
    trainingTypeLink = "[[!]]"
    contactTypeLink = "[[@]]"

End Sub
