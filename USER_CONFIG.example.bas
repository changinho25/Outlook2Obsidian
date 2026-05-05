Attribute VB_Name = "USER_CONFIG"
Option Explicit

Public ObsidianFolder
Public baseFolder
Public vaultPathToSaveFileTo

Public Sub config(vaultPathToSaveFileTo As String, personNameStartChar As String, Optional emailFileNameStartChr As String, Optional emailTypeLink As String, Optional meetingFileNameStartChr As String, Optional meetingTypeLink As String, Optional trainingFileNameStartChr As String, Optional trainingTypeLink As String, Optional contactFileNameStartChr As String, Optional contactTypeLink As String)

    '================================================'
    '====DECLARE=YOUR=FILE=PATH=TO=SAVE=FILES=TO====='
    '================================================'
    ' 파일을 저장할 Obsidian vault 경로를 입력하세요
    ' 경로 끝에 반드시 백슬래시(\)를 붙여주세요
    ' ex) ObsidianFolder = "C:\Users\YourName\Obsidian\Notes\Mail\"
    ObsidianFolder = "C:\Users\YourName\Obsidian\Notes\Mail\"
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
