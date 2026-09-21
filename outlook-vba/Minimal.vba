Option Explicit

'==============================================================================
' Minimal version: open a Confidential email -> PDF in a folder.
'
' Alt+F11 -> Project1 -> Microsoft Outlook Objects -> ThisOutlookSession.
' Paste this, change SAVE_FOLDER, save, restart Outlook.
'
' Matches the classic Outlook sensitivity flag only
' (message Options -> Sensitivity -> Confidential).
' For Purview / AIP labels, rights-protected mail, arrival triggers and
' diagnostics, use ThisOutlookSession.vba instead.
'==============================================================================

Private Const SAVE_FOLDER As String = "C:\Users\YourName\Documents\Confidential PDFs"

Private WithEvents olInspectors As Outlook.Inspectors


Private Sub Application_Startup()
    Set olInspectors = Application.Inspectors
End Sub


Private Sub olInspectors_NewInspector(ByVal Inspector As Outlook.Inspector)
    On Error Resume Next

    Dim m As Object
    Set m = Inspector.CurrentItem
    If m Is Nothing Then Exit Sub
    If Not TypeOf m Is Outlook.MailItem Then Exit Sub

    ' This event also fires for messages you are COMPOSING. .Sent is False
    ' only for drafts -- without this you would export half-written mail.
    If m.Sent = False Then Exit Sub

    If m.Sensitivity <> olConfidential Then Exit Sub

    SaveAsPdf m
End Sub


Private Sub SaveAsPdf(ByVal m As Outlook.MailItem)
    Const olMHTML As Long = 10
    Const wdExportFormatPDF As Long = 17
    Const wdOpenFormatWebPages As Long = 7
    Const wdDoNotSaveChanges As Long = 0

    Dim fso As Object, wd As Object, doc As Object
    Dim startedWord As Boolean
    Dim baseName As String, mht As String, pdf As String

    On Error GoTo Done

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(SAVE_FOLDER) Then fso.CreateFolder SAVE_FOLDER

    baseName = Format$(m.ReceivedTime, "yyyy-mm-dd hhnnss") & " - " & _
           SafeName(m.SenderName, 40) & " - " & SafeName(m.Subject, 70)

    ' Same message reopened = same filename = no duplicate.
    pdf = SAVE_FOLDER & "\" & baseName & ".pdf"
    If fso.FileExists(pdf) Then Exit Sub

    ' Outlook cannot write PDF directly, so: save as MHTML, let Word export it.
    ' MHTML keeps the From / Sent / To / Subject header block.
    mht = Environ$("TEMP") & "\" & baseName & ".mht"
    m.SaveAs mht, olMHTML

    ' Reuse a running Word if there is one; only shut down a Word we started,
    ' otherwise we would close the user's open documents.
    On Error Resume Next
    Set wd = GetObject(, "Word.Application")
    On Error GoTo Done
    If wd Is Nothing Then
        Set wd = CreateObject("Word.Application")
        startedWord = True
        wd.Visible = False
    End If

    Set doc = wd.Documents.Open(FileName:=mht, _
                                ConfirmConversions:=False, _
                                ReadOnly:=True, _
                                AddToRecentFiles:=False, _
                                Format:=wdOpenFormatWebPages, _
                                Visible:=False)

    doc.ExportAsFixedFormat OutputFileName:=pdf, _
                            ExportFormat:=wdExportFormatPDF, _
                            OpenAfterExport:=False

Done:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    If startedWord Then wd.Quit wdDoNotSaveChanges
    If Len(mht) > 0 Then fso.DeleteFile mht, True
End Sub


Private Function SafeName(ByVal s As String, ByVal maxLen As Long) As String
    Dim c As Variant

    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbCr, vbLf, vbTab)
        s = Replace(s, c, "_")
    Next

    s = Trim$(s)
    If Len(s) > maxLen Then s = RTrim$(Left$(s, maxLen))
    If Len(s) = 0 Then s = "message"

    SafeName = s
End Function
