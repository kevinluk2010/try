Option Explicit

'==============================================================================
' Auto-export "Confidential" mail to PDF
'
' WHERE THIS GOES
'   Classic Outlook for Windows only (the desktop app with File > Options >
'   Trust Center). Alt+F11 -> Project1 -> Microsoft Outlook Objects ->
'   ThisOutlookSession. Paste this whole file there, save, restart Outlook.
'   The "new Outlook" for Windows, Outlook on the web, and Outlook for Mac
'   have no VBA -- see README.md for the Power Automate route.
'
' WHAT IT DOES
'   Watches for messages Outlook considers confidential and writes a PDF of
'   each one into SAVE_FOLDER. The PDF is produced by saving the message as
'   MHTML and having Word export it as a fixed-format PDF, which keeps the
'   From/To/Sent/Subject header block that Outlook's own printout has.
'
' WHAT IT DELIBERATELY DOES NOT DO
'   Rights-protected mail (IRM / "Do Not Forward" / encrypted) is skipped by
'   default. Exporting it to a plain file removes exactly the protection the
'   sender applied, and may be blocked by your tenant anyway. Flip
'   SKIP_RIGHTS_PROTECTED only if you know your policy permits it.
'==============================================================================


'--- CONFIG -------------------------------------------------------------------

' Destination folder. %ENVVARS% are expanded. Created if missing.
Private Const SAVE_FOLDER As String = "%USERPROFILE%\Documents\Confidential PDFs"

' Match on the classic Outlook sensitivity flag (Options > Sensitivity > Confidential).
Private Const MATCH_SENSITIVITY_FLAG As Boolean = True

' Match on Microsoft Purview / AIP sensitivity label names. Semicolon-separated,
' case-insensitive substring match, so "Confidential" also matches
' "Confidential \ Internal". Leave "" to disable label matching.
Private Const MATCH_LABEL_NAMES As String = "Confidential"

' Optional: match specific label GUIDs instead of / as well as names, for tenants
' whose labels do not stamp a readable name. Semicolon-separated, no braces.
Private Const MATCH_LABEL_GUIDS As String = ""

' Skip encrypted / rights-managed mail. Read the header comment before changing.
Private Const SKIP_RIGHTS_PROTECTED As Boolean = True

' Triggers.
Private Const EXPORT_ON_OPEN As Boolean = True      ' opened in its own window
Private Const EXPORT_ON_PREVIEW As Boolean = False  ' selected in the reading pane
Private Const EXPORT_ON_ARRIVAL As Boolean = False  ' lands in the Inbox, unopened

' Optional log file for troubleshooting. "" disables logging.
Private Const LOG_FILE As String = ""

'--- END CONFIG ---------------------------------------------------------------


Private WithEvents olInspectors As Outlook.Inspectors
Private WithEvents olExplorer As Outlook.Explorer
Private WithEvents olInboxItems As Outlook.Items


'==============================================================================
' Wiring
'==============================================================================

Private Sub Application_Startup()
    HookEvents
End Sub

' Public so you can re-run it from the VBA editor (F5) without restarting Outlook.
Public Sub HookEvents()
    On Error Resume Next

    Set olInspectors = Application.Inspectors
    Set olExplorer = Application.ActiveExplorer

    If EXPORT_ON_ARRIVAL Then
        Set olInboxItems = Application.Session _
            .GetDefaultFolder(olFolderInbox).Items
    End If

    LogLine "hooked; save folder = " & ExpandEnv(SAVE_FOLDER)
End Sub

Private Sub olInspectors_NewInspector(ByVal Inspector As Outlook.Inspector)
    If Not EXPORT_ON_OPEN Then Exit Sub
    On Error Resume Next
    Dim itm As Object
    Set itm = Inspector.CurrentItem
    If itm Is Nothing Then Exit Sub
    If TypeOf itm Is Outlook.MailItem Then HandleMail itm
End Sub

Private Sub olExplorer_SelectionChange()
    If Not EXPORT_ON_PREVIEW Then Exit Sub
    On Error Resume Next
    Dim sel As Outlook.Selection
    Set sel = olExplorer.Selection
    If sel Is Nothing Then Exit Sub
    If sel.Count <> 1 Then Exit Sub
    If TypeOf sel.Item(1) Is Outlook.MailItem Then HandleMail sel.Item(1)
End Sub

Private Sub olInboxItems_ItemAdd(ByVal Item As Object)
    On Error Resume Next
    If TypeOf Item Is Outlook.MailItem Then HandleMail Item
End Sub


'==============================================================================
' Manual commands -- add these to the ribbon or run from the VBA editor
'==============================================================================

' Export whatever is selected in the message list, ignoring the Confidential
' test. Useful for backfilling mail that arrived before the macro existed.
Public Sub ExportSelectedToPdf()
    Dim sel As Outlook.Selection, i As Long, n As Long
    On Error Resume Next
    Set sel = Application.ActiveExplorer.Selection
    On Error GoTo 0

    If sel Is Nothing Then
        MsgBox "Select one or more messages first.", vbInformation
        Exit Sub
    ElseIf sel.Count = 0 Then
        MsgBox "Select one or more messages first.", vbInformation
        Exit Sub
    End If

    For i = 1 To sel.Count
        If TypeOf sel.Item(i) Is Outlook.MailItem Then
            If ExportMail(sel.Item(i), True) Then n = n + 1
        End If
    Next

    MsgBox n & " of " & sel.Count & " message(s) exported to" & vbCrLf & _
           ExpandEnv(SAVE_FOLDER), vbInformation
End Sub

' Tell me why a message did or did not match, without exporting it.
Public Sub WhyNotSelected()
    Dim m As Outlook.MailItem
    On Error Resume Next
    Set m = Application.ActiveExplorer.Selection.Item(1)
    On Error GoTo 0
    If m Is Nothing Then Exit Sub

    MsgBox "Subject:      " & m.Subject & vbCrLf & _
           "Sensitivity:  " & m.Sensitivity & "  (3 = Confidential)" & vbCrLf & _
           "Permission:   " & SafeGetPermission(m) & "  (0 = unrestricted)" & vbCrLf & _
           "MessageClass: " & m.MessageClass & vbCrLf & vbCrLf & _
           "Label string:" & vbCrLf & _
           IIf(Len(GetLabelString(m)) = 0, "(none)", GetLabelString(m)) & vbCrLf & vbCrLf & _
           "Confidential: " & IsConfidential(m) & vbCrLf & _
           "Protected:    " & IsRightsProtected(m), vbInformation
End Sub


'==============================================================================
' Core
'==============================================================================

Private Sub HandleMail(ByVal Mail As Outlook.MailItem)
    On Error Resume Next

    ' NewInspector also fires for a message you are COMPOSING. Marking a draft
    ' Confidential as you write it would otherwise export the half-finished
    ' text. .Sent is False only for drafts, so this keeps us to real mail.
    If Mail.Sent = False Then Exit Sub

    If Not IsConfidential(Mail) Then Exit Sub
    ExportMail Mail, False
End Sub

Private Function ExportMail(ByVal Mail As Outlook.MailItem, _
                            ByVal Forced As Boolean) As Boolean
    Dim dest As String, mht As String, pdf As String
    Dim fso As Object

    On Error GoTo Fail

    If SKIP_RIGHTS_PROTECTED And IsRightsProtected(Mail) Then
        LogLine "skipped (rights-protected): " & Mail.Subject
        Exit Function
    End If

    dest = ExpandEnv(SAVE_FOLDER)
    EnsureFolder dest

    pdf = dest & "\" & BuildFileName(Mail) & ".pdf"

    ' Deterministic name = free deduplication. Re-opening the same message
    ' does not produce a second copy.
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(pdf) Then
        LogLine "already exported: " & pdf
        ExportMail = Forced
        Exit Function
    End If

    mht = ExpandEnv("%TEMP%") & "\olpdf_" & Format$(Now, "yyyymmdd_hhnnss") & _
          "_" & CLng(Rnd() * 100000) & ".mht"

    Mail.SaveAs mht, 10                 ' 10 = olMHTML
    MhtToPdf mht, pdf

    On Error Resume Next
    fso.DeleteFile mht, True
    On Error GoTo Fail

    LogLine "exported: " & pdf
    ExportMail = True
    Exit Function

Fail:
    Dim eNum As Long, eDesc As String, subj As String
    eNum = Err.Number
    eDesc = Err.Description

    On Error Resume Next
    subj = Mail.Subject

    LogLine "FAILED (" & eNum & " " & eDesc & "): " & subj
    If Len(mht) > 0 Then CreateObject("Scripting.FileSystemObject").DeleteFile mht, True

    If Forced Then
        MsgBox "Could not export:" & vbCrLf & subj & vbCrLf & vbCrLf & _
               eNum & " - " & eDesc, vbExclamation
    End If
End Function


'==============================================================================
' Is this message confidential?
'==============================================================================

Private Function IsConfidential(ByVal Mail As Outlook.MailItem) As Boolean
    Dim lbl As String, parts As Variant, i As Long

    On Error Resume Next

    If MATCH_SENSITIVITY_FLAG Then
        If Mail.Sensitivity = olConfidential Then
            IsConfidential = True
            Exit Function
        End If
    End If

    lbl = GetLabelString(Mail)
    If Len(lbl) = 0 Then Exit Function

    If Len(MATCH_LABEL_NAMES) > 0 Then
        parts = Split(MATCH_LABEL_NAMES, ";")
        For i = LBound(parts) To UBound(parts)
            If Len(Trim$(parts(i))) > 0 Then
                ' _Name=<label> is the readable part of the msip_labels blob.
                If InStr(1, lbl, "_Name=" & Trim$(parts(i)), vbTextCompare) > 0 Then
                    IsConfidential = True
                    Exit Function
                End If
            End If
        Next
    End If

    If Len(MATCH_LABEL_GUIDS) > 0 Then
        parts = Split(MATCH_LABEL_GUIDS, ";")
        For i = LBound(parts) To UBound(parts)
            If Len(Trim$(parts(i))) > 0 Then
                If InStr(1, lbl, Trim$(parts(i)), vbTextCompare) > 0 Then
                    IsConfidential = True
                    Exit Function
                End If
            End If
        Next
    End If
End Function

' The Purview/AIP label lives in a named MAPI property. Received mail carries it
' in the internet-headers namespace; locally created items in public strings.
Private Function GetLabelString(ByVal Mail As Outlook.MailItem) As String
    Const HDR As String = "http://schemas.microsoft.com/mapi/string/" & _
        "{00020386-0000-0000-C000-000000000046}/msip_labels"
    Const PUB As String = "http://schemas.microsoft.com/mapi/string/" & _
        "{00020329-0000-0000-C000-000000000046}/msip_labels"

    Dim s As String
    On Error Resume Next

    s = CStr(Mail.PropertyAccessor.GetProperty(HDR))
    If Err.Number <> 0 Or Len(s) = 0 Then
        Err.Clear
        s = CStr(Mail.PropertyAccessor.GetProperty(PUB))
    End If
    Err.Clear

    GetLabelString = s
End Function

Private Function IsRightsProtected(ByVal Mail As Outlook.MailItem) As Boolean
    Dim mc As String

    On Error Resume Next

    ' olUnrestricted = 0, olDoNotForward = 1, olPermissionTemplate = 2
    If SafeGetPermission(Mail) <> 0 Then
        IsRightsProtected = True
        Exit Function
    End If

    mc = Mail.MessageClass
    If InStr(1, mc, "rpmsg", vbTextCompare) > 0 Then
        IsRightsProtected = True
        Exit Function
    End If

    ' S/MIME encrypted. MultipartSigned is signed-only and exports fine.
    If InStr(1, mc, "IPM.Note.SMIME", vbTextCompare) = 1 And _
       InStr(1, mc, "MultipartSigned", vbTextCompare) = 0 Then
        IsRightsProtected = True
    End If
End Function

Private Function SafeGetPermission(ByVal Mail As Outlook.MailItem) As Long
    On Error Resume Next
    SafeGetPermission = Mail.Permission
    If Err.Number <> 0 Then
        Err.Clear
        SafeGetPermission = 0
    End If
End Function


'==============================================================================
' MHTML -> PDF via Word
'==============================================================================

Private Sub MhtToPdf(ByVal mhtPath As String, ByVal pdfPath As String)
    Const wdExportFormatPDF As Long = 17
    Const wdOpenFormatWebPages As Long = 7
    Const wdDoNotSaveChanges As Long = 0

    Dim wd As Object, doc As Object, createdWord As Boolean
    Dim errNum As Long, errDesc As String

    On Error Resume Next
    Set wd = GetObject(, "Word.Application")
    On Error GoTo 0

    If wd Is Nothing Then
        Set wd = CreateObject("Word.Application")
        createdWord = True
        wd.Visible = False
    End If

    On Error GoTo Cleanup

    Set doc = wd.Documents.Open(FileName:=mhtPath, _
                                ConfirmConversions:=False, _
                                ReadOnly:=True, _
                                AddToRecentFiles:=False, _
                                Format:=wdOpenFormatWebPages, _
                                Visible:=False)

    doc.ExportAsFixedFormat OutputFileName:=pdfPath, _
                            ExportFormat:=wdExportFormatPDF, _
                            OpenAfterExport:=False

Cleanup:
    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    ' Only shut down Word if we were the ones who started it.
    If createdWord Then wd.Quit wdDoNotSaveChanges
    Set doc = Nothing
    Set wd = Nothing
    On Error GoTo 0

    If errNum <> 0 Then Err.Raise errNum, "MhtToPdf", errDesc
End Sub


'==============================================================================
' Helpers
'==============================================================================

Private Function BuildFileName(ByVal Mail As Outlook.MailItem) As String
    Dim stamp As Date, who As String, subj As String

    On Error Resume Next

    stamp = Mail.ReceivedTime
    If Err.Number <> 0 Or stamp = 0 Then
        Err.Clear
        stamp = Mail.CreationTime
    End If
    If stamp = 0 Then stamp = Now

    who = Mail.SenderName
    If Len(who) = 0 Then who = Mail.SenderEmailAddress
    If Len(who) = 0 Then who = "unknown"

    subj = Mail.Subject
    Err.Clear

    BuildFileName = Format$(stamp, "yyyy-mm-dd hhnnss") & " - " & _
                    SafeName(who, 40) & " - " & SafeName(subj, 80)
End Function

Private Function SafeName(ByVal s As String, ByVal maxLen As Long) As String
    Dim bad As Variant, i As Long

    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbCr, vbLf, vbTab)
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), "_")
    Next

    s = Trim$(s)
    Do While Len(s) > 0
        If Right$(s, 1) = "." Or Right$(s, 1) = " " Then
            s = Left$(s, Len(s) - 1)
        Else
            Exit Do
        End If
    Loop

    If Len(s) > maxLen Then s = RTrim$(Left$(s, maxLen))
    If Len(s) = 0 Then s = "message"

    SafeName = s
End Function

Private Function ExpandEnv(ByVal p As String) As String
    On Error Resume Next
    ExpandEnv = CreateObject("WScript.Shell").ExpandEnvironmentStrings(p)
    If Err.Number <> 0 Or Len(ExpandEnv) = 0 Then
        Err.Clear
        ExpandEnv = p
    End If
End Function

Private Sub EnsureFolder(ByVal p As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    EnsureFolderRec fso, p
End Sub

Private Sub EnsureFolderRec(ByVal fso As Object, ByVal p As String)
    Dim parent As String

    If Len(p) = 0 Then Exit Sub
    If fso.FolderExists(p) Then Exit Sub

    parent = fso.GetParentFolderName(p)
    If Len(parent) = 0 Then Exit Sub          ' hit a drive or UNC root

    If Not fso.FolderExists(parent) Then EnsureFolderRec fso, parent
    fso.CreateFolder p
End Sub

Private Sub LogLine(ByVal msg As String)
    If Len(LOG_FILE) = 0 Then Exit Sub

    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open ExpandEnv(LOG_FILE) For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  " & msg
    Close #f
End Sub
