Option Explicit

'==============================================================================
' Auto-export "Confidential" mail to PDF
'
' WHERE THIS GOES
'   Classic Outlook for Windows only. Alt+F11 -> Project1 ->
'   Microsoft Outlook Objects -> ThisOutlookSession. Paste this whole file
'   there, save, restart Outlook.
'
' IF IT DOES NOT WORK
'   Select a confidential email in Outlook, then in the VBA editor press F5
'   and run RunDiagnostics. It reports which step is failing.
'   Everything is also logged to LOG_FILE (see CONFIG). OpenLog shows it.
'==============================================================================


'--- CONFIG -------------------------------------------------------------------

' Destination folder. %ENVVARS% are expanded. Created if missing.
Private Const SAVE_FOLDER As String = "C:\Users\lcy048\Downloads\Telegram Desktop"

' Match on the classic Outlook sensitivity flag (Options > Sensitivity > Confidential).
Private Const MATCH_SENSITIVITY_FLAG As Boolean = True

' Match on Microsoft Purview / AIP sensitivity label names. Semicolon-separated,
' case-insensitive substring match, so "Confidential" also matches
' "Confidential \ Internal". Leave "" to disable label matching.
Private Const MATCH_LABEL_NAMES As String = "Confidential"

' Optional: match specific label GUIDs as well as names, for tenants whose
' labels do not stamp a readable name. Semicolon-separated, no braces.
Private Const MATCH_LABEL_GUIDS As String = ""

' Skip encrypted / rights-managed mail.
Private Const SKIP_RIGHTS_PROTECTED As Boolean = True

' Triggers.
Private Const EXPORT_ON_OPEN As Boolean = True      ' opened in its own window
Private Const EXPORT_ON_PREVIEW As Boolean = False  ' selected in the reading pane
Private Const EXPORT_ON_ARRIVAL As Boolean = False  ' lands in the Inbox, unopened

' Log file. ON by default -- without it, failures are invisible.
Private Const LOG_FILE As String = "%TEMP%\outlook-pdf-export.log"

'--- END CONFIG ---------------------------------------------------------------


Private WithEvents olInspectors As Outlook.Inspectors
Private WithEvents olExplorer As Outlook.Explorer
Private WithEvents olInboxItems As Outlook.Items

' Why the last export attempt did what it did. Read by RunDiagnostics.
Private gLastReason As String


'==============================================================================
' Wiring
'==============================================================================

Private Sub Application_Startup()
    HookEvents
End Sub

' Public so you can run it with F5 instead of restarting Outlook.
Public Sub HookEvents()
    On Error GoTo Fail

    Set olInspectors = Application.Inspectors
    Set olExplorer = Application.ActiveExplorer

    If EXPORT_ON_ARRIVAL Then
        Set olInboxItems = Application.Session.GetDefaultFolder(olFolderInbox).Items
    End If

    LogLine "--- hooked. dest=" & ExpandEnv(SAVE_FOLDER) & _
            " onOpen=" & EXPORT_ON_OPEN & " onPreview=" & EXPORT_ON_PREVIEW & _
            " onArrival=" & EXPORT_ON_ARRIVAL
    Exit Sub

Fail:
    LogLine "HookEvents ERROR " & Err.Number & ": " & Err.Description
End Sub

Private Sub olInspectors_NewInspector(ByVal Inspector As Outlook.Inspector)
    If Not EXPORT_ON_OPEN Then Exit Sub

    Dim itm As Object
    On Error GoTo Fail

    Set itm = Inspector.CurrentItem
    If itm Is Nothing Then
        LogLine "NewInspector fired but CurrentItem is Nothing"
        Exit Sub
    End If

    If Not TypeOf itm Is Outlook.MailItem Then Exit Sub

    LogLine "NewInspector: mail opened"
    HandleMail itm
    Exit Sub

Fail:
    LogLine "NewInspector ERROR " & Err.Number & ": " & Err.Description
End Sub

Private Sub olExplorer_SelectionChange()
    If Not EXPORT_ON_PREVIEW Then Exit Sub

    Dim sel As Outlook.Selection
    On Error GoTo Fail

    Set sel = olExplorer.Selection
    If sel Is Nothing Then Exit Sub
    If sel.Count <> 1 Then Exit Sub
    If Not TypeOf sel.Item(1) Is Outlook.MailItem Then Exit Sub

    HandleMail sel.Item(1)
    Exit Sub

Fail:
    LogLine "SelectionChange ERROR " & Err.Number & ": " & Err.Description
End Sub

Private Sub olInboxItems_ItemAdd(ByVal Item As Object)
    On Error GoTo Fail

    If Not TypeOf Item Is Outlook.MailItem Then Exit Sub

    LogLine "ItemAdd: mail arrived"
    HandleMail Item
    Exit Sub

Fail:
    LogLine "ItemAdd ERROR " & Err.Number & ": " & Err.Description
End Sub


'==============================================================================
' DIAGNOSTICS -- run this with F5 when nothing is being saved
'==============================================================================

Public Sub RunDiagnostics()
    Dim r As String, dest As String, tmp As String
    Dim fso As Object, wd As Object, startedWord As Boolean
    Dim m As Object, sel As Outlook.Selection

    r = "PDF EXPORT DIAGNOSTICS" & vbCrLf & _
        "----------------------------------------" & vbCrLf & vbCrLf

    '-- 1. Did Application_Startup run? ---------------------------------------
    If olInspectors Is Nothing Then
        r = r & "1. Startup hook: NOT RUNNING" & vbCrLf & _
                "   Application_Startup never fired, so opening a mail" & vbCrLf & _
                "   triggers nothing. Cause is one of:" & vbCrLf & _
                "     - macros blocked in Trust Center, or" & vbCrLf & _
                "     - Outlook not restarted after saving the code." & vbCrLf & _
                "   Workaround right now: run HookEvents (F5)." & vbCrLf
    Else
        r = r & "1. Startup hook: running" & vbCrLf
    End If

    '-- 2. Destination folder -------------------------------------------------
    dest = ExpandEnv(SAVE_FOLDER)
    On Error Resume Next
    Err.Clear
    EnsureFolder dest
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(dest) Then
        r = r & "2. Folder: CANNOT CREATE" & vbCrLf & "   " & dest & vbCrLf
    Else
        tmp = dest & "\_writetest.tmp"
        Err.Clear
        fso.CreateTextFile(tmp, True).Close
        If Err.Number <> 0 Then
            r = r & "2. Folder: NOT WRITABLE (" & Err.Number & " " & _
                    Err.Description & ")" & vbCrLf
            Err.Clear
        Else
            fso.DeleteFile tmp, True
            r = r & "2. Folder: writable" & vbCrLf
        End If
    End If

    '-- 3. Word automation ----------------------------------------------------
    Err.Clear
    Set wd = GetObject(, "Word.Application")
    If wd Is Nothing Then
        Err.Clear
        Set wd = CreateObject("Word.Application")
        startedWord = True
    End If

    If wd Is Nothing Then
        r = r & "3. Word: UNAVAILABLE (" & Err.Number & " " & Err.Description & ")" & vbCrLf & _
                "   No Word means no PDF -- this method needs it." & vbCrLf
        Err.Clear
    Else
        r = r & "3. Word: version " & wd.Version & " OK" & vbCrLf
        If startedWord Then wd.Quit 0
        Set wd = Nothing
    End If

    '-- 4. The selected message ----------------------------------------------
    Err.Clear
    Set sel = Application.ActiveExplorer.Selection
    If Not sel Is Nothing Then
        If sel.Count > 0 Then Set m = sel.Item(1)
    End If
    Err.Clear

    If m Is Nothing Then
        r = r & vbCrLf & "4. No message selected." & vbCrLf & _
                "   Select a confidential email, then run this again." & vbCrLf
    ElseIf Not TypeOf m Is Outlook.MailItem Then
        r = r & vbCrLf & "4. Selected item is not an email." & vbCrLf
    Else
        r = r & vbCrLf & "4. Selected message:" & vbCrLf & _
                "   Sensitivity  = " & m.Sensitivity & "   (3 = Confidential)" & vbCrLf & _
                "   Permission   = " & SafeGetPermission(m) & "   (0 = unrestricted)" & vbCrLf & _
                "   MessageClass = " & m.MessageClass & vbCrLf & _
                "   Label        = " & Left$(IIf(Len(GetLabelString(m)) = 0, _
                                        "(none)", GetLabelString(m)), 120) & vbCrLf & _
                "   -> confidential? " & IsConfidential(m) & vbCrLf & _
                "   -> protected?    " & IsRightsProtected(m) & vbCrLf

        '-- 5. Real export attempt -------------------------------------------
        gLastReason = ""
        ExportMail m, False
        r = r & vbCrLf & "5. Export attempt:" & vbCrLf & "   " & gLastReason & vbCrLf
    End If

    r = r & vbCrLf & "Log: " & ExpandEnv(LOG_FILE)

    LogLine "DIAGNOSTICS" & vbCrLf & r
    If Len(r) > 1020 Then r = Left$(r, 1020) & vbCrLf & "... (full text in log)"
    MsgBox r, vbInformation, "PDF export diagnostics"
End Sub

Public Sub OpenLog()
    On Error Resume Next
    Shell "notepad.exe """ & ExpandEnv(LOG_FILE) & """", vbNormalFocus
End Sub


'==============================================================================
' Manual commands -- add these to the ribbon or run from the VBA editor
'==============================================================================

' Export whatever is selected, ignoring the Confidential test. Use it to
' backfill mail that arrived before the macro existed.
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


'==============================================================================
' Core
'==============================================================================

Private Sub HandleMail(ByVal Mail As Outlook.MailItem)
    On Error GoTo Fail

    ' NewInspector also fires for a message you are COMPOSING. Marking a draft
    ' Confidential as you write it would otherwise export half-finished text.
    ' .Sent is False only for drafts.
    If Mail.Sent = False Then
        LogLine "skipped: draft being composed"
        Exit Sub
    End If

    If Not IsConfidential(Mail) Then
        LogLine "skipped: not confidential. Sensitivity=" & Mail.Sensitivity & _
                " label=" & Left$(GetLabelString(Mail), 120)
        Exit Sub
    End If

    ExportMail Mail, False
    Exit Sub

Fail:
    LogLine "HandleMail ERROR " & Err.Number & ": " & Err.Description
End Sub

Private Function ExportMail(ByVal Mail As Outlook.MailItem, _
                            ByVal Forced As Boolean) As Boolean
    Dim dest As String, mht As String, pdf As String
    Dim fso As Object

    On Error GoTo Fail

    If SKIP_RIGHTS_PROTECTED And IsRightsProtected(Mail) Then
        gLastReason = "SKIPPED: rights-protected (encrypted / Do Not Forward)." & _
                      " Set SKIP_RIGHTS_PROTECTED = False to attempt anyway."
        LogLine gLastReason
        Exit Function
    End If

    dest = ExpandEnv(SAVE_FOLDER)
    EnsureFolder dest

    pdf = dest & "\" & BuildFileName(Mail) & ".pdf"

    ' Deterministic name = free deduplication. Re-opening the same message
    ' does not produce a second copy.
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(pdf) Then
        gLastReason = "SKIPPED: PDF already exists -- " & pdf
        LogLine gLastReason
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

    gLastReason = "OK: wrote " & pdf
    LogLine gLastReason
    ExportMail = True
    Exit Function

Fail:
    Dim eNum As Long, eDesc As String, subj As String
    eNum = Err.Number
    eDesc = Err.Description

    On Error Resume Next
    subj = Mail.Subject

    gLastReason = "ERROR " & eNum & ": " & eDesc
    LogLine gLastReason & "  [" & subj & "]"

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
