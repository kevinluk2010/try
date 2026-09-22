'==============================================================================
' QuickCheck -- paste this at the END of the code already in ThisOutlookSession.
'
' It reuses the helpers already there, so nothing else needs changing.
' Select a confidential email in Outlook, then put the cursor inside
' QuickCheck in the VBA editor and press F5.
'==============================================================================

Public Sub QuickCheck()
    Dim r As String, dest As String, tmp As String
    Dim fso As Object, wd As Object, madeWord As Boolean
    Dim sel As Outlook.Selection, m As Object

    '-- 1. Did Application_Startup actually run? ------------------------------
    If olInspectors Is Nothing Then
        r = "1. Startup hook: NOT RUNNING  <-- this is the problem" & vbCrLf & _
            "   Macros are blocked in Trust Center, or Outlook was" & vbCrLf & _
            "   not restarted after saving." & vbCrLf
    Else
        r = "1. Startup hook: running" & vbCrLf
    End If

    '-- 2. Is the destination folder actually writable? -----------------------
    On Error Resume Next
    dest = ExpandEnv(SAVE_FOLDER)
    EnsureFolder dest
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(dest) Then
        r = r & "2. Folder: MISSING / cannot create" & vbCrLf & "   " & dest & vbCrLf
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

    '-- 3. Can we drive Word? (no Word = no PDF) ------------------------------
    Err.Clear
    Set wd = GetObject(, "Word.Application")
    If wd Is Nothing Then
        Err.Clear
        Set wd = CreateObject("Word.Application")
        madeWord = True
    End If

    If wd Is Nothing Then
        r = r & "3. Word: UNAVAILABLE (" & Err.Number & " " & Err.Description & ")" & vbCrLf
        Err.Clear
    Else
        r = r & "3. Word: version " & wd.Version & " OK" & vbCrLf
        If madeWord Then wd.Quit 0
        Set wd = Nothing
    End If

    '-- 4. What does the selected message actually look like? -----------------
    Err.Clear
    Set sel = Application.ActiveExplorer.Selection
    If Not sel Is Nothing Then
        If sel.Count > 0 Then Set m = sel.Item(1)
    End If
    Err.Clear

    If m Is Nothing Then
        r = r & vbCrLf & "4. Nothing selected." & vbCrLf & _
                "   Select a confidential email and run this again."
    ElseIf Not TypeOf m Is Outlook.MailItem Then
        r = r & vbCrLf & "4. Selected item is not an email."
    Else
        r = r & vbCrLf & "4. Selected message:" & vbCrLf & _
            "   Sensitivity  = " & m.Sensitivity & "   (3 = Confidential)" & vbCrLf & _
            "   Permission   = " & SafeGetPermission(m) & "   (0 = unrestricted)" & vbCrLf & _
            "   MessageClass = " & m.MessageClass & vbCrLf & _
            "   Label        = " & IIf(Len(GetLabelString(m)) = 0, "(none)", _
                                       Left$(GetLabelString(m), 100)) & vbCrLf & _
            "   confidential? " & IsConfidential(m) & vbCrLf & _
            "   protected?    " & IsRightsProtected(m)
    End If

    MsgBox r, vbInformation, "QuickCheck"

    '-- 5. Now really try it. Any failure raises its own message box. ---------
    If Not m Is Nothing Then
        If TypeOf m Is Outlook.MailItem Then ExportMail m, True
    End If
End Sub
