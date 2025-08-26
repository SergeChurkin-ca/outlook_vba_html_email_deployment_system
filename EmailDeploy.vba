Option Explicit

'=============================
'  MAIN
'=============================
Sub LaunchEmailDeploy()
  Dim filePath As String
  Dim isMac As Boolean
  isMac = (InStr(1, Application.OperatingSystem, "Macintosh", vbTextCompare) > 0)

  ' 1) Pick HTML template
  If isMac Then
    On Error Resume Next
    filePath = MacScript("POSIX path of (choose file of type {""public.html""} with prompt ""Select HTML Template"")")
    On Error GoTo 0
    If Len(filePath) = 0 Then Exit Sub
  Else
    Dim tmp As Variant
    tmp = Application.GetOpenFilename("HTML Files (*.html), *.html", 1, "Select HTML Template")
    If VarType(tmp) = vbBoolean Then Exit Sub
    filePath = tmp
  End If

  ' 2) Preview in default browser
  ThisWorkbook.FollowHyperlink Address:=filePath, NewWindow:=True

  ' 3) Read HTML (plain VBA I/O)
  Dim rawHtml As String, oneLine As String
  Dim f As Integer: f = FreeFile
  Open filePath For Input As #f
  Do While Not EOF(f)
    Line Input #f, oneLine
    rawHtml = rawHtml & oneLine & vbCrLf
  Loop
  Close #f

  ' 4) Count recipients (for typed confirm)
  Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
  Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

  Dim i As Long, totalFound As Long, sample As String, shown As Long
  For i = 2 To lastRow
    Dim e As String: e = Trim(ws.Cells(i, "C").Value)
    If e <> "" Then
      totalFound = totalFound + 1
      If shown < 5 Then sample = sample & "• " & e & vbCrLf: shown = shown + 1
    End If
  Next i
  If totalFound = 0 Then
    MsgBox "No email addresses found in column C.", vbExclamation, "Nothing to send"
    Exit Sub
  End If

  If Not ConfirmSendTyped(totalFound, sample, Dir(filePath)) Then Exit Sub

  ' 5) Prep Outlook (Windows)
  Dim olApp As Object
  If Not isMac Then
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If olApp Is Nothing Then
      MsgBox "Outlook is not available. Please open Outlook and try again.", vbExclamation, "Outlook not running"
      Exit Sub
    End If
  End If

  ' 6) Loop & send with progress + robust summary
  Dim startT As Date: startT = Now
  Dim attempted As Long, sent As Long, skippedBlank As Long, failed As Long
  Dim errList As Collection: Set errList = New Collection

  Application.StatusBar = "Email deployment started…"
  For i = 2 To lastRow
    Dim toAddr As String: toAddr = Trim(ws.Cells(i, "C").Value)
    If toAddr = "" Then
      skippedBlank = skippedBlank + 1
      GoTo NextRow
    End If

    attempted = attempted + 1
    Dim fn As String:   fn = ws.Cells(i, "A").Value
    Dim ln As String:   ln = ws.Cells(i, "B").Value
    Dim subj As String: subj = ws.Cells(i, "D").Value

    Dim bodyHtml As String
    bodyHtml = Replace(rawHtml, "[First Name]", fn, 1, -1, vbTextCompare)
    bodyHtml = Replace(bodyHtml, "[Last Name]", ln, 1, -1, vbTextCompare)

    On Error Resume Next
    If isMac Then
      ' ---- Mac: AppleScript to Outlook (auto-send) ----
      Dim q As String: q = Chr(34)
      Dim s As String
      subj = Replace(subj, q, "\" & q)
      bodyHtml = Replace(bodyHtml, q, "\" & q)
      s = "tell application " & q & "Microsoft Outlook" & q & vbLf
      s = s & "set msg to make new outgoing message with properties {subject:" & q & subj & q & ", content:" & q & bodyHtml & q & "}" & vbLf
      s = s & "tell msg to make new recipient at end of to recipients with properties {email address:{address:" & q & toAddr & q & "}}" & vbLf
      s = s & "send msg" & vbLf & "end tell"
      MacScript s

      If Err.Number <> 0 Then
        failed = failed + 1
        errList.Add "Row " & i & " (" & toAddr & "): " & Err.Description
        Err.Clear
      Else
        sent = sent + 1
      End If

    Else
      ' ---- Windows: COM to Outlook (auto-send, choose account optional) ----
      Dim olMail As Object, acct As Object
      Set olMail = olApp.CreateItem(0)

      ' OPTIONAL: force a specific From account (leave blank to skip)
      Const WINDOWS_FROM_ADDRESS As String = ""    ' e.g. "you@yourcompany.com"
      If Len(WINDOWS_FROM_ADDRESS) > 0 Then
        Set acct = FindAccountBySmtp(olApp, WINDOWS_FROM_ADDRESS)
        If Not acct Is Nothing Then
          olMail.SendUsingAccount = acct   ' no error if set fails; Outlook will use default
        End If
      End If

      With olMail
        .To = toAddr
        .Subject = subj
        .htmlBody = bodyHtml
        .Send                      ' auto-send (use .Display only for debugging)
      End With

      If Err.Number <> 0 Then
        failed = failed + 1
        errList.Add "Row " & i & " (" & toAddr & "): " & Err.Description
        Err.Clear
      Else
        sent = sent + 1
      End If
    End If
    On Error GoTo 0

NextRow:
    If attempted Mod 10 = 0 Then
      Application.StatusBar = "Sending… " & attempted & " of " & totalFound & " (" & _
                              Format(attempted / totalFound, "0%") & ")"
      DoEvents
    End If
  Next i
  Application.StatusBar = False

  ' Push anything stuck in Outbox (Windows)
  If Not isMac Then KickSendReceive olApp

  ' 7) Summary
  Dim msg As String, duration As String
  duration = Format(Now - startT, "hh:nn:ss")
  msg = "Email deployment finished." & vbCrLf & _
        "Template: " & Dir(filePath) & vbCrLf & _
        "Duration: " & duration & vbCrLf & vbCrLf & _
        "Rows scanned: " & (lastRow - 1) & vbCrLf & _
        "Recipients found: " & totalFound & vbCrLf & _
        "Sent successfully: " & sent & vbCrLf & _
        "Skipped (blank email): " & skippedBlank & vbCrLf & _
        "Failed to send: " & failed

  If failed > 0 Then
    msg = msg & vbCrLf & vbCrLf & "First errors:" & vbCrLf & ListTop(errList, 5)
  End If

  MsgBox msg, vbInformation, "Email deployment — summary"
End Sub

'=============================
'  Typed confirmation dialog
'=============================
Private Function ConfirmSendTyped(total As Long, sample As String, templateName As String) As Boolean
  Dim prompt As String, resp As String
  prompt = "About to send " & total & " emails using template:" & vbCrLf & _
           "  " & templateName & vbCrLf & vbCrLf & _
           "First few recipients:" & vbCrLf & sample & vbCrLf & _
           "Type  " & Chr(34) & "SEND" & Chr(34) & "  to proceed."
  resp = InputBox(prompt, "Final confirmation")
  ConfirmSendTyped = (UCase$(Trim$(resp)) = "SEND")
End Function

'=============================
'  Helper: list first N items
'=============================
Private Function ListTop(col As Collection, ByVal n As Long) As String
  Dim i As Long, s As String
  If col Is Nothing Then Exit Function
  If n > col.Count Then n = col.Count
  For i = 1 To n
    s = s & "• " & col(i) & vbCrLf
  Next
  ListTop = s
End Function

'=============================
'  Windows helpers
'=============================
' Try to force a send/receive (helps push Outbox on some configs)
Private Sub KickSendReceive(ByVal olApp As Object)
  On Error Resume Next
  olApp.Session.SendAndReceive False
  If Err.Number <> 0 Then
    Err.Clear
    Dim i As Long
    For i = 1 To olApp.Session.SyncObjects.Count
      olApp.Session.SyncObjects.Item(i).Start
    Next
  End If
  On Error GoTo 0
End Sub

' Find a specific Outlook account by SMTP (optional From selection)
Private Function FindAccountBySmtp(ByVal olApp As Object, ByVal smtp As String) As Object
  On Error Resume Next
  Dim accts As Object, i As Long
  Set accts = olApp.Session.Accounts
  For i = 1 To accts.Count
    If LCase$(accts.Item(i).SmtpAddress) = LCase$(smtp) Then
      Set FindAccountBySmtp = accts.Item(i)
      Exit Function
    End If
  Next
  On Error GoTo 0
End Function

'=============================
'  One-time: add styled button
'=============================
Sub InstallStyledButton()
  Dim sh As Shape
  Set sh = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 240, 60, 300, 56)
  With sh
    .Name = "btnUploadPreview"
    On Error Resume Next
    .TextFrame2.TextRange.Characters.Text = "Upload & Preview (Step 1)"
    .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    With .TextFrame2.TextRange.Characters.Font
      .Bold = msoTrue
      .Size = 16
      .Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    On Error GoTo 0
    .Fill.ForeColor.RGB = RGB(26, 115, 232)
    .line.ForeColor.RGB = RGB(12, 64, 166)  ' << fix: .Line (not .line)
    .line.Weight = 2
    .Shadow.Visible = msoTrue
    .OnAction = "LaunchEmailDeploy"
  End With
  MsgBox "Button created. Drag it where you like.", vbInformation
End Sub


