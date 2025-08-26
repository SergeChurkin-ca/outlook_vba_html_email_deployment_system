Option Explicit

Sub LaunchEmailDeploy()
  Dim filePath As String
  Dim isMac As Boolean
  isMac = (InStr(1, Application.OperatingSystem, "Macintosh", vbTextCompare) > 0)

  ' 1) PICK HTML template
  If isMac Then
    On Error Resume Next
    filePath = MacScript( _
      "POSIX path of (choose file of type {""public.html""} with prompt ""Select HTML Template"")")
    On Error GoTo 0
    If Len(filePath) = 0 Then Exit Sub
  Else
    Dim tmp As Variant
    tmp = Application.GetOpenFilename("HTML Files (*.html), *.html", 1, "Select HTML Template")
    If VarType(tmp) = vbBoolean Then Exit Sub
    filePath = tmp
  End If

  ' 2) PREVIEW in default browser
  ThisWorkbook.FollowHyperlink Address:=filePath, NewWindow:=True

  ' 3) READ raw HTML (plain VBA I/O)
  Dim rawHtml As String, textLine As String
  Dim fNum As Integer: fNum = FreeFile
  Open filePath For Input As #fNum
    Do While Not EOF(fNum)
      Line Input #fNum, textLine
      rawHtml = rawHtml & textLine & vbCrLf
    Loop
  Close #fNum

  ' 4) CONFIRM send
  If MsgBox("Send emails using this template?", vbYesNo + vbQuestion, "Confirm Send") <> vbYes Then Exit Sub

  ' (Windows) ensure Outlook object
  Dim olApp As Object
  If Not isMac Then
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If olApp Is Nothing Then
      MsgBox "Outlook is not available. Please open Outlook and try again.", vbExclamation
      Exit Sub
    End If
  End If

  ' 5) LOOP rows & send
  Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
  Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
  Dim i As Long, sentCount As Long

  For i = 2 To lastRow
    Dim toAddr As String: toAddr = Trim(ws.Cells(i, "C").Value)
    If toAddr <> "" Then
      Dim fn As String: fn = ws.Cells(i, "A").Value
      Dim ln As String: ln = ws.Cells(i, "B").Value
      Dim subj As String: subj = ws.Cells(i, "D").Value
      Dim bodyHtml As String: bodyHtml = rawHtml

      bodyHtml = Replace(bodyHtml, "[First Name]", fn, 1, -1, vbTextCompare)
      bodyHtml = Replace(bodyHtml, "[Last Name]", ln, 1, -1, vbTextCompare)

      If isMac Then
        ' Build AppleScript safely (use Chr(34) for quotes)
        Dim script As String, q As String
        q = Chr(34)
        subj = Replace(subj, q, "\" & q)
        bodyHtml = Replace(bodyHtml, q, "\" & q)

        script = "tell application " & q & "Microsoft Outlook" & q & vbLf
        script = script & "  set msg to make new outgoing message with properties {subject:" & q & subj & q & ", content:" & q & bodyHtml & q & "}" & vbLf
        script = script & "  tell msg to make new recipient at end of to recipients with properties {email address:{address:" & q & toAddr & q & "}}" & vbLf
        script = script & "  send msg" & vbLf
        script = script & "end tell"

        On Error Resume Next
        MacScript script
        On Error GoTo 0
      Else
        ' Windows: show each mail for verification; change .Display to .Send when ready
        Dim olMail As Object
        Set olMail = olApp.CreateItem(0)
        With olMail
          .To = toAddr
          .Subject = subj
          .htmlBody = bodyHtml
          .Display   ' <- switch to .Send after you confirm it works
        End With
      End If

      sentCount = sentCount + 1
    End If
  Next i

  MsgBox "? Composed " & sentCount & " emails.", vbInformation, "Done"
End Sub


