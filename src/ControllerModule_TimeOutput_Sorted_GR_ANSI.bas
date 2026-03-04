Attribute VB_Name = "ControllerModule"
Option Explicit

' ===========================
'  ΡΥΘΜΙΣΕΙΣ / ΚΑΝΟΝΕΣ ΧΡΟΝΟΥ
' ===========================
Private Const DEFAULT_DURATION_MIN As Long = 10
Private Const POLICE_DEPOSITION_DURATION_MIN As Long = 20
Private Const DEFAULT_BREAK_MIN As Long = 5

' ========= Helpers =========

Private Function CleanCellText(ByVal s As String) As String
    s = Replace(s, Chr(13), "")
    s = Replace(s, Chr(7), "")
    CleanCellText = Trim$(s)
End Function

Private Function FileExists(ByVal p As String) As Boolean
    On Error GoTo NotThere
    Dim a As Long
    a = GetAttr(p)
    FileExists = True
    Exit Function
NotThere:
    FileExists = False
End Function

Private Function EnsureOutputFolder(ByVal baseFolder As String) As String
    Dim outPath As String
    outPath = baseFolder & "\OUTPUT"
    If Dir(outPath, vbDirectory) = "" Then
        MkDir outPath
    End If
    EnsureOutputFolder = outPath
End Function

Private Function SafeFilePart(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), "-")
    Next i
    SafeFilePart = Trim$(s)
End Function

Private Function BaseNameWithoutExt(ByVal filename As String) As String
    Dim p As Long
    p = InStrRev(filename, ".")
    If p > 0 Then
        BaseNameWithoutExt = Left$(filename, p - 1)
    Else
        BaseNameWithoutExt = filename
    End If
End Function

Private Function FileExt(ByVal filename As String) As String
    Dim p As Long
    p = InStrRev(filename, ".")
    If p > 0 Then
        FileExt = Mid$(filename, p)
    Else
        FileExt = ""
    End If
End Function

Private Function ReadMapFromFirstTable(ByVal ctrlDoc As Document) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    If ctrlDoc.Tables.Count = 0 Then
        Err.Raise vbObjectError + 101, , "Δεν βρέθηκε πίνακας στο Controller."
    End If

    Dim t As Table
    Set t = ctrlDoc.Tables(1)

    Dim r As Long
    For r = 2 To t.Rows.Count
        Dim k As String, v As String
        k = CleanCellText(t.Cell(r, 1).Range.Text)
        v = CleanCellText(t.Cell(r, 2).Range.Text)
        If Len(k) > 0 Then dict(k) = v
    Next r

    Set ReadMapFromFirstTable = dict
End Function

Private Sub ReplaceAllInRange(ByVal rng As Range, ByVal findText As String, ByVal replText As String)
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Private Sub ReplaceEverywhere(ByVal doc As Document, ByVal findText As String, ByVal replText As String)
    Dim s As Range
    For Each s In doc.StoryRanges
        ReplaceAllInRange s, findText, replText
        Do While Not (s.NextStoryRange Is Nothing)
            Set s = s.NextStoryRange
            ReplaceAllInRange s, findText, replText
        Loop
    Next s

    Dim shp As Shape
    For Each shp In doc.Shapes
        If shp.TextFrame.HasText Then
            ReplaceAllInRange shp.TextFrame.TextRange, findText, replText
        End If
    Next shp

    Dim i As Long, j As Long
    For i = 1 To doc.Sections.Count
        For j = 1 To 3
            On Error Resume Next
            For Each shp In doc.Sections(i).Headers(j).Shapes
                If shp.TextFrame.HasText Then
                    ReplaceAllInRange shp.TextFrame.TextRange, findText, replText
                End If
            Next shp
            For Each shp In doc.Sections(i).Footers(j).Shapes
                If shp.TextFrame.HasText Then
                    ReplaceAllInRange shp.TextFrame.TextRange, findText, replText
                End If
            Next shp
            On Error GoTo 0
        Next j
    Next i
End Sub

Private Function BuildUniqueOutputName(ByVal outFolder As String, ByVal caseId As String, ByVal baseName As String, ByVal ext As String) As String
    Dim candidate As String
    candidate = outFolder & "\" & caseId & "_" & baseName & ext

    Dim n As Long
    n = 1
    Do While FileExists(candidate)
        candidate = outFolder & "\" & caseId & "_" & baseName & "_" & n & ext
        n = n + 1
    Loop

    BuildUniqueOutputName = candidate
End Function

' ========= Time helpers =========

Private Function ParseTimeHHNN(ByVal s As String) As Date
    s = Trim$(s)
    If Len(s) = 0 Then
        ParseTimeHHNN = Time
    Else
        ParseTimeHHNN = TimeValue(s)
    End If
End Function

Private Function DurationMinutesFor(ByVal filename As String) As Long
    Dim u As String
    u = UCase$(filename)

    If (InStr(u, "ΚΑΤΑΘΕΣΗ") > 0) And (InStr(u, "ΑΣΤΥΝΟΜ") > 0) Then
        DurationMinutesFor = POLICE_DEPOSITION_DURATION_MIN
    Else
        DurationMinutesFor = DEFAULT_DURATION_MIN
    End If
End Function

Private Function GetBreakMinutes(ByVal map As Object) As Long
    Dim bm As Long
    bm = DEFAULT_BREAK_MIN
    If map.Exists("BreakMinutes") Then
        If IsNumeric(map("BreakMinutes")) Then bm = CLng(map("BreakMinutes"))
    End If
    GetBreakMinutes = bm
End Function

' ========= Sorting helpers (numeric order) =========

Private Function ExtractFirstNumber(ByVal filename As String) As Long
    Dim i As Long, ch As String, num As String
    For i = 1 To Len(filename)
        ch = Mid$(filename, i, 1)
        If ch Like "#" Then
            num = num & ch
        ElseIf Len(num) > 0 Then
            Exit For
        End If
    Next i

    If Len(num) = 0 Then
        ExtractFirstNumber = 999999
    Else
        ExtractFirstNumber = CLng(num)
    End If
End Function

Private Sub SortFilesByNumberThenName(ByRef arr() As String)
    Dim i As Long, j As Long
    Dim a As String, b As String
    Dim na As Long, nb As Long

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            a = arr(i): b = arr(j)
            na = ExtractFirstNumber(a): nb = ExtractFirstNumber(b)

            If (na > nb) Or ((na = nb) And (StrComp(a, b, vbTextCompare) > 0)) Then
                arr(i) = b
                arr(j) = a
            End If
        Next j
    Next i
End Sub

Private Function CollectAndSortTemplates(ByVal folderPath As String) As Variant
    Dim files() As String
    Dim cnt As Long
    cnt = 0

    Dim f As String
    f = Dir(folderPath & "\TEMPLATE_*.docx")

    Do While f <> ""
        If Left$(f, 2) <> "~$" Then
            cnt = cnt + 1
            ReDim Preserve files(1 To cnt)
            files(cnt) = f
        End If
        f = Dir()
    Loop

    If cnt = 0 Then
        CollectAndSortTemplates = Null
        Exit Function
    End If

    SortFilesByNumberThenName files
    CollectAndSortTemplates = files
End Function

' ========= Main Macro =========

Public Sub Generate_Reports_To_OUTPUT_From_Table()
    Dim folderPath As String
    folderPath = ThisDocument.Path
    If Len(folderPath) = 0 Then
        MsgBox "Αποθήκευσε πρώτα το 00_Controller.docm μέσα στον φάκελο με τις εκθέσεις.", vbExclamation
        Exit Sub
    End If

    Dim map As Object
    On Error GoTo EH
    Set map = ReadMapFromFirstTable(ThisDocument)
    On Error GoTo 0

    Dim outFolder As String
    outFolder = EnsureOutputFolder(folderPath)

    Dim caseId As String
    If map.Exists("CaseID") Then caseId = SafeFilePart(CStr(map("CaseID"))) Else caseId = ""
    If Len(caseId) = 0 Then caseId = Format(Now, "yyyymmdd_HHMMss")

    Dim breakMin As Long
    breakMin = GetBreakMinutes(map)

    Dim curStart As Date
    If map.Exists("OraStart") Then
        curStart = ParseTimeHHNN(CStr(map("OraStart")))
    Else
        curStart = Time
    End If

    Dim templates As Variant
    templates = CollectAndSortTemplates(folderPath)
    If IsNull(templates) Then
        MsgBox "Δεν βρέθηκαν templates (TEMPLATE_*.docx).", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone

    Dim producedCount As Long
    producedCount = 0

    Dim i As Long
    For i = LBound(templates) To UBound(templates)
        Dim f As String
        f = templates(i)

        Dim dur As Long
        Dim startT As Date, endT As Date

        startT = curStart
        dur = DurationMinutesFor(f)
        endT = DateAdd("n", dur, startT)

        map("OraEnarxis") = Format(startT, "hh:nn")
        map("OraPeratosis") = Format(endT, "hh:nn")

        Dim srcFull As String, dstFull As String
        Dim baseName As String, ext As String
        baseName = BaseNameWithoutExt(f)
        ext = FileExt(f)

        srcFull = folderPath & "\" & f
        dstFull = BuildUniqueOutputName(outFolder, caseId, baseName, ext)

        FileCopy srcFull, dstFull

        Dim doc As Document
        Set doc = Documents.Open(FileName:=dstFull, ReadOnly:=False, AddToRecentFiles:=False)

        Dim key As Variant
        For Each key In map.Keys
            ReplaceEverywhere doc, "{{" & CStr(key) & "}}", CStr(map(key))
        Next key

        doc.Save
        doc.Close SaveChanges:=False
        producedCount = producedCount + 1

        curStart = DateAdd("n", dur + breakMin, startT)
    Next i

    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True

    MsgBox "Έτοιμο. Παράχθηκαν " & producedCount & " εκθέσεις στον φάκελο OUTPUT.", vbInformation
    Exit Sub

EH:
    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True
    MsgBox "Σφάλμα: " & Err.Description, vbExclamation
End Sub
