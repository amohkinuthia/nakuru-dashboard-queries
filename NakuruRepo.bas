Attribute VB_Name = "NakuruRepo"
' =====================================================================
' NAKURU DASHBOARD – GitHub-hosted queries
'   1. ExportQueries        – writes every Power Query in this workbook to
'                              Documents\nakuru-dashboard-queries\queries\*.pq
'                              (upload that folder to your GitHub repo)
'   2. SwitchToGitHub       – replaces each query with a small loader that
'                              downloads its code from GitHub on every refresh.
'                              Original code is backed up to Documents\nakuru-dashboard-queries\backup-*.
'   3. RestoreLocalQueries  – puts the backed-up original code back.
' Run with Alt+F8. Requires Excel 2016 or newer (Windows).
' =====================================================================
Option Explicit

Private Const CORE As String = "fnParam,pDay,fnGetList,Facilities,Lib"   ' helpers, in dependency order
Private Const SHARED_DATA As String = "Billing,Receipts,Stock_Balance"    ' data queries other queries read
Private Const LOCAL_ONLY As String = "fnParam,fnRemote"                   ' never moved to GitHub

Private Function SafeName(n As String) As String
    SafeName = Replace(Trim(n), " ", "_")
End Function

Private Function InList(n As String, csv As String) As Boolean
    Dim p As Variant
    For Each p In Split(csv, ",")
        If StrComp(Trim(p), n, vbTextCompare) = 0 Then InList = True: Exit Function
    Next p
End Function

Private Function QueryExists(n As String) As Boolean
    Dim q As WorkbookQuery
    For Each q In ThisWorkbook.Queries
        If StrComp(q.Name, n, vbTextCompare) = 0 Then QueryExists = True: Exit Function
    Next q
End Function

Private Sub WriteUtf8(path As String, txt As String)
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText txt
    st.SaveToFile path, 2
    st.Close
End Sub

Private Sub MakeFolder(p As String)
    Dim parts() As String, cur As String, i As Long
    parts = Split(p, "\")
    cur = parts(0)
    For i = 1 To UBound(parts)
        cur = cur & "\" & parts(i)
        If Dir(cur, vbDirectory) = "" Then MkDir cur
    Next i
End Sub

' ---------------------------------------------------------------------
Public Sub ExportQueries()
    Dim folder As String, q As WorkbookQuery, n As Long, list As String
    folder = Environ("USERPROFILE") & "\Documents\nakuru-dashboard-queries\queries"
    MakeFolder folder
    For Each q In ThisWorkbook.Queries
        If Not InList(q.Name, LOCAL_ONLY) And InStr(q.Formula, "fnRemote(") = 0 Then
            WriteUtf8 folder & "\" & SafeName(q.Name) & ".pq", q.Formula
            list = list & SafeName(q.Name) & vbLf
            n = n + 1
        End If
    Next q
    WriteUtf8 folder & "\..\VERSION.txt", Format(Now, "yyyy-mm-dd hh:nn") & vbLf & list
    MsgBox n & " queries exported to:" & vbLf & folder & vbLf & vbLf & _
           "Upload the whole 'nakuru-dashboard-queries' folder to your GitHub repo " & _
           "(Add file -> Upload files), then run SwitchToGitHub." & vbLf & vbLf & _
           "Files:" & vbLf & list, vbInformation, "Export done"
    Shell "explorer.exe """ & folder & """", vbNormalFocus
End Sub

' ---------------------------------------------------------------------
Private Sub EnsureCodeUrl()
    Dim url As String, ws As Worksheet, nm As Name, have As Boolean
    On Error Resume Next
    Set nm = ThisWorkbook.Names("P_CodeUrl")
    have = Not nm Is Nothing
    On Error GoTo 0
    If have Then Exit Sub
    url = InputBox("Raw URL of your repo's main branch, e.g." & vbLf & _
                   "https://raw.githubusercontent.com/amohkinuthia/nakuru-dashboard-queries/main", _
                   "GitHub code location", "https://raw.githubusercontent.com/amohkinuthia/nakuru-dashboard-queries/main")
    If url = "" Then Err.Raise 513, , "No URL given"
    Set ws = ThisWorkbook.Worksheets("Settings")
    ws.Range("B30").Value = "Code location (GitHub raw URL)"
    ws.Range("C30").Value = url
    ThisWorkbook.Names.Add Name:="P_CodeUrl", RefersTo:="=Settings!$C$30"
End Sub

Private Function BackupFolder() As String
    BackupFolder = Environ("USERPROFILE") & "\Documents\nakuru-dashboard-queries\backup-" & Replace(ThisWorkbook.Name, ".", "_")
End Function

Private Function ReadUtf8(path As String) As String
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText
    st.Close
    If Len(ReadUtf8) > 0 Then
        If AscW(Left(ReadUtf8, 1)) = &HFEFF Then ReadUtf8 = Mid(ReadUtf8, 2)
    End If
End Function

Private Sub Backup()
    Dim q As WorkbookQuery, folder As String
    folder = BackupFolder()
    MakeFolder folder
    For Each q In ThisWorkbook.Queries
        If InStr(q.Formula, "fnRemote(") = 0 Then WriteUtf8 folder & "\" & q.Name & ".pq", q.Formula
    Next q
End Sub

Private Function EnvFor(n As String) As String
    ' record of other queries this one may use: earlier helpers only (no cycles);
    ' data queries get all helpers + the shared data queries (except themselves)
    Dim items As String, p As Variant, i As Long, pos As Long, coreArr As Variant
    coreArr = Split(CORE, ",")
    pos = -1
    For i = 0 To UBound(coreArr)
        If StrComp(coreArr(i), n, vbTextCompare) = 0 Then pos = i
    Next i
    For i = 0 To UBound(coreArr)
        If (pos = -1 Or i < pos) And QueryExists(CStr(coreArr(i))) Then items = items & IIf(items = "", "", ", ") & coreArr(i) & " = " & coreArr(i)
    Next i
    If pos = -1 Then
        For Each p In Split(SHARED_DATA, ",")
            If StrComp(p, n, vbTextCompare) <> 0 And QueryExists(CStr(p)) Then items = items & IIf(items = "", "", ", ") & p & " = " & p
        Next p
    End If
    EnvFor = "[" & items & "]"
End Function

Public Sub SwitchToGitHub()
    Dim q As WorkbookQuery, n As Long, f As String
    If MsgBox("This replaces every query's code with a loader that reads it from GitHub." & vbLf & _
              "A backup is kept (RestoreLocalQueries undoes it)." & vbLf & vbLf & _
              "Have you uploaded the exported .pq files to the repo?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    EnsureCodeUrl
    Backup
    ' fnRemote: downloads one query's code
    f = "(name as text) as text =>" & vbLf & _
        "let" & vbLf & _
        "    base = Text.TrimEnd(Text.From(fnParam(""P_CodeUrl"")), ""/"")," & vbLf & _
        "    bin  = Web.Contents(base & ""/queries/"" & name & "".pq"", [Headers = [#""Cache-Control"" = ""no-cache""]])," & vbLf & _
        "    t    = Text.Replace(Text.FromBinary(bin, TextEncoding.Utf8), Character.FromNumber(65279), """")," & vbLf & _
        "    t2   = Text.Trim(t)" & vbLf & _
        "in" & vbLf & _
        "    if Text.StartsWith(t2, ""="") then Text.Trim(Text.Middle(t2, 1)) else t2"
    If QueryExists("fnRemote") Then ThisWorkbook.Queries("fnRemote").Formula = f Else ThisWorkbook.Queries.Add "fnRemote", f
    For Each q In ThisWorkbook.Queries
        If Not InList(q.Name, LOCAL_ONLY) And InStr(q.Formula, "fnRemote(") = 0 Then
            q.Formula = "let" & vbLf & _
                        "    // code lives on GitHub: queries/" & SafeName(q.Name) & ".pq  (edit it there – everyone gets it on next refresh)" & vbLf & _
                        "    env = Record.Combine({#shared, " & EnvFor(q.Name) & "})," & vbLf & _
                        "    out = Expression.Evaluate(fnRemote(""" & SafeName(q.Name) & """), env)" & vbLf & _
                        "in" & vbLf & _
                        "    out"
            n = n + 1
        End If
    Next q
    MsgBox n & " queries now load their code from GitHub." & vbLf & "Click Data -> Refresh All to test." & vbLf & _
           "If something fails, run RestoreLocalQueries.", vbInformation, "Switched"
End Sub

Public Sub RestoreLocalQueries()
    Dim folder As String, q As WorkbookQuery, n As Long, f As String
    folder = BackupFolder()
    If Dir(folder, vbDirectory) = "" Then MsgBox "No backup found in " & folder, vbExclamation: Exit Sub
    For Each q In ThisWorkbook.Queries
        f = folder & "\" & q.Name & ".pq"
        If Dir(f) <> "" And InStr(q.Formula, "fnRemote(") > 0 Then q.Formula = ReadUtf8(f): n = n + 1
    Next q
    MsgBox n & " queries restored to their local code (from " & folder & ").", vbInformation
End Sub
