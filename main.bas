Option Explicit

Public Sub BuildWideFromLong()
    Dim wsIn As Worksheet, wsOut As Worksheet
    Dim lastRow As Long, lastCol As Long, r As Long, c As Long
    Dim hdr As Object: Set hdr = CreateObject("Scripting.Dictionary")
    Dim hasCategory As Boolean
    
    '--- Get input sheet
    On Error Resume Next
    Set wsIn = ThisWorkbook.Worksheets("Data")
    On Error GoTo 0
    If wsIn Is Nothing Then
        MsgBox "Missing sheet 'Data'. Add a sheet named Data with columns: System_ID, URL, (optional) Category.", vbExclamation
        Exit Sub
    End If
    
    '--- Detect header row/columns
    With wsIn
        If .UsedRange.Rows.Count < 2 Then
            MsgBox "No data found on sheet 'Data'.", vbExclamation
            Exit Sub
        End If
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
    End With
    
    Dim nameVal As String
    For c = 1 To lastCol
        nameVal = Trim$(CStr(wsIn.Cells(1, c).Value))
        If Len(nameVal) > 0 Then
            hdr(UCase$(nameVal)) = c
        End If
    Next c
    
    If Not hdr.Exists("SYSTEM_ID") Then
        MsgBox "Missing required column 'System_ID' in row 1.", vbCritical
        Exit Sub
    End If
    If Not hdr.Exists("URL") Then
        MsgBox "Missing required column 'URL' in row 1.", vbCritical
        Exit Sub
    End If
    hasCategory = hdr.Exists("CATEGORY")
    
    '--- Prepare data structures
    Dim slots As Object: Set slots = CreateObject("Scripting.Dictionary") ' per SID -> next slot (count)
    Dim mapCats As Object: Set mapCats = CreateObject("Scripting.Dictionary") ' SID -> dict(slot->Category)
    Dim mapUrls As Object: Set mapUrls = CreateObject("Scripting.Dictionary") ' SID -> dict(slot->URL)
    Dim maxSlot As Long: maxSlot = 0
    
    '--- Iterate rows in input order
    Dim SID As String, URLv As String, CATv As String
    Dim perSIDCats As Object, perSIDUrls As Object
    For r = 2 To lastRow
        SID = Trim$(CStr(wsIn.Cells(r, hdr("SYSTEM_ID")).Value))
        URLv = Trim$(CStr(wsIn.Cells(r, hdr("URL")).Value))
        If hasCategory Then
            CATv = Trim$(CStr(wsIn.Cells(r, hdr("CATEGORY")).Value))
        Else
            CATv = ""
        End If
        
        ' drop blank / "nan"
        If (SID = "" Or UCase$(SID) = "NAN") Then GoTo ContinueRow
        If (URLv = "" Or UCase$(URLv) = "NAN") Then GoTo ContinueRow
        
        ' allocate slot in input order
        If Not slots.Exists(SID) Then
            slots(SID) = 1
            Set perSIDCats = CreateObject("Scripting.Dictionary")
            Set perSIDUrls = CreateObject("Scripting.Dictionary")
            mapCats(SID) = perSIDCats
            mapUrls(SID) = perSIDUrls
        Else
            Set perSIDCats = mapCats(SID)
            Set perSIDUrls = mapUrls(SID)
            slots(SID) = CLng(slots(SID)) + 1
        End If
        
        Dim s As Long: s = CLng(slots(SID))
        ' store values at this slot
        perSIDCats(s) = CATv
        perSIDUrls(s) = URLv
        If s > maxSlot Then maxSlot = s
        
ContinueRow:
    Next r
    
    If mapUrls.Count = 0 Then
        MsgBox "No usable rows after cleaning (blank System_ID/URL removed).", vbExclamation
        Exit Sub
    End If
    
    '--- Ensure/clear Output sheet
    On Error Resume Next
    Set wsOut = ThisWorkbook.Worksheets("Output")
    On Error GoTo 0
    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Worksheets.Add(After:=wsIn)
        wsOut.Name = "Output"
    Else
        wsOut.Cells.Clear
    End If
    
    '--- Build header
    Dim col As Long: col = 1
    wsOut.Cells(1, col).Value = "SystemID": col = col + 1
    Dim sNum As Long
    For sNum = 1 To maxSlot
        wsOut.Cells(1, col).Value = "Category" & sNum: col = col + 1
        wsOut.Cells(1, col).Value = "ExternalFileField" & sNum: col = col + 1
    Next sNum
    
    '--- Sort SIDs for stable output
    Dim keys() As Variant: keys = mapUrls.Keys
    If UBound(keys) >= 1 Then QuickSortStrings keys, LBound(keys), UBound(keys)
    
    '--- Write rows
    Dim outRow As Long: outRow = 2
    For Each SID In keys
        wsOut.Cells(outRow, 1).Value = "id\" & CStr(SID)
        col = 2
        Set perSIDCats = mapCats(SID)
        Set perSIDUrls = mapUrls(SID)
        
        For sNum = 1 To maxSlot
            ' CategoryN
            If perSIDCats.Exists(sNum) Then
                wsOut.Cells(outRow, col).Value = perSIDCats(sNum)
            Else
                wsOut.Cells(outRow, col).Value = ""
            End If
            col = col + 1
            ' ExternalFileFieldN
            If perSIDUrls.Exists(sNum) Then
                wsOut.Cells(outRow, col).Value = perSIDUrls(sNum)
            Else
                wsOut.Cells(outRow, col).Value = ""
            End If
            col = col + 1
        Next sNum
        outRow = outRow + 1
    Next SID
    
    '--- Tidy up
    With wsOut
        .ListObjects.Add(xlSrcRange, .Range("A1").CurrentRegion, , xlYes).Name = "OutputTable"
        .Columns.AutoFit
    End With
    
    MsgBox "Wide table built on sheet 'Output'.", vbInformation
End Sub

'---------- helpers ----------
Private Sub QuickSortStrings(arr As Variant, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As String, tmp As String
    i = first: j = last
    pivot = CStr(arr((first + last) \ 2))
    Do While i <= j
        Do While CStr(arr(i)) < pivot: i = i + 1: Loop
        Do While CStr(arr(j)) > pivot: j = j - 1: Loop
        If i <= j Then
            tmp = CStr(arr(i)): arr(i) = arr(j): arr(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If first < j Then QuickSortStrings arr, first, j
    If i < last Then QuickSortStrings arr, i, last
End Sub
