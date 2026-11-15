Option Explicit
Option Compare Text

' ================================================================
'  Cross-Sheet Data Aggregator
'  Merge data from 2+ source sheets into a standardized output sheet,
'  with intelligent deduplication, entity-ID-based matching,
'  link sanitization, and batch-ID extraction.
'
'  TWO MODULES:
'    Module 1 — LinkCollector: merge rows from Source A + Source B
'                 by entity link, with 3-tier dedup.
'    Module 2 — IdentifierCollector: match rows from Source A + Source C
'                 by entity ID (extracted from links), merge metadata.
'
'  Author: ivy92888-glitch  |  License: MIT
'  Disclaimer: Unaffiliated with any spreadsheet vendor.
'              Generic clean-room implementation. Any resemblance to
'              specific business systems is coincidental.
' ================================================================

' --- SHEET NAME CONSTANTS ---
Private Const SHEET_SOURCE_A    As String = "SourceA"       ' 主数据表（含链接的那张）
Private Const SHEET_SOURCE_B    As String = "SourceB"       ' 备用数据源
Private Const SHEET_SOURCE_C    As String = "SourceC"       ' 参考库（IdentifierCollector 用）
Private Const SHEET_MERGED_LINKS As String = "MergedLinks"   ' 模块 1 输出
Private Const SHEET_MERGED_CODES As String = "MergedCodes"   ' 模块 2 输出

' --- COLUMN LETTERS (adjust to match your workbook layout) ---
Private Const COL_LINK          As String = "I"             ' 含实体链接的列（两个模块都用）
Private Const COL_ID_FIELD      As String = "B"             ' Source C 里存放候选值/编码的列
Private Const COL_ID_LINK       As String = "G"             ' Source C 里存放实体链接的列（用于提取 ID 做匹配）

' --- FORMATTING ---
Private Const CLR_HEADER_BG     As Long = RGB(31, 78, 121)  ' dark blue
Private Const CLR_HEADER_FG     As Long = RGB(255, 255, 255) ' white

' --- BATCH ID PATTERN (Module 1) ---
' Pattern: 4 uppercase letters + 7 digits = 11 chars total
Private Const BATCH_ID_REGEX    As String = "[A-Z]{4}\d{7}"

' --- ENTITY ID PATTERN (shared by both modules) ---
Private Const ENTITY_ID_REGEX   As String = "[A-Z][0-9A-Z]{9}"
Private Const TARGET_DOMAIN     As String = "amazon.example-domain"

' ================================================================
'  MAIN ENTRY
' ================================================================
Sub RunAllAggregations()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim msg As String
    msg = "=== Cross-Sheet Data Aggregator ===" & vbCrLf

    ' Module 1: Link Collection (Source A + Source B → MergedLinks)
    On Error Resume Next
    Dim hasA As Boolean, hasB As Boolean
    hasA = Not wb.Sheets(SHEET_SOURCE_A) Is Nothing
    hasB = Not wb.Sheets(SHEET_SOURCE_B) Is Nothing
    On Error GoTo 0

    If hasA And hasB Then
        Dim linkCount As Long
        linkCount = RunModule1_LinkCollector(wb)
        msg = msg & vbCrLf & _
              "Module 1 — Link Collector: collected " & linkCount & " rows into '" & SHEET_MERGED_LINKS & "'"
    Else
        msg = msg & vbCrLf & _
              "Module 1 skipped (need both '" & SHEET_SOURCE_A & "' and '" & SHEET_SOURCE_B & "' sheets)"
    End If

    ' Module 2: Identifier Collection (Source A + Source C → MergedCodes)
    On Error Resume Next
    Dim hasC As Boolean
    hasC = Not wb.Sheets(SHEET_SOURCE_C) Is Nothing
    On Error GoTo 0

    If hasA And hasC Then
        Dim codeCount As Long
        codeCount = RunModule2_IdentifierCollector(wb)
        msg = msg & vbCrLf & _
              "Module 2 — Identifier Collector: collected " & codeCount & " rows into '" & SHEET_MERGED_CODES & "'"
    Else
        msg = msg & vbCrLf & _
              "Module 2 skipped (need both '" & SHEET_SOURCE_A & "' and '" & SHEET_SOURCE_C & "' sheets)"
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox msg, vbInformation, "Aggregation Complete"
End Sub

' --- Individual module runners ---
Sub RunModule1()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim n As Long: n = RunModule1_LinkCollector(ThisWorkbook)
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Module 1 done: " & n & " rows → '" & SHEET_MERGED_LINKS & "'", vbInformation
End Sub

Sub RunModule2()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim n As Long: n = RunModule2_IdentifierCollector(ThisWorkbook)
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Module 2 done: " & n & " rows → '" & SHEET_MERGED_CODES & "'", vbInformation
End Sub

' ================================================================
'  MODULE 1 — Link Collector
'  Source A (primary) + Source B (secondary) → MergedLinks
'  Logic: collect rows from both sources, sanitize links,
'         extract batch IDs, build product-name lookup dict,
'         3-tier dedup (identical / same product diff name / diff product).
' ================================================================
Function RunModule1_LinkCollector(ByVal wb As Workbook) As Long
    Dim srcA As Worksheet, srcB As Worksheet, dest As Worksheet
    Set srcA = wb.Sheets(SHEET_SOURCE_A)
    Set srcB = wb.Sheets(SHEET_SOURCE_B)

    ' --- Step 1: Prepare / replace destination sheet ---
    On Error Resume Next
    Set dest = wb.Sheets(SHEET_MERGED_LINKS)
    If Not dest Is Nothing Then dest.Delete
    On Error GoTo 0
    Set dest = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    dest.Name = SHEET_MERGED_LINKS

    ' --- Step 2: Set up destination sheet header + formatting ---
    Call SetupDestinationSheet(dest)

    ' --- Step 3: Build product-name lookup from Source A ---
    ' Dict: ChineseName → Array(EnglishName, Code, PercentParam, Material)
    Dim nameDict As Object
    Set nameDict = CreateObject("Scripting.Dictionary")
    Call BuildNameDictionary(srcA, nameDict)

    ' --- Step 4: First pass — collect from Source A ---
    Dim linkFirstRow As Object   ' Dict: LinkURL → first appearance row #
    Set linkFirstRow = CreateObject("Scripting.Dictionary")

    Dim destRow As Long: destRow = 2
    Dim firstRowSrcA As Long: firstRowSrcA = 4   ' skip header rows
    Dim firstRowSrcB As Long: firstRowSrcB = 4

    Dim lastRowA As Long, lastRowB As Long
    lastRowA = srcA.Cells(srcA.Rows.Count, COL_LINK).End(xlUp).Row
    lastRowB = srcB.Cells(srcB.Rows.Count, COL_LINK).End(xlUp).Row

    Dim i As Long
    For i = firstRowSrcA To lastRowA
        Dim link As String
        link = Trim(srcA.Cells(i, COL_LINK).Value)
        If Len(link) = 0 Then GoTo M1_NextA

        Dim batchId As String, originName As String
        batchId = ExtractBatchId(link)         ' try from link; if empty try from filename
        originName = "SourceA"

        ' Sanitize the link
        link = SanitizeLink(link, TARGET_DOMAIN)

        ' If this link already appeared, check dedup rules
        If linkFirstRow.Exists(link) Then
            Dim firstRowNum As Long
            firstRowNum = linkFirstRow(link)
            Call ApplyDedupRules(dest, link, destRow, firstRowNum, srcA, srcB, i, "A", nameDict)
            GoTo M1_NextA
        End If

        ' Write row to destination
        Call WriteDestRow(dest, destRow, batchId, srcA, i, originName, link, nameDict)
        linkFirstRow.Add link, destRow
        destRow = destRow + 1

M1_NextA:
    Next i

    ' --- Step 5: Second pass — collect from Source B (append, mark origin) ---
    For i = firstRowSrcB To lastRowB
        link = Trim(srcB.Cells(i, COL_LINK).Value)
        If Len(link) = 0 Then GoTo M1_NextB

        batchId = ExtractBatchId(link)
        originName = "SourceB"

        link = SanitizeLink(link, TARGET_DOMAIN)

        If linkFirstRow.Exists(link) Then
            Dim firstRowNumB As Long
            firstRowNumB = linkFirstRow(link)
            Call ApplyDedupRules(dest, link, destRow, firstRowNumB, srcA, srcB, i, "B", nameDict)
            GoTo M1_NextB
        End If

        Call WriteDestRow(dest, destRow, batchId, srcB, i, originName, link, nameDict)
        linkFirstRow.Add link, destRow
        destRow = destRow + 1

M1_NextB:
    Next i

    ' --- Step 6: Finalize formatting ---
    dest.Range("E:E").NumberFormat = "0.00%"   ' percent column
    On Error Resume Next
    Dim r As Long
    For r = 2 To destRow - 1
        If dest.Cells(r, 9).Value <> "" Then    ' column I = link column
            dest.Hyperlinks.Add Anchor:=dest.Cells(r, 9), Address:=dest.Cells(r, 9).Value
        End If
    Next r
    On Error GoTo 0

    dest.Activate

    RunModule1_LinkCollector = destRow - 2   ' return row count (excluding header)
End Function

' --- Module 1 helpers ---

Sub SetupDestinationSheet(ByVal dest As Worksheet)
    Dim headers As Variant
    headers = Array("Batch ID", "Name (CN)", "Name (EN)", "Identifier", _
                    "Percent", "Description", "Price", "Weight", "Link", _
                    "Material", "Source", "Price-per-kg")
    dest.Range("A1:L1").Value = headers

    With dest.Range("A1:L1")
        .Font.Bold = True
        .Font.Color = CLR_HEADER_FG
        .Interior.Color = CLR_HEADER_BG
        .HorizontalAlignment = xlCenter
    End With

    ' Column widths (characters)
    dest.Columns("A").ColumnWidth = 14
    dest.Columns("B").ColumnWidth = 15
    dest.Columns("C").ColumnWidth = 25
    dest.Columns("D").ColumnWidth = 15
    dest.Columns("E").ColumnWidth = 8
    dest.Columns("F").ColumnWidth = 12
    dest.Columns("G").ColumnWidth = 12
    dest.Columns("H").ColumnWidth = 8
    dest.Columns("I").ColumnWidth = 40
    dest.Columns("J").ColumnWidth = 15
    dest.Columns("K").ColumnWidth = 8
    dest.Columns("L").ColumnWidth = 10
End Sub

Sub BuildNameDictionary(ByVal ws As Worksheet, ByRef dict As Object)
    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row   ' col C = Chinese name
    For i = 4 To lastRow
        Dim cnName As String
        cnName = Trim(ws.Cells(i, 3).Value)
        If Len(cnName) > 0 And Not dict.Exists(cnName) Then
            Dim arr(1 To 4) As Variant
            arr(1) = Trim(ws.Cells(i, 4).Value)   ' English name (col D)
            arr(2) = Trim(ws.Cells(i, 5).Value)   ' Identifier    (col E)
            arr(3) = Trim(ws.Cells(i, 20).Value)   ' Percent       (col T, adjust)
            arr(4) = Trim(ws.Cells(i, 25).Value)   ' Material      (col Y, adjust)
            dict.Add cnName, arr
        End If
    Next i
End Sub

Sub WriteDestRow(ByVal dest As Worksheet, ByVal rowNum As Long, _
                 ByVal batchId As String, ByVal srcWs As Worksheet, _
                 ByVal srcRow As Long, ByVal origin As String, _
                 ByVal sanitizedLink As String, ByVal nameDict As Object)
    ' Primary fields from source sheet
    dest.Cells(rowNum, 1).Value = batchId                              ' Batch ID
    dest.Cells(rowNum, 2).Value = Trim(srcWs.Cells(srcRow, 3).Value)  ' Name CN
    dest.Cells(rowNum, 3).Value = Trim(srcWs.Cells(srcRow, 4).Value)  ' Name EN
    dest.Cells(rowNum, 4).Value = Trim(srcWs.Cells(srcRow, 5).Value)  ' Identifier
    dest.Cells(rowNum, 5).Value = Trim(srcWs.Cells(srcRow, 20).Value) ' Percent
    dest.Cells(rowNum, 6).Value = ""                                   ' Description (reserved)
    dest.Cells(rowNum, 7).Value = Trim(srcWs.Cells(srcRow, 18).Value) ' Price
    dest.Cells(rowNum, 8).Value = Trim(srcWs.Cells(srcRow, 10).Value) ' Weight
    dest.Cells(rowNum, 9).Value = sanitizedLink                        ' Sanitized Link
    dest.Cells(rowNum, 10).Value = Trim(srcWs.Cells(srcRow, 25).Value) ' Material
    dest.Cells(rowNum, 11).Value = origin                              ' Source tag
    dest.Cells(rowNum, 12).Value = ""                                   ' Price/kg (reserved)

    ' Fill missing fields using name dictionary (fallback)
    Dim cn As String
    cn = dest.Cells(rowNum, 2).Value
    If nameDict.Exists(cn) Then
        Dim arr As Variant
        arr = nameDict(cn)
        If Len(dest.Cells(rowNum, 3).Value) = 0 Then dest.Cells(rowNum, 3).Value = arr(1)
        If Len(dest.Cells(rowNum, 4).Value) = 0 Then dest.Cells(rowNum, 4).Value = arr(2)
        If Len(dest.Cells(rowNum, 5).Value) = 0 Then dest.Cells(rowNum, 5).Value = arr(3)
        If Len(dest.Cells(rowNum, 10).Value) = 0 Then dest.Cells(rowNum, 10).Value = arr(4)
    End If
End Sub

' --- 3-tier deduplication logic ---
' Rule 1: If all columns B,D-H identical → delete current duplicate row
' Rule 2: If columns D-H identical but B (Chinese name) differs → merge names into first row, delete current
' Rule 3: If columns D-H differ → keep both rows (different products sharing the same link)
Sub ApplyDedupRules(ByVal dest As Worksheet, ByVal link As String, _
                    ByVal destRow As Long, ByVal firstRowNum As Long, _
                    ByVal srcA As Worksheet, ByVal srcB As Worksheet, _
                    ByVal srcRow As Long, ByVal source As String, ByVal nameDict As Object)
    Dim rowNew As Object, rowFirst As Object
    Set rowNew = dest.Rows(destRow)
    Set rowFirst = dest.Rows(firstRowNum)

    Dim bMatch As Boolean, dMatch As Boolean, eMatch As Boolean, _
        fMatch As Boolean, gMatch As Boolean, hMatch As Boolean
    bMatch = (Trim(rowNew.Cells(1, 2).Value) = Trim(rowFirst.Cells(1, 2).Value))
    dMatch = (Trim(rowNew.Cells(1, 4).Value) = Trim(rowFirst.Cells(1, 4).Value))
    eMatch = (Trim(rowNew.Cells(1, 5).Value) = Trim(rowFirst.Cells(1, 5).Value))
    gMatch = (Trim(rowNew.Cells(1, 7).Value) = Trim(rowFirst.Cells(1, 7).Value))
    hMatch = (Trim(rowNew.Cells(1, 8).Value) = Trim(rowFirst.Cells(1, 8).Value))

    ' Rule 1: fully identical → skip (don't write current row)
    If bMatch And dMatch And eMatch And gMatch And hMatch Then
        Exit Sub
    End If

    ' Rule 2: same product (D-E-G-H match) but different Chinese name (B differs) → merge
    If dMatch And eMatch And gMatch And hMatch And Not bMatch Then
        Dim existingName As String, newName As String
        existingName = Trim(rowFirst.Cells(1, 2).Value)
        newName = Trim(rowNew.Cells(1, 2).Value)
        If InStr(1, existingName, newName, vbTextCompare) = 0 Then
            rowFirst.Cells(1, 2).Value = existingName & " / " & newName
        End If
        Exit Sub
    End If

    ' Rule 3: different product, same link → keep both
    Call WriteDestRow(dest, destRow, "", srcA, srcRow, source, link, nameDict)
End Sub

' --- Batch ID extraction ---
' Pattern: 4 uppercase letters + 7 digits (11 chars total)
' Also attempts to extract from a filename if the link has no pattern match
Function ExtractBatchId(ByVal source As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.Pattern = BATCH_ID_REGEX
    If re.Test(source) Then
        ExtractBatchId = re.Execute(source)(0).Value
    Else
        ExtractBatchId = ""
    End If
End Function

Function IsBatchId(ByVal s As String) As Boolean
    If Len(s) <> 11 Then IsBatchId = False: Exit Function
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = BATCH_ID_REGEX
    IsBatchId = re.Test(s)
End Function

' ================================================================
'  MODULE 2 — Identifier Collector
'  Source A (main, has links) + Source C (reference, has codes)
'  → MergedCodes sheet matched by entity ID extracted from links.
' ================================================================
Function RunModule2_IdentifierCollector(ByVal wb As Workbook) As Long
    Dim srcA As Worksheet, srcC As Worksheet, dest As Worksheet
    Set srcA = wb.Sheets(SHEET_SOURCE_A)
    Set srcC = wb.Sheets(SHEET_SOURCE_C)

    ' --- Step 1: Prepare destination sheet ---
    On Error Resume Next
    Set dest = wb.Sheets(SHEET_MERGED_CODES)
    If Not dest Is Nothing Then dest.Delete
    On Error GoTo 0
    Set dest = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    dest.Name = SHEET_MERGED_CODES

    Dim headers As Variant
    headers = Array("Name (CN)", "Name (EN)", "Identifier", "Percent", _
                    "Description", "Weight", "Price", "Link", "Material", "Source")
    dest.Range("A1:J1").Value = headers
    With dest.Range("A1:J1")
        .Font.Bold = True
        .Font.Color = CLR_HEADER_FG
        .Interior.Color = CLR_HEADER_BG
        .HorizontalAlignment = xlCenter
    End With
    dest.Columns("A").ColumnWidth = 15
    dest.Columns("B").ColumnWidth = 25
    dest.Columns("C").ColumnWidth = 15
    dest.Columns("D").ColumnWidth = 8
    dest.Columns("H").ColumnWidth = 40

    ' --- Step 2: Build index from Source A ---
    ' Key = extracted Entity ID, Value = row number in Source A
    Dim itemIndex As Object
    Set itemIndex = CreateObject("Scripting.Dictionary")
    Dim linkFirstRow As Object
    Set linkFirstRow = CreateObject("Scripting.Dictionary")

    Dim lastRowA As Long, i As Long
    lastRowA = srcA.Cells(srcA.Rows.Count, COL_LINK).End(xlUp).Row

    For i = 4 To lastRowA
        Dim link As String
        link = Trim(srcA.Cells(i, COL_LINK).Value)
        If Len(link) = 0 Then GoTo M2_NextA

        Dim entityId As String
        entityId = ExtractEntityId(link)
        If Len(entityId) > 0 Then
            If Not itemIndex.Exists(entityId) Then
                itemIndex.Add entityId, i
            End If
            If Not linkFirstRow.Exists(link) Then
                linkFirstRow.Add link, i
            End If
        End If

M2_NextA:
    Next i

    ' --- Step 3: Walk Source C, match by extracted entity ID ---
    Dim destRow As Long: destRow = 2
    Dim lastRowC As Long
    lastRowC = srcC.Cells(srcC.Rows.Count, COL_ID_LINK).End(xlUp).Row

    For i = 4 To lastRowC
        Dim cLink As String
        cLink = Trim(srcC.Cells(i, COL_ID_LINK).Value)
        If Len(cLink) = 0 Then GoTo M2_NextC

        Dim cId As String
        cId = ExtractEntityId(cLink)
        If Len(cId) = 0 Then GoTo M2_NextC

        If Not itemIndex.Exists(cId) Then GoTo M2_NextC   ' no match in Source A

        ' Found a match — write to destination
        Dim aRow As Long
        aRow = itemIndex(cId)

        Dim sanitized As String
        sanitized = SanitizeLink(cLink, TARGET_DOMAIN)

        dest.Cells(destRow, 1).Value = Trim(srcA.Cells(aRow, 3).Value)   ' CN name
        dest.Cells(destRow, 2).Value = Trim(srcA.Cells(aRow, 4).Value)   ' EN name
        dest.Cells(destRow, 3).Value = Trim(srcC.Cells(i, COL_ID_FIELD).Value) ' Identifier from Source C
        dest.Cells(destRow, 4).Value = Trim(srcC.Cells(i, 3).Value)      ' Percent
        dest.Cells(destRow, 5).Value = ""                                  ' Description reserved
        dest.Cells(destRow, 6).Value = Trim(srcA.Cells(aRow, 10).Value)   ' Weight from Source A
        dest.Cells(destRow, 7).Value = Trim(srcA.Cells(aRow, 18).Value)   ' Price from Source A
        dest.Cells(destRow, 8).Value = sanitized
        dest.Cells(destRow, 9).Value = Trim(srcA.Cells(aRow, 25).Value)   ' Material
        dest.Cells(destRow, 10).Value = "Matched by Entity ID"

        destRow = destRow + 1

M2_NextC:
    Next i

    ' --- Step 4: Dedup destination ---
    Dim r As Long
    For r = destRow - 1 To 3 Step -1
        ' Dedup based on Identifier + Link combo
        If dest.Cells(r, 3).Value = dest.Cells(r - 1, 3).Value And _
           dest.Cells(r, 8).Value = dest.Cells(r - 1, 8).Value Then
            dest.Rows(r).Delete
            destRow = destRow - 1
        End If
    Next r

    dest.Range("D:D").NumberFormat = "0.00%"
    dest.Activate

    RunModule2_IdentifierCollector = destRow - 2
End Function

' ================================================================
'  SHARED UTILITIES
' ================================================================

Function SanitizeLink(ByVal url As String, ByVal targetDomain As String) As String
    If Len(url) = 0 Then SanitizeLink = "": Exit Function

    Dim result As String
    result = url

    ' Remove query string
    Dim qPos As Long
    qPos = InStr(result, "?")
    If qPos > 0 Then result = Left(result, qPos - 1)

    ' Remove hash
    Dim hPos As Long
    hPos = InStr(result, "#")
    If hPos > 0 Then result = Left(result, hPos - 1)

    ' Extract entity identifier + rebuild
    Dim id As String
    id = ExtractEntityId(result)
    If Len(id) > 0 Then
        result = "https://" & targetDomain & "/dp/" & id
    End If

    SanitizeLink = result
End Function

Function ExtractEntityId(ByVal url As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.Pattern = ENTITY_ID_REGEX   ' [A-Z][0-9A-Z]{9}
    If re.Test(url) Then
        ExtractEntityId = re.Execute(url)(0).Value
    Else
        ExtractEntityId = ""
    End If
End Function
