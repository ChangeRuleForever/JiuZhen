Attribute VB_Name = "JiuzhenUniversalMacros"
Option Explicit

Private Const VALIDATION_LIST As Long = 3
Private Const VALIDATION_ALERT_STOP As Long = 1
Private Const VALIDATION_BETWEEN As Long = 1
Private Const SHEET_VISIBLE As Long = -1
Private Const SHEET_HIDDEN As Long = 0

Public Sub ApplyUniversalWorkbookTheme(ByVal themeKey As String)
    On Error GoTo CleanFail

    Dim settingsWs As Worksheet
    Dim bookWs As Worksheet
    Dim personalWs As Worksheet
    Dim traitWs As Worksheet
    Dim skillWs As Worksheet
    Dim equipmentWs As Worksheet
    Dim herbWs As Worksheet
    Dim followerWs As Worksheet

    Set settingsWs = TryGetWorksheet("基本信息")
    Set bookWs = TryGetWorksheet("书籍、特性与技能")
    Set personalWs = TryGetWorksheet("个人信息")
    Set traitWs = TryGetWorksheet("通用特性体系")
    Set skillWs = TryGetWorksheet("技能列表")
    Set equipmentWs = TryGetWorksheet("装备列表")
    Set herbWs = TryGetWorksheet("【药家】草药学列表")
    Set followerWs = TryGetWorksheet("随从列表")

    ApplyFamilySpecificSheetVisibility themeKey, herbWs
    ApplyExtendedSettingsTheme settingsWs, themeKey
    ApplyBookAndTraitsExtensionTheme bookWs, themeKey
    ApplyTraitCatalogTheme traitWs, themeKey
    ApplySkillCatalogTheme skillWs, themeKey
    ApplyEquipmentCatalogTheme equipmentWs, themeKey
    ApplyHerbCatalogTheme herbWs, themeKey
    ApplyFollowerSheetTheme followerWs, themeKey
    RefreshUniversalReferenceLists themeKey, traitWs, skillWs, bookWs, followerWs
    RefreshEquipmentReferenceLists equipmentWs, personalWs, followerWs
    Exit Sub

CleanFail:
    MsgBox "Apply universal workbook theme failed: " & Err.Description, vbExclamation
End Sub

Private Sub ApplyEquipmentCatalogTheme(ByVal ws As Worksheet, ByVal themeKey As String)
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim categoryName As String

    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 4 Then Exit Sub

    ws.Cells.Font.Name = "KaiTi"
    ws.Cells.Font.Size = 11
    ws.Range("A1:K" & lastRow).Interior.Color = ThemeTone(themeKey, "base")

    PaintUniversalBlock ws.Range("A1:K1"), ThemeTone(themeKey, "titleDark"), ThemeTone(themeKey, "textLight"), 13, True, True
    PaintUniversalBlock ws.Range("A2:K2"), ThemeTone(themeKey, "note"), ThemeTone(themeKey, "textDark"), 10.5, False, False
    PaintUniversalBlock ws.Range("A4:K4"), ThemeTone(themeKey, "header"), ThemeTone(themeKey, "textDark"), 11, True, True
    PaintUniversalBlock ws.Range("A5:K" & lastRow), ThemeTone(themeKey, "table"), ThemeTone(themeKey, "textDark"), 10.5, False, False

    For rowIndex = 5 To lastRow
        categoryName = Trim$(CStr(ws.Cells(rowIndex, "A").Value))
        If Len(categoryName) > 0 Then
            ws.Cells(rowIndex, "A").Interior.Color = ThemeTone(themeKey, "header")
            ws.Cells(rowIndex, "B").Interior.Color = ThemeTone(themeKey, "note")
        End If
    Next rowIndex

    ws.Columns("L:Z").Hidden = True
    ws.Rows("1:" & lastRow).RowHeight = 24
    ApplyUniversalGrid ws.Range("A1:K" & lastRow)
    ws.Tab.Color = ThemeTone(themeKey, "tab1")
End Sub

Private Sub ApplyFamilySpecificSheetVisibility(ByVal themeKey As String, ByVal herbWs As Worksheet)
    If herbWs Is Nothing Then Exit Sub

    If StrComp(themeKey, "药家", vbTextCompare) = 0 Then
        herbWs.Visible = SHEET_VISIBLE
    Else
        If ActiveSheet.Name = herbWs.Name Then
            Sheet1.Activate
        End If
        herbWs.Visible = SHEET_HIDDEN
    End If
End Sub

Private Sub ApplyExtendedSettingsTheme(ByVal ws As Worksheet, ByVal themeKey As String)
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim familyName As String

    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 21 Then lastRow = 21

    ws.Range("A19:F" & lastRow).Interior.Color = ThemeTone(themeKey, "base")
    PaintUniversalBlock ws.Range("A19:F19"), ThemeTone(themeKey, "titleDark"), ThemeTone(themeKey, "textLight"), 12.5, True, True
    PaintUniversalBlock ws.Range("A20:F20"), ThemeTone(themeKey, "header"), ThemeTone(themeKey, "textDark"), 11, True, True
    PaintUniversalBlock ws.Range("A21:F" & lastRow), ThemeTone(themeKey, "table"), ThemeTone(themeKey, "textDark"), 10.5, False, False

    For rowIndex = 21 To lastRow
        familyName = Trim$(CStr(ws.Cells(rowIndex, "A").Value))
        If Len(familyName) > 0 Then
            ws.Cells(rowIndex, "A").Interior.Color = UniversalPrimaryColor(familyName)
            ws.Cells(rowIndex, "A").Font.Color = ThemeTone(themeKey, "textLight")
            ws.Cells(rowIndex, "A").Font.Bold = True
            ws.Cells(rowIndex, "A").HorizontalAlignment = xlHAlignCenter
            ws.Cells(rowIndex, "A").VerticalAlignment = xlVAlignCenter
        End If
    Next rowIndex

    ws.Rows("19:" & lastRow).RowHeight = 30
    ApplyUniversalGrid ws.Range("A19:F" & lastRow)
End Sub

Private Sub ApplyBookAndTraitsExtensionTheme(ByVal ws As Worksheet, ByVal themeKey As String)
    If ws Is Nothing Then Exit Sub

    ws.Columns("W:AD").Hidden = True
    ws.Range("W1:AD200").Interior.Color = ThemeTone(themeKey, "base")
    ws.Range("W1:AD200").Font.Name = "KaiTi"
    ws.Range("W1:AD200").Font.Size = 10
End Sub

Private Sub ApplyTraitCatalogTheme(ByVal ws As Worksheet, ByVal themeKey As String)
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim familyName As String

    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ws.Cells.Font.Name = "KaiTi"
    ws.Cells.Font.Size = 11
    ws.Range("A1:D" & lastRow).Interior.Color = ThemeTone(themeKey, "base")

    PaintUniversalBlock ws.Range("A1:D1"), ThemeTone(themeKey, "titleDark"), ThemeTone(themeKey, "textLight"), 13, True, True
    PaintUniversalBlock ws.Range("A2:D2"), ThemeTone(themeKey, "header"), ThemeTone(themeKey, "textDark"), 11, True, True
    PaintUniversalBlock ws.Range("A3:D" & lastRow), ThemeTone(themeKey, "table"), ThemeTone(themeKey, "textDark"), 10.5, False, False

    For rowIndex = 3 To lastRow
        familyName = Trim$(CStr(ws.Cells(rowIndex, "A").Value))
        If Len(familyName) > 0 Then
            ws.Cells(rowIndex, "A").Interior.Color = UniversalPrimaryColor(familyName)
            ws.Cells(rowIndex, "A").Font.Color = ThemeTone(themeKey, "textLight")
        End If
    Next rowIndex

    ws.Rows("1:" & lastRow).RowHeight = 24
    ApplyUniversalGrid ws.Range("A1:D" & lastRow)
    ws.Tab.Color = ThemeTone(themeKey, "tab1")
End Sub

Private Sub ApplySkillCatalogTheme(ByVal ws As Worksheet, ByVal themeKey As String)
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim familyName As String

    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ws.Cells.Font.Name = "KaiTi"
    ws.Cells.Font.Size = 11
    ws.Range("A1:F" & lastRow).Interior.Color = ThemeTone(themeKey, "base")

    PaintUniversalBlock ws.Range("A1:F1"), ThemeTone(themeKey, "titleDark"), ThemeTone(themeKey, "textLight"), 13, True, True
    PaintUniversalBlock ws.Range("A2:F2"), ThemeTone(themeKey, "header"), ThemeTone(themeKey, "textDark"), 11, True, True
    PaintUniversalBlock ws.Range("A3:F" & lastRow), ThemeTone(themeKey, "table"), ThemeTone(themeKey, "textDark"), 10.5, False, False

    For rowIndex = 3 To lastRow
        familyName = Trim$(CStr(ws.Cells(rowIndex, "A").Value))
        If Len(familyName) > 0 Then
            ws.Cells(rowIndex, "A").Interior.Color = UniversalPrimaryColor(familyName)
            ws.Cells(rowIndex, "A").Font.Color = ThemeTone(themeKey, "textLight")
        End If
    Next rowIndex

    ws.Rows("1:" & lastRow).RowHeight = 24
    ApplyUniversalGrid ws.Range("A1:F" & lastRow)
    ws.Tab.Color = ThemeTone(themeKey, "tab1")
End Sub

Private Sub ApplyHerbCatalogTheme(ByVal ws As Worksheet, ByVal themeKey As String)
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim stageName As String

    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 4 Then Exit Sub

    ws.Cells.Font.Name = "KaiTi"
    ws.Cells.Font.Size = 11
    ws.Range("A1:E" & lastRow).Interior.Color = ThemeTone(themeKey, "base")

    PaintUniversalBlock ws.Range("A1:E1"), ThemeTone(themeKey, "titleDark"), ThemeTone(themeKey, "textLight"), 13, True, True
    PaintUniversalBlock ws.Range("A2:E2"), ThemeTone(themeKey, "note"), ThemeTone(themeKey, "textDark"), 10.5, False, False
    PaintUniversalBlock ws.Range("A4:E4"), ThemeTone(themeKey, "header"), ThemeTone(themeKey, "textDark"), 11, True, True
    PaintUniversalBlock ws.Range("A5:E" & lastRow), ThemeTone(themeKey, "table"), ThemeTone(themeKey, "textDark"), 10.5, False, False

    For rowIndex = 5 To lastRow
        stageName = Trim$(CStr(ws.Cells(rowIndex, "A").Value))
        If Len(stageName) > 0 Then
            ws.Cells(rowIndex, "A").Interior.Color = ThemeTone(themeKey, "header")
            ws.Cells(rowIndex, "B").Interior.Color = ThemeTone(themeKey, "note")
        End If
    Next rowIndex

    ws.Rows("1:" & lastRow).RowHeight = 24
    ApplyUniversalGrid ws.Range("A1:E" & lastRow)
    ws.Tab.Color = ThemeTone(themeKey, "tab1")
End Sub

Private Sub ApplyFollowerSheetTheme(ByVal ws As Worksheet, ByVal themeKey As String)
    Dim blockStarts As Variant
    Dim blockStart As Variant
    Dim startRow As Long
    Dim dataRow As Long

    If ws Is Nothing Then Exit Sub

    ws.Cells.Font.Name = "KaiTi"
    ws.Cells.Font.Size = 11
    ws.Range("A1:R56").Interior.Color = ThemeTone(themeKey, "base")

    blockStarts = Array(1, 15, 29, 43)
    For Each blockStart In blockStarts
        startRow = CLng(blockStart)

        PaintUniversalBlock ws.Range("A" & startRow & ":R" & startRow), ThemeTone(themeKey, "titleDark"), ThemeTone(themeKey, "textLight"), 13, True, True

        PaintUniversalBlock ws.Range("A" & (startRow + 1) & ":F" & (startRow + 4)), ThemeTone(themeKey, "table"), ThemeTone(themeKey, "textDark"), 11, False, False
        PaintUniversalBlock ws.Range("A" & (startRow + 1) & ":A" & (startRow + 4)), ThemeTone(themeKey, "header"), ThemeTone(themeKey, "textDark"), 11, True, True
        PaintUniversalBlock ws.Range("C" & (startRow + 1) & ":C" & (startRow + 4)), ThemeTone(themeKey, "header"), ThemeTone(themeKey, "textDark"), 11, True, True
        PaintUniversalBlock ws.Range("E" & (startRow + 1) & ":E" & (startRow + 4)), ThemeTone(themeKey, "header"), ThemeTone(themeKey, "textDark"), 11, True, True
        PaintUniversalBlock ws.Range("G" & (startRow + 1) & ":G" & (startRow + 4)), ThemeTone(themeKey, "header"), ThemeTone(themeKey, "textDark"), 10.5, True, True
        PaintUniversalBlock ws.Range("H" & (startRow + 1) & ":J" & (startRow + 4)), ThemeTone(themeKey, "table"), ThemeTone(themeKey, "textDark"), 10.5, False, True
        PaintUniversalBlock ws.Range("K" & (startRow + 1) & ":R" & (startRow + 4)), ThemeTone(themeKey, "note"), ThemeTone(themeKey, "textDark"), 10.5, False, False

        PaintUniversalBlock ws.Range("A" & (startRow + 6) & ":G" & (startRow + 6)), ThemeTone(themeKey, "titleMid"), ThemeTone(themeKey, "textLight"), 11.5, True, True
        PaintUniversalBlock ws.Range("H" & (startRow + 6) & ":M" & (startRow + 6)), ThemeTone(themeKey, "titleMid"), ThemeTone(themeKey, "textLight"), 11.5, True, True
        PaintUniversalBlock ws.Range("N" & (startRow + 6) & ":R" & (startRow + 6)), ThemeTone(themeKey, "titleAlt"), ThemeTone(themeKey, "textLight"), 11.5, True, True

        PaintUniversalBlock ws.Range("A" & (startRow + 7) & ":G" & (startRow + 7)), ThemeTone(themeKey, "header"), ThemeTone(themeKey, "textDark"), 10.5, True, True
        PaintUniversalBlock ws.Range("H" & (startRow + 7) & ":M" & (startRow + 7)), ThemeTone(themeKey, "header"), ThemeTone(themeKey, "textDark"), 10.5, True, True
        PaintUniversalBlock ws.Range("N" & (startRow + 7) & ":R" & (startRow + 7)), ThemeTone(themeKey, "header"), ThemeTone(themeKey, "textDark"), 10.5, True, True

        For dataRow = startRow + 8 To startRow + 11
            PaintUniversalBlock ws.Range("A" & dataRow & ":G" & dataRow), ThemeTone(themeKey, "table"), ThemeTone(themeKey, "textDark"), 10.5, False, False
            PaintUniversalBlock ws.Range("H" & dataRow & ":M" & dataRow), ThemeTone(themeKey, "table"), ThemeTone(themeKey, "textDark"), 10.5, False, False
            PaintUniversalBlock ws.Range("N" & dataRow & ":R" & dataRow), ThemeTone(themeKey, "table"), ThemeTone(themeKey, "textDark"), 10.5, False, False
        Next dataRow

        ApplyUniversalGrid ws.Range("A" & startRow & ":R" & (startRow + 11))
    Next blockStart

    ws.Rows("1:56").RowHeight = 24
    ws.Tab.Color = ThemeTone(themeKey, "tab1")
End Sub

Private Sub RefreshEquipmentReferenceLists(ByVal equipmentWs As Worksheet, ByVal personalWs As Worksheet, ByVal followerWs As Worksheet)
    Dim blockStarts As Variant
    Dim blockStart As Variant
    Dim dataRow As Long

    If equipmentWs Is Nothing Then Exit Sub

    If Not personalWs Is Nothing Then
        ApplyValidationFormula personalWs.Range("F13"), BuildValidationFormula(equipmentWs, "V", 5, 100)
        ApplyValidationFormula personalWs.Range("F15"), BuildValidationFormula(equipmentWs, "W", 5, 120)
        ApplyValidationFormula personalWs.Range("F17"), BuildValidationFormula(equipmentWs, "X", 5, 100)
        ApplyValidationFormula personalWs.Range("F19"), BuildValidationFormula(equipmentWs, "Y", 5, 150)
        ApplyValidationFormula personalWs.Range("F21"), BuildValidationFormula(equipmentWs, "Y", 5, 150)

        personalWs.Range("H13").Formula = "=IF(OR(F13="""",F13=""空""),"""",IFERROR(VLOOKUP(F13,'装备列表'!$M$5:$T$400,2,FALSE),""""))"
        personalWs.Range("H15").Formula = "=IF(OR(F15="""",F15=""空""),"""",IFERROR(VLOOKUP(F15,'装备列表'!$M$5:$T$400,2,FALSE),""""))"
        personalWs.Range("H17").Formula = "=IF(OR(F17="""",F17=""空""),"""",IFERROR(VLOOKUP(F17,'装备列表'!$M$5:$T$400,2,FALSE),""""))"
        personalWs.Range("H19").Formula = "=IF(OR(F19="""",F19=""空""),"""",IFERROR(VLOOKUP(F19,'装备列表'!$M$5:$T$400,2,FALSE),""""))"
        personalWs.Range("H21").Formula = "=IF(OR(F21="""",F21=""空""),"""",IFERROR(VLOOKUP(F21,'装备列表'!$M$5:$T$400,2,FALSE),""""))"
    End If

    If followerWs Is Nothing Then Exit Sub

    blockStarts = Array(1, 15, 29, 43)
    For Each blockStart In blockStarts
        ApplyValidationFormula followerWs.Range("H" & (CLng(blockStart) + 1)), BuildValidationFormula(equipmentWs, "V", 5, 100)
        ApplyValidationFormula followerWs.Range("H" & (CLng(blockStart) + 2)), BuildValidationFormula(equipmentWs, "W", 5, 120)
        ApplyValidationFormula followerWs.Range("H" & (CLng(blockStart) + 3)), BuildValidationFormula(equipmentWs, "X", 5, 100)
        ApplyValidationFormula followerWs.Range("H" & (CLng(blockStart) + 4)), BuildValidationFormula(equipmentWs, "Y", 5, 150)
        ApplyValidationFormula followerWs.Range("H" & (CLng(blockStart) + 8) & ":H" & (CLng(blockStart) + 11)), BuildValidationFormula(equipmentWs, "Z", 5, 200)

        followerWs.Range("K" & (CLng(blockStart) + 1)).Formula = "=IF(OR(H" & (CLng(blockStart) + 1) & "="""",H" & (CLng(blockStart) + 1) & "=""空""),"""",IFERROR(VLOOKUP(H" & (CLng(blockStart) + 1) & ",'装备列表'!$M$5:$T$400,2,FALSE),""""))"
        followerWs.Range("K" & (CLng(blockStart) + 2)).Formula = "=IF(OR(H" & (CLng(blockStart) + 2) & "="""",H" & (CLng(blockStart) + 2) & "=""空""),"""",IFERROR(VLOOKUP(H" & (CLng(blockStart) + 2) & ",'装备列表'!$M$5:$T$400,2,FALSE),""""))"
        followerWs.Range("K" & (CLng(blockStart) + 3)).Formula = "=IF(OR(H" & (CLng(blockStart) + 3) & "="""",H" & (CLng(blockStart) + 3) & "=""空""),"""",IFERROR(VLOOKUP(H" & (CLng(blockStart) + 3) & ",'装备列表'!$M$5:$T$400,2,FALSE),""""))"
        followerWs.Range("K" & (CLng(blockStart) + 4)).Formula = "=IF(OR(H" & (CLng(blockStart) + 4) & "="""",H" & (CLng(blockStart) + 4) & "=""空""),"""",IFERROR(VLOOKUP(H" & (CLng(blockStart) + 4) & ",'装备列表'!$M$5:$T$400,2,FALSE),""""))"

        For dataRow = CLng(blockStart) + 8 To CLng(blockStart) + 11
            followerWs.Range("J" & dataRow).Formula = "=IF(OR(H" & dataRow & "="""",H" & dataRow & "=""空""),"""",IFERROR(VLOOKUP(H" & dataRow & ",'装备列表'!$M$5:$T$400,2,FALSE),""""))"
        Next dataRow
    Next blockStart
End Sub

Private Sub RefreshUniversalReferenceLists(ByVal themeKey As String, ByVal traitWs As Worksheet, ByVal skillWs As Worksheet, ByVal bookWs As Worksheet, ByVal followerWs As Worksheet)
    Dim traitRow As Long
    Dim skillRow As Long
    Dim sourceRow As Long
    Dim lastTraitRow As Long
    Dim lastSkillRow As Long
    Dim familyName As String
    Dim blockStarts As Variant
    Dim blockStart As Variant
    Dim dataRow As Long

    If bookWs Is Nothing Then Exit Sub

    bookWs.Range("W1:AD400").ClearContents
    bookWs.Range("W1:X1").Value = Array("特性名称", "特性介绍")
    bookWs.Range("Y1:Z1").Value = Array("家族", "分组")
    bookWs.Range("AA1:AD1").Value = Array("技能名称", "技能介绍", "技能速度", "技能消耗")

    traitRow = 2
    bookWs.Cells(traitRow, "W").Value = "空"
    bookWs.Cells(traitRow, "X").Value = ""
    bookWs.Cells(traitRow, "Y").Value = "通用"
    bookWs.Cells(traitRow, "Z").Value = "占位"
    traitRow = traitRow + 1

    If Not traitWs Is Nothing Then
        lastTraitRow = traitWs.Cells(traitWs.Rows.Count, "A").End(xlUp).Row
        For sourceRow = 3 To lastTraitRow
            familyName = Trim$(CStr(traitWs.Cells(sourceRow, "A").Value))
            If StrComp(familyName, "通用", vbTextCompare) = 0 Or StrComp(familyName, themeKey, vbTextCompare) = 0 Then
                bookWs.Cells(traitRow, "W").Value = traitWs.Cells(sourceRow, "C").Value
                bookWs.Cells(traitRow, "X").Value = traitWs.Cells(sourceRow, "D").Value
                bookWs.Cells(traitRow, "Y").Value = familyName
                bookWs.Cells(traitRow, "Z").Value = traitWs.Cells(sourceRow, "B").Value
                traitRow = traitRow + 1
            End If
        Next sourceRow
    End If

    skillRow = 2
    bookWs.Cells(skillRow, "AA").Value = "空"
    bookWs.Cells(skillRow, "AB").Value = ""
    bookWs.Cells(skillRow, "AC").Value = ""
    bookWs.Cells(skillRow, "AD").Value = ""
    skillRow = skillRow + 1

    If Not skillWs Is Nothing Then
        lastSkillRow = skillWs.Cells(skillWs.Rows.Count, "A").End(xlUp).Row
        For sourceRow = 3 To lastSkillRow
            familyName = Trim$(CStr(skillWs.Cells(sourceRow, "A").Value))
            If StrComp(familyName, "通用", vbTextCompare) = 0 Or StrComp(familyName, themeKey, vbTextCompare) = 0 Then
                bookWs.Cells(skillRow, "AA").Value = skillWs.Cells(sourceRow, "C").Value
                bookWs.Cells(skillRow, "AB").Value = skillWs.Cells(sourceRow, "D").Value
                bookWs.Cells(skillRow, "AC").Value = skillWs.Cells(sourceRow, "E").Value
                bookWs.Cells(skillRow, "AD").Value = skillWs.Cells(sourceRow, "F").Value
                skillRow = skillRow + 1
            End If
        Next sourceRow
    End If

    ApplyValidationFormula bookWs.Range("L2:L17"), BuildValidationFormula(bookWs, "W", 2, 300)
    ApplyValidationFormula bookWs.Range("A20:A35"), BuildValidationFormula(bookWs, "AA", 2, 400)
    ApplyValidationFormula bookWs.Range("L20:L35"), BuildValidationFormula(bookWs, "AA", 2, 400)

    For sourceRow = 2 To 17
        bookWs.Range("M" & sourceRow).Formula = "=IF(OR(L" & sourceRow & "="""",L" & sourceRow & "=""空""),"""",IFERROR(VLOOKUP(L" & sourceRow & ",$W$2:$X$300,2,FALSE),""""))"
    Next sourceRow

    For sourceRow = 20 To 35
        bookWs.Range("B" & sourceRow).Formula = "=IF(OR(A" & sourceRow & "="""",A" & sourceRow & "=""空""),"""",IFERROR(VLOOKUP(A" & sourceRow & ",$AA$2:$AD$400,2,FALSE),""""))"
        bookWs.Range("I" & sourceRow).Formula = "=IF(OR(A" & sourceRow & "="""",A" & sourceRow & "=""空""),"""",IFERROR(VLOOKUP(A" & sourceRow & ",$AA$2:$AD$400,3,FALSE),""""))"
        bookWs.Range("J" & sourceRow).Formula = "=IF(OR(A" & sourceRow & "="""",A" & sourceRow & "=""空""),"""",IFERROR(VLOOKUP(A" & sourceRow & ",$AA$2:$AD$400,4,FALSE),""""))"
        bookWs.Range("M" & sourceRow).Formula = "=IF(OR(L" & sourceRow & "="""",L" & sourceRow & "=""空""),"""",IFERROR(VLOOKUP(L" & sourceRow & ",$AA$2:$AD$400,2,FALSE),""""))"
        bookWs.Range("T" & sourceRow).Formula = "=IF(OR(L" & sourceRow & "="""",L" & sourceRow & "=""空""),"""",IFERROR(VLOOKUP(L" & sourceRow & ",$AA$2:$AD$400,3,FALSE),""""))"
        bookWs.Range("U" & sourceRow).Formula = "=IF(OR(L" & sourceRow & "="""",L" & sourceRow & "=""空""),"""",IFERROR(VLOOKUP(L" & sourceRow & ",$AA$2:$AD$400,4,FALSE),""""))"
    Next sourceRow

    If followerWs Is Nothing Then Exit Sub

    blockStarts = Array(1, 15, 29, 43)
    For Each blockStart In blockStarts
        ApplyValidationFormula followerWs.Range("A" & (CLng(blockStart) + 8) & ":A" & (CLng(blockStart) + 11)), BuildValidationFormula(bookWs, "AA", 2, 400)

        For dataRow = CLng(blockStart) + 8 To CLng(blockStart) + 11
            followerWs.Range("B" & dataRow).Formula = "=IF(OR(A" & dataRow & "="""",A" & dataRow & "=""空""),"""",IFERROR(VLOOKUP(A" & dataRow & ",'书籍、特性与技能'!$AA$2:$AD$400,2,FALSE),""""))"
        Next dataRow
    Next blockStart
End Sub

Private Function BuildValidationFormula(ByVal sourceWs As Worksheet, ByVal helperColumn As String, ByVal firstRow As Long, ByVal lastRow As Long) As String
    BuildValidationFormula = "=OFFSET('" & sourceWs.Name & "'!$" & helperColumn & "$" & firstRow & ",0,0,COUNTA('" & sourceWs.Name & "'!$" & helperColumn & "$" & firstRow & ":$" & helperColumn & "$" & lastRow & "),1)"
End Function

Private Sub ApplyValidationFormula(ByVal targetRange As Range, ByVal formulaText As String)
    With targetRange.Validation
        .Delete
        .Add VALIDATION_LIST, VALIDATION_ALERT_STOP, VALIDATION_BETWEEN, formulaText
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

Private Function TryGetWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set TryGetWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function UniversalPrimaryColor(ByVal themeKey As String) As Long
    Select Case themeKey
        Case "亓家": UniversalPrimaryColor = RGB(140, 34, 52)
        Case "林家": UniversalPrimaryColor = RGB(63, 134, 120)
        Case "易家": UniversalPrimaryColor = RGB(69, 109, 181)
        Case "巫家": UniversalPrimaryColor = RGB(182, 104, 139)
        Case "岳家": UniversalPrimaryColor = RGB(181, 78, 52)
        Case "药家": UniversalPrimaryColor = RGB(74, 132, 62)
        Case "张家": UniversalPrimaryColor = RGB(116, 82, 196)
        Case "曹家": UniversalPrimaryColor = RGB(24, 140, 148)
        Case "墨家": UniversalPrimaryColor = RGB(77, 96, 112)
        Case "乐正家": UniversalPrimaryColor = RGB(186, 144, 56)
        Case "通用": UniversalPrimaryColor = RGB(86, 89, 94)
        Case Else: UniversalPrimaryColor = RGB(86, 89, 94)
    End Select
End Function

Private Function ThemeTone(ByVal themeKey As String, ByVal toneName As String) As Long
    Dim primary As Long
    primary = UniversalPrimaryColor(themeKey)

    Select Case toneName
        Case "base": ThemeTone = LightenColor(primary, 0.92)
        Case "panel": ThemeTone = LightenColor(primary, 0.82)
        Case "table": ThemeTone = LightenColor(primary, 0.96)
        Case "titleDark": ThemeTone = DarkenColor(primary, 0.48)
        Case "titleMid": ThemeTone = DarkenColor(primary, 0.72)
        Case "titleAlt": ThemeTone = DarkenColor(primary, 0.60)
        Case "header": ThemeTone = LightenColor(primary, 0.66)
        Case "note": ThemeTone = LightenColor(primary, 0.86)
        Case "textDark": ThemeTone = RGB(46, 46, 46)
        Case "textLight": ThemeTone = RGB(249, 249, 249)
        Case "tab1": ThemeTone = DarkenColor(primary, 0.72)
        Case "tab2": ThemeTone = DarkenColor(primary, 0.48)
        Case Else: ThemeTone = primary
    End Select
End Function

Private Function LightenColor(ByVal baseColor As Long, ByVal ratio As Double) As Long
    Dim redPart As Long
    Dim greenPart As Long
    Dim bluePart As Long

    redPart = baseColor Mod 256
    greenPart = (baseColor \ 256) Mod 256
    bluePart = (baseColor \ 65536) Mod 256

    redPart = CLng(redPart + (255 - redPart) * ratio)
    greenPart = CLng(greenPart + (255 - greenPart) * ratio)
    bluePart = CLng(bluePart + (255 - bluePart) * ratio)

    If redPart > 255 Then redPart = 255
    If greenPart > 255 Then greenPart = 255
    If bluePart > 255 Then bluePart = 255

    LightenColor = RGB(redPart, greenPart, bluePart)
End Function

Private Function DarkenColor(ByVal baseColor As Long, ByVal ratio As Double) As Long
    Dim redPart As Long
    Dim greenPart As Long
    Dim bluePart As Long

    redPart = baseColor Mod 256
    greenPart = (baseColor \ 256) Mod 256
    bluePart = (baseColor \ 65536) Mod 256

    redPart = CLng(redPart * ratio)
    greenPart = CLng(greenPart * ratio)
    bluePart = CLng(bluePart * ratio)

    If redPart < 0 Then redPart = 0
    If greenPart < 0 Then greenPart = 0
    If bluePart < 0 Then bluePart = 0

    DarkenColor = RGB(redPart, greenPart, bluePart)
End Function

Private Sub PaintUniversalBlock(ByVal rng As Range, ByVal fillColor As Long, ByVal fontColor As Long, ByVal fontSize As Double, Optional ByVal makeBold As Boolean = False, Optional ByVal centerText As Boolean = False)
    rng.Interior.Color = fillColor
    rng.Font.Color = fontColor
    rng.Font.Size = fontSize
    rng.Font.Bold = makeBold
    rng.WrapText = True

    If centerText Then
        rng.HorizontalAlignment = xlHAlignCenter
        rng.VerticalAlignment = xlVAlignCenter
    End If
End Sub

Private Sub ApplyUniversalGrid(ByVal rng As Range)
    rng.Borders(xlEdgeLeft).LineStyle = xlContinuous
    rng.Borders(xlEdgeTop).LineStyle = xlContinuous
    rng.Borders(xlEdgeBottom).LineStyle = xlContinuous
    rng.Borders(xlEdgeRight).LineStyle = xlContinuous
    rng.Borders(xlInsideVertical).LineStyle = xlContinuous
    rng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    rng.Borders(xlEdgeLeft).Weight = xlThin
    rng.Borders(xlEdgeTop).Weight = xlThin
    rng.Borders(xlEdgeBottom).Weight = xlThin
    rng.Borders(xlEdgeRight).Weight = xlThin
    rng.Borders(xlInsideVertical).Weight = xlThin
    rng.Borders(xlInsideHorizontal).Weight = xlThin
End Sub
