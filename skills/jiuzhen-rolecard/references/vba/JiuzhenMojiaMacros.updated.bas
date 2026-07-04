Attribute VB_Name = "JiuzhenMojiaMacros"
Option Explicit

Private Const FIRST_SLOT_ROW As Long = 5
Private Const BASE_SLOT_COUNT As Long = 15
Private Const MAX_SLOT_COUNT As Long = 300
Private Const LAST_SLOT_ROW As Long = FIRST_SLOT_ROW + MAX_SLOT_COUNT - 1
Private Const SLOT_TOTAL_CELL As String = "I5"
Private Const SLOT_SELF_USED_CELL As String = "I6"
Private Const SLOT_KUI_USED_CELL As String = "I7"
Private Const SLOT_REMAIN_CELL As String = "I8"
Private Const MOJIA_HELPER_COL As String = "F"

Private Const KUI_PLUGIN_FIRST_ROW As Long = 5
Private Const KUI_PLUGIN_LAST_ROW As Long = 14
Private Const KUI_PLUGIN_COL As String = "U"
Private Const KUI_PLUGIN_COST_COL As String = "V"
Private Const KUI_PLUGIN_DESC_COL As String = "W"
Private Const KUI_SINGLE_MODULE_SLOT_COST As Long = 5
Private Const KUI_MODULE_RENDER_PREFIX As String = "【魁人模组】"
Private Const KUI_WEAPON_RENDER_PREFIX As String = "【魁人武器】"

Private Const KUI_WEAPON_FIRST_ROW As Long = 17
Private Const KUI_WEAPON_LAST_ROW As Long = 19
Private Const KUI_WEAPON_COL As String = "P"
Private Const KUI_WEAPON_COST_COL As String = "Q"
Private Const KUI_WEAPON_DESC_COL As String = "R"

Private Const xlValidateList As Long = 3
Private Const xlValidAlertStop As Long = 1
Private Const xlBetween As Long = 1
Private Const xlHAlignCenter As Long = -4108
Private Const xlVAlignCenter As Long = -4108
Private Const xlValues As Long = -4163
Private Const xlEdgeLeft As Long = 7
Private Const xlEdgeTop As Long = 8
Private Const xlEdgeBottom As Long = 9
Private Const xlEdgeRight As Long = 10
Private Const xlInsideVertical As Long = 11
Private Const xlInsideHorizontal As Long = 12
Private Const xlContinuous As Long = 1
Private Const xlThin As Long = 2
Private Const xlSheetVisible As Long = -1
Private Const xlSheetHidden As Long = 0
Private Const xlSheetVeryHidden As Long = 2

Public Sub ApplyWorkbookTheme()
    On Error GoTo CleanFail

    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim themeKey As String

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    themeKey = CurrentThemeKey()

    ApplyMojiaSheetVisibility themeKey
    ApplyThemeSettings Sheet7, themeKey
    ApplyThemePersonalInfo Sheet1, themeKey
    ApplyThemeBookAndTraits Sheet2, themeKey
    ApplyThemeMojia Sheet3, themeKey
    ApplyThemeCatalog Sheet5, themeKey
    ApplyThemeFastCopy Sheet6, themeKey
    ApplyUniversalWorkbookTheme themeKey

    RefreshMojiaInventory

CleanExit:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

CleanFail:
    MsgBox "Apply workbook theme failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Public Sub RefreshMojiaInventory()
    On Error GoTo CleanFail

    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim ws As Worksheet
    Dim catalogWs As Worksheet
    Dim itemNames As Collection
    Dim totalSlots As Long
    Dim visibleSlots As Long
    Dim rejectedItems As String

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = Sheet3
    Set catalogWs = Sheet5
    Set itemNames = New Collection

    EnsureMojiaValidation ws, catalogWs
    EnsureGrowthAmountValidation ws
    EnsureBattleAbilityValidation ws, catalogWs
    EnsureKuiPluginValidation ws, catalogWs
    EnsureKuiWeaponValidation ws, catalogWs
    EnsureKuiDerivedFields ws

    CollectKuiEquipmentEntries ws, itemNames
    CollectItemEntries ws, catalogWs, "AE", FIRST_SLOT_ROW, LAST_SLOT_ROW, itemNames

    totalSlots = ReadLong(ws.Range(SLOT_TOTAL_CELL).Value, BASE_SLOT_COUNT)
    If totalSlots < BASE_SLOT_COUNT Then totalSlots = BASE_SLOT_COUNT
    If totalSlots > MAX_SLOT_COUNT Then totalSlots = MAX_SLOT_COUNT

    visibleSlots = totalSlots

    ResetSlotArea ws, FIRST_SLOT_ROW, LAST_SLOT_ROW, MOJIA_HELPER_COL
    ApplyInventoryValidation ws, catalogWs, FIRST_SLOT_ROW, LAST_SLOT_ROW
    rejectedItems = RenderItems(ws, catalogWs, itemNames, visibleSlots, FIRST_SLOT_ROW, LAST_SLOT_ROW, "AR", "AS", MOJIA_HELPER_COL)
    ApplySlotGrid ws, visibleSlots, FIRST_SLOT_ROW
    ApplyRowVisibility ws, visibleSlots, FIRST_SLOT_ROW, LAST_SLOT_ROW

    If Len(rejectedItems) > 0 Then
        MsgBox "工巧格不足，以下物品未能保留：" & vbCrLf & rejectedItems, vbExclamation
    End If

CleanExit:
    Application.CutCopyMode = False
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

CleanFail:
    MsgBox "Refresh Mojia inventory failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Public Sub RefreshKuiInventory()
    RefreshMojiaInventory
End Sub

Public Function CanSaveMojiaWorkbook(ByRef validationMessage As String) As Boolean
    CanSaveMojiaWorkbook = True

    If StrComp(CurrentThemeKey(), "墨家", vbTextCompare) <> 0 Then
        Exit Function
    End If

    CanSaveMojiaWorkbook = ValidateKuiModuleState(Sheet3, validationMessage)
End Function

Public Function ValidateKuiModuleSelectionChange(ByVal ws As Worksheet, ByVal Target As Range, ByRef validationMessage As String) As Boolean
    ValidateKuiModuleSelectionChange = True

    If Intersect(Target, ws.Range(KUI_PLUGIN_COL & KUI_PLUGIN_FIRST_ROW & ":" & KUI_PLUGIN_COL & KUI_PLUGIN_LAST_ROW)) Is Nothing Then
        Exit Function
    End If

    ValidateKuiModuleSelectionChange = ValidateKuiModuleState(ws, validationMessage)
End Function

Private Function ValidateKuiModuleState(ByVal ws As Worksheet, ByRef validationMessage As String) As Boolean
    Dim moduleCount As Long
    Dim moduleCap As Long
    Dim totalSlots As Long
    Dim personalUsed As Long
    Dim weaponUsed As Long
    Dim requiredSlots As Long

    moduleCount = CountNonBlankCells(ws.Range(KUI_PLUGIN_COL & KUI_PLUGIN_FIRST_ROW & ":" & KUI_PLUGIN_COL & KUI_PLUGIN_LAST_ROW))
    moduleCap = ReadLong(ws.Range("R6").Value, 1)

    If moduleCount > moduleCap Then
        validationMessage = "当前工巧格数量不足以解锁下一模组槽位"
        ValidateKuiModuleState = False
        Exit Function
    End If

    totalSlots = ReadLong(ws.Range(SLOT_TOTAL_CELL).Value, BASE_SLOT_COUNT)
    personalUsed = SumBagEntrySlots(ws, Sheet5, "AE", "AF")
    weaponUsed = SumSelectedEquipmentSlots(ws.Range(KUI_WEAPON_COST_COL & KUI_WEAPON_FIRST_ROW & ":" & KUI_WEAPON_COST_COL & KUI_WEAPON_LAST_ROW))
    requiredSlots = personalUsed + weaponUsed + moduleCount * KUI_SINGLE_MODULE_SLOT_COST

    If requiredSlots > totalSlots Then
        validationMessage = "当前工巧格剩余槽位不足（每个魁人模组需要消耗5工巧格）"
        ValidateKuiModuleState = False
        Exit Function
    End If

    ValidateKuiModuleState = True
End Function

Private Sub ApplyMojiaSheetVisibility(ByVal themeKey As String)
    Dim showMojia As Boolean

    showMojia = (StrComp(themeKey, "墨家", vbTextCompare) = 0)

    If Not showMojia Then
        If ActiveSheet.Name = Sheet3.Name Or ActiveSheet.Name = Sheet5.Name Then
            Sheet1.Activate
        End If
    End If

    Sheet3.Visible = IIf(showMojia, xlSheetVisible, xlSheetHidden)
    Sheet5.Visible = IIf(showMojia, xlSheetVisible, xlSheetHidden)
    Sheet4.Visible = xlSheetVeryHidden
End Sub

Private Sub CollectItemEntries(ByVal ws As Worksheet, ByVal catalogWs As Worksheet, ByVal helperNameColumn As String, ByVal firstRow As Long, ByVal lastRow As Long, ByRef itemNames As Collection)
    Dim rowIndex As Long
    Dim cellText As String

    For rowIndex = firstRow To lastRow
        cellText = Trim$(CStr(ws.Cells(rowIndex, "B").Value))
        If Len(cellText) > 0 Then
            If StrComp(cellText, "空", vbTextCompare) <> 0 Then
                If Not ItemExistsInHelper(catalogWs, cellText, helperNameColumn) Then
                    GoTo ContinueLoop
                End If
                If ws.Cells(rowIndex, "B").MergeCells Then
                    If ws.Cells(rowIndex, "B").MergeArea.Row = rowIndex Then
                        itemNames.Add cellText
                    End If
                Else
                    itemNames.Add cellText
                End If
            End If
        End If
ContinueLoop:
    Next rowIndex
End Sub

Private Sub CollectKuiEquipmentEntries(ByVal ws As Worksheet, ByRef itemNames As Collection)
    Dim rowIndex As Long
    Dim itemName As String

    For rowIndex = KUI_PLUGIN_FIRST_ROW To KUI_PLUGIN_LAST_ROW
        itemName = Trim$(CStr(ws.Range(KUI_PLUGIN_COL & rowIndex).Value))
        If Len(itemName) > 0 And StrComp(itemName, "空", vbTextCompare) <> 0 Then
            itemNames.Add KUI_MODULE_RENDER_PREFIX & itemName
        End If
    Next rowIndex

    For rowIndex = KUI_WEAPON_FIRST_ROW To KUI_WEAPON_LAST_ROW
        itemName = Trim$(CStr(ws.Range(KUI_WEAPON_COL & rowIndex).Value))
        If Len(itemName) > 0 And StrComp(itemName, "空", vbTextCompare) <> 0 Then
            itemNames.Add KUI_WEAPON_RENDER_PREFIX & itemName
        End If
    Next rowIndex
End Sub

Private Sub ResetSlotArea(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal helperColumn As String)
    Dim rowIndex As Long
    Dim bodyRange As Range
    Dim itemRange As Range
    Dim themeKey As String

    themeKey = CurrentThemeKey()

    ws.Rows(firstRow & ":" & lastRow).Hidden = False

    Set bodyRange = ws.Range("A" & firstRow & ":E" & lastRow)
    Set itemRange = ws.Range("B" & firstRow & ":E" & lastRow)

    itemRange.UnMerge
    bodyRange.ClearContents
    ws.Range(helperColumn & firstRow & ":" & helperColumn & lastRow).ClearContents

    bodyRange.Interior.Color = Palette(themeKey, "table")
    ws.Range("A" & firstRow & ":A" & lastRow).Interior.Color = Palette(themeKey, "slot")
    ws.Range("A" & firstRow & ":A" & lastRow).HorizontalAlignment = xlHAlignCenter
    ws.Range("A" & firstRow & ":A" & lastRow).VerticalAlignment = xlVAlignCenter
    ws.Range("B" & firstRow & ":E" & lastRow).HorizontalAlignment = xlHAlignCenter
    ws.Range("B" & firstRow & ":E" & lastRow).VerticalAlignment = xlVAlignCenter
    ws.Range("B" & firstRow & ":E" & lastRow).WrapText = True

    For rowIndex = firstRow To lastRow
        ws.Cells(rowIndex, "A").Value = rowIndex - firstRow + 1
    Next rowIndex
End Sub

Private Function RenderItems(ByVal ws As Worksheet, ByVal catalogWs As Worksheet, ByVal itemNames As Collection, ByVal totalSlots As Long, ByVal firstRow As Long, ByVal lastRow As Long, ByVal helperNameColumn As String, ByVal helperSlotsColumn As String, ByVal markerColumn As String) As String
    Dim itemName As Variant
    Dim slotSize As Long
    Dim currentRow As Long
    Dim endRow As Long
    Dim rowIndex As Long
    Dim usedSlots As Long
    Dim fillRange As Range
    Dim themeKey As String
    Dim rejected As String
    Dim itemIndex As Long

    themeKey = CurrentThemeKey()
    currentRow = firstRow
    usedSlots = 0
    rejected = vbNullString

    If totalSlots <= 0 Then
        RenderItems = vbNullString
        Exit Function
    End If

    For itemIndex = 1 To itemNames.Count
        itemName = itemNames(itemIndex)
        slotSize = GetItemSlotsFromHelper(catalogWs, CStr(itemName), helperNameColumn, helperSlotsColumn)
        If usedSlots + slotSize > totalSlots Then
            If Len(rejected) = 0 Then
                rejected = CStr(itemName)
            Else
                rejected = rejected & "、" & CStr(itemName)
            End If
        Else
            endRow = currentRow + slotSize - 1
            If endRow > lastRow Then Exit For

            Set fillRange = ws.Range("B" & currentRow & ":E" & endRow)
            fillRange.Merge
            fillRange.Value = CStr(itemName)
            fillRange.HorizontalAlignment = xlHAlignCenter
            fillRange.VerticalAlignment = xlVAlignCenter
            fillRange.WrapText = True
            fillRange.Interior.Color = Palette(themeKey, "fill")
            ws.Range(markerColumn & currentRow & ":" & markerColumn & endRow).Value = 1

            currentRow = endRow + 1
            usedSlots = usedSlots + slotSize
        End If
    Next itemIndex

    For rowIndex = firstRow To firstRow + totalSlots - 1
        If rowIndex > lastRow Then Exit For
        If Not ws.Cells(rowIndex, "B").MergeCells Then
            ws.Range("B" & rowIndex & ":E" & rowIndex).Merge
        End If
    Next rowIndex

    RenderItems = rejected
End Function

Private Sub ApplyInventoryValidation(ByVal ws As Worksheet, ByVal catalogWs As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    SetValidationFromHelper ws.Range("B" & firstRow & ":B" & lastRow), catalogWs, "AE", 5, 200
End Sub

Private Sub EnsureMojiaValidation(ByVal ws As Worksheet, ByVal catalogWs As Worksheet)
    SetValidationFromHelper ws.Range("K4:K13"), catalogWs, "AV", 5, 20
End Sub

Private Sub EnsureGrowthAmountValidation(ByVal ws As Worksheet)
    With ws.Range("L4:L13").Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, "1,2,3,4,5"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

Private Sub EnsureBattleAbilityValidation(ByVal ws As Worksheet, ByVal catalogWs As Worksheet)
    SetValidationFromHelper ws.Range("P7"), catalogWs, "AW", 5, 20
End Sub

Private Sub EnsureKuiPluginValidation(ByVal ws As Worksheet, ByVal catalogWs As Worksheet)
    Dim formulaText As String

    SetValidationFromHelper ws.Range("U5:U14"), catalogWs, "AJ", 5, 100

    formulaText = "=IF(OR(U5="""",U5=""空""),0,IFERROR(VLOOKUP(U5,'" & catalogWs.Name & "'!$AJ$5:$AK$100,2,FALSE),0))"
    ws.Range("V5").Formula = formulaText
    ws.Range("V5:V14").FillDown

    formulaText = "=IF(OR(U5="""",U5=""空""),"""",IFERROR(VLOOKUP(U5,'" & catalogWs.Name & "'!$AJ$5:$AL$100,3,FALSE),""""))"
    ws.Range("W5").Formula = formulaText
    ws.Range("W5:W14").FillDown
End Sub

Private Sub EnsureKuiWeaponValidation(ByVal ws As Worksheet, ByVal catalogWs As Worksheet)
    Dim formulaText As String

    SetValidationFromHelper ws.Range("P17:P19"), catalogWs, "AN", 5, 50

    formulaText = "=IF(OR(P17="""",P17=""空""),0,IFERROR(VLOOKUP(P17,'" & catalogWs.Name & "'!$AN$5:$AO$50,2,FALSE),0))"
    ws.Range("Q17").Formula = formulaText
    ws.Range("Q17:Q19").FillDown

    formulaText = "=IF(OR(P17="""",P17=""空""),"""",IFERROR(VLOOKUP(P17,'" & catalogWs.Name & "'!$AN$5:$AP$50,3,FALSE),""""))"
    ws.Range("R17").Formula = formulaText
    ws.Range("R17:R19").FillDown
End Sub

Private Sub EnsureKuiDerivedFields(ByVal ws As Worksheet)
    ws.Range("I3").Formula = "=15"
    ws.Range("I4").Formula = "=SUM(L4:L13)"
    ws.Range("I5").Formula = "=I3+I4"
    ws.Range("I6").Formula = "=SUM(F5:F304)"
    ws.Range("I7").Formula = "=R12"
    ws.Range("I8").Formula = "=I5-I6"

    ws.Range("P4").Value = "同主人"
    ws.Range("P5").Formula = "=5+COUNTIF($U$5:$U$14,""特种护甲插件"")*10"
    ws.Range("P6").Value = 20
    ws.Range("P8").Value = 60
    ws.Range("P9").Value = "2d6"
    If Len(Trim$(CStr(ws.Range("P13").Value))) = 0 Then
        ws.Range("P13").Value = "基础值"
    End If

    ws.Range("R4").Formula = "=SUM(P10:P12)"
    ws.Range("R5").Value = 200
    ws.Range("R6").Formula = "=1+MAX(0,INT(($I$5-10)/10))"
    ws.Range("R7").Formula = "=COUNTA($U$5:$U$14)-COUNTIF($U$5:$U$14,""空"")"
    ws.Range("R8").Formula = "=R6-R7"
    ws.Range("R9").Value = KUI_SINGLE_MODULE_SLOT_COST
    ws.Range("R10").Formula = "=R7*R9"
    ws.Range("R11").Formula = "=SUM($Q$17:$Q$19)"
    ws.Range("R12").Formula = "=R10+R11"
End Sub

Private Function CountNonBlankCells(ByVal rng As Range) As Long
    Dim cell As Range

    For Each cell In rng.Cells
        If Len(Trim$(CStr(cell.Value))) > 0 Then
            If StrComp(Trim$(CStr(cell.Value)), "空", vbTextCompare) <> 0 Then
                CountNonBlankCells = CountNonBlankCells + 1
            End If
        End If
    Next cell
End Function

Private Function SumBagEntrySlots(ByVal ws As Worksheet, ByVal catalogWs As Worksheet, ByVal helperNameColumn As String, ByVal helperSlotsColumn As String) As Long
    Dim rowIndex As Long
    Dim cellText As String

    For rowIndex = FIRST_SLOT_ROW To LAST_SLOT_ROW
        cellText = Trim$(CStr(ws.Cells(rowIndex, "B").Value))
        If Len(cellText) > 0 Then
            If StrComp(cellText, "空", vbTextCompare) <> 0 Then
                If ItemExistsInHelper(catalogWs, cellText, helperNameColumn) Then
                    If ws.Cells(rowIndex, "B").MergeCells Then
                        If ws.Cells(rowIndex, "B").MergeArea.Row = rowIndex Then
                            SumBagEntrySlots = SumBagEntrySlots + GetItemSlotsFromHelper(catalogWs, cellText, helperNameColumn, helperSlotsColumn)
                        End If
                    Else
                        SumBagEntrySlots = SumBagEntrySlots + GetItemSlotsFromHelper(catalogWs, cellText, helperNameColumn, helperSlotsColumn)
                    End If
                End If
            End If
        End If
    Next rowIndex
End Function

Private Function SumSelectedEquipmentSlots(ByVal rng As Range) As Long
    Dim cell As Range

    For Each cell In rng.Cells
        If IsNumeric(cell.Value) Then
            SumSelectedEquipmentSlots = SumSelectedEquipmentSlots + CLng(cell.Value)
        End If
    Next cell
End Function

Private Function ItemExistsInHelper(ByVal catalogWs As Worksheet, ByVal itemName As String, ByVal helperNameColumn As String) As Boolean
    Dim rowIndex As Long
    Dim currentName As String
    Dim lastCatalogRow As Long

    lastCatalogRow = catalogWs.Range(helperNameColumn & catalogWs.Rows.Count).End(xlUp).Row
    For rowIndex = 5 To lastCatalogRow
        currentName = Trim$(CStr(catalogWs.Range(helperNameColumn & rowIndex).Value))
        If StrComp(currentName, Trim$(itemName), vbTextCompare) = 0 Then
            ItemExistsInHelper = True
            Exit Function
        End If
    Next rowIndex
End Function

Private Function SumNumericRange(ByVal rng As Range) As Long
    Dim cell As Range

    For Each cell In rng.Cells
        If IsNumeric(cell.Value) Then
            SumNumericRange = SumNumericRange + CLng(cell.Value)
        End If
    Next cell
End Function

Private Sub SetValidationFromHelper(ByVal targetRange As Range, ByVal catalogWs As Worksheet, ByVal helperColumn As String, ByVal firstHelperRow As Long, ByVal lastHelperRow As Long)
    Dim formulaText As String

    formulaText = "=OFFSET('" & catalogWs.Name & "'!$" & helperColumn & "$" & firstHelperRow & ",0,0,COUNTA('" & catalogWs.Name & "'!$" & helperColumn & "$" & firstHelperRow & ":$" & helperColumn & "$" & lastHelperRow & "),1)"
    targetRange.Validation.Delete
    targetRange.Validation.Add xlValidateList, xlValidAlertStop, xlBetween, formulaText
    targetRange.Validation.IgnoreBlank = True
    targetRange.Validation.InCellDropdown = True
End Sub

Private Sub ApplyRowVisibility(ByVal ws As Worksheet, ByVal visibleSlots As Long, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim visibleEndRow As Long

    ws.Rows(firstRow & ":" & lastRow).Hidden = False

    If visibleSlots < 1 Then
        ws.Rows(firstRow & ":" & lastRow).Hidden = True
        Exit Sub
    End If

    visibleEndRow = firstRow + visibleSlots - 1
    If visibleEndRow > lastRow Then visibleEndRow = lastRow

    If visibleEndRow < lastRow Then
        ws.Rows(visibleEndRow + 1 & ":" & lastRow).Hidden = True
    End If
End Sub

Private Sub ApplySlotGrid(ByVal ws As Worksheet, ByVal visibleSlots As Long, ByVal firstRow As Long)
    Dim rowIndex As Long
    Dim rowRange As Range
    Dim themeKey As String

    themeKey = CurrentThemeKey()
    If visibleSlots < 1 Then Exit Sub

    For rowIndex = firstRow To firstRow + visibleSlots - 1
        Set rowRange = ws.Range("A" & rowIndex & ":E" & rowIndex)
        rowRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
        rowRange.Borders(xlEdgeTop).LineStyle = xlContinuous
        rowRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
        rowRange.Borders(xlEdgeRight).LineStyle = xlContinuous
        rowRange.Borders(xlEdgeLeft).Weight = xlThin
        rowRange.Borders(xlEdgeTop).Weight = xlThin
        rowRange.Borders(xlEdgeBottom).Weight = xlThin
        rowRange.Borders(xlEdgeRight).Weight = xlThin
        ws.Cells(rowIndex, "A").HorizontalAlignment = xlHAlignCenter
        ws.Cells(rowIndex, "A").VerticalAlignment = xlVAlignCenter
        ws.Cells(rowIndex, "A").Interior.Color = Palette(themeKey, "slot")
    Next rowIndex
End Sub

Private Function GetItemSlotsFromHelper(ByVal catalogWs As Worksheet, ByVal itemName As String, ByVal helperNameColumn As String, ByVal helperSlotsColumn As String) As Long
    Dim rowIndex As Long
    Dim currentName As String
    Dim slotValue As Variant
    Dim lastCatalogRow As Long

    If Len(Trim$(itemName)) = 0 Or StrComp(Trim$(itemName), "空", vbTextCompare) = 0 Then
        GetItemSlotsFromHelper = 0
        Exit Function
    End If

    lastCatalogRow = catalogWs.Range(helperNameColumn & catalogWs.Rows.Count).End(xlUp).Row
    For rowIndex = 5 To lastCatalogRow
        currentName = Trim$(CStr(catalogWs.Range(helperNameColumn & rowIndex).Value))
        If StrComp(currentName, Trim$(itemName), vbTextCompare) = 0 Then
            slotValue = catalogWs.Range(helperSlotsColumn & rowIndex).Value
            If IsNumeric(slotValue) Then
                GetItemSlotsFromHelper = CLng(slotValue)
            Else
                GetItemSlotsFromHelper = 1
            End If
            If GetItemSlotsFromHelper < 1 Then GetItemSlotsFromHelper = 1
            Exit Function
        End If
    Next rowIndex

    GetItemSlotsFromHelper = 1
End Function

Private Function ReadLong(ByVal cellValue As Variant, ByVal fallbackValue As Long) As Long
    If IsNumeric(cellValue) Then
        ReadLong = CLng(cellValue)
    Else
        ReadLong = fallbackValue
    End If
End Function

Private Function CurrentThemeKey() As String
    On Error Resume Next
    CurrentThemeKey = Trim$(CStr(Sheet7.Range("B2").Value))
    On Error GoTo 0
    If Len(CurrentThemeKey) = 0 Then CurrentThemeKey = "墨家"
End Function

Private Function PrimaryColor(ByVal themeKey As String) As Long
    Select Case themeKey
        Case "亓家": PrimaryColor = RGB(140, 34, 52)
        Case "林家": PrimaryColor = RGB(63, 134, 120)
        Case "易家": PrimaryColor = RGB(69, 109, 181)
        Case "巫家": PrimaryColor = RGB(182, 104, 139)
        Case "岳家": PrimaryColor = RGB(181, 78, 52)
        Case "药家": PrimaryColor = RGB(74, 132, 62)
        Case "张家": PrimaryColor = RGB(116, 82, 196)
        Case "曹家": PrimaryColor = RGB(24, 140, 148)
        Case "乐正家": PrimaryColor = RGB(186, 144, 56)
        Case "墨家": PrimaryColor = RGB(77, 96, 112)
        Case Else: PrimaryColor = RGB(86, 89, 94)
    End Select
End Function

Private Function Palette(ByVal themeKey As String, ByVal partName As String) As Long
    Select Case themeKey
        Case "亓家"
            Select Case partName
                Case "base": Palette = RGB(247, 235, 238)
                Case "panel": Palette = RGB(233, 208, 214)
                Case "table": Palette = RGB(252, 246, 247)
                Case "titleDark": Palette = RGB(92, 23, 37)
                Case "titleMid": Palette = RGB(140, 34, 52)
                Case "titleAlt": Palette = RGB(113, 45, 63)
                Case "header": Palette = RGB(220, 179, 188)
                Case "total": Palette = RGB(198, 94, 112)
                Case "note": Palette = RGB(241, 224, 228)
                Case "slot": Palette = RGB(228, 194, 201)
                Case "fill": Palette = RGB(205, 159, 169)
                Case "overflow": Palette = RGB(155, 100, 118)
                Case "textDark": Palette = RGB(60, 30, 36)
                Case "textLight": Palette = RGB(252, 247, 248)
                Case "tab1": Palette = RGB(140, 34, 52)
                Case "tab2": Palette = RGB(92, 23, 37)
            End Select
        Case "林家"
            Select Case partName
                Case "base": Palette = RGB(235, 244, 239)
                Case "panel": Palette = RGB(210, 227, 220)
                Case "table": Palette = RGB(246, 251, 248)
                Case "titleDark": Palette = RGB(42, 92, 81)
                Case "titleMid": Palette = RGB(72, 128, 113)
                Case "titleAlt": Palette = RGB(84, 112, 82)
                Case "header": Palette = RGB(172, 203, 192)
                Case "total": Palette = RGB(112, 176, 160)
                Case "note": Palette = RGB(224, 235, 229)
                Case "slot": Palette = RGB(191, 217, 206)
                Case "fill": Palette = RGB(154, 196, 181)
                Case "overflow": Palette = RGB(171, 152, 130)
                Case "textDark": Palette = RGB(38, 56, 48)
                Case "textLight": Palette = RGB(248, 250, 247)
                Case "tab1": Palette = RGB(72, 128, 113)
                Case "tab2": Palette = RGB(42, 92, 81)
            End Select
        Case "易家"
            Select Case partName
                Case "base": Palette = RGB(235, 240, 249)
                Case "panel": Palette = RGB(209, 220, 241)
                Case "table": Palette = RGB(246, 249, 254)
                Case "titleDark": Palette = RGB(49, 75, 131)
                Case "titleMid": Palette = RGB(78, 112, 176)
                Case "titleAlt": Palette = RGB(75, 95, 145)
                Case "header": Palette = RGB(178, 197, 229)
                Case "total": Palette = RGB(122, 156, 214)
                Case "note": Palette = RGB(224, 231, 244)
                Case "slot": Palette = RGB(195, 209, 236)
                Case "fill": Palette = RGB(164, 185, 224)
                Case "overflow": Palette = RGB(156, 138, 180)
                Case "textDark": Palette = RGB(35, 49, 78)
                Case "textLight": Palette = RGB(247, 249, 253)
                Case "tab1": Palette = RGB(78, 112, 176)
                Case "tab2": Palette = RGB(49, 75, 131)
            End Select
        Case "巫家"
            Select Case partName
                Case "base": Palette = RGB(249, 238, 243)
                Case "panel": Palette = RGB(236, 212, 223)
                Case "table": Palette = RGB(253, 247, 250)
                Case "titleDark": Palette = RGB(118, 63, 91)
                Case "titleMid": Palette = RGB(163, 93, 128)
                Case "titleAlt": Palette = RGB(143, 82, 111)
                Case "header": Palette = RGB(220, 180, 199)
                Case "total": Palette = RGB(212, 141, 176)
                Case "note": Palette = RGB(242, 228, 235)
                Case "slot": Palette = RGB(230, 197, 213)
                Case "fill": Palette = RGB(206, 159, 184)
                Case "overflow": Palette = RGB(180, 128, 156)
                Case "textDark": Palette = RGB(67, 37, 53)
                Case "textLight": Palette = RGB(252, 247, 250)
                Case "tab1": Palette = RGB(163, 93, 128)
                Case "tab2": Palette = RGB(118, 63, 91)
            End Select
        Case "岳家"
            Select Case partName
                Case "base": Palette = RGB(250, 240, 233)
                Case "panel": Palette = RGB(239, 217, 205)
                Case "table": Palette = RGB(254, 247, 243)
                Case "titleDark": Palette = RGB(124, 52, 32)
                Case "titleMid": Palette = RGB(181, 78, 52)
                Case "titleAlt": Palette = RGB(156, 91, 45)
                Case "header": Palette = RGB(226, 185, 166)
                Case "total": Palette = RGB(220, 122, 86)
                Case "note": Palette = RGB(244, 228, 219)
                Case "slot": Palette = RGB(236, 201, 184)
                Case "fill": Palette = RGB(210, 155, 128)
                Case "overflow": Palette = RGB(174, 116, 97)
                Case "textDark": Palette = RGB(74, 38, 28)
                Case "textLight": Palette = RGB(255, 249, 246)
                Case "tab1": Palette = RGB(181, 78, 52)
                Case "tab2": Palette = RGB(124, 52, 32)
            End Select
        Case "药家"
            Select Case partName
                Case "base": Palette = RGB(236, 246, 232)
                Case "panel": Palette = RGB(212, 230, 206)
                Case "table": Palette = RGB(247, 252, 244)
                Case "titleDark": Palette = RGB(53, 96, 42)
                Case "titleMid": Palette = RGB(81, 129, 66)
                Case "titleAlt": Palette = RGB(93, 118, 57)
                Case "header": Palette = RGB(179, 208, 166)
                Case "total": Palette = RGB(119, 173, 97)
                Case "note": Palette = RGB(225, 237, 219)
                Case "slot": Palette = RGB(196, 221, 187)
                Case "fill": Palette = RGB(161, 201, 145)
                Case "overflow": Palette = RGB(152, 139, 106)
                Case "textDark": Palette = RGB(38, 63, 32)
                Case "textLight": Palette = RGB(248, 252, 247)
                Case "tab1": Palette = RGB(81, 129, 66)
                Case "tab2": Palette = RGB(53, 96, 42)
            End Select
        Case "张家"
            Select Case partName
                Case "base": Palette = RGB(242, 237, 249)
                Case "panel": Palette = RGB(227, 216, 243)
                Case "table": Palette = RGB(250, 247, 253)
                Case "titleDark": Palette = RGB(75, 46, 133)
                Case "titleMid": Palette = RGB(111, 74, 179)
                Case "titleAlt": Palette = RGB(90, 70, 140)
                Case "header": Palette = RGB(204, 187, 231)
                Case "total": Palette = RGB(154, 119, 212)
                Case "note": Palette = RGB(236, 229, 245)
                Case "slot": Palette = RGB(220, 208, 238)
                Case "fill": Palette = RGB(189, 170, 223)
                Case "overflow": Palette = RGB(150, 125, 191)
                Case "textDark": Palette = RGB(45, 34, 70)
                Case "textLight": Palette = RGB(250, 248, 253)
                Case "tab1": Palette = RGB(111, 74, 179)
                Case "tab2": Palette = RGB(75, 46, 133)
            End Select
        Case "曹家"
            Select Case partName
                Case "base": Palette = RGB(233, 246, 247)
                Case "panel": Palette = RGB(205, 229, 231)
                Case "table": Palette = RGB(245, 252, 252)
                Case "titleDark": Palette = RGB(20, 86, 92)
                Case "titleMid": Palette = RGB(24, 140, 148)
                Case "titleAlt": Palette = RGB(44, 112, 118)
                Case "header": Palette = RGB(170, 210, 214)
                Case "total": Palette = RGB(96, 184, 191)
                Case "note": Palette = RGB(223, 238, 239)
                Case "slot": Palette = RGB(190, 219, 221)
                Case "fill": Palette = RGB(148, 198, 201)
                Case "overflow": Palette = RGB(130, 161, 164)
                Case "textDark": Palette = RGB(27, 60, 62)
                Case "textLight": Palette = RGB(248, 252, 252)
                Case "tab1": Palette = RGB(24, 140, 148)
                Case "tab2": Palette = RGB(20, 86, 92)
            End Select
        Case "乐正家"
            Select Case partName
                Case "base": Palette = RGB(249, 244, 232)
                Case "panel": Palette = RGB(239, 228, 204)
                Case "table": Palette = RGB(253, 249, 242)
                Case "titleDark": Palette = RGB(112, 84, 29)
                Case "titleMid": Palette = RGB(186, 144, 56)
                Case "titleAlt": Palette = RGB(154, 121, 52)
                Case "header": Palette = RGB(226, 206, 166)
                Case "total": Palette = RGB(214, 178, 96)
                Case "note": Palette = RGB(245, 236, 219)
                Case "slot": Palette = RGB(236, 221, 184)
                Case "fill": Palette = RGB(214, 193, 136)
                Case "overflow": Palette = RGB(173, 145, 93)
                Case "textDark": Palette = RGB(72, 56, 25)
                Case "textLight": Palette = RGB(255, 251, 245)
                Case "tab1": Palette = RGB(186, 144, 56)
                Case "tab2": Palette = RGB(112, 84, 29)
            End Select
        Case "墨家"
            Select Case partName
                Case "base": Palette = RGB(232, 236, 240)
                Case "panel": Palette = RGB(207, 214, 221)
                Case "table": Palette = RGB(244, 247, 249)
                Case "titleDark": Palette = RGB(43, 55, 66)
                Case "titleMid": Palette = RGB(77, 96, 112)
                Case "titleAlt": Palette = RGB(73, 86, 98)
                Case "header": Palette = RGB(172, 183, 194)
                Case "total": Palette = RGB(125, 142, 158)
                Case "note": Palette = RGB(223, 229, 234)
                Case "slot": Palette = RGB(194, 204, 212)
                Case "fill": Palette = RGB(156, 169, 180)
                Case "overflow": Palette = RGB(112, 121, 132)
                Case "textDark": Palette = RGB(29, 36, 43)
                Case "textLight": Palette = RGB(248, 249, 250)
                Case "tab1": Palette = RGB(77, 96, 112)
                Case "tab2": Palette = RGB(43, 55, 66)
            End Select
        Case Else
            Select Case partName
                Case "base": Palette = RGB(236, 237, 239)
                Case "panel": Palette = RGB(212, 215, 219)
                Case "table": Palette = RGB(246, 247, 248)
                Case "titleDark": Palette = RGB(60, 64, 70)
                Case "titleMid": Palette = RGB(89, 95, 105)
                Case "titleAlt": Palette = RGB(111, 93, 71)
                Case "header": Palette = RGB(182, 188, 195)
                Case "total": Palette = RGB(154, 162, 173)
                Case "note": Palette = RGB(228, 230, 233)
                Case "slot": Palette = RGB(197, 201, 207)
                Case "fill": Palette = RGB(164, 171, 180)
                Case "overflow": Palette = RGB(154, 133, 118)
                Case "textDark": Palette = RGB(42, 45, 50)
                Case "textLight": Palette = RGB(248, 248, 248)
                Case "tab1": Palette = RGB(89, 95, 105)
                Case "tab2": Palette = RGB(60, 64, 70)
            End Select
    End Select
End Function

Private Sub ApplyThemePersonalInfo(ByVal ws As Worksheet, ByVal themeKey As String)
    ws.Cells.Font.Name = "KaiTi"
    ws.Cells.Font.Size = 11.5
    ws.Range("A1:V55").Interior.Color = Palette(themeKey, "base")

    PaintBlock ws.Range("A1:D1"), Palette(themeKey, "titleDark"), Palette(themeKey, "textLight"), 13.5, True, True
    PaintBlock ws.Range("F1:K1"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 13.5, True, True
    PaintBlock ws.Range("M1:O1"), Palette(themeKey, "titleAlt"), Palette(themeKey, "textLight"), 13.5, True, True
    PaintBlock ws.Range("Q1:V1"), Palette(themeKey, "titleDark"), Palette(themeKey, "textLight"), 13.5, True, True
    PaintBlock ws.Range("A10:P10"), Palette(themeKey, "titleAlt"), Palette(themeKey, "textLight"), 12, True, True
    PaintBlock ws.Range("Q10:V10"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 12, True, True
    PaintBlock ws.Range("A23:N23"), Palette(themeKey, "titleDark"), Palette(themeKey, "textLight"), 12.5, True, True
    PaintBlock ws.Range("P23:V23"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 12.5, True, True
    PaintBlock ws.Range("P30:V30"), Palette(themeKey, "titleAlt"), Palette(themeKey, "textLight"), 12, True, True
    PaintBlock ws.Range("A41:V41"), Palette(themeKey, "titleDark"), Palette(themeKey, "textLight"), 12.5, True, True

    PaintBlock ws.Range("A2:D9"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11.5, False, False
    PaintBlock ws.Range("F2:K9"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11.5, False, False
    PaintBlock ws.Range("M2:O9"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11.5, False, False
    PaintBlock ws.Range("Q2:V9"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11.5, False, False
    PaintBlock ws.Range("Q11:V21"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11.5, False, False

    PaintBlock ws.Range("A11:D11"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("F11:J11"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("L11:O11"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("A12:D20"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11.5, False, False
    PaintBlock ws.Range("F12:J20"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11.5, False, False
    PaintBlock ws.Range("L12:O20"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11.5, False, False
    PaintBlock ws.Range("A22:V22"), Palette(themeKey, "panel"), Palette(themeKey, "textDark"), 10.5, False, False
    PaintBlock ws.Range("A40:V40"), Palette(themeKey, "panel"), Palette(themeKey, "textDark"), 10.5, False, False

    PaintBlock ws.Range("A24:G24"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("H24:N24"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("A25:B39"), Palette(themeKey, "panel"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("C25:F39"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("G25:G39"), Palette(themeKey, "total"), Palette(themeKey, "textDark"), 11, True, True
    PaintBlock ws.Range("H25:I39"), Palette(themeKey, "panel"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("J25:M39"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("N25:N39"), Palette(themeKey, "total"), Palette(themeKey, "textDark"), 11, True, True

    PaintBlock ws.Range("P24:R24"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("S24:V24"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("P25:R29"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("S25:V29"), Palette(themeKey, "note"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("P31:V39"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False

    PaintBlock ws.Range("B42:K55"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("M42:V55"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False

    ws.Range("E1:E55").Interior.Color = Palette(themeKey, "base")
    ws.Range("P1:P22").Interior.Color = Palette(themeKey, "base")
    ws.Range("K11:K21").Interior.Color = Palette(themeKey, "base")
    ws.Range("L42:L55").Interior.Color = Palette(themeKey, "base")

    ApplyGrid ws.Range("A2:D9")
    ApplyGrid ws.Range("F2:K9")
    ApplyGrid ws.Range("M2:O9")
    ApplyGrid ws.Range("Q2:V21")
    ApplyGrid ws.Range("A11:D20")
    ApplyGrid ws.Range("F11:J20")
    ApplyGrid ws.Range("L11:O20")
    ApplyGrid ws.Range("A24:N39")
    ApplyGrid ws.Range("P24:V29")
    ApplyGrid ws.Range("P30:V39")
    ApplyGrid ws.Range("B42:K55")
    ApplyGrid ws.Range("M42:V55")
    ws.Tab.Color = Palette(themeKey, "tab1")
End Sub

Private Sub ApplyThemeBookAndTraits(ByVal ws As Worksheet, ByVal themeKey As String)
    Dim rowIndex As Long

    ws.Cells.Font.Name = "KaiTi"
    ws.Cells.Font.Size = 11
    ws.Range("A1:U35").Interior.Color = Palette(themeKey, "base")

    PaintBlock ws.Range("A1:B17"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("C1:J17"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("K1:K35"), Palette(themeKey, "base"), Palette(themeKey, "textDark"), 10.5, False, False
    PaintBlock ws.Range("L1:L17"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("M1:U17"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("A18:U18"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 12, True, True

    PaintBlock ws.Range("A19:A35"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("B19:H19"), Palette(themeKey, "titleDark"), Palette(themeKey, "textLight"), 11.5, True, True
    PaintBlock ws.Range("I19:J19"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("L19:L35"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("M19:S19"), Palette(themeKey, "titleDark"), Palette(themeKey, "textLight"), 11.5, True, True
    PaintBlock ws.Range("T19:U19"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True

    For rowIndex = 20 To 35
        If rowIndex Mod 2 = 0 Then
            PaintBlock ws.Range("B" & rowIndex & ":H" & rowIndex), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
            PaintBlock ws.Range("I" & rowIndex & ":J" & rowIndex), Palette(themeKey, "note"), Palette(themeKey, "textDark"), 11, True, True
            PaintBlock ws.Range("M" & rowIndex & ":S" & rowIndex), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
            PaintBlock ws.Range("T" & rowIndex & ":U" & rowIndex), Palette(themeKey, "note"), Palette(themeKey, "textDark"), 11, True, True
        Else
            PaintBlock ws.Range("B" & rowIndex & ":H" & rowIndex), Palette(themeKey, "panel"), Palette(themeKey, "textDark"), 11, False, False
            PaintBlock ws.Range("I" & rowIndex & ":J" & rowIndex), Palette(themeKey, "total"), Palette(themeKey, "textDark"), 11, True, True
            PaintBlock ws.Range("M" & rowIndex & ":S" & rowIndex), Palette(themeKey, "panel"), Palette(themeKey, "textDark"), 11, False, False
            PaintBlock ws.Range("T" & rowIndex & ":U" & rowIndex), Palette(themeKey, "total"), Palette(themeKey, "textDark"), 11, True, True
        End If
    Next rowIndex

    ApplyGrid ws.Range("A1:J17")
    ApplyGrid ws.Range("L1:U17")
    ApplyGrid ws.Range("A18:U18")
    ApplyGrid ws.Range("A19:J35")
    ApplyGrid ws.Range("L19:U35")
    ws.Tab.Color = Palette(themeKey, "tab2")
End Sub

Private Sub ApplyThemeMojia(ByVal ws As Worksheet, ByVal themeKey As String)
    ws.Cells.Font.Name = "KaiTi"
    ws.Cells.Font.Size = 11
    ws.Range("A1:W320").Interior.Color = Palette(themeKey, "base")

    PaintBlock ws.Range("A1:E1"), Palette(themeKey, "titleDark"), Palette(themeKey, "textLight"), 13.5, True, True
    PaintBlock ws.Range("H1:I1"), Palette(themeKey, "titleAlt"), Palette(themeKey, "textLight"), 12, True, True
    PaintBlock ws.Range("K1:M1"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 12, True, True
    PaintBlock ws.Range("O1:W1"), Palette(themeKey, "titleDark"), Palette(themeKey, "textLight"), 12.5, True, True

    PaintBlock ws.Range("A3:E3"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 12, True, True
    PaintBlock ws.Range("A4:E4"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True

    PaintBlock ws.Range("H3:I8"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, True
    PaintBlock ws.Range("H3:H8"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11, True, True
    PaintBlock ws.Range("H9:H16"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11, True, True
    PaintBlock ws.Range("I9:I16"), Palette(themeKey, "note"), Palette(themeKey, "textDark"), 11, False, False

    PaintBlock ws.Range("K3:M3"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("K4:M13"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, True

    PaintBlock ws.Range("O3:R3"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 11.5, True, True
    PaintBlock ws.Range("T3:W3"), Palette(themeKey, "titleAlt"), Palette(themeKey, "textLight"), 11.5, True, True
    PaintBlock ws.Range("O4:O13"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11, True, True
    PaintBlock ws.Range("P4:P13"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, True
    PaintBlock ws.Range("Q4:Q12"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11, True, True
    PaintBlock ws.Range("R4:R12"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, True
    PaintBlock ws.Range("Q13:R13"), Palette(themeKey, "note"), Palette(themeKey, "textDark"), 10.5, False, False

    PaintBlock ws.Range("T4:W4"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11, True, True
    PaintBlock ws.Range("T5:W14"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, True

    PaintBlock ws.Range("O15:R15"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 11.5, True, True
    PaintBlock ws.Range("O16:R16"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11, True, True
    PaintBlock ws.Range("O17:R19"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, True

    PaintBlock ws.Range("O21:W21"), Palette(themeKey, "titleAlt"), Palette(themeKey, "textLight"), 11.5, True, True
    PaintBlock ws.Range("O22:P28"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11, True, True
    PaintBlock ws.Range("Q22:W28"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False

    ws.Rows("5:304").RowHeight = 24
    ws.Rows("1:1").RowHeight = 28
    ws.Rows("21:28").RowHeight = 28
    ws.Range("A5:A304").HorizontalAlignment = xlHAlignCenter
    ws.Range("A5:A304").VerticalAlignment = xlVAlignCenter

    ApplyGrid ws.Range("A4:E19")
    ApplyGrid ws.Range("H3:I8")
    ApplyGrid ws.Range("H9:I16")
    ApplyGrid ws.Range("K3:M13")
    ApplyGrid ws.Range("O3:R13")
    ApplyGrid ws.Range("T3:W14")
    ApplyGrid ws.Range("O15:R19")
    ApplyGrid ws.Range("O21:W28")
    ws.Tab.Color = Palette(themeKey, "tab2")
End Sub

Private Sub ApplyThemeCatalog(ByVal ws As Worksheet, ByVal themeKey As String)
    ws.Cells.Font.Name = "KaiTi"
    ws.Cells.Font.Size = 11
    ws.Range("A1:AC30").Interior.Color = Palette(themeKey, "base")

    PaintBlock ws.Range("A1:AC1"), Palette(themeKey, "titleDark"), Palette(themeKey, "textLight"), 13, True, True

    PaintBlock ws.Range("A3:D3"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 11.5, True, True
    PaintBlock ws.Range("F3:I3"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 11.5, True, True
    PaintBlock ws.Range("K3:N3"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 11.5, True, True
    PaintBlock ws.Range("P3:S3"), Palette(themeKey, "titleAlt"), Palette(themeKey, "textLight"), 11.5, True, True
    PaintBlock ws.Range("U3:X3"), Palette(themeKey, "titleAlt"), Palette(themeKey, "textLight"), 11.5, True, True
    PaintBlock ws.Range("Z3:AC3"), Palette(themeKey, "titleAlt"), Palette(themeKey, "textLight"), 11.5, True, True

    PaintBlock ws.Range("A4:D4"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("F4:I4"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("K4:N4"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("P4:S4"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("U4:X4"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("Z4:AC4"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True

    PaintBlock ws.Range("A5:D14"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("F5:I19"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("K5:N16"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("P5:S17"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("U5:X9"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("Z5:AC8"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False

    PaintBlock ws.Range("A22:AC22"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 11.5, True, True
    PaintBlock ws.Range("A23:AC26"), Palette(themeKey, "note"), Palette(themeKey, "textDark"), 11, False, False

    ws.Rows("1:30").RowHeight = 24
    ws.Rows("1:1").RowHeight = 28
    ws.Rows("22:26").RowHeight = 30
    ws.Tab.Color = Palette(themeKey, "tab2")

    ApplyGrid ws.Range("A3:D14")
    ApplyGrid ws.Range("F3:I19")
    ApplyGrid ws.Range("K3:N16")
    ApplyGrid ws.Range("P3:S17")
    ApplyGrid ws.Range("U3:X9")
    ApplyGrid ws.Range("Z3:AC8")
    ApplyGrid ws.Range("A22:AC26")
End Sub

Private Sub ApplyThemeFastCopy(ByVal ws As Worksheet, ByVal themeKey As String)
    ws.Cells.Font.Name = "KaiTi"
    ws.Cells.Font.Size = 11
    ws.Range("B1:G11").Interior.Color = Palette(themeKey, "base")

    PaintBlock ws.Range("B1:G2"), Palette(themeKey, "titleDark"), Palette(themeKey, "textLight"), 12.5, True, True
    PaintBlock ws.Range("B3:B11"), Palette(themeKey, "panel"), Palette(themeKey, "textDark"), 11, True, False
    PaintBlock ws.Range("D3:D11"), Palette(themeKey, "panel"), Palette(themeKey, "textDark"), 11, True, False
    PaintBlock ws.Range("F3:F11"), Palette(themeKey, "panel"), Palette(themeKey, "textDark"), 11, True, False
    PaintBlock ws.Range("C3:C11"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("E3:E11"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("G3:G11"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    ApplyGrid ws.Range("B3:G11")
    ws.Tab.Color = Palette(themeKey, "tab1")
End Sub

Private Sub ApplyThemeSettings(ByVal ws As Worksheet, ByVal themeKey As String)
    Dim rowIndex As Long
    Dim rowTheme As String

    ws.Cells.Font.Name = "KaiTi"
    ws.Cells.Font.Size = 11
    ws.Range("A1:D17").Interior.Color = Palette(themeKey, "base")

    PaintBlock ws.Range("A1:D1"), Palette(themeKey, "titleDark"), Palette(themeKey, "textLight"), 13, True, True
    PaintBlock ws.Range("A2:B2"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("A4:D4"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 11.5, True, True
    PaintBlock ws.Range("A5:D5"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("A6:D15"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("A16:D17"), Palette(themeKey, "note"), Palette(themeKey, "textDark"), 11, False, False

    For rowIndex = 6 To 15
        rowTheme = CStr(ws.Cells(rowIndex, "A").Value)
        ws.Cells(rowIndex, "D").Interior.Color = PrimaryColor(rowTheme)
        ws.Cells(rowIndex, "D").Font.Color = Palette(themeKey, "textLight")
        ws.Cells(rowIndex, "D").HorizontalAlignment = xlHAlignCenter
        ws.Cells(rowIndex, "D").VerticalAlignment = xlVAlignCenter
    Next rowIndex

    ws.Rows("4:17").RowHeight = 28
    ApplyGrid ws.Range("A2:B2")
    ApplyGrid ws.Range("A5:D15")
    ApplyGrid ws.Range("A16:D17")
    ws.Tab.Color = Palette(themeKey, "tab2")
End Sub

Private Sub PaintBlock(ByVal rng As Range, ByVal fillColor As Long, ByVal fontColor As Long, ByVal fontSize As Double, Optional ByVal makeBold As Boolean = False, Optional ByVal centerText As Boolean = False)
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

Private Sub ApplyGrid(ByVal rng As Range)
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
