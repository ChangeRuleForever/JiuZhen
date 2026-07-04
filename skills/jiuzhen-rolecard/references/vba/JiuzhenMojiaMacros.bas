Attribute VB_Name = "JiuzhenMojiaMacros"
Option Explicit

Private Const MOJIA_SHEET_INDEX As Long = 3
Private Const KUI_SHEET_INDEX As Long = 4
Private Const CATALOG_SHEET_INDEX As Long = 5
Private Const FASTCOPY_SHEET_INDEX As Long = 6
Private Const THEME_SHEET_INDEX As Long = 7

Private Const FIRST_SLOT_ROW As Long = 5
Private Const BASE_SLOT_COUNT As Long = 15
Private Const MAX_SLOT_COUNT As Long = 300
Private Const LAST_SLOT_ROW As Long = FIRST_SLOT_ROW + MAX_SLOT_COUNT - 1
Private Const SLOT_TOTAL_CELL As String = "I5"
Private Const MOJIA_HELPER_COL As String = "F"

Private Const KUI_FIRST_SLOT_ROW As Long = 31
Private Const KUI_BASE_SLOT_COUNT As Long = 5
Private Const KUI_MAX_SLOT_COUNT As Long = 300
Private Const KUI_LAST_SLOT_ROW As Long = KUI_FIRST_SLOT_ROW + KUI_MAX_SLOT_COUNT - 1
Private Const KUI_TOTAL_CELL As String = "D12"
Private Const KUI_HELPER_COL As String = "I"

Private Const xlValidateList As Long = 3
Private Const xlValidAlertStop As Long = 1
Private Const xlBetween As Long = 1
Private Const xlHAlignCenter As Long = -4108
Private Const xlVAlignCenter As Long = -4108
Private Const xlValues As Long = -4163
Private Const xlWhole As Long = 1
Private Const xlEdgeLeft As Long = 7
Private Const xlEdgeTop As Long = 8
Private Const xlEdgeBottom As Long = 9
Private Const xlEdgeRight As Long = 10
Private Const xlInsideVertical As Long = 11
Private Const xlInsideHorizontal As Long = 12
Private Const xlContinuous As Long = 1
Private Const xlThin As Long = 2

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

    ApplyThemePersonalInfo ThisWorkbook.Worksheets(1), themeKey
    ApplyThemeBookAndTraits ThisWorkbook.Worksheets(2), themeKey
    ApplyThemeMojia ThisWorkbook.Worksheets(MOJIA_SHEET_INDEX), themeKey
    ApplyThemeKui ThisWorkbook.Worksheets(KUI_SHEET_INDEX), themeKey
    ApplyThemeCatalog ThisWorkbook.Worksheets(CATALOG_SHEET_INDEX), themeKey
    ApplyThemeFastCopy ThisWorkbook.Worksheets(FASTCOPY_SHEET_INDEX), themeKey
    ApplyThemeSettings ThisWorkbook.Worksheets(THEME_SHEET_INDEX), themeKey

    RefreshMojiaInventory
    RefreshKuiInventory

CleanExit:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

CleanFail:
    MsgBox "应用主题失败：" & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Public Sub RefreshMojiaInventory()
    On Error GoTo CleanFail

    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim ws As Worksheet
    Dim catalogWs As Worksheet
    Dim itemNames As Collection
    Dim itemRows As Collection
    Dim totalSlots As Long
    Dim visibleSlots As Long
    Dim rejectedItems As String

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = ThisWorkbook.Worksheets(MOJIA_SHEET_INDEX)
    Set catalogWs = ThisWorkbook.Worksheets(CATALOG_SHEET_INDEX)
    Set itemNames = New Collection
    Set itemRows = New Collection

    CollectItemEntries ws, FIRST_SLOT_ROW, LAST_SLOT_ROW, itemNames, itemRows

    totalSlots = ReadLong(ws.Range(SLOT_TOTAL_CELL).Value, BASE_SLOT_COUNT)
    If totalSlots < BASE_SLOT_COUNT Then totalSlots = BASE_SLOT_COUNT
    If totalSlots > MAX_SLOT_COUNT Then totalSlots = MAX_SLOT_COUNT
    visibleSlots = totalSlots

    ResetSlotArea ws, FIRST_SLOT_ROW, LAST_SLOT_ROW, MOJIA_HELPER_COL
    ApplyInventoryValidation ws, catalogWs, FIRST_SLOT_ROW, LAST_SLOT_ROW, "A"
    rejectedItems = RenderItems(ws, catalogWs, itemNames, itemRows, totalSlots, visibleSlots, FIRST_SLOT_ROW, LAST_SLOT_ROW, "A", MOJIA_HELPER_COL)
    ApplySlotGrid ws, visibleSlots, FIRST_SLOT_ROW
    ApplyRowVisibility ws, visibleSlots, FIRST_SLOT_ROW, LAST_SLOT_ROW

    If Len(rejectedItems) > 0 Then
        MsgBox "本人剩余槽位不足，以下物品未携带：" & vbCrLf & rejectedItems, vbExclamation
    End If

CleanExit:
    Application.CutCopyMode = False
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

CleanFail:
    MsgBox "刷新本人背包失败：" & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Public Sub RefreshKuiInventory()
    On Error GoTo CleanFail

    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim ws As Worksheet
    Dim catalogWs As Worksheet
    Dim itemNames As Collection
    Dim itemRows As Collection
    Dim totalSlots As Long
    Dim visibleSlots As Long
    Dim rejectedItems As String

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = ThisWorkbook.Worksheets(KUI_SHEET_INDEX)
    Set catalogWs = ThisWorkbook.Worksheets(CATALOG_SHEET_INDEX)
    Set itemNames = New Collection
    Set itemRows = New Collection

    EnsureKuiPluginValidation ws, catalogWs
    EnsureKuiPluginEffect ws
    CollectItemEntries ws, KUI_FIRST_SLOT_ROW, KUI_LAST_SLOT_ROW, itemNames, itemRows

    totalSlots = ReadLong(ws.Range(KUI_TOTAL_CELL).Value, KUI_BASE_SLOT_COUNT)
    If totalSlots < KUI_BASE_SLOT_COUNT Then totalSlots = KUI_BASE_SLOT_COUNT
    If totalSlots > KUI_MAX_SLOT_COUNT Then totalSlots = KUI_MAX_SLOT_COUNT
    visibleSlots = totalSlots

    ResetSlotArea ws, KUI_FIRST_SLOT_ROW, KUI_LAST_SLOT_ROW, KUI_HELPER_COL
    ApplyInventoryValidation ws, catalogWs, KUI_FIRST_SLOT_ROW, KUI_LAST_SLOT_ROW, "F"
    rejectedItems = RenderItems(ws, catalogWs, itemNames, itemRows, totalSlots, visibleSlots, KUI_FIRST_SLOT_ROW, KUI_LAST_SLOT_ROW, "F", KUI_HELPER_COL)
    ApplySlotGrid ws, visibleSlots, KUI_FIRST_SLOT_ROW
    ApplyRowVisibility ws, visibleSlots, KUI_FIRST_SLOT_ROW, KUI_LAST_SLOT_ROW

    If Len(rejectedItems) > 0 Then
        MsgBox "魁人剩余槽位不足，以下物品未携带：" & vbCrLf & rejectedItems, vbExclamation
    End If

CleanExit:
    Application.CutCopyMode = False
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

CleanFail:
    MsgBox "刷新魁人面板失败：" & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Sub CollectItemEntries(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByRef itemNames As Collection, ByRef itemRows As Collection)
    Dim rowIndex As Long
    Dim cellText As String

    For rowIndex = firstRow To lastRow
        cellText = Trim$(CStr(ws.Cells(rowIndex, "B").Value))
        If Len(cellText) > 0 Then
            If StrComp(cellText, "空", vbTextCompare) <> 0 Then
                If ws.Cells(rowIndex, "B").MergeCells Then
                    If ws.Cells(rowIndex, "B").MergeArea.Row = rowIndex Then
                        itemNames.Add cellText
                        itemRows.Add rowIndex
                    End If
                Else
                    itemNames.Add cellText
                    itemRows.Add rowIndex
                End If
            End If
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

Private Function RenderItems(ByVal ws As Worksheet, ByVal catalogWs As Worksheet, ByVal itemNames As Collection, ByVal itemRows As Collection, ByVal totalSlots As Long, ByVal visibleSlots As Long, ByVal firstRow As Long, ByVal lastRow As Long, ByVal catalogItemColumn As String, ByVal helperColumn As String) As String
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

    For itemIndex = 1 To itemNames.Count
        itemName = itemNames(itemIndex)
        slotSize = GetItemSlots(catalogWs, CStr(itemName), catalogItemColumn)
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
            ws.Range(helperColumn & currentRow & ":" & helperColumn & endRow).Value = 1

            currentRow = endRow + 1
            usedSlots = usedSlots + slotSize
        End If
    Next itemIndex

    For rowIndex = firstRow To firstRow + visibleSlots - 1
        If rowIndex > lastRow Then Exit For
        If Not ws.Cells(rowIndex, "B").MergeCells Then
            ws.Range("B" & rowIndex & ":E" & rowIndex).Merge
        End If
    Next rowIndex

    RenderItems = rejected
End Function

Private Sub ApplyInventoryValidation(ByVal ws As Worksheet, ByVal catalogWs As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal catalogItemColumn As String)
    Dim targetRange As Range
    Dim formulaText As String
    Dim sourceCol As String
    Dim helperCol As String

    Select Case catalogItemColumn
        Case "F"
            sourceCol = "F"
            helperCol = "V"
        Case Else
            sourceCol = "A"
            helperCol = "U"
    End Select

    formulaText = "=OFFSET('" & catalogWs.Name & "'!$" & helperCol & "$5,0,0,COUNTA('" & catalogWs.Name & "'!$" & sourceCol & "$5:$" & sourceCol & "$300)+1,1)"
    Set targetRange = ws.Range("B" & firstRow & ":B" & lastRow)

    targetRange.Validation.Delete
    targetRange.Validation.Add xlValidateList, xlValidAlertStop, xlBetween, formulaText
    targetRange.Validation.IgnoreBlank = True
    targetRange.Validation.InCellDropdown = True
    targetRange.Validation.ErrorTitle = "槽位不足"
    targetRange.Validation.ErrorMessage = "如选择的物品超出剩余槽位，宏会自动取消本次携带。"
End Sub

Private Sub EnsureKuiPluginValidation(ByVal ws As Worksheet, ByVal catalogWs As Worksheet)
    Dim targetRange As Range
    Dim formulaText As String

    formulaText = "=OFFSET('" & catalogWs.Name & "'!$W$5,0,0,COUNTA('" & catalogWs.Name & "'!$K$5:$K$300)+1,1)"
    Set targetRange = ws.Range("G4:G13")

    targetRange.Validation.Delete
    targetRange.Validation.Add xlValidateList, xlValidAlertStop, xlBetween, formulaText
    targetRange.Validation.IgnoreBlank = True
    targetRange.Validation.InCellDropdown = True
End Sub

Private Sub EnsureKuiPluginEffect(ByVal ws As Worksheet)
    ws.Range("H4").Formula = "=IF(OR(G4="""",G4=""空""),"""",IFERROR(VLOOKUP(G4,'物品库'!$K$5:$M$300,3,FALSE),""""))"
    ws.Range("H4:H13").FillDown
End Sub

Private Sub ApplyRowVisibility(ByVal ws As Worksheet, ByVal visibleSlots As Long, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim visibleEndRow As Long

    visibleEndRow = firstRow + visibleSlots - 1
    If visibleEndRow > lastRow Then visibleEndRow = lastRow

    ws.Rows(firstRow & ":" & lastRow).Hidden = False
    If visibleEndRow < lastRow Then
        ws.Rows(visibleEndRow + 1 & ":" & lastRow).Hidden = True
    End If
End Sub

Private Sub ApplySlotGrid(ByVal ws As Worksheet, ByVal visibleSlots As Long, ByVal firstRow As Long)
    Dim rowIndex As Long
    Dim rowRange As Range
    Dim themeKey As String

    themeKey = CurrentThemeKey()

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

Private Function GetItemSlots(ByVal catalogWs As Worksheet, ByVal itemName As String, ByVal itemColumn As String) As Long
    Dim foundCell As Range
    Dim slotValue As Variant
    Dim lastCatalogRow As Long

    If Len(Trim$(itemName)) = 0 Or StrComp(Trim$(itemName), "空", vbTextCompare) = 0 Then
        GetItemSlots = 0
        Exit Function
    End If

    lastCatalogRow = catalogWs.Cells(catalogWs.Rows.Count, itemColumn).End(xlUp).Row
    Set foundCell = catalogWs.Range(itemColumn & "5:" & itemColumn & lastCatalogRow).Find(What:=itemName, LookIn:=xlValues, LookAt:=xlWhole)
    If foundCell Is Nothing Then
        GetItemSlots = 1
        Exit Function
    End If

    slotValue = foundCell.Offset(0, 1).Value
    If IsNumeric(slotValue) Then
        GetItemSlots = CLng(slotValue)
    Else
        GetItemSlots = 1
    End If

    If GetItemSlots < 1 Then GetItemSlots = 1
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
    CurrentThemeKey = Trim$(CStr(ThisWorkbook.Worksheets(THEME_SHEET_INDEX).Range("B2").Value))
    On Error GoTo 0
    If Len(CurrentThemeKey) = 0 Then CurrentThemeKey = "墨家"
End Function

Private Function PrimaryColor(ByVal themeKey As String) As Long
    Select Case themeKey
        Case "张家": PrimaryColor = RGB(214, 177, 51)
        Case "岳家": PrimaryColor = RGB(153, 103, 51)
        Case "林家": PrimaryColor = RGB(63, 134, 120)
        Case "易家": PrimaryColor = RGB(69, 109, 181)
        Case "曹家": PrimaryColor = RGB(194, 154, 63)
        Case "药家": PrimaryColor = RGB(74, 132, 62)
        Case "巫家": PrimaryColor = RGB(182, 104, 139)
        Case "亓家": PrimaryColor = RGB(44, 44, 44)
        Case Else: PrimaryColor = RGB(86, 89, 94)
    End Select
End Function

Private Function Palette(ByVal themeKey As String, ByVal partName As String) As Long
    Select Case themeKey
        Case "张家"
            Select Case partName
                Case "base": Palette = RGB(248, 244, 225)
                Case "panel": Palette = RGB(236, 226, 190)
                Case "table": Palette = RGB(255, 251, 236)
                Case "titleDark": Palette = RGB(128, 90, 21)
                Case "titleMid": Palette = RGB(174, 132, 38)
                Case "titleAlt": Palette = RGB(152, 115, 49)
                Case "header": Palette = RGB(225, 205, 138)
                Case "total": Palette = RGB(238, 190, 94)
                Case "note": Palette = RGB(243, 235, 206)
                Case "slot": Palette = RGB(245, 225, 167)
                Case "fill": Palette = RGB(229, 204, 132)
                Case "overflow": Palette = RGB(201, 153, 108)
                Case "textDark": Palette = RGB(70, 49, 21)
                Case "textLight": Palette = RGB(255, 250, 240)
                Case "tab1": Palette = RGB(174, 132, 38)
                Case "tab2": Palette = RGB(128, 90, 21)
            End Select
        Case "岳家"
            Select Case partName
                Case "base": Palette = RGB(246, 238, 228)
                Case "panel": Palette = RGB(228, 210, 188)
                Case "table": Palette = RGB(252, 246, 240)
                Case "titleDark": Palette = RGB(108, 63, 27)
                Case "titleMid": Palette = RGB(145, 92, 42)
                Case "titleAlt": Palette = RGB(128, 79, 36)
                Case "header": Palette = RGB(205, 172, 137)
                Case "total": Palette = RGB(218, 154, 90)
                Case "note": Palette = RGB(237, 223, 208)
                Case "slot": Palette = RGB(223, 191, 158)
                Case "fill": Palette = RGB(201, 160, 114)
                Case "overflow": Palette = RGB(180, 127, 102)
                Case "textDark": Palette = RGB(61, 39, 20)
                Case "textLight": Palette = RGB(251, 246, 241)
                Case "tab1": Palette = RGB(145, 92, 42)
                Case "tab2": Palette = RGB(108, 63, 27)
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
        Case "曹家"
            Select Case partName
                Case "base": Palette = RGB(250, 244, 230)
                Case "panel": Palette = RGB(237, 224, 188)
                Case "table": Palette = RGB(255, 250, 239)
                Case "titleDark": Palette = RGB(119, 88, 25)
                Case "titleMid": Palette = RGB(166, 128, 36)
                Case "titleAlt": Palette = RGB(150, 109, 42)
                Case "header": Palette = RGB(225, 205, 145)
                Case "total": Palette = RGB(235, 182, 76)
                Case "note": Palette = RGB(242, 234, 209)
                Case "slot": Palette = RGB(243, 221, 166)
                Case "fill": Palette = RGB(225, 196, 116)
                Case "overflow": Palette = RGB(191, 150, 97)
                Case "textDark": Palette = RGB(70, 51, 20)
                Case "textLight": Palette = RGB(255, 251, 241)
                Case "tab1": Palette = RGB(166, 128, 36)
                Case "tab2": Palette = RGB(119, 88, 25)
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
        Case "亓家"
            Select Case partName
                Case "base": Palette = RGB(231, 231, 234)
                Case "panel": Palette = RGB(202, 203, 209)
                Case "table": Palette = RGB(244, 244, 247)
                Case "titleDark": Palette = RGB(28, 28, 31)
                Case "titleMid": Palette = RGB(62, 62, 70)
                Case "titleAlt": Palette = RGB(89, 56, 76)
                Case "header": Palette = RGB(164, 165, 174)
                Case "total": Palette = RGB(126, 127, 139)
                Case "note": Palette = RGB(221, 221, 226)
                Case "slot": Palette = RGB(188, 189, 197)
                Case "fill": Palette = RGB(144, 145, 158)
                Case "overflow": Palette = RGB(126, 98, 112)
                Case "textDark": Palette = RGB(26, 26, 30)
                Case "textLight": Palette = RGB(247, 247, 249)
                Case "tab1": Palette = RGB(62, 62, 70)
                Case "tab2": Palette = RGB(28, 28, 31)
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
    ws.Range("A1:M320").Interior.Color = Palette(themeKey, "base")

    PaintBlock ws.Range("A1:E1"), Palette(themeKey, "titleDark"), Palette(themeKey, "textLight"), 13.5, True, True
    PaintBlock ws.Range("H1:I1"), Palette(themeKey, "titleAlt"), Palette(themeKey, "textLight"), 12, True, True
    PaintBlock ws.Range("K1:M1"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 12, True, True
    PaintBlock ws.Range("A3:E3"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 12, True, True
    PaintBlock ws.Range("A4:E4"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("H3:I7"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, True
    PaintBlock ws.Range("H3:H7"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11, True, True
    PaintBlock ws.Range("K3:M3"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("K4:M13"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, True
    PaintBlock ws.Range("H9:I10"), Palette(themeKey, "note"), Palette(themeKey, "textDark"), 11, False, False

    ws.Range("A5:A304").HorizontalAlignment = xlHAlignCenter
    ws.Range("A5:A304").VerticalAlignment = xlVAlignCenter
    ws.Rows("5:304").RowHeight = 24

    ApplyGrid ws.Range("A4:E19")
    ApplyGrid ws.Range("H3:I7")
    ApplyGrid ws.Range("K3:M13")
    ApplyGrid ws.Range("H9:I10")
    ws.Tab.Color = Palette(themeKey, "tab2")
End Sub

Private Sub ApplyThemeKui(ByVal ws As Worksheet, ByVal themeKey As String)
    ws.Cells.Font.Name = "KaiTi"
    ws.Cells.Font.Size = 11
    ws.Range("A1:J330").Interior.Color = Palette(themeKey, "base")

    PaintBlock ws.Range("A1:H1"), Palette(themeKey, "titleDark"), Palette(themeKey, "textLight"), 13.5, True, True
    PaintBlock ws.Range("A2:B2"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("C2:D2"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("F2:H2"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 11.5, True, True
    PaintBlock ws.Range("F3:H3"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("A3:B9"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("C3:D9"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("F4:H13"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False

    PaintBlock ws.Range("A11:B14"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("C11:D15"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("A11:A14"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11, True, True
    PaintBlock ws.Range("C11:C15"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11, True, True
    PaintBlock ws.Range("A16:E16"), Palette(themeKey, "titleAlt"), Palette(themeKey, "textLight"), 12, True, True
    PaintBlock ws.Range("A17:E17"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("A18:E27"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, True
    PaintBlock ws.Range("F16:H17"), Palette(themeKey, "note"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("A29:E29"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 12, True, True
    PaintBlock ws.Range("A30:E30"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True

    ws.Rows("1:330").RowHeight = 24
    ws.Rows("1:1").RowHeight = 28
    ws.Rows("16:17").RowHeight = 26
    ws.Rows("29:30").RowHeight = 26
    ws.Range("A31:A330").HorizontalAlignment = xlHAlignCenter
    ws.Range("A31:A330").VerticalAlignment = xlVAlignCenter

    ApplyGrid ws.Range("A2:B9")
    ApplyGrid ws.Range("C2:D9")
    ApplyGrid ws.Range("F3:H13")
    ApplyGrid ws.Range("A11:B14")
    ApplyGrid ws.Range("C11:D15")
    ApplyGrid ws.Range("A17:E27")
    ApplyGrid ws.Range("F16:H17")
    ApplyGrid ws.Range("A30:E35")
    ws.Tab.Color = Palette(themeKey, "tab1")
End Sub

Private Sub ApplyThemeCatalog(ByVal ws As Worksheet, ByVal themeKey As String)
    ws.Cells.Font.Name = "KaiTi"
    ws.Cells.Font.Size = 11
    ws.Range("A1:W41").Interior.Color = Palette(themeKey, "base")

    PaintBlock ws.Range("A1:N2"), Palette(themeKey, "titleDark"), Palette(themeKey, "textLight"), 13, True, True
    PaintBlock ws.Range("A3:D3"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 11.5, True, True
    PaintBlock ws.Range("F3:I3"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 11.5, True, True
    PaintBlock ws.Range("K3:M3"), Palette(themeKey, "titleAlt"), Palette(themeKey, "textLight"), 11.5, True, True
    PaintBlock ws.Range("P3:Q3"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 11.5, True, True
    PaintBlock ws.Range("R3:S3"), Palette(themeKey, "titleMid"), Palette(themeKey, "textLight"), 11.5, True, True
    PaintBlock ws.Range("P11:S11"), Palette(themeKey, "titleAlt"), Palette(themeKey, "textLight"), 11.5, True, True

    PaintBlock ws.Range("A4:D17"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("F4:I17"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("K4:M17"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("P4:Q9"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("R4:S9"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("P12:S17"), Palette(themeKey, "note"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("A4:D4"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("F4:I4"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("K4:M4"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("P4:Q4"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True
    PaintBlock ws.Range("R4:S4"), Palette(themeKey, "header"), Palette(themeKey, "textDark"), 11.5, True, True

    ws.Rows("1:2").RowHeight = 28
    ws.Rows("3:41").RowHeight = 24
    ws.Tab.Color = Palette(themeKey, "tab2")

    ApplyGrid ws.Range("A3:D17")
    ApplyGrid ws.Range("F3:I17")
    ApplyGrid ws.Range("K3:M17")
    ApplyGrid ws.Range("P3:Q9")
    ApplyGrid ws.Range("R3:S9")
    ApplyGrid ws.Range("P11:S17")
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
    PaintBlock ws.Range("A6:D14"), Palette(themeKey, "table"), Palette(themeKey, "textDark"), 11, False, False
    PaintBlock ws.Range("A16:D17"), Palette(themeKey, "note"), Palette(themeKey, "textDark"), 11, False, False

    For rowIndex = 6 To 14
        rowTheme = CStr(ws.Cells(rowIndex, "A").Value)
        ws.Cells(rowIndex, "D").Interior.Color = PrimaryColor(rowTheme)
        ws.Cells(rowIndex, "D").Font.Color = Palette(themeKey, "textLight")
        ws.Cells(rowIndex, "D").HorizontalAlignment = xlHAlignCenter
        ws.Cells(rowIndex, "D").VerticalAlignment = xlVAlignCenter
    Next rowIndex

    ws.Rows("4:17").RowHeight = 28
    ApplyGrid ws.Range("A2:B2")
    ApplyGrid ws.Range("A5:D14")
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

