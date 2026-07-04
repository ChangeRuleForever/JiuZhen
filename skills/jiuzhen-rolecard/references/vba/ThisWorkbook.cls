VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    ApplyWorkbookTheme
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim validationMessage As String

    If Application.EnableEvents = False Then
        Exit Sub
    End If

    On Error GoTo CleanExit

    If Not CanSaveMojiaWorkbook(validationMessage) Then
        Cancel = True
        MsgBox validationMessage, vbExclamation
    End If

CleanExit:
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Application.EnableEvents = False Then
        Exit Sub
    End If

    On Error GoTo CleanExit

    If Sh.Name = "通用特性体系" Then
        If Not Intersect(Target, Sh.Range("E1:G2")) Is Nothing Then
            Application.EnableEvents = False
            ToggleTraitCatalogFamilyFilter
            Sh.Range("A1").Select
        End If
    ElseIf Sh.Name = "技能列表" Then
        If Not Intersect(Target, Sh.Range("G1:I2")) Is Nothing Then
            Application.EnableEvents = False
            ToggleSkillCatalogFamilyFilter
            Sh.Range("A1").Select
        End If
    End If

CleanExit:
    Application.EnableEvents = True
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Application.EnableEvents = False Then
        Exit Sub
    End If

    On Error GoTo CleanExit

    If Sh Is Sheet7 Then
        If Not Intersect(Target, Sheet7.Range("B2")) Is Nothing Then
            Application.EnableEvents = False
            ApplyWorkbookTheme
        End If
    ElseIf Sh.Name = "装备列表" Then
        If Not Intersect(Target, Sh.Range("A:K")) Is Nothing Then
            Application.EnableEvents = False
            ApplyWorkbookTheme
        End If
    ElseIf Sh.Name = "通用特性体系" Then
        If Not Intersect(Target, Sh.Range("A:D")) Is Nothing Then
            Application.EnableEvents = False
            ApplyWorkbookTheme
        End If
    ElseIf Sh.Name = "技能列表" Then
        If Not Intersect(Target, Sh.Range("A:F")) Is Nothing Then
            Application.EnableEvents = False
            ApplyWorkbookTheme
        End If
    End If

CleanExit:
    Application.EnableEvents = True
End Sub
