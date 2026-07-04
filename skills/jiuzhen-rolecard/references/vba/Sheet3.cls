VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim validationMessage As String

    If Application.EnableEvents = False Then
        Exit Sub
    End If

    If Intersect(Target, Me.Range("B5:E304,K4:M13,P7,P10:P12,U5:U14,P17:P19")) Is Nothing Then
        Exit Sub
    End If

    On Error GoTo CleanFail

    If Not ValidateKuiModuleSelectionChange(Me, Target, validationMessage) Then
        Application.EnableEvents = False
        Application.Undo
        Application.EnableEvents = True
        MsgBox validationMessage, vbExclamation
        Exit Sub
    End If

    Application.EnableEvents = False
    RefreshMojiaInventory

CleanExit:
    Application.EnableEvents = True
    Exit Sub

CleanFail:
    Application.EnableEvents = True
    MsgBox "Refresh inventory failed: " & Err.Description, vbExclamation
End Sub
