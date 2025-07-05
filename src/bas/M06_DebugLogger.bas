Option Explicit

'==================================================================================================
' [Module] M06_DebugLogger
' [Description] デバッグログ出力関連モジュール
'==================================================================================================

' デバッグモードのON/OFFを切り替える定数
Public Const IS_DEBUG_MODE As Boolean = True

Private Const DEBUG_LOG_SHEET_NAME As String = "DebugLog"

'--------------------------------------------------------------------------------------------------
' [Sub] InitializeDebugLog
' [Description] DebugLogシートを初期化する
'--------------------------------------------------------------------------------------------------
Public Sub InitializeDebugLog()
    If Not IS_DEBUG_MODE Then Exit Sub
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DEBUG_LOG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = DEBUG_LOG_SHEET_NAME
    End If

    ws.Cells.Clear
    ws.Range("A1").Resize(1, 2).Value = Array("日時", "デバッグメッセージ")
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 100
End Sub

'--------------------------------------------------------------------------------------------------
' [Sub] WriteDebugLog
' [Description] DebugLogシートにデバッグ情報を書き込む
' [Args] message: ログメッセージ
'--------------------------------------------------------------------------------------------------
Public Sub WriteDebugLog(ByVal message As String)
    If Not IS_DEBUG_MODE Then Exit Sub

    Dim debugSheet As Worksheet
    Dim newRow As Long

    On Error Resume Next
    Set debugSheet = ThisWorkbook.Worksheets(DEBUG_LOG_SHEET_NAME)
    On Error GoTo 0
    If debugSheet Is Nothing Then Exit Sub ' 初期化に失敗している場合は何もしない

    newRow = debugSheet.Cells(debugSheet.Rows.Count, 1).End(xlUp).Row + 1

    With debugSheet
        .Cells(newRow, 1).Value = Now
        .Cells(newRow, 2).Value = message
    End With
End Sub
