Option Explicit

'==================================================================================================
' [Module] M06_DebugLogger
' [Description] デバッグログ出力モジュール
'==================================================================================================

' デバッグロガーを初期化する
'================================================================================================'
Public Sub InitializeDebugLog()
    Dim debugSheet As Worksheet
    
    ' DebugLogシートの存在確認・作成
    On Error Resume Next
    Set debugSheet = ThisWorkbook.Worksheets("DebugLog")
    On Error GoTo 0
    
    If debugSheet Is Nothing Then
        Set debugSheet = ThisWorkbook.Worksheets.Add
        debugSheet.Name = "DebugLog"
    End If
    
    ' ヘッダー行の設定
    debugSheet.Cells.Clear
    debugSheet.Cells(1, 1).Value = "日時"
    debugSheet.Cells(1, 2).Value = "デバッグメッセージ"
    
    ' 列幅の調整
    debugSheet.Columns("A:A").ColumnWidth = 20
    debugSheet.Columns("B:B").ColumnWidth = 100
    
    ' ヘッダー行の書式設定
    With debugSheet.Range("A1:B1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
End Sub

' デバッグログを出力する
' @param message ログメッセージ
'================================================================================================'
Public Sub WriteDebugLog(ByVal message As String)
    Dim debugSheet As Worksheet
    Dim lastRow As Long
    
    On Error Resume Next
    Set debugSheet = ThisWorkbook.Worksheets("DebugLog")
    On Error GoTo 0
    
    If Not debugSheet Is Nothing Then
        lastRow = debugSheet.Cells(debugSheet.Rows.Count, 1).End(xlUp).Row
        debugSheet.Cells(lastRow + 1, 1).Value = Format(Now(), "yyyy/m/d h:mm")
        debugSheet.Cells(lastRow + 1, 2).Value = message
    End If
End Sub

