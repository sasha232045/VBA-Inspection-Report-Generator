Option Explicit

'==================================================================================================
' [Module] M04_Logger
' [Description] ログ出力関連モジュール
'==================================================================================================

Private Const LOG_SHEET_NAME As String = "Log"
Private Const ERROR_SHEET_NAME As String = "ErrorLog"

'--------------------------------------------------------------------------------------------------
' [Sub] InitializeLogs
' [Description] LogシートとErrorシートを初期化する
'--------------------------------------------------------------------------------------------------
Public Sub InitializeLogs()
    M06_DebugLogger.WriteDebugLog "LogシートとErrorシートを初期化します。"
    InitializeSheet LOG_SHEET_NAME, Array("日時", "レベル", "処理内容", "詳細")
    InitializeSheet ERROR_SHEET_NAME, Array("日時", "エラーレベル", "処理シートNo", "処理No.", "エラー内容", "エラー詳細")
End Sub

'--------------------------------------------------------------------------------------------------
' [Sub] WriteLog
' [Description] Logシートに情報を書き込む
' [Args] content: 処理内容, details: 詳細
'--------------------------------------------------------------------------------------------------
Public Sub WriteLog(ByVal content As String, Optional ByVal details As String = "")
    Dim logSheet As Worksheet
    Dim newRow As Long

    On Error Resume Next
    Set logSheet = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    On Error GoTo 0
    If logSheet Is Nothing Then Exit Sub ' シートが存在しない場合は処理を抜ける

    newRow = logSheet.Cells(logSheet.Rows.Count, 1).End(xlUp).Row + 1

    With logSheet
        .Cells(newRow, 1).Value = Now
        .Cells(newRow, 2).Value = "INFO"
        .Cells(newRow, 3).Value = content
        .Cells(newRow, 4).Value = details
    End With
End Sub

'--------------------------------------------------------------------------------------------------
' [Sub] WriteError
' [Description] Errorシートにエラー情報を書き込む
' [Args] level, sheetNo, procNo, content, description
'--------------------------------------------------------------------------------------------------
Public Sub WriteError(ByVal level As String, ByVal sheetNo As String, ByVal procNo As String, ByVal content As String, ByVal description As String)
    Dim errorSheet As Worksheet
    Dim newRow As Long

    On Error Resume Next
    Set errorSheet = ThisWorkbook.Worksheets(ERROR_SHEET_NAME)
    On Error GoTo 0
    If errorSheet Is Nothing Then Exit Sub ' シートが存在しない場合は処理を抜ける

    newRow = errorSheet.Cells(errorSheet.Rows.Count, 1).End(xlUp).Row + 1

    With errorSheet
        .Cells(newRow, 1).Value = Now
        .Cells(newRow, 2).Value = level
        .Cells(newRow, 3).Value = sheetNo
        .Cells(newRow, 4).Value = procNo
        .Cells(newRow, 5).Value = content
        .Cells(newRow, 6).Value = description

        Select Case level
            Case "[致命的エラー]"
                .Rows(newRow).Interior.Color = vbRed
            Case "[警告]"
                .Rows(newRow).Interior.Color = vbYellow
        End Select
    End With
End Sub

'--------------------------------------------------------------------------------------------------
' [Private Sub] InitializeSheet
' [Description] 指定されたシートを初期化し、ヘッダーを書き込む。シートがなければ作成する。
' [Args] sheetName: シート名, headers: ヘッダー配列
'--------------------------------------------------------------------------------------------------
Private Sub InitializeSheet(ByVal sheetName As String, ByVal headers As Variant)
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        M06_DebugLogger.WriteDebugLog "シート '" & sheetName & "' が存在しないため、新規に作成します。"
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        M06_DebugLogger.WriteDebugLog "シート '" & sheetName & "' をクリアします。"
    End If

    ws.Cells.Clear
    ws.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
End Sub
