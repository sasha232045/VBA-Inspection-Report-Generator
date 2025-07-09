Option Explicit

' グローバル変数
'================================================================================================'
Private settings As Object ' Scripting.Dictionary

' メイン処理
'================================================================================================'
Public Sub ProcessReports()
    DebugLog "M02_Processor", "ProcessReports", "Start"
    
    ' 変数宣言
    Dim dataSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim templateSheetName As String
    Dim outputSheetName As String
    
    ' 初期設定
    Set dataSheet = ThisWorkbook.Sheets("List")
    lastRow = GetLastRow(dataSheet, 1) ' A列を基準に行数を取得
    DebugLog "M02_Processor", "ProcessReports", "Data sheet 'List' last row: " & lastRow
    
    ' 設定を辞書から取得
    templateSheetName = settings("TemplateSheetName")
    outputSheetName = settings("OutputSheetName")

    ' デバッグログ
    DebugLog "M02_Processor", "ProcessReports", "TemplateSheetName from settings: '" & templateSheetName & "'"
    DebugLog "M02_Processor", "ProcessReports", "OutputSheetName from settings: '" & outputSheetName & "'"

    ' データ行をループしてレポートを作成
    For i = 2 To lastRow
        DebugLog "M02_Processor", "ProcessReports", "Processing row: " & i
        ' レポート作成
        CreateReport dataSheet, i, templateSheetName, outputSheetName
    Next i
    
    DebugLog "M02_Processor", "ProcessReports", "End"
End Sub

' 設定シートから設定を読み込み、グローバルな辞書に格納する
'================================================================================================'
Public Sub LoadSettings()
    DebugLog "M02_Processor", "LoadSettings", "Start"
    
    Set settings = CreateObject("Scripting.Dictionary")
    Dim settingsSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim key As String
    Dim Value As String
    
    ' "Settings"シートを取得
    On Error Resume Next
    Set settingsSheet = ThisWorkbook.Sheets("Settings")
    On Error GoTo 0
    
    If settingsSheet Is Nothing Then
        Log "エラー: Settingsシートが見つかりません。"
        DebugLog "M02_Processor", "LoadSettings", "Error: 'Settings' sheet not found. Exiting sub."
        Exit Sub
    End If
    DebugLog "M02_Processor", "LoadSettings", "'Settings' sheet found."
    
    ' C列を基準に最終行を取得
    lastRow = GetLastRow(settingsSheet, 3)
    DebugLog "M02_Processor", "LoadSettings", "Settings sheet last row (Column C): " & lastRow
    
    ' C列とD列から設定を読み込む
    For i = 2 To lastRow
        key = settingsSheet.Cells(i, 3).Value
        Value = settingsSheet.Cells(i, 4).Value
        If key <> "" Then
            settings(key) = Value
            DebugLog "M02_Processor", "LoadSettings", "Loaded setting: Key='" & key & "', Value='" & Value & "'"
        End If
    Next i
    
    DebugLog "M02_Processor", "LoadSettings", "End"
End Sub

'--------------------------------------------------------------------------------------------------
' [Private Sub] CopyData
' [Description] コピー処理を実行
'--------------------------------------------------------------------------------------------------
Private Sub CopyData(oldWb As Workbook, newWb As Workbook, oldShtName As String, newShtName As String, srcAddr As String, dstAddr As String, sheetNo As String, procNo As String)
    M06_DebugLogger.WriteDebugLog "コピー処理実行: 旧[" & oldShtName & "!" & srcAddr & "] -> 新[" & newShtName & "!" & dstAddr & "]"
    oldWb.Worksheets(oldShtName).Range(srcAddr).Copy newWb.Worksheets(newShtName).Range(dstAddr)
    M04_Logger.WriteLog "コピー処理", "旧: '" & oldShtName & "'!" & srcAddr & " -> 新: '" & newShtName & "'!" & dstAddr
End Sub

'--------------------------------------------------------------------------------------------------
' [Private Sub] DeleteData
' [Description] 削除処理を実行
'--------------------------------------------------------------------------------------------------
Private Sub DeleteData(wb As Workbook, shtName As String, addr As String, sheetNo As String, procNo As String)
    Dim targetRange As Range
    Dim cell As Range
    
    On Error GoTo DeleteError
    M06_DebugLogger.WriteDebugLog "削除処理実行: 対象[" & shtName & "!" & addr & "]"
    Set targetRange = M05_Utility.GetRangeFromAddressString(wb.Worksheets(shtName), addr)
    If Not targetRange Is Nothing Then
        ' 結合セルを考慮した削除処理
        For Each cell In targetRange
            On Error Resume Next
            If cell.MergeCells Then
                cell.MergeArea.ClearContents
                M06_DebugLogger.WriteDebugLog "結合セル " & cell.MergeArea.Address & " の内容を削除しました。"
            Else
                cell.ClearContents
                M06_DebugLogger.WriteDebugLog "セル " & cell.Address & " の内容を削除しました。"
            End If
            On Error GoTo DeleteError
        Next cell
        M04_Logger.WriteLog "削除処理", "対象: '" & shtName & "'!" & addr
        M06_DebugLogger.WriteDebugLog "削除処理が正常に完了しました。"
    Else
        M06_DebugLogger.WriteDebugLog "削除処理スキップ: アドレス変換に失敗しました。"
    End If
    Exit Sub
    
DeleteError:
    M06_DebugLogger.WriteDebugLog "削除処理でエラーが発生しました: " & Err.Description
    M04_Logger.WriteError "[警告]", sheetNo, procNo, "削除処理エラー", "アドレス: " & addr & ", エラー: " & Err.Description
End Sub

'--------------------------------------------------------------------------------------------------
' [Private Sub] InputData
' [Description] 入力処理を実行
'--------------------------------------------------------------------------------------------------
Private Sub InputData(wb As Workbook, shtName As String, addr As String, val As String, sheetNo As String, procNo As String)
    Dim targetRange As Range
    M06_DebugLogger.WriteDebugLog "入力処理実行: 対象[" & shtName & "!" & addr & "] に '" & val & "' を入力"
    Set targetRange = M05_Utility.GetRangeFromAddressString(wb.Worksheets(shtName), addr)
    If Not targetRange Is Nothing Then
        targetRange.Value = val
        M04_Logger.WriteLog "入力処理", "対象: '" & shtName & "'!" & addr & ", 内容: " & val
    Else
        M06_DebugLogger.WriteDebugLog "入力処理スキップ: アドレス変換に失敗しました。"
        ' M04_Logger.WriteError を削除
    End If
End Sub