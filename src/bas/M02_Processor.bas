Option Explicit

'==================================================================================================
' [Module] M02_Processor
' [Description] データ処理コアロジックモジュール
'==================================================================================================

'--------------------------------------------------------------------------------------------------
' [Sub] ExecuteAllTasks
' [Description] Settingsシートの定義に基づき、全てのデータ移行タスクを実行する
' [Args] oldBookPath: 旧ブックのフルパス, newBookPath: 新ブックのフルパス
'--------------------------------------------------------------------------------------------------
Public Sub ExecuteAllTasks(ByVal oldBookPath As String, ByVal newBookPath As String)
    Dim settingsSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim oldWb As Workbook, newWb As Workbook
    Dim oldSheetName As String, newSheetName As String
    Dim copyFromAddr As String, copyToAddr As String
    Dim deleteAddr As String
    Dim inputAddr As String, inputValue As String
    Dim sheetNo As String, procNo As String, procContent As String, procValue As String

    M06_DebugLogger.WriteDebugLog "データ移行処理(ExecuteAllTasks)を開始します。"
    M06_DebugLogger.WriteDebugLog "旧ブックパス: " & oldBookPath
    M06_DebugLogger.WriteDebugLog "新ブックパス: " & newBookPath

    Set settingsSheet = ThisWorkbook.Worksheets("Settings")
    lastRow = settingsSheet.Cells(settingsSheet.Rows.Count, "A").End(xlUp).Row
    M06_DebugLogger.WriteDebugLog "Settingsシートの最終行: " & lastRow

    Set oldWb = M03_FileHandler.OpenWorkbook(oldBookPath)
    Set newWb = M03_FileHandler.OpenWorkbook(newBookPath)

    If oldWb Is Nothing Or newWb Is Nothing Then
        M04_Logger.WriteError "[致命的エラー]", "-", "-", "ブックオープン失敗", "処理対象のブックが開けませんでした。"
        M06_DebugLogger.WriteDebugLog "ブックが開けなかったため、処理を中断します。"
        Exit Sub
    End If

    M06_DebugLogger.WriteDebugLog "SettingsシートのA51からループを開始します。"
    For i = 51 To lastRow
        On Error GoTo TaskErrorHandler

        sheetNo = settingsSheet.Cells(i, "A").Value
        procNo = settingsSheet.Cells(i, "B").Value
        procContent = settingsSheet.Cells(i, "C").Value
        procValue = settingsSheet.Cells(i, "D").Value
        M06_DebugLogger.WriteDebugLog i & "行目: " & sheetNo & ", " & procNo & ", " & procContent & ", " & procValue

        Select Case CInt(procNo) ' B列の数値で分岐
            Case 1 '旧ブック_シート名'
                oldSheetName = procValue
            Case 2 '新ブック_シート名'
                newSheetName = procValue
            Case 3 '旧ブック_コピー元アドレス'
                copyFromAddr = procValue
            Case 4 '新ブック_コピー先アドレス'
                copyToAddr = procValue
                CopyData oldWb, newWb, oldSheetName, newSheetName, copyFromAddr, copyToAddr, sheetNo, procNo
            Case 5 '新ブック_削除するアドレス'
                deleteAddr = procValue
                DeleteData newWb, newSheetName, deleteAddr, sheetNo, procNo
            Case 6 '新ブック_入力するアドレス'
                inputAddr = procValue
            Case 7 '新ブック_入力する内容'
                inputValue = procValue
                InputData newWb, newSheetName, inputAddr, inputValue, sheetNo, procNo
        End Select

ContinueNextTask:
    Next i

    M06_DebugLogger.WriteDebugLog "すべてのタスクが完了しました。ブックを保存して閉じます。"
    oldWb.Close SaveChanges:=False
    newWb.Close SaveChanges:=True
    Exit Sub

TaskErrorHandler:
    M06_DebugLogger.WriteDebugLog "タスク実行中にエラーが発生しました。エラーを記録し、次のタスクへ進みます。 エラー: " & Err.Description
    Dim errorProcContent As String
    Select Case CInt(procNo)
        Case 1: errorProcContent = "旧ブック_シート名"
        Case 2: errorProcContent = "新ブック_シート名"
        Case 3: errorProcContent = "旧ブック_コピー元アドレス"
        Case 4: errorProcContent = "新ブック_コピー先アドレス"
        Case 5: errorProcContent = "新ブック_削除するアドレス"
        Case 6: errorProcContent = "新ブック_入力するアドレス"
        Case 7: errorProcContent = "新ブック_入力する内容"
        Case Else: errorProcContent = "不明な処理内容"
    End Select
    M04_Logger.WriteError "[警告]", sheetNo, procNo, errorProcContent, Err.Description
    hasWarnings = True ' 警告フラグを立てる
    ExecuteAllTasks = True
    Resume ContinueNextTask
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
    M06_DebugLogger.WriteDebugLog "削除処理実行: 対象[" & shtName & "!" & addr & "]"
    Set targetRange = M05_Utility.GetRangeFromAddressString(wb.Worksheets(shtName), addr)
    If Not targetRange Is Nothing Then
        targetRange.MergeArea.ClearContents
        M04_Logger.WriteLog "削除処理", "対象: '" & shtName & "'!" & addr
    Else
        M06_DebugLogger.WriteDebugLog "削除処理スキップ: アドレス変換に失敗しました。"
        ' M04_Logger.WriteError を削除
    End If
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