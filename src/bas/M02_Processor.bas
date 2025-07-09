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
    Dim originalDisplayAlerts As Boolean
    Dim originalAskToUpdateLinks As Boolean
    ' Application.CheckCompatibilityは古いExcelバージョンでサポートされていない可能性があるため削除

    M06_DebugLogger.WriteDebugLog "データ移行処理(ExecuteAllTasks)を開始します。"
    M06_DebugLogger.WriteDebugLog "旧ブックパス: " & oldBookPath
    M06_DebugLogger.WriteDebugLog "新ブックパス: " & newBookPath

    ' 強力なダイアログ抑制設定を保存
    originalDisplayAlerts = Application.DisplayAlerts
    originalAskToUpdateLinks = Application.AskToUpdateLinks
    
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    
    ' CheckCompatibilityプロパティの安全な設定
    On Error Resume Next
    Application.CheckCompatibility = False ' 互換性チェックを無効化（サポートされている場合のみ）
    On Error GoTo 0
    
    M06_DebugLogger.WriteDebugLog "ダイアログ抑制とアラート無効化を設定しました。"

    Set settingsSheet = ThisWorkbook.Worksheets("Settings")
    lastRow = settingsSheet.Cells(settingsSheet.Rows.Count, "A").End(xlUp).Row
    M06_DebugLogger.WriteDebugLog "Settingsシートの最終行: " & lastRow

    Set oldWb = M03_FileHandler.OpenWorkbook(oldBookPath)
    Set newWb = M03_FileHandler.OpenWorkbook(newBookPath)

    If oldWb Is Nothing Or newWb Is Nothing Then
        M04_Logger.WriteError "[致命的エラー]", "-", "-", "ブックオープン失敗", "処理対象のブックが開けませんでした。"
        M06_DebugLogger.WriteDebugLog "ブックが開けなかったため、処理を中断します。"
        GoTo RestoreSettings
    End If

    M06_DebugLogger.WriteDebugLog "SettingsシートのA19からループを開始します。"
    For i = 19 To lastRow
        On Error GoTo TaskErrorHandler

        sheetNo = settingsSheet.Cells(i, "A").Value
        procNo = settingsSheet.Cells(i, "B").Value
        procContent = settingsSheet.Cells(i, "C").Value
        procValue = settingsSheet.Cells(i, "D").Value
        M06_DebugLogger.WriteDebugLog i & "行目: " & sheetNo & ", " & procNo & ", " & procContent & ", " & procValue

        ' 有効な処理行のみ実行（シートNoと処理種類が両方入力されている行）
        If Trim(sheetNo) <> "" And IsNumeric(procNo) And Trim(procNo) <> "" Then
            Select Case CInt(procNo)
                Case 1  ' 旧ブック_シート名
                    oldSheetName = procValue
                Case 2  ' 新ブック_シート名
                    newSheetName = procValue
                Case 3  ' 旧ブック_コピー元アドレス
                    copyFromAddr = procValue
                Case 4  ' 新ブック_コピー先アドレス
                    copyToAddr = procValue
                    ' 4が来たらコピー実行
                    If oldSheetName <> "" And newSheetName <> "" And copyFromAddr <> "" And copyToAddr <> "" Then
                        CopyData oldWb, newWb, oldSheetName, newSheetName, copyFromAddr, copyToAddr, sheetNo, procNo
                    End If
                Case 5  ' 新ブック_削除するアドレス
                    deleteAddr = procValue
                    If newSheetName <> "" And deleteAddr <> "" Then
                        DeleteData newWb, newSheetName, deleteAddr, sheetNo, procNo
                    End If
                Case 6  ' 新ブック_入力するアドレス
                    inputAddr = procValue
                Case 7  ' 新ブック_入力する内容
                    inputValue = procValue
                    ' 7が来たら入力実行
                    If newSheetName <> "" And inputAddr <> "" And inputValue <> "" Then
                        InputData newWb, newSheetName, inputAddr, inputValue, sheetNo, procNo
                    End If
            End Select
        Else
            ' 処理対象外の行はスキップ（エラーログに出力しない）
            M06_DebugLogger.WriteDebugLog "処理対象外の行をスキップしました: " & i & "行目"
        End If

ContinueNextTask:
    Next i

    M06_DebugLogger.WriteDebugLog "すべてのタスクが完了しました。ブックを保存して閉じます。"
    oldWb.Close SaveChanges:=False
    
    ' 新ブックを強制保存（ダイアログ抑制）
    On Error Resume Next
    newWb.Save ' 強制保存
    If Err.Number <> 0 Then
        M06_DebugLogger.WriteDebugLog "保存時にエラーが発生しましたが、処理を続行します: " & Err.Description
    End If
    On Error GoTo 0
    
    newWb.Close SaveChanges:=False ' 既に保存済みなのでFalse

RestoreSettings:
    ' アプリケーション設定を復元
    Application.DisplayAlerts = originalDisplayAlerts
    Application.AskToUpdateLinks = originalAskToUpdateLinks
    
    ' CheckCompatibilityプロパティの安全な復元
    On Error Resume Next
    Application.CheckCompatibility = True ' サポートされている場合のみ復元
    On Error GoTo 0
    
    M06_DebugLogger.WriteDebugLog "アプリケーション設定を復元しました。"
    Exit Sub

TaskErrorHandler:
    M06_DebugLogger.WriteDebugLog "タスク実行中にエラーが発生しました。エラーを記録し、次のタスクへ進みます。"
    M04_Logger.WriteError "[警告]", sheetNo, procNo, procContent, "エラー詳細: " & Err.Description
    Resume ContinueNextTask
End Sub

'--------------------------------------------------------------------------------------------------
' [Private Sub] CopyData
' [Description] コピー処理を実行
'--------------------------------------------------------------------------------------------------
Private Sub CopyData(oldWb As Workbook, newWb As Workbook, oldShtName As String, newShtName As String, srcAddr As String, dstAddr As String, sheetNo As String, procNo As String)
    On Error GoTo CopyError
    M06_DebugLogger.WriteDebugLog "コピー処理実行: 旧[" & oldShtName & "!" & srcAddr & "] -> 新[" & newShtName & "!" & dstAddr & "]"
    oldWb.Worksheets(oldShtName).Range(srcAddr).Copy newWb.Worksheets(newShtName).Range(dstAddr)
    M04_Logger.WriteLog "コピー処理", "旧: '" & oldShtName & "'!" & srcAddr & " -> 新: '" & newShtName & "'!" & dstAddr
    M06_DebugLogger.WriteDebugLog "コピー処理が正常に完了しました。"
    Exit Sub
    
CopyError:
    M06_DebugLogger.WriteDebugLog "コピー処理でエラーが発生しました: " & Err.Description
    M04_Logger.WriteError "[警告]", sheetNo, procNo, "コピー処理エラー", "旧: " & oldShtName & "!" & srcAddr & " -> 新: " & newShtName & "!" & dstAddr & ", エラー: " & Err.Description
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
    
    On Error GoTo InputError
    M06_DebugLogger.WriteDebugLog "入力処理実行: 対象[" & shtName & "!" & addr & "] に '" & val & "' を入力"
    Set targetRange = M05_Utility.GetRangeFromAddressString(wb.Worksheets(shtName), addr)
    If Not targetRange Is Nothing Then
        targetRange.Value = val
        M04_Logger.WriteLog "入力処理", "対象: '" & shtName & "'!" & addr & ", 内容: " & val
        M06_DebugLogger.WriteDebugLog "入力処理が正常に完了しました。"
    Else
        M06_DebugLogger.WriteDebugLog "入力処理スキップ: アドレス変換に失敗しました。"
    End If
    Exit Sub
    
InputError:
    M06_DebugLogger.WriteDebugLog "入力処理でエラーが発生しました: " & Err.Description
    M04_Logger.WriteError "[警告]", sheetNo, procNo, "入力処理エラー", "アドレス: " & addr & ", 値: " & val & ", エラー: " & Err.Description
End Sub