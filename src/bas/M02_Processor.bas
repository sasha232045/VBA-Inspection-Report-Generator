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
'--------------------------------------------------------------------------------------------------
' [Function] ExecuteAllTasks
' [Description] Settingsシートの定義に基づき、全てのデータ移行タスクを実行する
' [Args] oldBookPath: 旧ブックのフルパス, newBookPath: 新ブックのフルパス
' [Returns] 警告が発生した場合にTrueを返す
'--------------------------------------------------------------------------------------------------
Public Function ExecuteAllTasks(ByVal oldBookPath As String, ByVal newBookPath As String) As Boolean
    Dim settingsSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim oldWb As Workbook, newWb As Workbook
    Dim oldSheetName As String, newSheetName As String
    Dim copyFromAddr As String, copyToAddr As String
    Dim deleteAddr As String
    Dim inputAddr As String, inputValue As String
    Dim sheetNo As String, procNo As String, procContent As String, procValue As String

    ExecuteAllTasks = False ' 初期値はFalse (警告なし)

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
        Exit Function
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
    Exit Function

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
    ExecuteAllTasks = True ' 警告があったことを示す戻り値をセット
    Resume ContinueNextTask
End Function

'--------------------------------------------------------------------------------------------------
' [Private Sub] CopyData
' [Description] コピー処理を実行 (PasteSpecialによる堅牢な結合セル対応)
'--------------------------------------------------------------------------------------------------
Private Sub CopyData(oldWb As Workbook, newWb As Workbook, oldShtName As String, newShtName As String, srcAddr As String, dstAddr As String, sheetNo As String, procNo As String)
    Dim srcRange As Range
    Dim dstCell As Range

    M06_DebugLogger.WriteDebugLog "コピー処理実行: 旧[" & oldShtName & "!" & srcAddr & "] -> 新[" & newShtName & "!" & dstAddr & "]"

    Set srcRange = oldWb.Worksheets(oldShtName).Range(srcAddr)
    Set dstCell = newWb.Worksheets(newShtName).Range(dstAddr).Cells(1, 1)

    ' PasteSpecialを使用して、結合セルの問題を回避しつつ、書式ごとコピーする
    srcRange.Copy
    ' 最初に「値と数値の書式」を貼り付け
    dstCell.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    ' 次に「書式（結合状態、色、罫線など）」を貼り付け
    dstCell.PasteSpecial Paste:=xlPasteFormats

    Application.CutCopyMode = False ' コピーモードを解除

    M04_Logger.WriteLog "コピー処理", "旧: '" & oldShtName & "'!" & srcAddr & " -> 新: '" & newShtName & "'!" & dstAddr
End Sub

'--------------------------------------------------------------------------------------------------
' [Private Sub] DeleteData
' [Description] 削除処理を実行 (シート保護・詳細ログ対応)
'--------------------------------------------------------------------------------------------------
Private Sub DeleteData(wb As Workbook, shtName As String, addr As String, sheetNo As String, procNo As String)
    Dim targetSheet As Worksheet
    Dim targetRange As Range

    M06_DebugLogger.WriteDebugLog "[DeleteData] 開始: 対象[" & shtName & "!" & addr & "]"

    On Error GoTo GetSheetError
    Set targetSheet = wb.Worksheets(shtName)
    M06_DebugLogger.WriteDebugLog "[DeleteData] シートオブジェクト取得成功"
    On Error GoTo 0

    On Error GoTo GetRangeError
    Set targetRange = M05_Utility.GetRangeFromAddressString(targetSheet, addr)
    If targetRange Is Nothing Then
        M06_DebugLogger.WriteDebugLog "[DeleteData] 致命的エラー: アドレス変換失敗。アドレス: " & addr
        Exit Sub
    End If
    M06_DebugLogger.WriteDebugLog "[DeleteData] レンジオブジェクト取得成功"
    On Error GoTo 0

    On Error GoTo UnprotectError
    targetSheet.Unprotect
    M06_DebugLogger.WriteDebugLog "[DeleteData] シート保護解除成功"
    On Error GoTo 0

    On Error GoTo ClearContentsError
    targetRange.MergeArea.ClearContents
    M06_DebugLogger.WriteDebugLog "[DeleteData] ClearContents 正常終了"
    On Error GoTo 0

    On Error GoTo ProtectError
    targetSheet.Protect
    M06_DebugLogger.WriteDebugLog "[DeleteData] シート再保護成功"
    On Error GoTo 0

    M04_Logger.WriteLog "削除処理", "対象: '" & shtName & "'!" & addr
    Exit Sub

GetSheetError:
    M06_DebugLogger.WriteDebugLog "[DeleteData] 致命的エラー: シート取得失敗。シート名='" & shtName & "'. エラー: " & Err.Description
    Exit Sub
GetRangeError:
    M06_DebugLogger.WriteDebugLog "[DeleteData] 致命的エラー: レンジ取得失敗。アドレス='" & addr & "'. エラー: " & Err.Description
    Exit Sub
UnprotectError:
    M06_DebugLogger.WriteDebugLog "[DeleteData] 警告: 保護解除失敗。処理を続行します。エラー: " & Err.Description
    Resume Next
ClearContentsError:
    M06_DebugLogger.WriteDebugLog "[DeleteData] 致命的エラー: ClearContents失敗。エラー: " & Err.Description
    Resume ProtectLabel
ProtectError:
    M06_DebugLogger.WriteDebugLog "[DeleteData] 警告: 再保護失敗。エラー: " & Err.Description
ProtectLabel:
    targetSheet.Protect
End Sub

'--------------------------------------------------------------------------------------------------
' [Private Sub] InputData
' [Description] 入力処理を実行 (結合セル・シート保護対応)
'--------------------------------------------------------------------------------------------------
Private Sub InputData(wb As Workbook, shtName As String, addr As String, val As String, sheetNo As String, procNo As String)
    Dim targetSheet As Worksheet
    Dim targetRange As Range
    Dim cell As Range

    M06_DebugLogger.WriteDebugLog "入力処理実行: 対象[" & shtName & "!" & addr & "] に '" & val & "' を入力"
    On Error GoTo GetSheetError
    Set targetSheet = wb.Worksheets(shtName)
    On Error GoTo 0

    Set targetRange = M05_Utility.GetRangeFromAddressString(targetSheet, addr)

    If Not targetRange Is Nothing Then
        On Error GoTo UnprotectError
        targetSheet.Unprotect
        M06_DebugLogger.WriteDebugLog "シート '" & shtName & "' の保護を解除しました。"
        On Error GoTo 0

        On Error GoTo InputValueError
        For Each cell In targetRange
            If cell.MergeCells Then
                cell.MergeArea.Cells(1, 1).Value = val
            Else
                cell.Value = val
            End If
        Next cell
        M06_DebugLogger.WriteDebugLog "入力処理が完了しました。"
        On Error GoTo 0

        targetSheet.Protect
        M06_DebugLogger.WriteDebugLog "シート '" & shtName & "' を再度保護しました。"
        M04_Logger.WriteLog "入力処理", "対象: '" & shtName & "'!" & addr & ", 内容: " & val
    Else
        M06_DebugLogger.WriteDebugLog "入力処理スキップ: アドレス変換に失敗しました。"
    End If
    Exit Sub

GetSheetError:
    M06_DebugLogger.WriteDebugLog "入力処理エラー: シート '" & shtName & "' の取得に失敗しました。"
    Exit Sub
UnprotectError:
    M06_DebugLogger.WriteDebugLog "入力処理警告: シート '" & shtName & "' の保護解除に失敗しました。処理を続行します。"
    Resume Next
InputValueError:
    M06_DebugLogger.WriteDebugLog "入力処理エラー: 値の入力に失敗しました。セル: " & cell.Address & ", エラー: " & Err.Description
    targetSheet.Protect ' エラーが発生してもシートは再度保護する
End Sub