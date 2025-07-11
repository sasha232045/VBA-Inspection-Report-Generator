Option Explicit

'==================================================================================================
' [Module] M01_Main
' [Description] メインコントローラーモジュール
'==================================================================================================

'--------------------------------------------------------------------------------------------------
' [Sub] StartProcess
' [Description] 処理全体の流れを制御するメインプロシージャ
'--------------------------------------------------------------------------------------------------
Public Sub StartProcess()
    Dim inputSheet As Worksheet
    Dim settingsSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long

    M06_DebugLogger.InitializeDebugLog
    M06_DebugLogger.WriteDebugLog "メイン処理を開始します。"

    ' アプリケーション設定（ダイアログ無効化）
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = False ' ステータスバーをリセット
    Application.EnableEvents = False ' イベント無効化
    Application.AskToUpdateLinks = False ' リンク更新ダイアログ無効化
    
    ' AlertBeforeOverwritingの安全な設定（サポートされている場合のみ）
    On Error Resume Next
    Application.AlertBeforeOverwriting = False ' 上書き警告無効化
    On Error GoTo 0
    
    M06_DebugLogger.WriteDebugLog "アプリケーション設定: ダイアログとアラートを無効化しました。"
    
    M04_Logger.InitializeLogs
    M04_Logger.WriteLog "処理開始"

    Set inputSheet = ThisWorkbook.Worksheets("入力")
    Set settingsSheet = ThisWorkbook.Worksheets("Settings")
    lastRow = inputSheet.Cells(inputSheet.Rows.Count, "B").End(xlUp).Row

    M06_DebugLogger.WriteDebugLog "入力シートの最終行: " & lastRow

    Dim lastNewBookPath As String

    For i = 2 To lastRow
        M06_DebugLogger.WriteDebugLog "入力シートの " & i & "行目の処理を開始します。"
        
        ' 入力シートの内容をSettingsシートに転記
        settingsSheet.Range("D7").Value = inputSheet.Cells(i, "B").Value
        settingsSheet.Range("D8").Value = inputSheet.Cells(i, "C").Value
        settingsSheet.Range("D9").Value = inputSheet.Cells(i, "D").Value
        settingsSheet.Range("D10").Value = inputSheet.Cells(i, "E").Value
        settingsSheet.Range("D11").Value = inputSheet.Cells(i, "G").Value
        
        M06_DebugLogger.WriteDebugLog "入力シートからSettingsシートへの転記が完了しました。"

        ' 一行分の処理を実行
        lastNewBookPath = ProcessInputRow()
        
        M06_DebugLogger.WriteDebugLog "入力シートの " & i & "行目の処理が完了しました。"
    Next i

    M04_Logger.WriteLog "全処理正常終了"
    M06_DebugLogger.WriteDebugLog "メイン処理が正常に終了しました。"
    
    ' 処理完了メッセージと保存フォルダを開く
    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("すべての処理が完了しました。" & vbCrLf & vbCrLf & _
                         "最後に作成されたファイル: " & Dir(lastNewBookPath) & vbCrLf & _
                         "保存場所: " & lastNewBookPath & vbCrLf & vbCrLf & _
                         "保存フォルダを開きますか？", _
                         vbYesNo + vbInformation, "処理完了")
    
    If userResponse = vbYes Then
        M06_DebugLogger.WriteDebugLog "保存フォルダを開きます。"
        M03_FileHandler.OpenFileLocation lastNewBookPath
    End If

Finally:
    ' アプリケーション設定を復元
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = True
    Application.EnableEvents = True
    Application.AskToUpdateLinks = True
    
    ' AlertBeforeOverwritingの安全な復元
    On Error Resume Next
    Application.AlertBeforeOverwriting = True
    On Error GoTo 0
    
    Application.Calculation = xlCalculationAutomatic ' 計算を自動に戻す
    
    M06_DebugLogger.WriteDebugLog "アプリケーション設定を復元し、画面描画を再開し、処理を終了します。"
End Sub

'--------------------------------------------------------------------------------------------------
' [Sub] ProcessInputRow
' [Description] 入力シート1行分のデータ処理を実行する
'--------------------------------------------------------------------------------------------------
Private Function ProcessInputRow() As String
    Dim oldBookPath As String
    Dim judgeAddress As String
    Dim modelType As String ' 「型式」を格納する変数
    Dim templatePath As String
    Dim newBookPath As String
    Dim oldWb As Workbook
    Dim settingsSheet As Worksheet
    Dim retryCount As Integer

    On Error GoTo FatalErrorHandler

    Set settingsSheet = ThisWorkbook.Worksheets("Settings")
    M06_DebugLogger.WriteDebugLog "設定シートから値を取得します。"
    With settingsSheet
        oldBookPath = Trim(.Range("D7").Value)
        judgeAddress = Trim(.Range("D8").Value)
    End With
    
    ' 独自ダイアログ無視オプションの確認
    Dim ignoreCustomDialog As Boolean
    ignoreCustomDialog = (UCase(Trim(CStr(settingsSheet.Range("B28").Value))) = "非表示")
    If ignoreCustomDialog Then
        M06_DebugLogger.WriteDebugLog "独自ダイアログ非表示オプションが有効です。"
        ' さらに強力なダイアログ無効化
        Application.Calculation = xlCalculationManual ' 計算を手動に
        Application.ScreenUpdating = False
    End If
    
    M06_DebugLogger.WriteDebugLog "旧ブックパス: " & oldBookPath
    M06_DebugLogger.WriteDebugLog "判定アドレス: " & judgeAddress

    ' 必須項目のチェックを追加
    If oldBookPath = "" Then
        M04_Logger.WriteError "[致命的エラー]", "-", "-", "設定エラー", "旧ブック_ファイルパス (D7) が入力されていません。"
        GoTo FatalErrorHandler
    End If
    
    If judgeAddress = "" Then
        M04_Logger.WriteError "[致命的エラー]", "-", "-", "設定エラー", "旧ブック_新ブック名判定アドレス (D8) が入力されていません。"
        GoTo FatalErrorHandler
    End If

    M06_DebugLogger.WriteDebugLog "旧ブックを開きます。"
    Set oldWb = M03_FileHandler.OpenWorkbook(oldBookPath)
    If oldWb Is Nothing Then
        M04_Logger.WriteError "[致命的エラー]", "-", "-", "旧ブックオープン失敗", "指定されたパスのファイルが開けません: " & oldBookPath
        GoTo FatalErrorHandler
    End If

    M06_DebugLogger.WriteDebugLog "旧ブックから型式を取得します。"
    modelType = GetValueFromOldBook(oldWb, judgeAddress)
    M06_DebugLogger.WriteDebugLog "取得した型式: " & modelType
    oldWb.Close SaveChanges:=False
    Set oldWb = Nothing

    M06_DebugLogger.WriteDebugLog "SettingsシートのD22に型式を書き込みます。"
    settingsSheet.Range("D22").Value = modelType
    
    ' XLOOKUPの計算完了を待つ
    Application.Calculate
    DoEvents
    Application.Wait (Now + TimeValue("00:00:01") / 2) ' 500ms待機
    
    M06_DebugLogger.WriteDebugLog "Excelの数式を再計算しました。"

    ' 再計算された結果を読み取る
    templatePath = settingsSheet.Range("D24").Value
    M06_DebugLogger.WriteDebugLog "再計算後のテンプレートパス: " & templatePath

    If templatePath = "" Or Not M03_FileHandler.FileExists(templatePath) Then
        M04_Logger.WriteError "[致命的エラー]", "-", "-", "テンプレート特定失敗", "D24セルから有効なテンプレートパスが取得できませんでした。パス: " & templatePath
        GoTo FatalErrorHandler
    End If

    M06_DebugLogger.WriteDebugLog "新ブックを作成します。"
    newBookPath = M03_FileHandler.CreateNewBook(templatePath)
    M06_DebugLogger.WriteDebugLog "新ブックパス: " & newBookPath

    M06_DebugLogger.WriteDebugLog "データ移行処理を開始します。"
    M02_Processor.ExecuteAllTasks oldBookPath, newBookPath

    M04_Logger.WriteLog "処理正常終了"
    M06_DebugLogger.WriteDebugLog "メイン処理が正常に終了しました。"
    
    ProcessInputRow = newBookPath
    
    ' 処理完了後、次の行へ
    GoTo Finally

FatalErrorHandler:
    M06_DebugLogger.WriteDebugLog "致命的なエラーが発生しました。処理を中断します。 エラー: " & Err.Description
    M04_Logger.WriteError "[致命的エラー]", "-", "-", "実行時エラー: " & Err.Number, Err.Description
    ' MsgBox "致命的なエラーが発生しました。処理を中断します。詳細はErrorシートを確認してください。"

Finally:
    ' このレベルでのFinallyは不要。StartProcessのFinallyで一括処理
End Function


'--------------------------------------------------------------------------------------------------
' [Function] GetValueFromOldBook
' [Description] 旧ブックから指定されたアドレスの値を取得する
' [Args] wb: 対象ワークブック, address: セルアドレス
' [Returns] セルの値
'--------------------------------------------------------------------------------------------------
Private Function GetValueFromOldBook(ByVal wb As Workbook, ByVal address As String) As String
    Dim sheetName As String
    Dim rangeAddress As String

    M06_DebugLogger.WriteDebugLog "旧ブックの値取得を開始: " & address
    On Error GoTo GetValueError

    If M03_FileHandler.ParseAddress(address, sheetName, rangeAddress) Then
        GetValueFromOldBook = wb.Worksheets(sheetName).Range(rangeAddress).Value
        M06_DebugLogger.WriteDebugLog "値を取得しました: " & GetValueFromOldBook
    Else
        GetValueFromOldBook = "" ' 解析失敗
    End If
    Exit Function

GetValueError:
    GetValueFromOldBook = ""
    M06_DebugLogger.WriteDebugLog "値の取得でエラーが発生しました。"
    M04_Logger.WriteError "[警告]", "-", "-", "値の取得失敗", "旧ブックの '" & address & "' から値を取得できませんでした。エラー: " & Err.description
End Function


