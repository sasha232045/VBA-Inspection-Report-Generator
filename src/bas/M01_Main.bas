Option Explicit

<<<<<<< HEAD
'==================================================================================================
' [Module] M01_Main
' [Description] メインコントローラーモジュール
'==================================================================================================

'--------------------------------------------------------------------------------------------------
' [Sub] StartProcess
' [Description] 処理全体の流れを制御するメインプロシージャ
'--------------------------------------------------------------------------------------------------
Public Sub StartProcess()
    Dim oldBookPath As String
    Dim judgeAddress As String
    Dim modelType As String ' 「型式」を格納する変数
    Dim templatePath As String
    Dim newBookPath As String
    Dim oldWb As Workbook
    Dim settingsSheet As Worksheet

    M06_DebugLogger.InitializeDebugLog
    M06_DebugLogger.WriteDebugLog "メイン処理を開始します。"

    Application.ScreenUpdating = False
    M04_Logger.InitializeLogs
    M04_Logger.WriteLog "処理開始"

    On Error GoTo FatalErrorHandler

    Set settingsSheet = ThisWorkbook.Worksheets("Settings")
    M06_DebugLogger.WriteDebugLog "設定シートから値を取得します。"
    With settingsSheet
        oldBookPath = Trim(.Range("D7").Value)
        judgeAddress = Trim(.Range("D8").Value)
    End With
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

    M06_DebugLogger.WriteDebugLog "SettingsシートのD21に型式を書き込みます。"
    settingsSheet.Range("D21").Value = modelType
=======
' メインプロシージャ
'================================================================================================'
Sub Main()
    DebugLog "M01_Main", "Main", "Start"
>>>>>>> 4a358f29dfa0286bcc247843225878f0af2a21d0
    
    ' 初期化
    Initialize

<<<<<<< HEAD
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
    MsgBox "処理が完了しました。"

    GoTo Finally

FatalErrorHandler:
    M06_DebugLogger.WriteDebugLog "致命的なエラーが発生しました。処理を中断します。 エラー: " & Err.Description
    M04_Logger.WriteError "[致命的エラー]", "-", "-", "実行時エラー: " & Err.Number, Err.Description
    MsgBox "致命的なエラーが発生しました。処理を中断します。詳細はErrorシートを確認してください。"

Finally:
    Application.ScreenUpdating = True
    M06_DebugLogger.WriteDebugLog "画面描画を再開し、処理を終了します。"
=======
    ' メイン処理
    DebugLog "M01_Main", "Main", "Calling ProcessReports"
    ProcessReports
    DebugLog "M01_Main", "Main", "Returned from ProcessReports"

    ' 終了処理
    Finalize
    
    DebugLog "M01_Main", "Main", "End"
>>>>>>> 4a358f29dfa0286bcc247843225878f0af2a21d0
End Sub

' 初期化プロシージャ
'================================================================================================'
Private Sub Initialize()
    DebugLog "M01_Main", "Initialize", "Start"
    
    ' ロガーの初期化
    InitializeLogger
    InitializeDebugLogger ' デバッグロガーの初期化を追加

    ' ログ開始
    Log "処理開始"
    
    ' 設定の読み込み
    DebugLog "M01_Main", "Initialize", "Calling LoadSettings"
    LoadSettings
    DebugLog "M01_Main", "Initialize", "Returned from LoadSettings"
    
    DebugLog "M01_Main", "Initialize", "End"
End Sub

' 終了処理プロシージャ
'================================================================================================'
Private Sub Finalize()
    DebugLog "M01_Main", "Finalize", "Start"
    
    ' ログ終了
    Log "処理終了"

<<<<<<< HEAD
GetValueError:
    GetValueFromOldBook = ""
    M06_DebugLogger.WriteDebugLog "値の取得でエラーが発生しました。"
    M04_Logger.WriteError "[警告]", "-", "-", "値の取得失敗", "旧ブックの '" & address & "' から値を取得できませんでした。エラー: " & Err.Description
End Function

'--------------------------------------------------------------------------------------------------
' [Function] GetTemplatePathFromList
' [Description] Listシートからブックの種類に対応するテンプレートパスを取得する
' [Args] bookType: ブックの種類
' [Returns] テンプレートファイルのパス
'--------------------------------------------------------------------------------------------------
Private Function GetTemplatePathFromList(ByVal bookType As String) As String
    Dim listSheet As Worksheet
    Dim findRange As Range

    M06_DebugLogger.WriteDebugLog "テンプレートパスの検索を開始: " & bookType
    Set listSheet = ThisWorkbook.Worksheets("List")
    Set findRange = listSheet.Columns("F").Find(What:=bookType, LookIn:=xlValues, lookat:=xlWhole)

    If Not findRange Is Nothing Then
        GetTemplatePathFromList = listSheet.Cells(findRange.Row, "G").Value
        M06_DebugLogger.WriteDebugLog "テンプレートパスが見つかりました: " & GetTemplatePathFromList
    Else
        GetTemplatePathFromList = ""
        M06_DebugLogger.WriteDebugLog "テンプレートパスが見つかりませんでした。"
    End If
End Function
=======
    ' ロガーのクリーンアップ
    CleanupLogger
    CleanupDebugLogger ' デバッグロガーのクリーンアップを追加
    
    DebugLog "M01_Main", "Finalize", "End"
End Sub
>>>>>>> 4a358f29dfa0286bcc247843225878f0af2a21d0
