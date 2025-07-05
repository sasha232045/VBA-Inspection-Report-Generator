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
    Dim oldBookPath As String
    Dim judgeAddress As String
    Dim bookType As String
    Dim templatePath As String
    Dim newBookPath As String
    Dim oldWb As Workbook

    M06_DebugLogger.InitializeDebugLog
    M06_DebugLogger.WriteDebugLog "メイン処理を開始します。"

    Application.ScreenUpdating = False
    M04_Logger.InitializeLogs
    M04_Logger.WriteLog "処理開始"

    On Error GoTo FatalErrorHandler

    M06_DebugLogger.WriteDebugLog "設定シートから値を取得します。"
    With ThisWorkbook.Worksheets("Settings")
        oldBookPath = .Range("D4").Value
        judgeAddress = .Range("D5").Value
    End With
    M06_DebugLogger.WriteDebugLog "旧ブックパス: " & oldBookPath
    M06_DebugLogger.WriteDebugLog "判定アドレス: " & judgeAddress

    M06_DebugLogger.WriteDebugLog "旧ブックを開きます。"
    Set oldWb = M03_FileHandler.OpenWorkbook(oldBookPath)
    If oldWb Is Nothing Then
        M04_Logger.WriteError "[致命的エラー]", "-", "-", "旧ブックオープン失敗", "指定されたパスのファイルが開けません: " & oldBookPath
        GoTo FatalErrorHandler
    End If

    M06_DebugLogger.WriteDebugLog "旧ブックから種類を取得します。"
    bookType = GetValueFromOldBook(oldWb, judgeAddress)
    M06_DebugLogger.WriteDebugLog "取得した種類: " & bookType
    oldWb.Close SaveChanges:=False
    Set oldWb = Nothing

    M06_DebugLogger.WriteDebugLog "Listシートからテンプレートパスを検索します。"
    templatePath = GetTemplatePathFromList(bookType)
    If templatePath = "" Then
        M04_Logger.WriteError "[致命的エラー]", "-", "-", "テンプレート特定失敗", "種類 '" & bookType & "' に一致するテンプレートが見つかりません。"
        GoTo FatalErrorHandler
    End If
    M06_DebugLogger.WriteDebugLog "テンプレートパス: " & templatePath

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
    M06_DebugLogger.WriteDebugLog "致命的なエラーが発生しました。処理を中断します。"
    MsgBox "致命的なエラーが発生しました。処理を中断します。詳細はErrorシートを確認してください。"

Finally:
    Application.ScreenUpdating = True
    M06_DebugLogger.WriteDebugLog "画面描画を再開し、処理を終了します。"
End Sub

'--------------------------------------------------------------------------------------------------
' [Function] GetValueFromOldBook
' [Description] 旧ブックから指定されたアドレスの値を取得する
' [Args] wb: 対象ワークブック, address: セルアドレス
' [Returns] セルの値
'--------------------------------------------------------------------------------------------------
Private Function GetValueFromOldBook(ByVal wb As Workbook, ByVal address As String) As String
    Dim sheetName As String
    Dim rangeAddress As String
    Dim parts() As String

    M06_DebugLogger.WriteDebugLog "旧ブックの値取得を開始: " & address
    On Error GoTo GetValueError

    parts = Split(address, "!")
    sheetName = Replace(parts(0), "'", "")
    rangeAddress = parts(1)
    M06_DebugLogger.WriteDebugLog "シート名: " & sheetName & ", レンジ: " & rangeAddress

    GetValueFromOldBook = wb.Worksheets(sheetName).Range(rangeAddress).Value
    M06_DebugLogger.WriteDebugLog "値を取得しました: " & GetValueFromOldBook
    Exit Function

GetValueError:
    GetValueFromOldBook = ""
    M06_DebugLogger.WriteDebugLog "値の取得でエラーが発生しました。"
    M04_Logger.WriteError "[警告]", "-", "-", "値の取得失敗", "旧ブックの '" & address & "' から値を取得できませんでした。"
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