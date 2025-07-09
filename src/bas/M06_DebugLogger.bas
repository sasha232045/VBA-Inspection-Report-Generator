Option Explicit

Private Const DEBUG_LOG_FILE_NAME As String = "DebugLog.txt"
Private fsoDebug As Object
Private logStreamDebug As Object

' デバッグロガーを初期化する
'================================================================================================'
Public Sub InitializeDebugLogger()
    Set fsoDebug = CreateObject("Scripting.FileSystemObject")
    Dim logFolderPath As String
    logFolderPath = ThisWorkbook.Path & "\dist"
    
    If Not fsoDebug.FolderExists(logFolderPath) Then
        fsoDebug.CreateFolder logFolderPath
    End If
    
    Dim logFilePath As String
    logFilePath = logFolderPath & "\" & DEBUG_LOG_FILE_NAME
    
    ' ログファイルを新規作成（上書き）
    Set logStreamDebug = fsoDebug.CreateTextFile(logFilePath, True)
End Sub

' デバッグロガーをクリーンアップする
'================================================================================================'
Public Sub CleanupDebugLogger()
    If Not logStreamDebug Is Nothing Then
        logStreamDebug.Close
    End If
    Set logStreamDebug = Nothing
    Set fsoDebug = Nothing
End Sub

' デバッグログを出力する
' @param moduleName モジュール名
' @param procedureName プロシージャ名
' @param message ログメッセージ
'================================================================================================'
Public Sub DebugLog(ByVal moduleName As String, ByVal procedureName As String, ByVal message As String)
    If logStreamDebug Is Nothing Then
        ' 初期化されていない場合は何もしない（またはエラーを発生させる）
        Exit Sub
    End If
    
    logStreamDebug.WriteLine Now() & " [Debug] - " & moduleName & " - " & procedureName & " - " & message
End Sub
