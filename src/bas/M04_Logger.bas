Option Explicit

Private Const LOG_FILE_NAME As String = "Log.txt"
Private fso As Object
Private logStream As Object

' ロガーを初期化する
'================================================================================================'
Public Sub InitializeLogger()
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim logFolderPath As String
    logFolderPath = ThisWorkbook.Path & "\dist"
    
    If Not fso.FolderExists(logFolderPath) Then
        fso.CreateFolder logFolderPath
    End If
    
    Dim logFilePath As String
    logFilePath = logFolderPath & "\" & LOG_FILE_NAME
    
    ' ログファイルを追記モードで開く
    Set logStream = fso.OpenTextFile(logFilePath, 8, True) ' 8 = ForAppending, True = CreateIfNeeded
End Sub

' ロガーをクリーンアップする
'================================================================================================'
Public Sub CleanupLogger()
    If Not logStream Is Nothing Then
        logStream.Close
    End If
    Set logStream = Nothing
    Set fso = Nothing
End Sub

' ログを出力する
' @param message ログメッセージ
'================================================================================================'
Public Sub Log(ByVal message As String)
    If logStream Is Nothing Then
        Exit Sub
    End If
    
    logStream.WriteLine Now() & " - " & message
End Sub