Option Explicit

'==================================================================================================
' [Module] M03_FileHandler
' [Description] ファイル操作関連モジュール
'==================================================================================================

'--------------------------------------------------------------------------------------------------
' [Function] OpenWorkbook
' [Description] 指定されたパスのワークブックを開く
' [Args] path: ファイルパス
' [Returns] Workbookオブジェクト (失敗時: Nothing)
'--------------------------------------------------------------------------------------------------
Public Function OpenWorkbook(ByVal path As String) As Workbook
    M06_DebugLogger.WriteDebugLog "ワークブックを開きます: " & path
    On Error Resume Next
    Set OpenWorkbook = Workbooks.Open(path)
    On Error GoTo 0
    If OpenWorkbook Is Nothing Then
        M06_DebugLogger.WriteDebugLog "ワークブックを開けませんでした。"
    End If
End Function

'--------------------------------------------------------------------------------------------------
' [Function] CreateNewBook
' [Description] テンプレートをコピーして新しいブックを作成する
' [Args] templatePath: テンプレートのフルパス
' [Returns] 新しいブックのフルパス
'--------------------------------------------------------------------------------------------------
Public Function CreateNewBook(ByVal templatePath As String) As String
    Dim fso As Object
    Dim newBookName As String
    Dim newBookPath As String
    Dim settingsSheet As Worksheet

    M06_DebugLogger.WriteDebugLog "新しいブックの作成を開始します。"
    M06_DebugLogger.WriteDebugLog "テンプレートパス: " & templatePath

    Set settingsSheet = ThisWorkbook.Worksheets("Settings")
    
    newBookName = settingsSheet.Range("D10").Value & "-" & _
                  settingsSheet.Range("D11").Value & " Ry-" & _
                  Format(settingsSheet.Range("D12").Value, "yyyy.mm.dd") & ".xls"
    M06_DebugLogger.WriteDebugLog "生成された新しいファイル名: " & newBookName

    newBookPath = ThisWorkbook.path & "\dist\" & newBookName
    M06_DebugLogger.WriteDebugLog "新しいブックの保存先パス: " & newBookPath

    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile templatePath, newBookPath, True
    Set fso = Nothing

    CreateNewBook = newBookPath
    M04_Logger.WriteLog "新ブック作成", "作成パス: " & newBookPath
    M06_DebugLogger.WriteDebugLog "新しいブックの作成が完了しました。"
End Function