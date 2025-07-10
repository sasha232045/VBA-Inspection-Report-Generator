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
    Dim originalDisplayAlerts As Boolean
    Dim originalAskToUpdateLinks As Boolean
    Dim originalEnableEvents As Boolean
    
    M06_DebugLogger.WriteDebugLog "ワークブックを開きます: " & path
    
    ' 現在の設定を保存
    originalDisplayAlerts = Application.DisplayAlerts
    originalAskToUpdateLinks = Application.AskToUpdateLinks
    originalEnableEvents = Application.EnableEvents
    
    ' ダイアログを無効化
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.EnableEvents = False
    
    On Error Resume Next
    Set OpenWorkbook = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=False, _
                                      IgnoreReadOnlyRecommended:=True, _
                                      Notify:=False, _
                                      AddToMru:=False)
    On Error GoTo 0
    
    ' 設定を復元
    Application.DisplayAlerts = originalDisplayAlerts
    Application.AskToUpdateLinks = originalAskToUpdateLinks
    Application.EnableEvents = originalEnableEvents
    
    If OpenWorkbook Is Nothing Then
        M06_DebugLogger.WriteDebugLog "ワークブックを開けませんでした。"
    Else
        M06_DebugLogger.WriteDebugLog "ワークブックを正常に開きました。"
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
    Dim saveFolderPath As String
    Dim originalDisplayAlerts As Boolean

    M06_DebugLogger.WriteDebugLog "新しいブックの作成を開始します。"
    M06_DebugLogger.WriteDebugLog "テンプレートパス: " & templatePath

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set settingsSheet = ThisWorkbook.Worksheets("Settings")
    
    ' ファイル名の生成ロジックを修正 (D11から直接取得)
    newBookName = Trim(settingsSheet.Range("D11").Value)
    If newBookName = "" Then
        ' D11が空の場合は従来のロジックを使用
        newBookName = "_" & settingsSheet.Range("D17").Value & "-" & _
                      settingsSheet.Range("D18").Value & "-" & _
                      Format(settingsSheet.Range("D20").Value, "yyyy.mm.dd") & ".xls"
        M06_DebugLogger.WriteDebugLog "D11が空のため、自動生成されたファイル名を使用: " & newBookName
    Else
        M06_DebugLogger.WriteDebugLog "D11から取得したファイル名を使用: " & newBookName
    End If

    ' 保存先フォルダの決定
    saveFolderPath = Trim(settingsSheet.Range("D9").Value)
    If saveFolderPath = "" Then
        saveFolderPath = ThisWorkbook.path ' D9が空白の場合はVBA実行ファイルと同じフォルダ
        M06_DebugLogger.WriteDebugLog "SettingsシートD9が空白のため、VBA実行ファイルと同じフォルダを保存先とします: " & saveFolderPath
    Else
        M06_DebugLogger.WriteDebugLog "SettingsシートD9から保存先フォルダを取得しました: " & saveFolderPath
    End If

    ' 安全なパス結合
    newBookPath = fso.BuildPath(saveFolderPath, newBookName)
    M06_DebugLogger.WriteDebugLog "新しいブックの保存先パス: " & newBookPath

    ' 保存先フォルダが存在しない場合は作成
    If Not fso.FolderExists(saveFolderPath) Then
        M06_DebugLogger.WriteDebugLog "保存先フォルダが存在しないため作成します: " & saveFolderPath
        fso.CreateFolder saveFolderPath
    End If

    ' ダイアログを無効化して強制コピー
    originalDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    On Error Resume Next
    fso.CopyFile templatePath, newBookPath, True ' True = 上書き許可
    If Err.Number <> 0 Then
        M06_DebugLogger.WriteDebugLog "ファイルコピー中にエラーが発生しました: " & Err.description
        M04_Logger.WriteError "[エラー]", "-", "-", "ファイルコピー失敗", "テンプレートファイルのコピーに失敗しました。エラー: " & Err.description
    End If
    On Error GoTo 0
    
    Application.DisplayAlerts = originalDisplayAlerts
    
    Set fso = Nothing

    CreateNewBook = newBookPath
    M04_Logger.WriteLog "新ブック作成", "作成パス: " & newBookPath
    M06_DebugLogger.WriteDebugLog "新しいブックの作成が完了しました。"
End Function

'--------------------------------------------------------------------------------------------------
' [Function] FileExists
' [Description] 指定されたパスのファイルが存在するか確認する
' [Args] path: ファイルパス
' [Returns] 存在する場合 True, しない場合 False
'--------------------------------------------------------------------------------------------------
Public Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(path) <> "")
    On Error GoTo 0
End Function

'--------------------------------------------------------------------------------------------------
' [Public Function] ParseAddress
' [Description] 正規表現を使用して、'シート名'!セル範囲 の形式のアドレス文字列を解析する
' [Args] address: 解析対象のアドレス文字列
' [Args] outSheetName: (出力) 抽出されたシート名
' [Args] outRangeAddress: (出力) 抽出されたセル範囲
' [Returns] 成功した場合 True, 失敗した場合 False
'--------------------------------------------------------------------------------------------------
Public Function ParseAddress(ByVal address As String, ByRef outSheetName As String, ByRef outRangeAddress As String) As Boolean
    Dim regex As Object
    Dim matches As Object

    M06_DebugLogger.WriteDebugLog "アドレス解析を開始: " & address
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^'?(.*?)'?!([^!]+)$" ' シート名とセル範囲を抽出する正規表現

    If regex.Test(address) Then
        Set matches = regex.Execute(address)
        outSheetName = matches(0).SubMatches(0)
        outRangeAddress = matches(0).SubMatches(1)
        ParseAddress = True
        M06_DebugLogger.WriteDebugLog "解析成功: シート名='" & outSheetName & "', アドレス='" & outRangeAddress & "'"
    Else
        ParseAddress = False
        M06_DebugLogger.WriteDebugLog "解析失敗: 有効なアドレス形式ではありません。"
        M04_Logger.WriteError "[警告]", "-", "-", "アドレス解析失敗", "'" & address & "' は有効な形式ではありません。"
    End If
End Function

'--------------------------------------------------------------------------------------------------
' [Public Sub] OpenFileLocation
' [Description] 指定されたファイルパスの保存フォルダをエクスプローラーで開く（ファイル選択状態）
' [Args] filePath: ファイルのフルパス
'--------------------------------------------------------------------------------------------------
Public Sub OpenFileLocation(ByVal filePath As String)
    Dim fso As Object
    Dim folderPath As String
    Dim fileName As String
    
    On Error GoTo OpenFolderError
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' ファイルが存在するかチェック
    If fso.FileExists(filePath) Then
        ' ファイルを選択状態でエクスプローラーを開く
        Shell "explorer.exe /select,""" & filePath & """", vbNormalFocus
        M06_DebugLogger.WriteDebugLog "ファイルを選択状態でエクスプローラーを開きました: " & filePath
    Else
        ' ファイルが存在しない場合はフォルダのみ開く
        folderPath = fso.GetParentFolderName(filePath)
        If fso.FolderExists(folderPath) Then
            Shell "explorer.exe """ & folderPath & """", vbNormalFocus
            M06_DebugLogger.WriteDebugLog "保存フォルダを開きました: " & folderPath
        Else
            MsgBox "保存フォルダが見つかりません: " & folderPath, vbExclamation, "フォルダエラー"
            M06_DebugLogger.WriteDebugLog "保存フォルダが見つかりません: " & folderPath
        End If
    End If
    
    Set fso = Nothing
    Exit Sub
    
OpenFolderError:
    MsgBox "フォルダを開くことができませんでした。" & vbCrLf & "エラー: " & Err.description, vbCritical, "エラー"
    M06_DebugLogger.WriteDebugLog "フォルダを開く際にエラーが発生しました: " & Err.description
    Set fso = Nothing
End Sub

