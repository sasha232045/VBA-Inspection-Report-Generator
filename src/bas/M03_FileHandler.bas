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
    Dim saveFolderPath As String

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

    fso.CopyFile templatePath, newBookPath, True
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