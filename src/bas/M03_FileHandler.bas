Option Explicit

' テンプレートシートをコピーして新しいブックを作成する
' @param templateSheetName テンプレートシート名
' @return 新しいブック
'================================================================================================'
Public Function CreateWorkbookFromTemplate(ByVal templateSheetName As String) As Workbook
    DebugLog "M03_FileHandler", "CreateWorkbookFromTemplate", "Start - templateSheetName: '" & templateSheetName & "'"
    
<<<<<<< HEAD
    ' ファイル名の生成ロジックを修正 (Ryの重複を回避)
    newBookName = "_" & settingsSheet.Range("D17").Value & "-" & _
                  settingsSheet.Range("D18").Value & "-" & _
                  Format(settingsSheet.Range("D20").Value, "yyyy.mm.dd") & ".xls"
    M06_DebugLogger.WriteDebugLog "生成された新しいファイル名: " & newBookName

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
=======
    Dim templateSheet As Worksheet
    
    ' テンプレートシートを取得
>>>>>>> 4a358f29dfa0286bcc247843225878f0af2a21d0
    On Error Resume Next
    Set templateSheet = ThisWorkbook.Sheets(templateSheetName)
    On Error GoTo 0
    
    If templateSheet Is Nothing Then
        Log "エラー: テンプレートシート '" & templateSheetName & "' が見つかりません。"
        DebugLog "M03_FileHandler", "CreateWorkbookFromTemplate", "Error: Template sheet '" & templateSheetName & "' not found. Returning Nothing."
        Set CreateWorkbookFromTemplate = Nothing
        Exit Function
    End If
    
    DebugLog "M03_FileHandler", "CreateWorkbookFromTemplate", "Template sheet '" & templateSheetName & "' found. Copying to new workbook."
    
    ' テンプレートシートを新しいブックにコピー
    templateSheet.Copy
    
    ' 新しいブックを返す
    Set CreateWorkbookFromTemplate = ActiveWorkbook
    DebugLog "M03_FileHandler", "CreateWorkbookFromTemplate", "End - Workbook created successfully."
End Function

' ブックを保存する
' @param wb 保存するブック
' @param fileName ファイル名
'================================================================================================'
Public Sub SaveWorkbook(ByVal wb As Workbook, ByVal fileName As String)
    DebugLog "M03_FileHandler", "SaveWorkbook", "Start - fileName: '" & fileName & "'"
    
    Dim saveFolderPath As String
    saveFolderPath = ThisWorkbook.Path & "\dist"
    DebugLog "M03_FileHandler", "SaveWorkbook", "Save folder path: '" & saveFolderPath & "'"
    
    ' distフォルダが存在しない場合は作成
    If Dir(saveFolderPath, vbDirectory) = "" Then
        DebugLog "M03_FileHandler", "SaveWorkbook", "'dist' folder not found. Creating it."
        MkDir saveFolderPath
    End If
    
    Dim fullPath As String
    fullPath = saveFolderPath & "\" & fileName
    DebugLog "M03_FileHandler", "SaveWorkbook", "Saving workbook to: '" & fullPath & "'"
    
    ' ブックを保存
    On Error Resume Next
    wb.SaveAs fileName:=fullPath, FileFormat:=xlOpenXMLWorkbook
    If Err.Number <> 0 Then
        Log "エラー: ファイルの保存に失敗しました。 " & fullPath
        DebugLog "M03_FileHandler", "SaveWorkbook", "Error on SaveAs: " & Err.Description
        On Error GoTo 0
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    On Error GoTo 0
    
    ' ブックを閉じる
    wb.Close SaveChanges:=False
    DebugLog "M03_FileHandler", "SaveWorkbook", "Workbook saved and closed successfully."
    DebugLog "M03_FileHandler", "SaveWorkbook", "End"
End Sub