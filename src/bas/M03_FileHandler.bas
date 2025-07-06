Option Explicit

' テンプレートシートをコピーして新しいブックを作成する
' @param templateSheetName テンプレートシート名
' @return 新しいブック
'================================================================================================'
Public Function CreateWorkbookFromTemplate(ByVal templateSheetName As String) As Workbook
    DebugLog "M03_FileHandler", "CreateWorkbookFromTemplate", "Start - templateSheetName: '" & templateSheetName & "'"
    
    Dim templateSheet As Worksheet
    
    ' テンプレートシートを取得
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