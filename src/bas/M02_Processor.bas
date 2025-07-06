Option Explicit

' グローバル変数
'================================================================================================'
Private settings As Object ' Scripting.Dictionary

' メイン処理
'================================================================================================'
Public Sub ProcessReports()
    DebugLog "M02_Processor", "ProcessReports", "Start"
    
    ' 変数宣言
    Dim dataSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim templateSheetName As String
    Dim outputSheetName As String
    
    ' 初期設定
    Set dataSheet = ThisWorkbook.Sheets("List")
    lastRow = GetLastRow(dataSheet, 1) ' A列を基準に行数を取得
    DebugLog "M02_Processor", "ProcessReports", "Data sheet 'List' last row: " & lastRow
    
    ' 設定を辞書から取得
    templateSheetName = settings("TemplateSheetName")
    outputSheetName = settings("OutputSheetName")

    ' デバッグログ
    DebugLog "M02_Processor", "ProcessReports", "TemplateSheetName from settings: '" & templateSheetName & "'"
    DebugLog "M02_Processor", "ProcessReports", "OutputSheetName from settings: '" & outputSheetName & "'"

    ' データ行をループしてレポートを作成
    For i = 2 To lastRow
        DebugLog "M02_Processor", "ProcessReports", "Processing row: " & i
        ' レポート作成
        CreateReport dataSheet, i, templateSheetName, outputSheetName
    Next i
    
    DebugLog "M02_Processor", "ProcessReports", "End"
End Sub

' 設定シートから設定を読み込み、グローバルな辞書に格納する
'================================================================================================'
Public Sub LoadSettings()
    DebugLog "M02_Processor", "LoadSettings", "Start"
    
    Set settings = CreateObject("Scripting.Dictionary")
    Dim settingsSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim key As String
    Dim Value As String
    
    ' "Settings"シートを取得
    On Error Resume Next
    Set settingsSheet = ThisWorkbook.Sheets("Settings")
    On Error GoTo 0
    
    If settingsSheet Is Nothing Then
        Log "エラー: Settingsシートが見つかりません。"
        DebugLog "M02_Processor", "LoadSettings", "Error: 'Settings' sheet not found. Exiting sub."
        Exit Sub
    End If
    DebugLog "M02_Processor", "LoadSettings", "'Settings' sheet found."
    
    ' C列を基準に最終行を取得
    lastRow = GetLastRow(settingsSheet, 3)
    DebugLog "M02_Processor", "LoadSettings", "Settings sheet last row (Column C): " & lastRow
    
    ' C列とD列から設定を読み込む
    For i = 2 To lastRow
        key = settingsSheet.Cells(i, 3).Value
        Value = settingsSheet.Cells(i, 4).Value
        If key <> "" Then
            settings(key) = Value
            DebugLog "M02_Processor", "LoadSettings", "Loaded setting: Key='" & key & "', Value='" & Value & "'"
        End If
    Next i
    
    DebugLog "M02_Processor", "LoadSettings", "End"
End Sub

' レポート作成
' @param dataSheet データシート
' @param row 行番号
' @param templateSheetName テンプレートシート名
' @param outputSheetName 出力シート名
'================================================================================================'
Private Sub CreateReport(ByVal dataSheet As Worksheet, ByVal row As Long, ByVal templateSheetName As String, ByVal outputSheetName As String)
    DebugLog "M02_Processor", "CreateReport", "Start for row: " & row
    
    ' 変数宣言
    Dim outputBook As Workbook
    Dim outputSheet As Worksheet
    Dim fileName As String
    
    ' 出力用の新しいブックを作成
    DebugLog "M02_Processor", "CreateReport", "Calling CreateWorkbookFromTemplate with TemplateSheetName: '" & templateSheetName & "'"
    Set outputBook = CreateWorkbookFromTemplate(templateSheetName)
    
    If outputBook Is Nothing Then
        DebugLog "M02_Processor", "CreateReport", "Failed to create workbook from template. Exiting sub."
        Exit Sub
    End If
    
    Set outputSheet = outputBook.Sheets(1)
    outputSheet.Name = outputSheetName
    DebugLog "M02_Processor", "CreateReport", "Copied template sheet and renamed to '" & outputSheetName & "'"
    
    ' データを転記
    DebugLog "M02_Processor", "CreateReport", "Transferring data from source row " & row
    outputSheet.Range("C4").Value = dataSheet.Cells(row, 1).Value ' No.
    outputSheet.Range("C5").Value = dataSheet.Cells(row, 2).Value ' Date
    outputSheet.Range("C6").Value = dataSheet.Cells(row, 3).Value ' Title
    outputSheet.Range("C7").Value = dataSheet.Cells(row, 4).Value ' Inspector
    DebugLog "M02_Processor", "CreateReport", "Data transfer complete"
    
    ' ファイル名を生成
    fileName = "InspectionReport_" & dataSheet.Cells(row, 1).Value & ".xlsx"
    DebugLog "M02_Processor", "CreateReport", "Generated filename: '" & fileName & "'"
    
    ' ブックを保存
    DebugLog "M02_Processor", "CreateReport", "Calling SaveWorkbook"
    SaveWorkbook outputBook, fileName
    
    ' ログ出力
    Log "レポートを作成しました: " & fileName
    
    DebugLog "M02_Processor", "CreateReport", "End for row: " & row
End Sub