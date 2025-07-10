Option Explicit

' 指定されたシートと列の最終行を取得する
' @param ws 対象のワークシート
' @param column 最終行を取得する列番号
' @return 最終行番号
'================================================================================================'
Public Function GetLastRow(ByVal ws As Worksheet, ByVal column As Long) As Long
    M06_DebugLogger.WriteDebugLog "GetLastRow - Sheet: '" & ws.Name & "', Column: " & column
    GetLastRow = ws.Cells(ws.Rows.Count, column).End(xlUp).Row
    M06_DebugLogger.WriteDebugLog "GetLastRow - Last row is " & GetLastRow
End Function

'--------------------------------------------------------------------------------------------------
' [Public Function] GetRangeFromAddressString
' [Description] アドレス文字列から対象のRangeオブジェクトを取得する
' [Args] ws: 対象ワークシート, addressStr: アドレス文字列 (例: "A1", "A1:C10")
' [Returns] Rangeオブジェクト (失敗時: Nothing)
'--------------------------------------------------------------------------------------------------
Public Function GetRangeFromAddressString(ByVal ws As Worksheet, ByVal addressStr As String) As Range
    On Error Resume Next
    Set GetRangeFromAddressString = ws.Range(addressStr)
    If Err.Number <> 0 Then
        M06_DebugLogger.WriteDebugLog "アドレス文字列の解析に失敗しました: " & addressStr & ", エラー: " & Err.description
        Set GetRangeFromAddressString = Nothing
    Else
        M06_DebugLogger.WriteDebugLog "アドレス文字列を正常に解析しました: " & addressStr
    End If
    On Error GoTo 0
End Function
