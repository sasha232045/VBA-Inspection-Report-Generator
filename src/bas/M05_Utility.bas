Option Explicit

' 指定されたシートと列の最終行を取得する
' @param ws 対象のワークシート
' @param column 最終行を取得する列番号
' @return 最終行番号
'================================================================================================'
Public Function GetLastRow(ByVal ws As Worksheet, ByVal column As Long) As Long
    DebugLog "M05_Utility", "GetLastRow", "Start - Sheet: '" & ws.Name & "', Column: " & column
    GetLastRow = ws.Cells(ws.Rows.Count, column).End(xlUp).row
    DebugLog "M05_Utility", "GetLastRow", "End - Last row is " & GetLastRow
End Function