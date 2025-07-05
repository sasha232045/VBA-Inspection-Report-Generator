Option Explicit

'==================================================================================================
' [Module] M05_Utility
' [Description] 汎用ユーティリティモジュール
'==================================================================================================

'--------------------------------------------------------------------------------------------------
' [Function] GetRangeFromAddressString
' [Description] カンマ区切りのアドレス文字列をRangeオブジェクトに変換する
' [Args] sht: 対象シート, addressString: アドレス文字列
' [Returns] 複数の範囲を含む単一のRangeオブジェクト
'--------------------------------------------------------------------------------------------------
Public Function GetRangeFromAddressString(ByVal sht As Worksheet, ByVal addressString As String) As Range
    Dim combinedRange As Range
    Dim addresses() As String
    Dim addr As Variant

    M06_DebugLogger.WriteDebugLog "アドレス文字列をRangeオブジェクトに変換します: " & addressString
    On Error GoTo ErrorHandler

    addresses = Split(addressString, ",")

    For Each addr In addresses
        If combinedRange Is Nothing Then
            Set combinedRange = sht.Range(Trim(addr))
        Else
            Set combinedRange = Union(combinedRange, sht.Range(Trim(addr)))
        End If
    Next addr

    Set GetRangeFromAddressString = combinedRange
    M06_DebugLogger.WriteDebugLog "Rangeオブジェクトの変換に成功しました。"
    Exit Function

ErrorHandler:
    Set GetRangeFromAddressString = Nothing
    M06_DebugLogger.WriteDebugLog "アドレス文字列の変換中にエラーが発生しました: " & addressString
    M04_Logger.WriteError "[警告]", "-", "-", "アドレス変換エラー", "'" & addressString & "' は有効なアドレスではありません。"
End Function