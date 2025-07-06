Option Explicit

' メインプロシージャ
'================================================================================================'
Sub Main()
    DebugLog "M01_Main", "Main", "Start"
    
    ' 初期化
    Initialize

    ' メイン処理
    DebugLog "M01_Main", "Main", "Calling ProcessReports"
    ProcessReports
    DebugLog "M01_Main", "Main", "Returned from ProcessReports"

    ' 終了処理
    Finalize
    
    DebugLog "M01_Main", "Main", "End"
End Sub

' 初期化プロシージャ
'================================================================================================'
Private Sub Initialize()
    DebugLog "M01_Main", "Initialize", "Start"
    
    ' ロガーの初期化
    InitializeLogger
    InitializeDebugLogger ' デバッグロガーの初期化を追加

    ' ログ開始
    Log "処理開始"
    
    ' 設定の読み込み
    DebugLog "M01_Main", "Initialize", "Calling LoadSettings"
    LoadSettings
    DebugLog "M01_Main", "Initialize", "Returned from LoadSettings"
    
    DebugLog "M01_Main", "Initialize", "End"
End Sub

' 終了処理プロシージャ
'================================================================================================'
Private Sub Finalize()
    DebugLog "M01_Main", "Finalize", "Start"
    
    ' ログ終了
    Log "処理終了"

    ' ロガーのクリーンアップ
    CleanupLogger
    CleanupDebugLogger ' デバッグロガーのクリーンアップを追加
    
    DebugLog "M01_Main", "Finalize", "End"
End Sub