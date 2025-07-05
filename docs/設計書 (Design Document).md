
### **【成果物3/3】 設計書 (Design Document)**

これは、仕様書で定義された要件を、VBAのコードやExcelの機能レベルでどのように実現するかを具体的に記述した技術ドキュメントです。

---

**文書管理**

| ドキュメント名 | 点検報告書 旧データ移行ツール 設計書 |
| :--- | :--- |
| バージョン | 1.0 |
| 作成日 | 2025年7月5日 |
| 作成者 | (あなたのためのVBAプロフェッショナル) |
| 承認者 | (お客様名) |

---

### **1. 概要**

本設計書は、「点検報告書 旧データ移行ツール 仕様書 Ver.1.0」で定義された要件を実現するための、技術的な実装方法を定義するものである。VBAのモジュール構成、各機能の実装ロジック、Excelの数式設計、およびエラーハンドリングの具体的な手法について詳述する。

### **2. システムアーキテクチャ**

#### **2.1. Excelブックとシート構成**

本ツール（以下、マクロブック）は、以下のシートで構成される。

-   **`Control`**: ユーザーインターフェースシート。マクロ実行ボタンを配置する。
-   **`Settings`**: 処理設定シート。ユーザーが処理内容を定義する。数式による自動計算エリアと、手動設定エリアに分かれる。
-   **`List`**: テンプレートファイルのマスタリストシート。
-   **`Log`**: 処理ログ出力シート。`[日時] | [レベル] | [モジュール.プロシージャ] | [メッセージ]` の形式で記録する。
-   **`Error`**: エラーログ出力シート。`[日時] | [モジュール.プロシージャ] | [影響範囲] | [エラーNo] | [エラー内容]` の形式で記録する。

#### **2.2. VBAプロジェクト構成**

VBAコードは機能ごとにモジュールを分割し、保守性と可読性を高める。

-   **`M_Main` (標準モジュール)**: メインコントローラー。処理全体の流れを制御する。`Control`シートのボタンから呼び出される。
-   **`M_Processor` (標準モジュール)**: データ処理のコアロジックを実装するモジュール。`Settings`シートの処理詳細設定を解釈し、コピー、削除、入力の各機能を実行する。
-   **`M_FileHandler` (標準モジュール)**: ファイル操作（旧ブックを開く、新ブックを作成・保存する等）に関する機能を実装する。
-   **`M_SheetReader` (標準モジュール)**: `Settings`シートや`List`シートから設定値を読み取る機能を実装する。
-   **`M_Logger` (標準モジュール)**: `Log`シートおよび`Error`シートへの書き込み処理を実装する。
-   **`M_Utility` (標準モジュール)**: 汎用的な便利機能（カンマ区切りのアドレス文字列をRangeオブジェクトに変換する機能など）を実装する。

### **3. シート設計と数式**

#### **3.1. `Settings`シートのレイアウトと数式**

| | A | B | C | D |
|---|---|---|---|---|
| **1** | **全体設定** | | | |
| **2** | 処理対象 | 旧ブック_ファイルパス | | `(ユーザーが入力)` |
| **3** | | 旧ブック_新ブック名判定アドレス | | `(ユーザーが入力)` |
| **4** | ↓ | | | |
| **5** | 判断結果 | 旧ブック_ファイル名 | | `=IFERROR(TEXTAFTER(D2,"\",-1),"")` |
| **6** | | 単価契約明細ＮＯ | | `=IFERROR(TEXTBEFORE(D5,"神戸港変電所"),"")` |
| **7** | | 変電所- | | `=IFERROR(LET(s,D5, t,"変電所-", start,FIND(t,s)+LEN(t), end,FIND("TrB Ry-",s), LEFT(MID(s,start,99),end-start)),"")` |
| **8** | | TrB Ry- | | `=IFERROR(LET(s,D5, t,"TrB Ry-", start,FIND(t,s)+LEN(t), end,FIND("-",s,start), LEFT(MID(s,start,99),end-start)),"")` |
| **9** | | 日付 | | `=IFERROR(TEXTBEFORE(TEXTAFTER(D5,"-",-1),".xls"),"")` |
| **10**| | 種類 | | `=IFERROR(INDIRECT(D3,TRUE),"要旧ブックオープン")` |
| **11**| | 新ブック_ファイルパス | | `=IFERROR(XLOOKUP(D10,List!F:F,List!G:G,"該当なし"),"")`|
| **12**| | 本部 | | `=IFERROR(XLOOKUP(D7,List!C:C,List!B:B,"該当なし",0),"")`|
| **13**| | 管内 | | `=IFERROR(XLOOKUP(D7,List!C:C,List!C:C,"該当なし",0),"")`|
| **14**| | | | |
| **15**| **処理シートNo** | **処理No.** | **処理内容** | **内容** |
| **16**| 1 | 1 | 旧ブック_シート名 | 統一表紙 |
| **...**|...|...|...|...|

*   **数式に関する注釈:**
    *   `D5`の数式は、`D2`のフルパスからファイル名のみを抽出します。
    *   `D6`～`D9`の数式は、`D5`のファイル名を特定のキーワード（例：「変電所-」）を区切り文字として分割し、必要な情報を抽出します。（注: 上記はご提示のファイル名例に基づく一例です。命名規則の変動に強い、より汎用的な`TEXTSPLIT`関数等への変更も可能です。）
    *   `D10`の`INDIRECT`関数は、旧ブックが開かれている状態でのみ正しく値を参照します。マクロ実行時にVBAで値を取得するため、このセルはユーザーの確認用と位置付けます。
    *   `D11`の`XLOOKUP`関数は、`D10`で特定した「種類」をキーに`List`シートのF列を検索し、対応するG列の「ファイルパス」（＝テンプレートのパス）を取得します。

### **4. VBAモジュール詳細設計**

#### **4.1. `M_Main`**
-   **`Public Sub StartProcess()`**
    1.  処理開始をユーザーに通知 (`Application.StatusBar`, `MsgBox`)
    2.  `M_Logger.InitializeLogs` を呼び出し、Log/Errorシートを初期化。
    3.  `M_Logger.WriteLog "処理開始"`
    4.  `On Error GoTo ErrorHandler` で大域的なエラーハンドラを設定。
    5.  **設定読み込みフェーズ:**
        -   `M_SheetReader.GetOldBookPath()` で旧ブックパス(D2)を取得。
        -   `M_SheetReader.GetJudgeAddress()` で判定アドレス(D3)を取得。
    6.  **新ブック準備フェーズ:**
        -   `M_FileHandler.OpenOldBook()` で旧ブックを開く。
        -   `M_SheetReader.GetBookType()` で旧ブックの判定アドレスから「種類」を取得。
        -   `M_SheetReader.GetTemplatePath()` で`List`シートからテンプレートパスを取得。
        -   `M_FileHandler.CreateNewBookFromTemplate()` でテンプレートをコピーし、新ブックを作成・保存。
        -   旧ブックを閉じる (`M_FileHandler.CloseOldBook`)。
    7.  **データ移行フェーズ:**
        -   `M_Processor.ExecuteAllTasks` を呼び出し、新旧ブックオブジェクトを渡して全処理を実行。
    8.  **終了処理:**
        -   新ブックを保存して閉じる (`M_FileHandler.CloseNewBook`)。
        -   `M_Logger.WriteLog "処理正常終了"`
        -   `MsgBox`で完了通知。`Error`シートに書き込みがあればその旨も通知。
    9.  `Exit Sub`
    -   `ErrorHandler:`
        -   致命的なエラー（旧ブックが開けない等）の場合、`M_Logger.WriteError`で記録し、`MsgBox`で処理中断を通知。

#### **4.2. `M_Processor`**
-   **`Public Sub ExecuteAllTasks(oldWb As Workbook, newWb As Workbook)`**
    1.  `Settings`シートの16行目から最終行までループ。
    2.  各行で `処理シートNo` をキーに、処理をグループ化する。（`Dictionary`オブジェクト等を利用）
    3.  グループ化した処理`Dictionary`をループ。
    4.  ループ内で`On Error GoTo TaskErrorHandler`を設定。これによりタスク単位でエラーを捕捉し、処理を継続する。
    5.  `処理内容`列の値で`Select Case`分岐。
        -   `Case "旧ブック_シート名":` 変数に格納
        -   `Case "新ブック_シート名":` 変数に格納
        -   `Case "旧ブック_コピー元アドレス", "新ブック_コピー先アドレス":` 変数に格納し、ペアが揃ったら`CopyData`を呼び出す。
        -   `Case "新ブック_削除するアドレス":` `DeleteData`を呼び出す。
        -   `Case "新ブック_入力するアドレス", "新ブック_入力する内容":` 変数に格納し、ペアが揃ったら`InputData`を呼び出す。
    6.  `ContinueNextTask:` `Resume Next`の飛び先。
    7.  `Exit Sub`
    -   `TaskErrorHandler:`
        -   `M_Logger.WriteError`を呼び出し、エラー内容を記録。
        -   `Resume ContinueNextTask`で次の処理へ移行。

-   **`Private Sub CopyData(...)`**: `oldWb.Sheets(oldSheet).Range(srcAddr).Copy newWb.Sheets(newSheet).Range(dstAddr)` を実行。結合セルも考慮される。
-   **`Private Sub DeleteData(...)`**: `M_Utility.GetRangeFromAddressString`を呼び出して対象Rangeオブジェクトを取得後、`targetRange.MergeArea.ClearContents`を実行。
-   **`Private Sub InputData(...)`**: `M_Utility.GetRangeFromAddressString`を呼び出して対象Rangeオブジェクトを取得後、`targetRange.Value = valueToInput`を実行。

#### **4.3. `M_Utility`**
-   **`Public Function GetRangeFromAddressString(sht As Worksheet, ByVal addressString As String) As Range`**
    1.  `Union`メソッドを使って複数の範囲を結合するための`Range`オブジェクトを宣言 (`Dim combinedRange As Range`)。
    2.  `addressString`をカンマ（`,`）で分割し、配列にする (`Split`関数)。
    3.  配列をループ。
    4.  各アドレス文字列に対し、`sht.Range(addr)`で`Range`オブジェクトを取得。
    5.  `combinedRange`が`Nothing`なら、最初のRangeオブジェクトをセット。`Nothing`でなければ、`Set combinedRange = Union(combinedRange, currentRange)`で既存の範囲に結合する。
    6.  ループ終了後、結合された`combinedRange`を関数の戻り値として返す。

#### **4.4. `M_Logger`**
-   **`Public Sub InitializeLogs()`**: `Log`と`Error`シートの既存の内容をクリアし、ヘッダー行を書き込む。
-   **`Public Sub WriteLog(message As String, Optional procName As String = "")`**: `Log`シートの最終行に、タイムスタンプ、レベル(INFO)、プロシージャ名、メッセージを書き込む。
-   **`Public Sub WriteError(procName As String, scope As String, errNum As Long, errDesc As String)`**: `Error`シートの最終行に、タイムスタンプ、プロシージャ名、影響範囲、エラー番号、エラー内容を書き込む。

---

以上が、ご要望いただいた全ての要件を反映した設計書です。
この設計に基づき、VBAコードの実装に着手いたします。ご提示いただいた仕様と設計に沿って、堅牢かつメンテナンス性の高いツールを構築することをお約束します。

もし、この内容で実装を進めてよろしければ、その旨をお知らせください。直ちにコーディングを開始いたします。