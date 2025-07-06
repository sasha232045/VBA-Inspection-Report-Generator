---
title: 設計書 (Design Document)
---

# 設計書 (Design Document)

## 1. 概要

Excel VBA を使用して、検査報告書を自動生成するシステムの設計書。

## 2. システム構成

- **入力**:
  - `_input.xlsm`: メインの Excel ブック。
    - `List` シート: 報告書データのリスト。
    - `Settings` シート: 処理に必要な設定値。
    - `TemplateSheet` シート: 報告書のテンプレート。
- **出力**:
  - `dist/` フォルダ:
    - `InspectionReport_*.xlsx`: 生成された検査報告書。
    - `Log.txt`: 処理ログ。
    - `DebugLog.txt`: 詳細なデバッグログ。
- **ソースコード**:
  - `src/bas/`: VBA モジュール。

## 3. モジュール設計

| モジュール名        | 説明                               |
| ------------------- | ---------------------------------- |
| `M01_Main`          | メイン処理の制御                   |
| `M02_Processor`     | データ処理とレポート作成のコアロジック |
| `M03_FileHandler`   | ファイル・ブック・シート操作       |
| `M04_Logger`        | 通常ログの出力                     |
| `M05_Utility`       | 汎用的なユーティリティ関数         |
| `M06_DebugLogger`   | デバッグログの出力                 |

## 4. データフロー

1.  `M01_Main` が処理を開始。
2.  `M02_Processor` の `LoadSettings` が `Settings` シートから設定を読み込む。
3.  `M02_Processor` の `ProcessReports` が `List` シートのデータを一行ずつ処理。
4.  各行について、`M03_FileHandler` が `TemplateSheet` をコピーして新しいブックを作成。
5.  `M02_Processor` が新しいブックにデータを転記。
6.  `M03_FileHandler` が新しいブックを指定されたファイル名で `dist/` フォルダに保存。
7.  `M04_Logger` と `M06_DebugLogger` が随時ログを記録。

## 5. 詳細設計

### M01_Main

- `Main`: メインプロシージャ。`Initialize`, `ProcessReports`, `Finalize` を順に呼び出す。
- `Initialize`: ロガーの初期化、設定の読み込みを行う。
- `Finalize`: ロガーのクリーンアップを行う。

### M02_Processor

- `settings` (Private, Dictionary): 設定値を保持するグローバル変数。
- `LoadSettings`: `Settings` シートの C, D 列からキーと値を読み取り、`settings` に格納する。
- `ProcessReports`: `List` シートのデータに基づいてループ処理を行い、`CreateReport` を呼び出す。
- `CreateReport`: `M03_FileHandler` を使ってレポートブックを作成し、データを転記し、保存する。

### M03_FileHandler

- `CreateWorkbookFromTemplate`: `TemplateSheetName` を引数に取り、そのシートをコピーして新しいワークブックオブジェクトを返す。
- `SaveWorkbook`: ワークブックオブジェクトとファイル名を引数に取り、`dist/` フォルダに `.xlsx` 形式で保存する。

### M04_Logger

- `Log`: 文字列を引数に取り、タイムスタンプと共に `Log.txt` に追記する。

### M05_Utility

- `GetLastRow`: ワークシートと列番号を引数に取り、最終行の行番号を返す。

### M06_DebugLogger

- `DebugLog`: モジュール名、プロシージャ名、メッセージを引数に取り、詳細なデバッグ情報を `DebugLog.txt` に追記する。

## ログ仕様

本システムは、処理の進行状況やエラーを追跡するために2種類のログを出力する。

### 1. 通常ログ (Log.txt)

- **目的**: ユーザーが処理の主要な流れを把握するため。
- **出力内容**: 処理の開始・終了、生成されたファイル名など、主要なイベント。
- **出力場所**: `dist` フォルダ内の `Log.txt`。
- **実装**: `M04_Logger` モジュールの `Log` プロシージャを使用する。

### 2. デバッグログ (DebugLog.txt)

- **目的**: 開発者が詳細な処理内容を確認し、問題解決（デバッグ）に役立てるため。
- **出力内容**:
    - 各プロシージャ・関数の開始と終了。
    - 受け渡された引数の値。
    - ループ内のカウンターや重要な変数の値。
    - ファイル操作やシート操作の直前のパラメータ。
    - 設定ファイルから読み込んだキーと値のペア。
- **出力場所**: `dist` フォルダ内の `DebugLog.txt`。
- **実装**: `M06_DebugLogger` モジュールの `DebugLog` プロシージャを使用する。
- **フォーマット**: `[Debug] - <モジュール名> - <プロシージャ名> - <メッセージ>`

---

## 著作権と免責事項

- **著作権**: このVBAプログラムの著作権は作成者に帰属します。
- **免責事項**: このプログラムの使用によって生じたいかなる損害についても、作成者は責任を負いません。自己責任でご使用ください。