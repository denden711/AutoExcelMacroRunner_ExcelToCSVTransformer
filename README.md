# AutoExcelMacroRunner_ExcelToCSVTransformer
 
---

# AutoExcelMacroRunner

AutoExcelMacroRunnerは、指定されたExcelファイル内のマクロを自動的に実行し、プロセスが完了した後にExcelを閉じるPythonスクリプトです。このツールは、特に繰り返し実行が必要なExcelマクロの自動化に適しています。

## 機能

- 指定したExcelファイルを開く
- 指定したマクロを自動実行
- マクロ実行後、ファイルを保存しExcelを自動で閉じる
- 実行過程をログファイルに記録

## 前提条件

このプログラムを使用するには、以下が必要です：

- Python 3.x
- Microsoft Excel
- Pythonライブラリ `win32com.client` (pywin32)

## セットアップ

1. **Pythonのインストール**: [Python公式サイト](https://www.python.org/downloads/)からPythonをダウンロードしてインストールします。

2. **必要なパッケージのインストール**: コマンドラインまたはターミナルを開き、以下のコマンドを実行して必要なパッケージをインストールします。

    ```bash
    pip install pywin32
    ```

3. **プログラムのダウンロード**: このリポジトリをクローンまたはダウンロードし、スクリプトが含まれるディレクトリに移動します。

## 使用方法

1. `ExcelToCSVTransformer.xlsm`という名前のExcelファイルにマクロを用意します。
2. スクリプトが置かれているディレクトリにExcelファイルを配置します。
3. コマンドラインまたはターミナルで、以下のコマンドを実行してスクリプトを起動します。

    ```bash
    python AutoExcelMacroRunner.py
    ```

4. プログラムが自動的にExcelファイルを開き、マクロを実行後、Excelを閉じます。

## ログ

- プログラムの実行中に発生するすべての重要なイベントは、同じディレクトリに生成される `ExcelToCSVTransformer.log` ファイルに記録されます。
- エラーや異常が発生した場合は、このログファイルを確認して問題の診断に役立ててください。

## トラブルシューティング

- **Excelファイルが見つからない場合**: Excelファイルがスクリプトと同じディレクトリにあることを確認してください。
- **マクロが実行されない場合**: マクロの設定が正しく配置されているか、またマクロのセキュリティ設定が適切であるかを確認してください。
- **依存関係エラーが発生する場合**: Pythonとpywin32が正しくインストールされているか再確認してください。

## サポート

プログラムに関するさらなる質問やサポートが必要な場合は、[GitHub Issues](#)で質問を投稿してください。

---
