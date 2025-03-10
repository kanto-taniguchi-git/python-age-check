# Python Age Check

## 目次

- [Python Age Check](#python-age-check)
  - [目次](#目次)
  - [概要](#概要)
  - [特徴](#特徴)
  - [動作環境](#動作環境)
    - [インストール](#インストール)
      - [確認](#確認)
  - [使い方](#使い方)
    - [1. Pythonスクリプトを直接実行](#1-pythonスクリプトを直接実行)
    - [2. バッチファイルで実行する（Windows）](#2-バッチファイルで実行するwindows)
  - [出力ファイル](#出力ファイル)

## 概要

Python Age Checkは、ユーザーの生年月日を入力すると、100歳までの年齢リストをExcelファイルに出力するPythonスクリプトです。さらに、今年の誕生日が過ぎているかどうかを判定し、該当するセルの色を変更します。

## 特徴

- ユーザーの入力に基づいてExcelファイルを自動生成
- 今年の誕生日が過ぎている場合は緑色、まだ来ていない場合は赤色でセルをハイライト
- 氏名ごとにフォルダを作成し、データを整理
- シンプルなバッチファイルで簡単に実行可能

## 動作環境

- Python 3.13.2
- openpyxlライブラリ

### インストール

```powershell
pip install -U openpyxl
          or
python3 -m pip install openpyxl
```

#### 確認

```python
import openpyxl
openpyxl.__version__
```

## 使い方

スクリプトを実行するには、以下のいずれかの方法を使用してください。

### 1. Pythonスクリプトを直接実行

```bash
python age_check.py
```

### 2. バッチファイルで実行する（Windows）

`age_check.bat`をダブルクリックすると、スクリプトが実行されます。

## 出力ファイル

- `./[入力した名前]/age_check.xlsx`にExcelファイルが作成されます。
- Excelファイルには100歳までの年齢リストが記録され、誕生日のセルは色分けされます。
