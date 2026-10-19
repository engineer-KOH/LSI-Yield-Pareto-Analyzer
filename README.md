# 🚀 LSI-Yield-Pareto-Analyzer

## 📌 概要（Overview）
このシステムは、半導体LSIのテストプログラム評価およびテスト結果データ（多重ロット・複数シートのExcel等）を自動集計し、歩留まり（Yield）の解析とパレート図の生成を効率化・自動化するためのPythonツールです。  
手動での集計や不良原因の切り分け作業を自動化し、エンジニアリング業務の迅速化と精度向上を実現します。

## 🎯 主な目的（Objective）
- **歩留まり自動解析**: テストカテゴリーごとの合否結果や不良発生状況を瞬時に集計
- **可視化の効率化**: パレート分析による不良要因の特定とレポート自動生成
- **半導体テスト評価の高度化**: アナログ、コンタクト、リーク、機能テスト等の各種テストデータの横断的評価

## 🛠️ 使用技術（Tech Stack）
- **言語**: Python
- **主要ライブラリ**: `pandas`, `openpyxl`, `matplotlib`, `seaborn` など
- **対象ドメイン**: 半導体LSIテストプログラム評価、歩留まり向上・品質保証業務

## ⚙️ システム構成・処理フロー（Workflow）
- **入力 (Inputs)**: フォルダに格納された半導体LSIのテスト結果データ（Excelファイル等）の読み込み
- **処理 (Process)**: データ前処理、テスト項目ごとの歩留まり計算、不良要因の分類とパレート分析データの算出
- **出力 (Outputs)**: フォルダへの集計結果の保存およびパレート図・歩留まりサマリーレポートの自動生成

## 🚀 使い方（Quick Start）
- **1.リポジトリのクローン**: 「git clone [https://github.com/mochapaparocku-crypto/LSI-Yield-Pareto-Analyzer.git](https://github.com/mochapaparocku-crypto/LSI-Yield-Pareto-Analyzer.git)」
- **2.依存ライブラリのインストール**: 「pip install -r requirements.txt」
- **3.スクリプトの実行**: 「python yield_analyzer.py」

## 📁 フォルダ構成（Repository Structure）
```text
LSI-Yield-Pareto-Analyzer/
├── inputs/              # 入力データ（テスト結果エクセル、サンプルデータ等）
├── outputs/             # 解析結果・生成されたパレート図・レポートの出力先
├── logs/                # 実行ログの保存先
├── yield_analyzer.py    # メインの解析・集計スクリプト
├── run_yield_analyzer   # 一括実行用バッチファイル
└── .gitignore           # 機密情報・ログ・出力ファイルの除外設定
