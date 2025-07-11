# DICOM ROI Analyzer

## 概要
DICOM ROI Analyzerは、複数のDICOM画像フォルダを読み込み、可変サイズのROI（Region of Interest）を設定して画素値の統計解析を行うデスクトップアプリケーションです。

- Python + matplotlib + tkinterベース
- DICOM画像の表示・ROI操作・統計解析・CSV出力対応
- Windows, macOS, Linux対応

## ディレクトリ構成

```
DICOM_ROI_Analyzer/
├── main.py
├── gui/
│   ├── main_window.py
│   ├── dicom_viewer.py
│   ├── roi_controller.py
│   ├── statistics_panel.py
│   ├── metadata_window.py
│   └── settings_dialog.py
├── core/
│   ├── dicom_loader.py
│   ├── roi_manager.py
│   ├── statistics_engine.py
│   └── export_manager.py
├── utils/
│   ├── dicom_utils.py
│   ├── image_utils.py
│   ├── math_utils.py
│   └── file_utils.py
└── resources/
    ├── icons/
    └── config/
```

## セットアップ

1. Python 3.8以上をインストールしてください。
2. 必要なライブラリをインストールします：

```bash
pip install -r requirements.txt
```

## 実行方法

```bash
python main.py
```

## 主な機能
- DICOM画像・フォルダ読み込み
- ROI（領域）設定・編集・統計解析
- 統計情報のCSV出力
- メタデータ表示
- 直感的なGUI操作

---

詳細な仕様は `dicom_roi_requirements.md` を参照してください。 