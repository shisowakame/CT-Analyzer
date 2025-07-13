# CT-Analyzer: Advanced DICOM ROI Analysis Platform

## 🚀 概要

CT-Analyzerは、複数のDICOM画像シリーズを同時に読み込み、高度なROI（Region of Interest）解析を行う次世代の医療画像解析プラットフォームです。Webベースのモダンなインターフェースと強力な解析機能を備え、医療従事者や研究者の画像解析作業を大幅に効率化します。

### ✨ 主要特徴

- **🔍 マルチシリーズ同時解析**: 複数のDICOMフォルダを同時に読み込み、比較解析が可能
- **🎯 高度なROI機能**: 可変サイズのROI設定、同期/独立モード切り替え
- **📊 リアルタイム統計解析**: 平均値・標準偏差の即座計算と表示
- **📈 履歴管理システム**: ROI統計の履歴保存・表示・分析機能
- **🔄 動的フォルダ切り替え**: フォルダ構造に応じた柔軟な画像切り替え
- **💾 データ保存機能**: 解析結果の保存とエクスポート
- **🎨 モダンWeb UI**: 直感的で美しいユーザーインターフェース

## 📁 プロジェクト構造

```
CT-Analyzer/
├── 📄 main.py                    # アプリケーションエントリーポイント
├── 📄 requirements.txt            # 依存関係定義
├── 📄 README.md                  # このファイル
├── 🖥️ gui/                      # GUI関連モジュール
│   ├── 📄 dicom_webview.py      # メインWebビューア（37KB, 806行）
│   ├── 📄 dicom_viewer.py       # DICOMビューア
│   ├── 📄 main_window.py        # メインウィンドウ
│   ├── 📄 roi_controller.py     # ROI制御
│   ├── 📄 statistics_panel.py   # 統計パネル
│   ├── 📄 metadata_window.py    # メタデータウィンドウ
│   ├── 📄 settings_dialog.py    # 設定ダイアログ
│   └── 📄 __init__.py
├── ⚙️ core/                     # コア機能モジュール
│   ├── 📄 dicom_loader.py       # DICOM読み込み
│   ├── 📄 roi_manager.py        # ROI管理
│   ├── 📄 statistics_engine.py  # 統計解析エンジン
│   ├── 📄 export_manager.py     # エクスポート管理
│   └── 📄 __init__.py
├── 🛠️ utils/                    # ユーティリティモジュール
│   ├── 📄 dicom_utils.py        # DICOMユーティリティ
│   ├── 📄 file_utils.py         # ファイル操作
│   ├── 📄 image_utils.py        # 画像処理
│   ├── 📄 math_utils.py         # 数学計算
│   └── 📄 __init__.py
└── 📦 resources/                # リソースファイル
    ├── 📁 config/               # 設定ファイル
    └── 📁 icons/                # アイコンファイル
```

## 🛠️ 技術スタック

### バックエンド
- **Python 3.8+**: メイン開発言語
- **NumPy**: 数値計算・配列操作
- **Pillow (PIL)**: 画像処理
- **PyDICOM**: DICOMファイル読み込み
- **SciPy**: 科学計算
- **Pandas**: データ分析

### フロントエンド
- **pywebview**: WebベースGUIフレームワーク
- **HTML5/CSS3**: モダンなユーザーインターフェース
- **JavaScript (ES6+)**: インタラクティブ機能
- **Canvas API**: ROI描画・操作

### 開発・デプロイ
- **Git**: バージョン管理
- **pip**: パッケージ管理
- **Cross-platform**: Windows, macOS, Linux対応

## 🚀 セットアップ

### 前提条件
- Python 3.8以上
- pip（Pythonパッケージマネージャー）

### インストール手順

1. **リポジトリのクローン**
   ```bash
   git clone <repository-url>
   cd CT-Analyzer
   ```

2. **依存関係のインストール**
   ```bash
   pip install -r requirements.txt
   ```

3. **アプリケーションの起動**
   ```bash
   # メインアプリケーション（tkinter版）
   python main.py
   
   # Webビューア（推奨）
   python gui/dicom_webview.py <folder1> [<folder2> ...]
   ```

## 📖 使用方法

### 基本的な使用方法

1. **DICOMフォルダの指定**
   ```bash
   python gui/dicom_webview.py /path/to/dicom/folder1 /path/to/dicom/folder2
   ```

2. **フォルダ構造の理解**
   - **folder1**: 単一のDICOMフォルダ
   - **folder2**: 複数のDICOMフォルダを含む親フォルダ

### 主要機能の操作方法

#### 🎯 ROI（Region of Interest）操作
- **ROI設定**: 画像上でドラッグしてROIを設定
- **サイズ調整**: ツールバーの数値入力でROIサイズを変更
- **プリセット**: 5×5, 10×10, 20×20, 50×50のプリセットサイズ
- **リセット**: 「ROIリセット」ボタンで全ROIをクリア

#### 🔄 モード切り替え
- **同期モード**: 全シリーズで同じROI位置を使用
- **独立モード**: 各シリーズで独立したROI設定が可能

#### 📊 統計解析
- **リアルタイム計算**: ROI設定時に即座に平均値・標準偏差を計算
- **HU値解析**: 正規化前の元データ（HU値）を使用した正確な解析
- **履歴管理**: 統計結果の履歴保存・表示

#### 📁 フォルダ操作
- **動的切り替え**: フォルダ2の場合、ドロップダウンでフォルダを切り替え
- **ファイル名表示**: 現在のスライス番号とファイル名を表示
- **スライス番号**: 1ベースの直感的な番号表示

#### 💾 データ管理
- **履歴保存**: Ctrl+Sまたは保存ボタンで統計履歴を保存
- **履歴表示**: 別ウィンドウで履歴テーブルを表示
- **行単位削除**: 履歴の個別行を削除可能
- **平均計算**: 履歴データの平均値を自動計算

## 🔧 高度な機能

### マルチシリーズ解析
- 最大4つのDICOMシリーズを同時表示
- グリッドレイアウトで効率的な比較解析
- 全体同期スライダーで全シリーズを同時操作

### フォルダ構造対応
- **folder1**: 単一フォルダの直接読み込み
- **folder2**: 複数フォルダの動的切り替え
- 自動フォルダタイプ判定

### 画像処理
- **正規化**: 表示用の画像正規化
- **HU値保持**: 解析用の元データ保持
- **Base64変換**: Web表示用の画像エンコード

### 統計解析エンジン
- **平均値計算**: ROI内の画素値平均
- **標準偏差計算**: ROI内の画素値標準偏差
- **精度**: 8桁の高精度計算

## 🐛 トラブルシューティング

### よくある問題

#### pywebviewエラー
```
System.ArgumentException: 'DicomWebApi' value cannot be converted to System.Drawing.Rectangle
```
**解決方法**: 
- `debug=False`で起動
- 最新版のpywebviewを使用

#### SyntaxWarning
```
SyntaxWarning: invalid escape sequence '\d'
```
**解決方法**: 
- 正規表現のエスケープを修正済み

#### メモリ不足
**解決方法**:
- 大きなDICOMファイルの場合は、画像サイズを調整
- 一度に読み込むフォルダ数を制限

### パフォーマンス最適化
- 大量のDICOMファイルがある場合は、フォルダを分割
- 高解像度画像の場合は、表示サイズを調整
- 履歴データが多い場合は、定期的にクリア

## 📊 機能詳細

### ROI統計解析
```python
def get_roi_stats(self, series_idx, slice_idx, x, y, w, h):
    """
    ROI統計を計算
    - series_idx: シリーズ番号
    - slice_idx: スライス番号
    - x, y: ROI座標
    - w, h: ROIサイズ
    - 戻り値: {'mean': 平均値, 'std': 標準偏差}
    """
```

### フォルダ管理
```python
def switch_folder(self, series_idx, folder_idx):
    """
    フォルダを切り替え
    - series_idx: シリーズ番号
    - folder_idx: フォルダ番号
    - 戻り値: {'success': True/False, 'max_idx': 最大スライス番号}
    """
```

### 履歴管理
```python
def add_history_data(self, row_data):
    """
    履歴データを追加
    - row_data: 統計データの配列
    """

def save_history_data(self, history_data):
    """
    履歴データを保存
    - history_data: 履歴データの配列
    """
```

## 🔮 今後の開発予定

### 短期目標
- [ ] エクスポート機能の強化（Excel, JSON形式）
- [ ] バッチ処理機能の追加
- [ ] プラグインシステムの実装

### 中期目標
- [ ] 3D表示機能の追加
- [ ] AI支援解析機能
- [ ] クラウド連携機能

### 長期目標
- [ ] 医療機関向けエンタープライズ版
- [ ] モバイルアプリ版
- [ ] リアルタイム解析機能

## 🤝 コントリビューション

### 開発環境のセットアップ
1. リポジトリをフォーク
2. 開発ブランチを作成
3. 変更をコミット
4. プルリクエストを作成

### コーディング規約
- Python: PEP 8準拠
- JavaScript: ESLint準拠
- コミットメッセージ: Conventional Commits

### テスト
```bash
# ユニットテスト
python -m pytest tests/

# 統合テスト
python -m pytest tests/integration/
```

## 📄 ライセンス

このプロジェクトはMITライセンスの下で公開されています。詳細は[LICENSE](LICENSE)ファイルを参照してください。

## 📞 サポート

### 問題報告
- GitHub Issues: [問題報告](https://github.com/your-repo/CT-Analyzer/issues)
- 機能要求: [機能要求](https://github.com/your-repo/CT-Analyzer/issues/new)

### ドキュメント
- [API ドキュメント](docs/api.md)
- [ユーザーガイド](docs/user-guide.md)
- [開発者ガイド](docs/developer-guide.md)

### コミュニティ
- [Discord](https://discord.gg/ct-analyzer)
- [Twitter](https://twitter.com/ct_analyzer)
- [ブログ](https://blog.ct-analyzer.com)

---

## 🏆 謝辞

このプロジェクトは以下のオープンソースプロジェクトに支えられています：

- **PyDICOM**: DICOMファイル処理
- **NumPy**: 数値計算
- **pywebview**: WebベースGUI
- **Pillow**: 画像処理

---

**CT-Analyzer** - 次世代の医療画像解析プラットフォーム 🚀

*Version 2.0.0 | Last Updated: 2024年12月* 