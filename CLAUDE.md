# CLAUDE.md

対話は日本語でして下さい。
目標はテスト結果が向上するようにこのライブラリを改良する事です。

## プロジェクト概要

AnyToMarkdownは、ドキュメント（現在はPDFとDOCX）をMarkdown形式に変換する.NETライブラリです。このプロジェクトは、各ドキュメントタイプに特化したコンバーターを持つファクトリーパターンアーキテクチャを使用し、変換中にドキュメント構造、テーブル、書式、テキストレイアウトを保持することに重点を置いています。

## ビルドとテストコマンド

```bash
# ソリューションをビルド
dotnet build

# 全テストを実行
dotnet test

# 特定のテストを実行
dotnet test --filter "TestMethodName"

# 特定のファイルタイプのテストを実行
dotnet test --filter "RoundTripConversionTest"

# リリースパッケージをビルド
dotnet pack
ls ./AnyToMarkdown/bin/Release
```

## アーキテクチャ概要

### コアコンポーネント

- **AnyConverter**: ファイル変換のための静的メソッドを提供するメインエントリーポイント
- **ConverterFactory**: ファイル拡張子を適切なコンバーターにルーティングするファクトリーパターン実装
- **IConverter/ConverterBase**: 全ドキュメントコンバーターのインターフェースとベースクラス
- **ConvertResult**: 変換出力と警告のコンテナ

### PDF処理パイプライン

PDFコンバーターは洗練された構造解析アプローチを使用：

1. **PdfStructureAnalyzer**: 座標とフォント情報を使用してPDFページを解析
   - フォント分布分析を実行（FontAnalysisクラス）
   - 位置データを使用して単語を行にグループ化
   - タイポグラフィとレイアウトに基づいて要素をヘッダー、段落、テーブル、リストとして分類

2. **PdfWordProcessor**: 単語のグループ化と行検出を処理
   - 垂直許容値を使用して単語を行にグループ化
   - 水平許容値内で単語をマージ
   - 位置によってコンテンツをソート（上から下、左から右）

3. **MarkdownGenerator**: 構造化された要素をMarkdownに変換
   - フォントサイズとコンテンツ長に基づいてヘッダーレベルを決定
   - 適切なMarkdownテーブル構文でテーブル構造を保持
   - リストフォーマットと番号付きリストを処理

4. **フォントベース書式設定**: PdfPigのFontNameプロパティを使用して太字/斜体テキストを検出
   - "bold"、"italic"、"heavy"などのキーワードでフォント名を分析
   - 適切なMarkdown書式設定（**太字**、*斜体*）を適用

### 主要な設計原則

- **汎用的アプローチ**: 文字列の長さによる推測やハードコードされたキーワードを避け、座標とタイポグラフィ分析に依存
- **Unicode対応**: 日本語テキスト（ひらがな、カタカナ、漢字）と多言語コンテンツを処理
- **レイアウトベース検出**: 構造認識に単語位置、ギャップ、フォント分析を使用
- **本番対応**: テスト固有の仮定なしに任意のPDFで動作するよう設計

### ⚠️ 重要：テストケース特化コードの禁止

テストケースのスコア向上という目標はテストケースに対するメタ的な実装をする訳ではありません。
これは非常に当たり前のことでです。

**絶対に避けるべきパターン：**
- 特定のテスト文書の内容をハードコード（例：「田中太郎」「エンジニア」「東京都」）
- 特定言語の職業・地名・人名パターンマッチング
- テストファイル名に基づく条件分岐
- マークダウン記法に用いないフレーズでの判定

**許可される汎用的アプローチ：**
- 座標・フォントサイズ・ギャップに基づく分析
- マークダウン表記に用いる記号「# 」、「- 」、「* 」、「!\[]()」、「1. 」、「> 」などに基づいた分析
- 統計的分布分析（平均、中央値、四分位数）
- 空白・改行・インデントなどの物理的レイアウト情報

このシステムは世界中の任意のPDFで動作する必要があります。特定のテストケースでのみ機能するコードは本番環境で完全に無用です。

## テストフレームワーク

テストスイートには包括的な精度テストが含まれます：

- **ConversionAccuracyTest**: 元のMarkdownと変換されたMarkdownの構造的類似性を測定
- **ラウンドトリップテスト**: Markdown → PDF → Markdownを変換し、構造保持を比較
- **多言語サポート**: 日本語テキスト、複雑なテーブル、財務レポートでテスト
- **構造解析**: 合格/不合格閾値でヘッダー、テーブル、リスト、書式保持を評価

テストファイルは`AnyToMarkdown.Tests/Resources/`にあり、対応するPDFバージョンはpandocで生成されます。

## 主要な依存関係

- **UglyToad.PdfPig**: PDF解析とテキスト抽出
- **DocumentFormat.OpenXml**: DOCXファイル処理
- **xUnit**: テストフレームワーク

## PDF構造解析の作業

PDF処理を変更する際：
- フォントの情報が重要 - 人はグラフィックから読み取っているため、フォントの種類やサイズは文字のまとまりや意味を読み取るための重要な情報を持っています
- 線や四角のブロックといった記号の座標やサイズの情報が重要 - 人は線で囲われている内容を表や引用、スプリッターや下線などと認識しています
- 座標ベースのテーブル検出はコンテンツパターンではなく単語ギャップ分析を使用
- ヘッダー検出はコンテンツキーワードよりもフォントサイズを優先
- システムは言語に依存せず、任意のPDFレイアウトで動作するよう設計

**❌ 禁止事項：**
- 「田中太郎30」などの具体的内容に基づくセル境界検出
- 日本語の職業名・地名での意味的境界判定
- テスト文書の期待値に合わせた特別処理

## 開発ノート

- コードベースは現代的なC#機能（コレクション式、パターンマッチング）を使用
- 全PDF処理は本番使用のためハードコードされたコンテンツ仮定を避ける
- テスト精度は設定可能な閾値で構造的類似性分析を使用して測定
- ファクトリーパターンにより新しいドキュメントタイプの簡単な拡張が可能

## 専門的ドキュメント

- **PDF テーブル処理**: `Documents/PDF/Table.md` - PDF座標ベースのテーブル検出・処理の詳細な設計方針と実装ガイドライン