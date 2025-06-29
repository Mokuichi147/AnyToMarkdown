# PDF テーブル処理の開発方針

## 概要

AnyToMarkdownライブラリにおけるPDFテーブル処理は、PDF仕様に基づく座標解析とグラフィック情報を活用した構造検出を基盤としています。この文書は、テーブル変換処理の設計原則、実装方針、および禁止事項を定義します。

## 基本原則

### 1. 座標ベースアプローチ
- **PDF座標情報の活用**: 単語の位置（BoundingBox）、フォント情報、グラフィック要素を主要な判定基準とする
- **レイアウト構造の分析**: 垂直・水平の配置パターンから論理的なテーブル構造を推定
- **物理的境界の検出**: PDFの線分情報（LineSegment）を用いた実際のテーブル境界識別

### 2. 汎用性の確保
- **言語非依存**: 特定言語のキーワードや文字パターンに依存しない処理
- **コンテンツ非依存**: テスト固有の内容や特定のデータ形式に特化しない汎用的なアルゴリズム
- **レイアウト適応性**: 様々なPDFレイアウトスタイルに対応可能な柔軟な構造解析

## 技術的実装

### 座標ベーステーブル検出
```csharp
// PostProcessor.cs - ConsolidateTableElementsByCoordinates
// グラフィック情報とテーブルパターンを使用した境界検出
public static List<DocumentElement> ConsolidateTableElementsByCoordinates(
    List<DocumentElement> elements, GraphicsInfo graphicsInfo)
```

### 行・列の境界分析
- **行グループ化**: 水平線情報と垂直位置による行の識別
- **列グループ化**: 垂直線情報と水平位置による列の識別
- **列配置分析**: 座標の分散値による左寄せ・中央寄せ・右寄せの判定

### テーブル要素統合
```csharp
// TableProcessor.cs - ShouldIntegrateIntoPreviousTableRow
// 段落要素のテーブル行への統合判定
private static bool ShouldIntegrateIntoPreviousTableRow(
    DocumentElement paragraphElement, DocumentElement previousTableRow)
```

## 処理フロー

### 1. 構造解析段階
1. **グラフィック情報収集**: PDF内の線分、矩形などの図形要素を抽出
2. **テーブルパターン検出**: 線分の配置からテーブル候補領域を特定
3. **境界領域定義**: テーブルの物理的な境界範囲を座標で定義

### 2. 要素分類段階
1. **要素位置分析**: 各テキスト要素の座標とテーブル境界との関係を評価
2. **行列マッピング**: 座標情報に基づく論理的な行・列への要素配置
3. **セル境界検出**: 隣接要素間の距離とフォントサイズによるセル分割

### 3. 統合処理段階
1. **行内統合**: 同一行内での要素結合とセル内容構築
2. **行間統合**: 分離された段落要素のテーブル行への統合判定
3. **Markdown変換**: 統合されたテーブル構造のMarkdown形式変換

## 設定可能パラメータ

### 距離閾値
- **垂直距離閾値**: `フォントサイズ * 1.5` - テーブル行間の段落統合判定
- **水平距離閾値**: `フォントサイズ * 0.3` - セル内単語間のスペース挿入判定
- **クラスタリング閾値**: `20.0` - 列境界のクラスタリング距離

### 統合条件
- **テキスト長制限**: 段落要素の長さが50文字以下の場合に統合候補とする
- **位置重複条件**: 水平位置の重複または包含関係による統合適格性判定

## 禁止事項

### ❌ 絶対に避けるべき実装パターン

#### 1. テスト固有のハードコーディング
```csharp
// 禁止例
if (cellContent.Contains("田中太郎") || cellContent.Contains("30"))
{
    // テスト固有の処理
}
```

#### 2. 言語固有のパターンマッチング
```csharp
// 禁止例
if (IsJapaneseName(text) || IsJobTitle(text))
{
    // 特定言語の意味的判定
}
```

#### 3. コンテンツベースの境界検出
```csharp
// 禁止例
var cellBoundary = text.Split("エンジニア")[0].Length;
```

#### 4. ファイル名による条件分岐
```csharp
// 禁止例
if (fileName.Contains("test-multiline-table"))
{
    // テストファイル固有の処理
}
```

### ✅ 推奨される実装パターン

#### 1. 座標ベース判定
```csharp
// 推奨例
var verticalGap = Math.Abs(paragraphTop - tableBottom);
var maxVerticalGap = avgFontSize * 1.5;
return verticalGap <= maxVerticalGap;
```

#### 2. 統計的分析
```csharp
// 推奨例
var columnVariance = CalculatePositionVariance(columnElements);
var alignment = columnVariance < threshold ? ColumnAlignment.Left : ColumnAlignment.Center;
```

#### 3. 構造的パターン検出
```csharp
// 推奨例
bool isTableStructure = hasHorizontalLines && hasVerticalLines && 
                       elementCount > minTableSize;
```

## パフォーマンス考慮事項

### 効率的な処理順序
1. **粗い篩い分け**: 明らかにテーブルでない要素の早期除外
2. **段階的精緻化**: 候補領域の段階的な境界精緻化
3. **キャッシュ活用**: 座標計算結果の再利用

### メモリ使用量最適化
- **遅延評価**: 必要時のみ詳細な座標計算を実行
- **インデックス管理**: 処理済み要素のインデックス追跡による重複処理回避

## テスト方針

### 構造的検証
- **座標精度**: 境界検出の座標精度測定
- **統合率**: 正しく統合された要素の割合
- **Markdown品質**: 生成されたMarkdownテーブルの構文正確性

### 汎用性検証
- **多言語対応**: 日本語、英語、その他言語での一貫した動作
- **レイアウト多様性**: 異なるテーブルスタイルでの適応性
- **スケーラビリティ**: 大規模テーブルでの処理効率

## 関連ファイル

### 実装ファイル
- `AnyToMarkdown/Pdf/PostProcessor.cs` - 座標ベース統合処理
- `AnyToMarkdown/Pdf/TableProcessor.cs` - テーブル行処理とMarkdown変換
- `AnyToMarkdown/Pdf/PdfStructureAnalyzer.cs` - 構造解析とグラフィック情報抽出

### データ構造
- `GraphicsInfo.cs` - グラフィック要素とテーブルパターン定義
- `DocumentElement.cs` - 要素の座標と内容情報
- `ColumnBoundary.cs` - 列境界の座標定義

## 今後の拡張方針

### 短期的改善
- **セル結合処理**: 複数セルにまたがる要素の検出と処理
- **ネストテーブル**: テーブル内テーブルの階層構造対応
- **複雑レイアウト**: 不規則な境界線を持つテーブルへの対応

### 長期的発展
- **機械学習統合**: 座標パターンの学習による検出精度向上
- **動的閾値**: 文書特性に応じた自動閾値調整
- **構造テンプレート**: 一般的なテーブルパターンのテンプレート化

---

この方針に従うことで、特定のテストケースに依存しない汎用的で堅牢なテーブル処理システムを維持・発展させることができます。