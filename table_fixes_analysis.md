# テーブル変換問題の解析と修正案

## 問題概要

現在のテーブル変換において以下の問題が発生：

1. ヘッダー行「列1 列2 列3」がテーブル外に出力される
2. テーブル区切り線の形式が期待値と異なる

## 原因分析

### 1. IsStandaloneHeaderInTable メソッドの問題 (2505-2530行目)

```csharp
private static bool IsStandaloneHeaderInTable(DocumentElement element)
{
    var content = element.Content.Trim();
    
    // 太字フォーマットされたテキストの汎用的分析
    if (content.Contains("**"))
    {
        // ... 太字テキストの判定ロジック
        return boldRatio > 0.7 && cells.Count <= 2 && emptyRatio >= 0.5;
    }
    
    return false;
}
```

**問題**: 「列1 列2 列3」は太字でない場合、このメソッドが`false`を返し、テーブル行として処理されるはずだが、他の箇所で除外されている可能性がある。

### 2. GenerateMarkdownTable メソッドの区切り線生成 (452-459行目)

```csharp
// ヘッダー行の後に区切り行を追加
if (rowIndex == 0)
{
    sb.Append("|");
    for (int i = 0; i < maxColumns; i++)
    {
        sb.Append(" --- |");  // 問題: スペースが入っている
    }
    sb.AppendLine();
}
```

**期待値**: `|-----|-----|-----|`
**実際**: `| --- | --- | --- |`

### 3. ConvertTableRow メソッドの処理フロー (325-362行目)

```csharp
for (int i = currentIndex; i < allElements.Count; i++)
{
    var currentElement = allElements[i];
    
    if (currentElement.Type == ElementType.TableRow)
    {
        // 汎用的なヘッダー的内容をテーブル行から除外
        if (IsStandaloneHeaderInTable(currentElement))
        {
            break; // ヘッダー的要素に遭遇したらテーブル終了
        }
        tableRows.Add(currentElement);
    }
}
```

**問題**: `IsStandaloneHeaderInTable`で除外されたヘッダー行が、テーブル外に出力される。

## 修正案

### 修正1: 区切り線の形式修正

```csharp
// 修正前
sb.Append(" --- |");

// 修正後  
sb.Append("-----|");
```

### 修正2: IsStandaloneHeaderInTable の条件見直し

通常のテーブルヘッダー行（「列1 列2 列3」など）を除外しないよう、より厳密な条件を設定：

```csharp
private static bool IsStandaloneHeaderInTable(DocumentElement element)
{
    var content = element.Content.Trim();
    
    // 太字フォーマットされたテキストの汎用的分析
    if (content.Contains("**"))
    {
        var boldMatch = System.Text.RegularExpressions.Regex.Match(content, @"\*\*([^*]+)\*\*");
        if (boldMatch.Success)
        {
            var boldText = boldMatch.Groups[1].Value.Trim();
            var cells = ParseTableCells(element);
            
            // より厳密な条件: 太字部分が大部分を占め、かつ明らかにヘッダータイトル的
            var boldRatio = (double)boldText.Length / content.Replace("**", "").Length;
            var emptyRatio = (double)cells.Count(c => string.IsNullOrWhiteSpace(c)) / Math.Max(cells.Count, 1);
            
            // 条件を厳しくして、通常のテーブルヘッダーを除外しないように
            return boldRatio > 0.9 && cells.Count == 1 && emptyRatio >= 0.8;
        }
    }
    
    return false;
}
```

### 修正3: テーブル構造の統一性確保

テーブル行の連続性を保持し、ヘッダー行が適切にテーブル内に含まれるようにする。

## 実装順序

1. 区切り線の形式修正（単純）
2. IsStandaloneHeaderInTable の条件見直し（慎重に）
3. テスト実行での確認
4. 必要に応じてさらなる調整