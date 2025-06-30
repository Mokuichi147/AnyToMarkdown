using System.Linq;
using System.Text;
using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

internal static class TableProcessor
{
    public static string ConvertTableRow(DocumentElement element, List<DocumentElement> allElements, int currentIndex)
    {
        // 既に処理済みのテーブル行かチェック
        if (currentIndex > 0 && allElements[currentIndex - 1].Type == ElementType.TableRow)
        {
            return ""; // 前の行で既に処理済み
        }
        
        // 連続するテーブル行を検出
        var consecutiveTableRows = new List<DocumentElement> { element };
        
        // 現在の行から続くテーブル行を収集（段落も含めて統合検討）
        for (int i = currentIndex + 1; i < allElements.Count; i++)
        {
            if (allElements[i].Type == ElementType.TableRow)
            {
                consecutiveTableRows.Add(allElements[i]);
            }
            else if (allElements[i].Type == ElementType.Empty)
            {
                continue; // 空行は無視して続行
            }
            else if (allElements[i].Type == ElementType.Paragraph && ShouldIntegrateIntoPreviousTableRow(allElements[i], consecutiveTableRows.Last()))
            {
                // 前のテーブル行に統合
                var lastRow = consecutiveTableRows.Last();
                lastRow.Content = lastRow.Content + "<br>" + allElements[i].Content.Trim();
            }
            else
            {
                break; // テーブル行以外の要素が出現したら終了
            }
        }
        
        // 複数のテーブル行がある場合、Markdownテーブルを生成
        if (consecutiveTableRows.Count > 1)
        {
            return GenerateMarkdownTableWithHeaders(consecutiveTableRows);
        }
        
        // 単一行の場合は通常の行として処理
        return element.Content;
    }

    public static string GenerateMarkdownTable(List<DocumentElement> tableRows)
    {
        if (tableRows.Count == 0) return "";
        
        var tableBuilder = new StringBuilder();
        var allCells = new List<List<string>>();
        
        // 各行のセルを解析
        foreach (var row in tableRows)
        {
            var cells = ParseTableCells(row);
            if (cells.Count > 0)
            {
                allCells.Add(cells);
            }
        }
        
        if (allCells.Count == 0) return "";
        
        // 最大列数を決定
        var maxColumns = allCells.Max(row => row.Count);
        
        // 各行の列数を統一
        foreach (var row in allCells)
        {
            while (row.Count < maxColumns)
            {
                row.Add("");
            }
        }
        
        // ヘッダー行
        tableBuilder.Append("| ");
        for (int i = 0; i < maxColumns; i++)
        {
            var cellContent = allCells[0][i].Trim();
            
            // セル内の<br>を適切に処理し、空白セルをプレースホルダーで保持
            if (string.IsNullOrWhiteSpace(cellContent))
            {
                cellContent = " "; // 空白セルのプレースホルダー
            }
            else
            {
                // セル内容のクリーンアップ
                cellContent = CleanTableCell(cellContent);
            }
            
            tableBuilder.Append($"{cellContent} |");
            if (i < maxColumns - 1) tableBuilder.Append(" ");
        }
        tableBuilder.AppendLine();
        
        // 区切り行
        tableBuilder.Append("|");
        for (int i = 0; i < maxColumns; i++)
        {
            tableBuilder.Append("-----|");
        }
        tableBuilder.AppendLine();
        
        // データ行
        for (int rowIndex = 1; rowIndex < allCells.Count; rowIndex++)
        {
            tableBuilder.Append("| ");
            for (int colIndex = 0; colIndex < maxColumns; colIndex++)
            {
                var cellContent = allCells[rowIndex][colIndex].Trim();
                
                if (string.IsNullOrWhiteSpace(cellContent))
                {
                    cellContent = " "; // 空白セルのプレースホルダー
                }
                else
                {
                    cellContent = CleanTableCell(cellContent);
                }
                
                tableBuilder.Append($"{cellContent} |");
                if (colIndex < maxColumns - 1) tableBuilder.Append(" ");
            }
            tableBuilder.AppendLine();
        }
        
        return tableBuilder.ToString();
    }

    private static string CleanTableCell(string cellContent)
    {
        if (string.IsNullOrEmpty(cellContent)) return "";
        
        // テーブルセルでの<br>をMarkdown改行（スペース2個+改行）に変換
        cellContent = cellContent.Replace("<br>", "  \n").Replace("<br/>", "  \n").Replace("<BR>", "  \n").Replace("<BR/>", "  \n").Replace("<br />", "  \n");
        cellContent = System.Text.RegularExpressions.Regex.Replace(cellContent, @"<br\s*/?>\s*", "  \n", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        
        // パイプ文字をエスケープ
        cellContent = cellContent.Replace("|", "\\|");
        
        // 通常の改行は空白に変換（<br>以外の改行）
        cellContent = cellContent.Replace("\r\n", " ").Replace("\r", " ");
        // \nは<br>変換後のものなので保持
        
        // 余分なスペースを除去（但し、行末の2個スペース+\nは保持）
        cellContent = System.Text.RegularExpressions.Regex.Replace(cellContent, @"(?<!  )\s+(?!\n)", " ");
        
        return cellContent.Trim();
    }

    public static List<string> ParseTableCells(DocumentElement row)
    {
        var cells = new List<string>();
        
        // パイプで区切られたMarkdownテーブル形式の場合
        if (row.Content.Contains("|"))
        {
            var pipeSeparated = row.Content.Split('|')
                .Where(cell => !string.IsNullOrWhiteSpace(cell))
                .Select(cell => cell.Trim())
                .ToList();
            
            if (pipeSeparated.Count > 1)
            {
                return pipeSeparated;
            }
        }
        
        // 座標ベースのセル分割
        if (row.Words?.Count > 0)
        {
            var coordinateCells = ParseTableCellsWithBoundaries(row);
            if (coordinateCells.Count > 1)
            {
                return coordinateCells;
            }
        }
        
        // フォールバック：テキストベースの分割
        var textCells = SplitTextIntoTableCells(row.Content);
        
        // 単一セルの場合は無理に分割しない
        if (textCells.Count <= 1)
        {
            return [row.Content.Trim()];
        }
        
        return textCells;
    }

    private static List<string> ParseTableCellsWithBoundaries(DocumentElement row)
    {
        var cells = new List<string>();
        if (row.Words == null || row.Words.Count == 0) return cells;
        
        // 単語を左から右にソート
        var sortedWords = row.Words.OrderBy(w => w.BoundingBox.Left).ToList();
        
        // 列境界を分析
        var boundaries = AnalyzeTableColumnBoundaries(sortedWords);
        
        // 境界に基づいてセルを生成
        foreach (var boundary in boundaries)
        {
            var wordsInCell = sortedWords.Where(w => 
                w.BoundingBox.Left >= boundary.Left && 
                w.BoundingBox.Right <= boundary.Right).ToList();
            
            var cellText = BuildCellTextWithSpacing(wordsInCell);
            cells.Add(cellText);
        }
        
        return cells.Where(cell => !string.IsNullOrWhiteSpace(cell)).ToList();
    }

    private static string BuildCellTextWithSpacing(List<Word> words)
    {
        if (words.Count == 0) return "";
        if (words.Count == 1) return words[0].Text?.Trim() ?? "";
        
        var result = new StringBuilder();
        for (int i = 0; i < words.Count; i++)
        {
            var word = words[i];
            var text = word.Text?.Trim() ?? "";
            
            if (string.IsNullOrEmpty(text)) continue;
            
            if (result.Length > 0)
            {
                // 単語間の距離を計算してスペースを挿入するかどうか決定
                var previousWord = words[i - 1];
                var gap = word.BoundingBox.Left - previousWord.BoundingBox.Right;
                var avgFontSize = (word.BoundingBox.Height + previousWord.BoundingBox.Height) / 2;
                
                // フォントサイズの30%以上の間隔がある場合はスペースを挿入
                if (gap > avgFontSize * 0.3)
                {
                    result.Append(" ");
                }
            }
            
            result.Append(text);
        }
        
        return result.ToString().Trim();
    }

    private static List<ColumnBoundary> AnalyzeTableColumnBoundaries(List<Word> words)
    {
        var boundaries = new List<ColumnBoundary>();
        if (words.Count == 0) return boundaries;
        
        // 単語の位置をクラスタリング
        var positions = words.Select(w => w.BoundingBox.Left).Distinct().OrderBy(p => p).ToList();
        var clusters = ClusterPositions(positions);
        
        // 各クラスタから境界を生成
        foreach (var cluster in clusters)
        {
            var left = cluster.Min();
            var right = words.Where(w => Math.Abs(w.BoundingBox.Left - left) < 10)
                            .Max(w => w.BoundingBox.Right);
            
            boundaries.Add(new ColumnBoundary { Left = left, Right = right });
        }
        
        return boundaries;
    }

    private static List<List<double>> ClusterPositions(List<double> positions)
    {
        var clusters = new List<List<double>>();
        if (positions.Count == 0) return clusters;
        
        // 重複を除去してソート
        var uniquePositions = positions.Distinct().OrderBy(p => p).ToList();
        
        const double threshold = 20.0; // クラスタリング閾値
        var currentCluster = new List<double> { uniquePositions[0] };
        
        for (int i = 1; i < uniquePositions.Count; i++)
        {
            if (uniquePositions[i] - uniquePositions[i - 1] <= threshold)
            {
                currentCluster.Add(uniquePositions[i]);
            }
            else
            {
                clusters.Add(currentCluster);
                currentCluster = new List<double> { uniquePositions[i] };
            }
        }
        
        clusters.Add(currentCluster);
        return clusters;
    }

    private static List<string> SplitTextIntoTableCells(string content)
    {
        if (string.IsNullOrWhiteSpace(content)) return [];
        
        // タブ区切りを優先
        if (content.Contains('\t'))
        {
            return content.Split('\t')
                .Where(cell => !string.IsNullOrWhiteSpace(cell))
                .Select(cell => cell.Trim())
                .ToList();
        }
        
        // 複数スペースによる区切り
        var multiSpaceSplit = System.Text.RegularExpressions.Regex.Split(content, @"\s{2,}")
            .Where(cell => !string.IsNullOrWhiteSpace(cell))
            .Select(cell => cell.Trim())
            .ToList();
        
        if (multiSpaceSplit.Count > 1)
        {
            return multiSpaceSplit;
        }
        
        // 単一スペースによる区切り（最後の手段）
        return content.Split(' ')
            .Where(cell => !string.IsNullOrWhiteSpace(cell))
            .Select(cell => cell.Trim())
            .ToList();
    }

    private static bool ShouldIntegrateIntoPreviousTableRow(DocumentElement paragraphElement, DocumentElement previousTableRow)
    {
        if (paragraphElement?.Words == null || !paragraphElement.Words.Any() ||
            previousTableRow?.Words == null || !previousTableRow.Words.Any())
        {
            return false;
        }

        // 段落要素の位置情報を取得
        var paragraphWords = paragraphElement.Words;
        var tableWords = previousTableRow.Words;

        // 垂直距離を計算（段落が前のテーブル行の直下にあるか）
        var paragraphTop = paragraphWords.Min(w => w.BoundingBox.Top);
        var tableBottom = tableWords.Max(w => w.BoundingBox.Bottom);
        var verticalGap = Math.Abs(paragraphTop - tableBottom);

        // フォントサイズを使用した距離閾値
        var avgFontSize = paragraphWords.Average(w => w.BoundingBox.Height);
        var maxVerticalGap = avgFontSize * 1.5; // フォントサイズの1.5倍以内

        if (verticalGap > maxVerticalGap)
        {
            return false; // 距離が離れすぎている
        }

        // 水平位置の重複をチェック（段落の単語がテーブル行の列範囲内にあるか）
        var tableLeft = tableWords.Min(w => w.BoundingBox.Left);
        var tableRight = tableWords.Max(w => w.BoundingBox.Right);
        var paragraphLeft = paragraphWords.Min(w => w.BoundingBox.Left);
        var paragraphRight = paragraphWords.Max(w => w.BoundingBox.Right);

        // 段落の範囲がテーブルの範囲と重複または包含される場合
        bool hasHorizontalOverlap = (paragraphLeft >= tableLeft && paragraphLeft <= tableRight) ||
                                   (paragraphRight >= tableLeft && paragraphRight <= tableRight) ||
                                   (paragraphLeft <= tableLeft && paragraphRight >= tableRight);

        if (!hasHorizontalOverlap)
        {
            return false; // 水平位置が全く重複しない
        }

        // 段落テキストがテーブルの継続として妥当かチェック
        var paragraphText = paragraphElement.Content?.Trim() ?? "";
        
        // マークダウンヘッダー記号で始まる場合は統合しない
        if (paragraphText.StartsWith("#"))
        {
            return false;
        }

        // 短いテキストで、位置的にテーブル内にある場合は統合候補
        return paragraphText.Length <= 50 && hasHorizontalOverlap && verticalGap <= maxVerticalGap;
    }

    public static string[] MergeDisconnectedTableCells(string[] lines)
    {
        var result = new List<string>();
        
        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            
            // テーブル行の場合、次の行が分離されたセル内容かチェック
            if (line.StartsWith("|") && line.EndsWith("|") && i + 1 < lines.Length)
            {
                var nextLine = lines[i + 1].Trim();
                
                // 次の行がテーブル行でもヘッダーでもない短いテキストの場合、セル内容として結合
                if (!nextLine.StartsWith("|") && !nextLine.StartsWith("#") && 
                    !nextLine.Contains("---") && nextLine.Length > 0 && nextLine.Length <= 15)
                {
                    // セル内容を前の行に結合 - より智能的に適切なセルを選択
                    var cells = line.Split('|').ToList();
                    if (cells.Count >= 3)  // 最低限のテーブル構造
                    {
                        // 各セルの長さを分析して最も短いセルに追加（通常は不完全なセル）
                        var contentCells = cells.Skip(1).Take(cells.Count - 2).ToList();  // 最初と最後の空要素を除く
                        if (contentCells.Count > 0)
                        {
                            // 最も短いセルまたは最後のセルに追加
                            var shortestCellIndex = 0;
                            var shortestLength = int.MaxValue;
                            
                            for (int k = 0; k < contentCells.Count; k++)
                            {
                                var cellLength = contentCells[k].Trim().Length;
                                if (cellLength < shortestLength)
                                {
                                    shortestLength = cellLength;
                                    shortestCellIndex = k;
                                }
                            }
                            
                            // インデックスを調整（最初の空要素分）
                            var targetCellIndex = shortestCellIndex + 1;
                            cells[targetCellIndex] = cells[targetCellIndex].Trim() + nextLine;
                            line = string.Join("|", cells);
                            i++; // 次の行をスキップ
                        }
                    }
                }
            }
            
            result.Add(line);
        }
        
        return result.ToArray();
    }
    
    public static string GenerateMarkdownTableWithHeaders(List<DocumentElement> tableRows)
    {
        if (tableRows.Count == 0) return "";
        
        var result = new StringBuilder();
        
        // 最初の行でヘッダー候補をチェック
        var firstRow = tableRows[0];
        
        // "テーブルテスト"のようなヘッダーテキストが含まれている場合
        if (ContainsHeaderText(firstRow.Content))
        {
            // ヘッダーテキストを抽出してMarkdownヘッダーとして出力
            var headerText = ExtractHeaderFromTableRow(firstRow.Content);
            if (!string.IsNullOrWhiteSpace(headerText))
            {
                result.AppendLine($"# {headerText}");
                result.AppendLine();
            }
            
            // 残りの行でテーブルを生成
            var remainingRows = ExtractTableRowsOnly(tableRows);
            if (remainingRows.Count > 0)
            {
                result.Append(GenerateMarkdownTable(remainingRows));
            }
        }
        else
        {
            // 通常のテーブル処理
            result.Append(GenerateMarkdownTable(tableRows));
        }
        
        return result.ToString();
    }
    
    private static bool ContainsHeaderText(string content)
    {
        // ヘッダーテキストの特徴を検出
        var lines = content.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length == 0) return false;
        
        var firstLine = lines[0].Trim();
        
        // 短いテキストで、テーブル区切り文字を含まない場合はヘッダー候補
        return firstLine.Length <= 20 && 
               !firstLine.Contains("|") && 
               !string.IsNullOrWhiteSpace(firstLine) &&
               lines.Length > 1; // 複数行がある場合
    }
    
    private static string ExtractHeaderFromTableRow(string content)
    {
        var lines = content.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
        return lines.Length > 0 ? lines[0].Trim() : "";
    }
    
    private static List<DocumentElement> ExtractTableRowsOnly(List<DocumentElement> tableRows)
    {
        var result = new List<DocumentElement>();
        
        foreach (var row in tableRows)
        {
            var lines = row.Content.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length > 1)
            {
                // 最初の行（ヘッダー）を除外して残りを結合
                var tableContent = string.Join("\n", lines.Skip(1).ToArray());
                if (!string.IsNullOrWhiteSpace(tableContent))
                {
                    result.Add(new DocumentElement
                    {
                        Type = ElementType.TableRow,
                        Content = tableContent,
                        FontSize = row.FontSize,
                        LeftMargin = row.LeftMargin,
                        Words = row.Words
                    });
                }
            }
            else
            {
                result.Add(row);
            }
        }
        
        return result;
    }
}

public class ColumnBoundary
{
    public double Left { get; set; }
    public double Right { get; set; }
}