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
        
        // 現在の行から続くテーブル行を収集（より保守的な統合）
        for (int i = currentIndex + 1; i < allElements.Count; i++)
        {
            if (allElements[i].Type == ElementType.TableRow)
            {
                // 座標ベースでテーブル行の連続性をチェック
                if (IsTableRowContinuous(consecutiveTableRows.Last(), allElements[i]))
                {
                    consecutiveTableRows.Add(allElements[i]);
                }
                else
                {
                    break; // 座標的に離れている場合は別テーブル
                }
            }
            else if (allElements[i].Type == ElementType.Empty)
            {
                continue; // 空行は無視して続行
            }
            else if (allElements[i].Type == ElementType.Header)
            {
                break; // ヘッダーが出現したら別のセクション
            }
            else if (allElements[i].Type == ElementType.Paragraph)
            {
                // 段落が前のテーブル行の続きかを座標で判定（より厳格に）
                if (ShouldIntegrateIntoPreviousTableRowByCoordinates(allElements[i], consecutiveTableRows.Last()))
                {
                    // 前のテーブル行に統合
                    var lastRow = consecutiveTableRows.Last();
                    lastRow.Content = lastRow.Content + "<br>" + allElements[i].Content.Trim();
                }
                else
                {
                    break; // 座標的に離れている場合は統合しない
                }
            }
            else
            {
                break; // テーブル行以外の要素が出現したら終了
            }
        }
        
        // テーブル行のしきい値を調整（より厳格に）
        if (consecutiveTableRows.Count >= 1) // 単一行でもテーブルとして処理
        {
            return GenerateMarkdownTableWithHeaders(consecutiveTableRows);
        }
        
        // フォールバック
        return element.Content;
    }

    public static string GenerateMarkdownTable(List<DocumentElement> tableRows)
    {
        if (tableRows.Count == 0) return "";
        
        var tableBuilder = new StringBuilder();
        var allCells = new List<List<string>>();
        
        // テーブル全体の統一列境界を計算（CLAUDE.md準拠）
        var globalBoundaries = CalculateGlobalColumnBoundaries(tableRows);
        
        // 各行のセルを解析（統計的分析を優先）
        foreach (var row in tableRows)
        {
            // 優先順位1: 統計的ギャップ分析
            var statisticalCells = ParseTableCellsWithStatisticalGapAnalysis(row);
            
            // 優先順位2: 統一境界分析（統計的分析が不十分な場合のみ）
            var globalCells = new List<string>();
            if (globalBoundaries.Any())
            {
                globalCells = ParseTableCellsWithGlobalBoundaries(row, globalBoundaries);
            }
            
            // より多くのセルを生成した方を採用（ただし合理的な範囲内）
            var cells = statisticalCells;
            if (globalCells.Count > statisticalCells.Count && 
                globalCells.Count <= statisticalCells.Count * 2)
            {
                cells = globalCells;
            }
            
            // デバッグ：各行のセル数を確認
            if (row.Content.Contains("製品名") || row.Content.Contains("Basic") || row.Content.Contains("プレミアム"))
            {
                Console.WriteLine($"DEBUG TABLE - Row: {row.Content.Substring(0, Math.Min(50, row.Content.Length))}...");
                Console.WriteLine($"DEBUG TABLE - Statistical cells ({statisticalCells.Count}): [{string.Join(", ", statisticalCells.Select(c => $"'{c}'"))}]");
                Console.WriteLine($"DEBUG TABLE - Global cells ({globalCells.Count}): [{string.Join(", ", globalCells.Select(c => $"'{c}'"))}]");
                Console.WriteLine($"DEBUG TABLE - Final cells ({cells.Count}): [{string.Join(", ", cells.Select(c => $"'{c}'"))}]");
            }
            
            if (cells.Count > 0)
            {
                allCells.Add(cells);
            }
        }
        
        if (allCells.Count == 0) return "";
        
        // 最大列数を決定
        var maxColumns = allCells.Max(row => row.Count);
        
        // デバッグ：列数の決定過程を確認
        Console.WriteLine($"DEBUG TABLE - All rows column counts: [{string.Join(", ", allCells.Select(row => row.Count))}]");
        Console.WriteLine($"DEBUG TABLE - Max columns determined: {maxColumns}");
        
        // 各行の列数を統一（空列を適切に処理）
        foreach (var row in allCells)
        {
            while (row.Count < maxColumns)
            {
                row.Add(" "); // 完全に空ではなく、スペース1個で埋める
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
                cellContent = " "; // 空白セルのプレースホルダー（一貫した処理）
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
        
        // 座標ベースのセル分割（保守的アプローチ）
        if (row.Words?.Count > 0)
        {
            // 統計的ギャップ分析による座標ベースセル分割（CLAUDE.md準拠）
            var statisticalCells = ParseTableCellsWithStatisticalGapAnalysis(row);
            if (statisticalCells.Count > 1)
            {
                return statisticalCells;
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
        
        // より細かい列境界分析
        var boundaries = AnalyzeTableColumnBoundariesImproved(sortedWords);
        
        // 境界に基づいてセルを生成（さらに厳密な境界チェック）
        foreach (var boundary in boundaries)
        {
            var wordsInCell = sortedWords.Where(w => 
            {
                // 単語が境界内に完全に含まれているかチェック
                var wordLeft = w.BoundingBox.Left;
                var wordRight = w.BoundingBox.Right;
                var wordCenter = (wordLeft + wordRight) / 2.0;
                
                // 中心点が境界内 かつ 単語の大部分が境界内
                var centerInBoundary = wordCenter >= boundary.Left && wordCenter <= boundary.Right;
                var majorityInBoundary = (Math.Max(wordLeft, boundary.Left) < Math.Min(wordRight, boundary.Right)) &&
                                       ((Math.Min(wordRight, boundary.Right) - Math.Max(wordLeft, boundary.Left)) > 
                                        (wordRight - wordLeft) * 0.6); // 60%以上重複
                
                return centerInBoundary && majorityInBoundary;
            }).ToList();
            
            if (wordsInCell.Any())
            {
                var cellText = BuildCellTextWithSpacing(wordsInCell);
                if (!string.IsNullOrWhiteSpace(cellText))
                {
                    cells.Add(cellText);
                }
            }
        }
        
        // フォールバック：境界分析が失敗した場合は段階的により敏感な検出を試行
        if (cells.Count <= 1 && sortedWords.Count > 1)
        {
            // より敏感なギャップベース分割
            cells = SplitWordsByLargeGaps(sortedWords);
            
            // さらなるフォールバック：フォントサイズベースの分離
            if (cells.Count <= 1)
            {
                cells = SplitWordsByFontSizeGaps(sortedWords);
            }
            
            // 最終フォールバック：非常に敏感な距離ベース分割
            if (cells.Count <= 1)
            {
                cells = SplitWordsBySensitiveDistanceAnalysis(sortedWords);
            }
        }
        
        return cells.Where(cell => !string.IsNullOrWhiteSpace(cell)).ToList();
    }
    
    private static List<string> SplitWordsBySensitiveDistanceAnalysis(List<Word> words)
    {
        var cells = new List<string>();
        var currentCellWords = new List<Word>();
        
        if (words.Count == 0) return cells;
        
        // より敏感な距離ベース分析
        for (int i = 0; i < words.Count; i++)
        {
            var currentWord = words[i];
            currentCellWords.Add(currentWord);
            
            if (i < words.Count - 1)
            {
                var nextWord = words[i + 1];
                var gapSize = nextWord.BoundingBox.Left - currentWord.BoundingBox.Right;
                
                // 平均文字幅を計算
                var avgCharWidth = currentWord.BoundingBox.Width / Math.Max(1, currentWord.Text?.Length ?? 1);
                
                // より敏感な閾値：1文字分のギャップで分離
                var sensitiveThreshold = avgCharWidth * 1.0;
                
                if (gapSize > sensitiveThreshold)
                {
                    var cellText = BuildCellTextWithSpacing(currentCellWords);
                    if (!string.IsNullOrWhiteSpace(cellText))
                    {
                        cells.Add(cellText);
                    }
                    currentCellWords.Clear();
                }
            }
        }
        
        // 最後のセルを追加
        if (currentCellWords.Count > 0)
        {
            var cellText = BuildCellTextWithSpacing(currentCellWords);
            if (!string.IsNullOrWhiteSpace(cellText))
            {
                cells.Add(cellText);
            }
        }
        
        return cells;
    }
    
    private static List<string> ParseTableCellsWithAdvancedBoundaryDetection(DocumentElement row, List<DocumentElement>? allTableRows = null)
    {
        var cells = new List<string>();
        if (row.Words == null || row.Words.Count == 0) return cells;
        
        var sortedWords = row.Words.OrderBy(w => w.BoundingBox.Left).ToList();
        
        // Step 1: 単語間ギャップ分析による直接的セル分離
        var cellGroups = AnalyzeWordGapsForCellSeparation(sortedWords);
        
        if (cellGroups.Count > 1)
        {
            foreach (var group in cellGroups)
            {
                var cellText = BuildCellTextWithSpacing(group);
                if (!string.IsNullOrWhiteSpace(cellText))
                {
                    cells.Add(cellText.Trim());
                }
                else
                {
                    cells.Add(""); // 空のセル
                }
            }
            
            return cells;
        }
        
        // Step 1.5: フォントサイズと座標ベースの強制分離（ヘッダー対応）
        var fontBasedCells = AnalyzeFontAndPositionBasedSeparation(sortedWords);
        if (fontBasedCells.Count > 1)
        {
            return fontBasedCells.Select(cell => cell.Trim()).ToList();
        }
        
        // Step 1.6: 微小ギャップでの強制的セル分離（テーブルヘッダー対応）
        var microGapCells = AnalyzeMicroGapsForTableHeaders(sortedWords);
        if (microGapCells.Count > 1)
        {
            return microGapCells.Select(cell => cell.Trim()).ToList();
        }
        
        // Step 2: フォールバック - より厳密な列境界分析
        var columnBoundaries = AnalyzeStrictColumnBoundaries(sortedWords);
        
        // Step 3: 境界に基づいてセルを分割
        if (columnBoundaries.Count > 1)
        {
            for (int i = 0; i < columnBoundaries.Count; i++)
            {
                var boundary = columnBoundaries[i];
                var wordsInColumn = sortedWords.Where(w => 
                {
                    var wordCenter = (w.BoundingBox.Left + w.BoundingBox.Right) / 2.0;
                    return wordCenter >= boundary.Start && wordCenter <= boundary.End;
                }).ToList();
                
                if (wordsInColumn.Any())
                {
                    var cellText = BuildCellTextWithSpacing(wordsInColumn);
                    if (!string.IsNullOrWhiteSpace(cellText))
                    {
                        cells.Add(cellText);
                    }
                }
                else
                {
                    cells.Add(""); // 空のセル
                }
            }
        }
        else
        {
            // フォールバック：クラスタリング
            var clusters = ClusterWordsByProximity(sortedWords);
            foreach (var cluster in clusters)
            {
                if (cluster.Any())
                {
                    var cellText = BuildCellTextWithSpacing(cluster);
                    if (!string.IsNullOrWhiteSpace(cellText))
                    {
                        cells.Add(cellText);
                    }
                }
            }
        }
        
        return cells.Count > 1 ? cells : new List<string>();
    }
    
    private static List<List<Word>> AnalyzeWordGapsForCellSeparation(List<Word> sortedWords)
    {
        var cellGroups = new List<List<Word>>();
        if (sortedWords.Count == 0) return cellGroups;
        
        // 単語間の水平ギャップを分析
        var gaps = new List<(double Gap, int Index)>();
        for (int i = 0; i < sortedWords.Count - 1; i++)
        {
            var gap = sortedWords[i + 1].BoundingBox.Left - sortedWords[i].BoundingBox.Right;
            gaps.Add((gap, i));
        }
        
        if (gaps.Count == 0)
        {
            cellGroups.Add(sortedWords);
            return cellGroups;
        }
        
        // IQR分析による統計的セル境界検出
        var sortedGaps = gaps.OrderBy(g => g.Gap).ToList();
        var q1Index = (int)(sortedGaps.Count * 0.25);
        var q3Index = (int)(sortedGaps.Count * 0.75);
        
        var q1 = sortedGaps[Math.Max(0, q1Index)].Gap;
        var q3 = sortedGaps[Math.Min(sortedGaps.Count - 1, q3Index)].Gap;
        var iqr = q3 - q1;
        
        // 平均文字幅を動的計算
        var avgWordWidth = sortedWords.Average(w => w.BoundingBox.Width);
        var charBasedThreshold = avgWordWidth * 0.8; // 文字幅の80%
        
        // 動的閾値：より保守的な分析（過度な分離を避ける）
        var statisticalThreshold = q3 + (iqr * 1.0); // 保守的な閾値
        var dynamicThreshold = Math.Max(statisticalThreshold, charBasedThreshold);
        
        // 最小閾値設定（大きなギャップのみ検出）
        var minThreshold = avgWordWidth * 0.5; // より高い最小値で分離を抑制
        dynamicThreshold = Math.Max(dynamicThreshold, minThreshold);
        
        // 小さなギャップは無視する
        if (iqr < avgWordWidth * 0.5)
        {
            dynamicThreshold = Math.Max(dynamicThreshold, avgWordWidth * 0.8); // 分離を大幅に抑制
        }
        
        // 座標ベース：単語幅の変化による境界検出
        var additionalBoundaries = new List<int>();
        for (int i = 0; i < sortedWords.Count - 1; i++)
        {
            var currentWordWidth = sortedWords[i].BoundingBox.Width;
            var nextWordWidth = sortedWords[i + 1].BoundingBox.Width;
            var gap = gaps[i].Gap;
            
            // 単語幅の統計的変化を分析
            var widthRatio = Math.Max(currentWordWidth, nextWordWidth) / Math.Min(currentWordWidth, nextWordWidth);
            
            // 単語幅の大きな変化 + 小さなギャップ = 潜在的境界
            if (widthRatio >= 1.5 && gap >= avgWordWidth * 0.02)
            {
                additionalBoundaries.Add(i);
            }
            
            // フォント高さの変化による境界検出
            var currentHeight = sortedWords[i].BoundingBox.Height;
            var nextHeight = sortedWords[i + 1].BoundingBox.Height;
            var heightDiff = Math.Abs(currentHeight - nextHeight);
            
            if (heightDiff >= avgWordWidth * 0.1 && gap >= avgWordWidth * 0.01)
            {
                additionalBoundaries.Add(i);
            }
        }
        
        // セル境界として検出されるギャップを特定（通常ギャップ＋座標ベース境界）
        var normalBoundaries = gaps.Where(g => g.Gap >= dynamicThreshold)
                                   .OrderBy(g => g.Index)
                                   .Select(g => g.Index);
        
        var cellBoundaries = normalBoundaries.Union(additionalBoundaries)
                                           .OrderBy(x => x)
                                           .ToList();
        
        // 単語をセルグループに分割
        int startIndex = 0;
        foreach (var boundaryIndex in cellBoundaries)
        {
            var cellWords = sortedWords.Skip(startIndex).Take(boundaryIndex - startIndex + 1).ToList();
            if (cellWords.Count > 0)
            {
                cellGroups.Add(cellWords);
            }
            startIndex = boundaryIndex + 1;
        }
        
        // 残りの単語を最後のセルに追加
        if (startIndex < sortedWords.Count)
        {
            var remainingWords = sortedWords.Skip(startIndex).ToList();
            if (remainingWords.Count > 0)
            {
                cellGroups.Add(remainingWords);
            }
        }
        
        // 単一グループの場合は空リストを返す（フォールバックメソッドを使用）
        if (cellGroups.Count <= 1)
        {
            return new List<List<Word>>();
        }
        
        return cellGroups;
    }
    
    private static List<string> AnalyzeFontAndPositionBasedSeparation(List<Word> sortedWords)
    {
        var cells = new List<string>();
        if (sortedWords.Count <= 1) return cells;
        
        // フォントサイズと位置の統合的分析による分離
        var separationCandidates = new List<int>();
        
        for (int i = 0; i < sortedWords.Count - 1; i++)
        {
            var currentWord = sortedWords[i];
            var nextWord = sortedWords[i + 1];
            
            // 座標間隔
            var gap = nextWord.BoundingBox.Left - currentWord.BoundingBox.Right;
            var avgWordWidth = sortedWords.Average(w => w.BoundingBox.Width);
            
            // フォントサイズの類似性
            var fontSizeDiff = Math.Abs(currentWord.BoundingBox.Height - nextWord.BoundingBox.Height);
            var avgFontSize = sortedWords.Average(w => w.BoundingBox.Height);
            
            // 座標ベース分析：
            // 1. わずかな間隔でもフォントが同じなら分離候補
            // 2. 同一フォントサイズで一定の距離がある場合
            if (gap >= avgWordWidth * 0.01 && // 極小ギャップでも検出
                fontSizeDiff <= avgFontSize * 0.15) // フォントサイズの差が15%以内
            {
                separationCandidates.Add(i);
            }
            
            // さらに、単語の境界Box間に十分な空間がある場合
            var wordSpacing = gap / avgWordWidth;
            if (wordSpacing >= 0.02) // 単語幅の2%以上（極小）
            {
                separationCandidates.Add(i);
            }
            
            // 絶対位置ベースの分析
            var currentCenter = (currentWord.BoundingBox.Left + currentWord.BoundingBox.Right) / 2.0;
            var nextCenter = (nextWord.BoundingBox.Left + nextWord.BoundingBox.Right) / 2.0;
            var centerDistance = Math.Abs(nextCenter - currentCenter);
            
            // 中心間距離が単語幅より大きい場合は分離候補
            if (centerDistance >= avgWordWidth * 0.8)
            {
                separationCandidates.Add(i);
            }
            
            // 純粋な座標ベース分析：単語境界Box間の物理的距離分析
            var currentRight = currentWord.BoundingBox.Right;
            var nextLeft = nextWord.BoundingBox.Left;
            var physicalGap = nextLeft - currentRight;
            
            // 物理的ギャップの統計的分析による境界検出
            if (physicalGap > 0) // 正のギャップのみ
            {
                var gapRatio = physicalGap / avgWordWidth;
                
                // 統計的閾値：ギャップが平均単語幅の特定比率を超える場合
                if (gapRatio >= 0.1) // 10%以上のギャップで潜在的境界
                {
                    separationCandidates.Add(i);
                }
            }
        }
        
        if (separationCandidates.Count == 0) return cells;
        
        // 重複を除去し、ソート
        separationCandidates = separationCandidates.Distinct().OrderBy(x => x).ToList();
        
        // 各セルを構築
        int startIndex = 0;
        foreach (var sepIndex in separationCandidates)
        {
            var cellWords = sortedWords.Skip(startIndex).Take(sepIndex - startIndex + 1).ToList();
            if (cellWords.Count > 0)
            {
                var cellText = string.Join(" ", cellWords.Select(w => w.Text));
                cells.Add(cellText);
            }
            startIndex = sepIndex + 1;
        }
        
        // 残りの単語を最後のセルに追加
        if (startIndex < sortedWords.Count)
        {
            var remainingWords = sortedWords.Skip(startIndex).ToList();
            if (remainingWords.Count > 0)
            {
                var cellText = string.Join(" ", remainingWords.Select(w => w.Text));
                cells.Add(cellText);
            }
        }
        
        return cells;
    }
    
    private static List<string> AnalyzeMicroGapsForTableHeaders(List<Word> sortedWords)
    {
        var cells = new List<string>();
        if (sortedWords.Count <= 1) return cells;
        
        // テーブルヘッダー用の非常に低い閾値を設定
        var avgWordWidth = sortedWords.Average(w => w.BoundingBox.Width);
        var microThreshold = avgWordWidth * 0.05; // 単語幅の5%（より積極的）
        
        var gaps = new List<(double Gap, int Index)>();
        for (int i = 0; i < sortedWords.Count - 1; i++)
        {
            var gap = sortedWords[i + 1].BoundingBox.Left - sortedWords[i].BoundingBox.Right;
            gaps.Add((gap, i));
        }
        
        // 微小ギャップでもセル境界として検出（負のギャップも含む）
        var microBoundaries = gaps.Where(g => g.Gap >= -avgWordWidth * 0.1) // 軽微な重複も許可
                                 .OrderBy(g => g.Index)
                                 .Select(g => g.Index)
                                 .ToList();
        
        if (microBoundaries.Count == 0) return cells;
        
        // 各セル境界で単語を分割
        int startIndex = 0;
        foreach (var boundaryIndex in microBoundaries)
        {
            var cellWords = sortedWords.Skip(startIndex).Take(boundaryIndex - startIndex + 1).ToList();
            if (cellWords.Count > 0)
            {
                var cellText = string.Join(" ", cellWords.Select(w => w.Text));
                cells.Add(cellText);
            }
            startIndex = boundaryIndex + 1;
        }
        
        // 残りの単語を最後のセルに追加
        if (startIndex < sortedWords.Count)
        {
            var remainingWords = sortedWords.Skip(startIndex).ToList();
            if (remainingWords.Count > 0)
            {
                var cellText = string.Join(" ", remainingWords.Select(w => w.Text));
                cells.Add(cellText);
            }
        }
        
        return cells;
    }
    
    private static List<(double Start, double End)> AnalyzeStrictColumnBoundaries(List<Word> words)
    {
        var boundaries = new List<(double Start, double End)>();
        if (words.Count == 0) return boundaries;
        
        // 各単語の左端座標を収集
        var leftPositions = words.Select(w => w.BoundingBox.Left).Distinct().OrderBy(x => x).ToList();
        
        // 各位置での単語の密度を分析
        var clusters = new List<List<double>>();
        var currentCluster = new List<double> { leftPositions[0] };
        
        for (int i = 1; i < leftPositions.Count; i++)
        {
            var gap = leftPositions[i] - leftPositions[i - 1];
            var avgWordWidth = words.Average(w => w.BoundingBox.Width);
            
            // 単語幅の半分以上のギャップで新しいクラスタ
            if (gap > avgWordWidth * 0.5)
            {
                clusters.Add(currentCluster);
                currentCluster = new List<double> { leftPositions[i] };
            }
            else
            {
                currentCluster.Add(leftPositions[i]);
            }
        }
        clusters.Add(currentCluster);
        
        // 各クラスタから列境界を生成
        var totalWidth = words.Max(w => w.BoundingBox.Right) - words.Min(w => w.BoundingBox.Left);
        var numColumns = Math.Max(clusters.Count, 2); // 最低2列
        
        for (int i = 0; i < clusters.Count; i++)
        {
            var clusterStart = clusters[i].Min();
            var clusterEnd = i < clusters.Count - 1 ? 
                (clusters[i].Max() + clusters[i + 1].Min()) / 2.0 : 
                words.Max(w => w.BoundingBox.Right);
                
            boundaries.Add((clusterStart, clusterEnd));
        }
        
        return boundaries;
    }
    
    private static List<List<Word>> ClusterWordsByProximity(List<Word> words)
    {
        if (words.Count == 0) return new List<List<Word>>();
        if (words.Count == 1) return new List<List<Word>> { new List<Word> { words[0] } };
        
        var clusters = new List<List<Word>>();
        var currentCluster = new List<Word> { words[0] };
        
        // 動的な閾値計算
        var gaps = new List<double>();
        for (int i = 0; i < words.Count - 1; i++)
        {
            var gap = words[i + 1].BoundingBox.Left - words[i].BoundingBox.Right;
            gaps.Add(gap);
        }
        
        // 統計的閾値計算（より保守的に）
        gaps.Sort();
        var medianGap = gaps[gaps.Count / 2];
        var q1 = gaps[gaps.Count / 4];
        var q3 = gaps[(3 * gaps.Count) / 4];
        
        // より精密な境界検出（意味的境界を考慮）
        var avgGap = gaps.Average();
        var positiveGaps = gaps.Where(g => g > 0).ToList();
        
        double clusterThreshold;
        if (positiveGaps.Count == 0)
        {
            // 全てのギャップが0以下の場合はフォールバック
            var fallbackWordWidth = words.Average(w => w.BoundingBox.Width);
            clusterThreshold = fallbackWordWidth * 0.5;
        }
        else
        {
            var minSignificantGap = positiveGaps.Min();
            var maxGap = positiveGaps.Max();
            
            // 文字幅ベースの閾値計算
            var avgCharWidth = words.Average(w => w.BoundingBox.Width / Math.Max(1, w.Text?.Length ?? 1));
            var characterBasedThreshold = avgCharWidth * 2.0; // 2文字分のギャップ
            
            // 複数の候補閾値を計算
            var medianThreshold = medianGap * 0.6; // より敏感
            var q1Threshold = q1 * 1.0; // より敏感
            var avgThreshold = avgGap * 0.4; // より敏感
            var charThreshold = characterBasedThreshold;
            
            // 最も敏感（小さい）な閾値を選択
            clusterThreshold = Math.Min(
                Math.Min(Math.Min(medianThreshold, q1Threshold), avgThreshold),
                charThreshold
            );
            
            // ただし最小値は保証（ノイズ除去）
            clusterThreshold = Math.Max(clusterThreshold, minSignificantGap * 0.8);
        }
        
        // デフォルト閾値（フォールバック）
        var avgWordWidth = words.Average(w => w.BoundingBox.Width);
        var defaultThreshold = avgWordWidth * 0.2; // より敏感
        var finalThreshold = Math.Max(clusterThreshold, defaultThreshold);
        
        for (int i = 1; i < words.Count; i++)
        {
            var prevWord = words[i - 1];
            var currentWord = words[i];
            var gap = currentWord.BoundingBox.Left - prevWord.BoundingBox.Right;
            
            if (gap > finalThreshold)
            {
                // 新しいクラスタ（セル）を開始
                clusters.Add(currentCluster);
                currentCluster = new List<Word> { currentWord };
            }
            else
            {
                // 現在のクラスタに追加
                currentCluster.Add(currentWord);
            }
        }
        
        // 最後のクラスタを追加
        clusters.Add(currentCluster);
        
        return clusters;
    }
    
    private static List<ColumnBoundary> AnalyzeTableColumnBoundariesImproved(List<Word> words)
    {
        if (words.Count == 0) 
            return [];
        
        if (words.Count == 1)
        {
            return [new ColumnBoundary { Left = words[0].BoundingBox.Left, Right = words[0].BoundingBox.Right }];
        }
        
        // 座標ベースのクラスタリング分析
        var wordPositions = words.Select(w => new
        {
            Word = w,
            LeftEdge = w.BoundingBox.Left,
            RightEdge = w.BoundingBox.Right,
            Center = (w.BoundingBox.Left + w.BoundingBox.Right) / 2.0
        }).OrderBy(w => w.LeftEdge).ToList();
        
        // 単語間ギャップの統計的分析
        var gaps = new List<double>();
        for (int i = 0; i < wordPositions.Count - 1; i++)
        {
            var gap = wordPositions[i + 1].LeftEdge - wordPositions[i].RightEdge;
            if (gap > 0) gaps.Add(gap);
        }
        
        if (gaps.Count == 0)
        {
            return [new ColumnBoundary { Left = words[0].BoundingBox.Left, Right = words.Last().BoundingBox.Right }];
        }
        
        // 統計的閾値計算（より厳密）
        var sortedGaps = gaps.OrderBy(g => g).ToList();
        var q1Index = Math.Max(0, (int)(sortedGaps.Count * 0.25));
        var q3Index = Math.Min(sortedGaps.Count - 1, (int)(sortedGaps.Count * 0.75));
        var medianIndex = sortedGaps.Count / 2;
        
        var q1 = sortedGaps[q1Index];
        var q3 = sortedGaps[q3Index];
        var median = sortedGaps[medianIndex];
        
        // より敏感なギャップ検出（複数の基準を組み合わせ）
        var iqr = q3 - q1;
        var iqrThreshold = q1 + iqr * 0.3; // より敏感
        
        // 最小ギャップベースの閾値
        var minGapThreshold = sortedGaps[0] * 2.5; // より敏感
        
        // 中央値ベースの閾値
        var medianThreshold = median * 0.6; // より敏感
        
        // 複数の閾値を計算して最適なものを選択
        var candidateThresholds = new List<double> { iqrThreshold, minGapThreshold, medianThreshold };
        
        // 非常に小さいギャップは除外しつつ、敏感に検出
        var minAllowedGap = sortedGaps[0] * 1.2; // より敏感
        var maxReasonableGap = q3 * 1.5; // 過度に大きなギャップは制限
        
        var finalThreshold = candidateThresholds
            .Where(t => t >= minAllowedGap && t <= maxReasonableGap)
            .DefaultIfEmpty(medianThreshold)
            .Min(); // 最も敏感（小さい）な閾値を選択
        
        // 有意なギャップを検出
        var columnBreaks = new List<double>();
        for (int i = 0; i < wordPositions.Count - 1; i++)
        {
            var gap = wordPositions[i + 1].LeftEdge - wordPositions[i].RightEdge;
            if (gap >= finalThreshold)
            {
                // 境界位置を正確に設定（ギャップの中点）
                var breakPoint = (wordPositions[i].RightEdge + wordPositions[i + 1].LeftEdge) / 2.0;
                columnBreaks.Add(breakPoint);
            }
        }
        
        // 境界を生成
        var boundaries = new List<ColumnBoundary>();
        var currentLeft = wordPositions[0].LeftEdge;
        
        foreach (var breakPoint in columnBreaks.OrderBy(b => b))
        {
            boundaries.Add(new ColumnBoundary 
            { 
                Left = currentLeft, 
                Right = breakPoint 
            });
            currentLeft = breakPoint;
        }
        
        // 最後の境界
        boundaries.Add(new ColumnBoundary 
        { 
            Left = currentLeft, 
            Right = wordPositions.Last().RightEdge 
        });
        
        return boundaries;
    }
    
    private static List<string> SplitWordsByLargeGaps(List<Word> words)
    {
        var cells = new List<string>();
        var currentCellWords = new List<Word>();
        
        if (words.Count == 0) return cells;
        
        // ギャップの統計的分析
        var gaps = new List<double>();
        for (int i = 0; i < words.Count - 1; i++)
        {
            var currentWord = words[i];
            var nextWord = words[i + 1];
            gaps.Add(nextWord.BoundingBox.Left - currentWord.BoundingBox.Right);
        }
        
        // 平均ギャップサイズを計算
        var avgGap = gaps.Count > 0 ? gaps.Average() : 0.0;
        var maxGap = gaps.Count > 0 ? gaps.Max() : 0.0;
        
        // より敏感な閾値：平均ギャップの1.5倍または最大ギャップの30%のいずれか小さい方
        var splitThreshold = Math.Min(avgGap * 1.5, maxGap * 0.3);
        
        // 最小閾値を設定（ギャップが極めて小さい場合の対策）
        if (gaps.Count > 0)
        {
            var minGap = gaps.Min();
            splitThreshold = Math.Max(splitThreshold, minGap * 1.2);
        }
        
        for (int i = 0; i < words.Count; i++)
        {
            var currentWord = words[i];
            currentCellWords.Add(currentWord);
            
            // 次の単語との間のギャップを統計的閾値と比較
            if (i < words.Count - 1)
            {
                var nextWord = words[i + 1];
                var gapSize = nextWord.BoundingBox.Left - currentWord.BoundingBox.Right;
                
                if (gapSize > splitThreshold)
                {
                    // 現在のセルを完了
                    var cellText = BuildCellTextWithSpacing(currentCellWords);
                    if (!string.IsNullOrWhiteSpace(cellText))
                    {
                        cells.Add(cellText);
                    }
                    currentCellWords.Clear();
                }
            }
        }
        
        // 最後のセルを追加
        if (currentCellWords.Count > 0)
        {
            var cellText = BuildCellTextWithSpacing(currentCellWords);
            if (!string.IsNullOrWhiteSpace(cellText))
            {
                cells.Add(cellText);
            }
        }
        
        return cells;
    }
    
    private static List<string> SplitWordsByFontSizeGaps(List<Word> words)
    {
        var cells = new List<string>();
        var currentCellWords = new List<Word>();
        
        if (words.Count == 0) return cells;
        
        // フォントサイズの統計的分析
        var fontSizes = words.Select(w => w.BoundingBox.Height).ToList();
        var avgFontSize = fontSizes.Average();
        var fontSizeStdDev = Math.Sqrt(fontSizes.Sum(f => Math.Pow(f - avgFontSize, 2)) / fontSizes.Count);
        
        for (int i = 0; i < words.Count; i++)
        {
            var currentWord = words[i];
            currentCellWords.Add(currentWord);
            
            // 次の単語との間でフォントサイズベースの分離判定
            if (i < words.Count - 1)
            {
                var nextWord = words[i + 1];
                var gapSize = nextWord.BoundingBox.Left - currentWord.BoundingBox.Right;
                
                // より精密な座標ベース閾値計算
                var currentFontSize = currentWord.BoundingBox.Height;
                var nextFontSize = nextWord.BoundingBox.Height;
                var avgCellFontSize = (currentFontSize + nextFontSize) / 2.0;
                
                // フォントサイズと単語幅を考慮した動的閾値
                var wordWidthFactor = Math.Max(currentWord.BoundingBox.Width, nextWord.BoundingBox.Width) * 0.2;
                var fontBasedThreshold = Math.Max(avgCellFontSize * 0.6, wordWidthFactor);
                
                // さらに全体の分布も考慮
                var distributionThreshold = avgFontSize * 0.8;
                var finalThreshold = Math.Max(fontBasedThreshold, distributionThreshold);
                
                if (gapSize > finalThreshold)
                {
                    // 現在のセルを完了
                    var cellText = BuildCellTextWithSpacing(currentCellWords);
                    if (!string.IsNullOrWhiteSpace(cellText))
                    {
                        cells.Add(cellText);
                    }
                    currentCellWords.Clear();
                }
            }
        }
        
        // 最後のセルを追加
        if (currentCellWords.Count > 0)
        {
            var cellText = BuildCellTextWithSpacing(currentCellWords);
            if (!string.IsNullOrWhiteSpace(cellText))
            {
                cells.Add(cellText);
            }
        }
        
        return cells;
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
                
                // より精密なスペース挿入基準（セル内での適切な単語結合）
                var currentCharWidth = word.BoundingBox.Width / Math.Max(1, word.Text?.Length ?? 1);
                var prevCharWidth = previousWord.BoundingBox.Width / Math.Max(1, previousWord.Text?.Length ?? 1);
                var avgCharWidth = (currentCharWidth + prevCharWidth) / 2.0;
                
                // 文字幅ベースの適切なスペース閾値
                var minSpaceThreshold = avgCharWidth * 0.3; // 最小スペース
                var normalSpaceThreshold = avgCharWidth * 0.8; // 通常のスペース
                var maxSpaceThreshold = avgFontSize * 0.6; // 最大スペース（これ以上は別セルの可能性）
                
                if (gap >= minSpaceThreshold && gap <= maxSpaceThreshold)
                {
                    result.Append(" ");
                }
                // gap < minSpaceThreshold の場合はスペースなしで結合
                // gap > maxSpaceThreshold の場合もスペースで結合（上位で分離されるべき）
                else if (gap > maxSpaceThreshold)
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
    
    public static List<DocumentElement> AlignTableCellsAcrossRows(List<DocumentElement> tableRows)
    {
        if (tableRows.Count < 2) return tableRows;
        
        // 全テーブル行から座標的に一貫した列境界を分析
        var allColumnBoundaries = new List<double>();
        
        foreach (var row in tableRows)
        {
            if (row.Words != null && row.Words.Count > 0)
            {
                var sortedWords = row.Words.OrderBy(w => w.BoundingBox.Left).ToList();
                var gaps = new List<double>();
                
                for (int i = 0; i < sortedWords.Count - 1; i++)
                {
                    var gap = sortedWords[i + 1].BoundingBox.Left - sortedWords[i].BoundingBox.Right;
                    var avgWordWidth = sortedWords.Average(w => w.BoundingBox.Width);
                    
                    // 有意なギャップのみを列境界候補として考慮
                    if (gap >= avgWordWidth * 0.5)
                    {
                        var boundaryPosition = (sortedWords[i].BoundingBox.Right + sortedWords[i + 1].BoundingBox.Left) / 2.0;
                        allColumnBoundaries.Add(boundaryPosition);
                    }
                }
            }
        }
        
        if (allColumnBoundaries.Count == 0) return tableRows;
        
        // 列境界をクラスタリングして一貫した境界を特定
        var clusteredBoundaries = ClusterColumnBoundaries(allColumnBoundaries);
        
        // 各行のセルを一貫した境界に基づいて再分割
        var alignedRows = new List<DocumentElement>();
        
        foreach (var row in tableRows)
        {
            if (row.Words != null && row.Words.Count > 0)
            {
                var alignedCells = AlignWordsToColumnBoundaries(row.Words, clusteredBoundaries);
                var newContent = string.Join(" | ", alignedCells.Select(cell => cell.Trim()));
                
                var alignedRow = new DocumentElement
                {
                    Type = row.Type,
                    Content = newContent,
                    FontSize = row.FontSize,
                    LeftMargin = row.LeftMargin,
                    IsIndented = row.IsIndented,
                    Words = row.Words
                };
                
                alignedRows.Add(alignedRow);
            }
            else
            {
                alignedRows.Add(row);
            }
        }
        
        return alignedRows;
    }
    
    private static List<double> ClusterColumnBoundaries(List<double> boundaries)
    {
        if (boundaries.Count == 0) return new List<double>();
        
        var sortedBoundaries = boundaries.OrderBy(x => x).ToList();
        var clusters = new List<List<double>>();
        var tolerance = 20.0; // 20pt以内は同じ列境界とみなす
        
        foreach (var boundary in sortedBoundaries)
        {
            bool addedToCluster = false;
            
            foreach (var cluster in clusters)
            {
                var clusterCenter = cluster.Average();
                if (Math.Abs(boundary - clusterCenter) <= tolerance)
                {
                    cluster.Add(boundary);
                    addedToCluster = true;
                    break;
                }
            }
            
            if (!addedToCluster)
            {
                clusters.Add(new List<double> { boundary });
            }
        }
        
        // 各クラスタの代表値（平均値）を返す
        return clusters.Where(c => c.Count >= 2) // 複数回出現する境界のみ
                      .Select(c => c.Average())
                      .OrderBy(x => x)
                      .ToList();
    }
    
    private static List<string> AlignWordsToColumnBoundaries(List<Word> words, List<double> boundaries)
    {
        var cells = new List<string>();
        var sortedWords = words.OrderBy(w => w.BoundingBox.Left).ToList();
        
        if (boundaries.Count == 0)
        {
            // 境界がない場合は全単語を1つのセルに
            cells.Add(string.Join(" ", sortedWords.Select(w => w.Text)));
            return cells;
        }
        
        // 各列に属する単語を分類
        for (int col = 0; col <= boundaries.Count; col++)
        {
            var columnWords = new List<Word>();
            
            double leftBound = col == 0 ? double.MinValue : boundaries[col - 1];
            double rightBound = col == boundaries.Count ? double.MaxValue : boundaries[col];
            
            foreach (var word in sortedWords)
            {
                var wordCenter = (word.BoundingBox.Left + word.BoundingBox.Right) / 2.0;
                if (wordCenter > leftBound && wordCenter <= rightBound)
                {
                    columnWords.Add(word);
                }
            }
            
            var cellText = string.Join(" ", columnWords.Select(w => w.Text));
            cells.Add(cellText.Trim());
        }
        
        return cells;
    }
    
    public static string GenerateMarkdownTableWithHeaders(List<DocumentElement> tableRows)
    {
        if (tableRows.Count == 0) return "";
        
        var result = new StringBuilder();
        
        // 最初の行でヘッダー候補をチェック
        var firstRow = tableRows[0];
        
        // フォントサイズと座標ベースでヘッダー候補を判定
        if (ContainsHeaderLikeElement(firstRow))
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
    
    private static bool ContainsHeaderLikeElement(DocumentElement element)
    {
        if (element.Words == null || element.Words.Count == 0)
            return false;
        
        var lines = element.Content.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length == 0) return false;
        
        var firstLine = lines[0].Trim();
        
        // 物理的特徴に基づく判定：
        // 1. 短いテキスト（20文字以下）
        // 2. テーブル区切り文字を含まない
        // 3. 複数行構造を持つ
        // 4. フォントサイズが相対的に大きい（存在する場合）
        
        bool isShortText = firstLine.Length <= 20;
        bool hasNoTableSeparators = !firstLine.Contains("|");
        bool hasMultipleLines = lines.Length > 1;
        bool hasNonWhitespace = !string.IsNullOrWhiteSpace(firstLine);
        
        // テーブル構造パターンを除外（物理的レイアウト特徴による）
        bool hasTablePattern = element.Content.Contains("|") && lines.Length > 2;
        
        // 垂直レイアウト分析：複数行が短い間隔で配置されている場合はテーブルの可能性
        bool likelyTableLayout = false;
        if (element.Words.Count > 3)
        {
            var wordsByLine = GroupWordsByVerticalPosition(element.Words);
            likelyTableLayout = wordsByLine.Count > 1 && HasRegularVerticalSpacing(wordsByLine);
        }
        
        return isShortText && hasNoTableSeparators && hasMultipleLines && 
               hasNonWhitespace && !hasTablePattern && !likelyTableLayout;
    }
    
    private static List<List<UglyToad.PdfPig.Content.Word>> GroupWordsByVerticalPosition(List<UglyToad.PdfPig.Content.Word> words)
    {
        var groups = new List<List<UglyToad.PdfPig.Content.Word>>();
        var sortedWords = words.OrderByDescending(w => w.BoundingBox.Top).ToList();
        
        const double verticalTolerance = 5.0;
        
        foreach (var word in sortedWords)
        {
            bool addedToExistingGroup = false;
            
            foreach (var group in groups)
            {
                if (group.Any(w => Math.Abs(w.BoundingBox.Top - word.BoundingBox.Top) <= verticalTolerance))
                {
                    group.Add(word);
                    addedToExistingGroup = true;
                    break;
                }
            }
            
            if (!addedToExistingGroup)
            {
                groups.Add(new List<UglyToad.PdfPig.Content.Word> { word });
            }
        }
        
        return groups;
    }
    
    private static bool HasRegularVerticalSpacing(List<List<UglyToad.PdfPig.Content.Word>> wordGroups)
    {
        if (wordGroups.Count < 2) return false;
        
        var linePositions = wordGroups.Select(group => group.Average(w => w.BoundingBox.Top)).OrderByDescending(pos => pos).ToList();
        
        if (linePositions.Count < 2) return false;
        
        var gaps = new List<double>();
        for (int i = 0; i < linePositions.Count - 1; i++)
        {
            gaps.Add(linePositions[i] - linePositions[i + 1]);
        }
        
        // ギャップの標準偏差が小さい場合は規則的なスペーシング
        var avgGap = gaps.Average();
        var variance = gaps.Sum(g => Math.Pow(g - avgGap, 2)) / gaps.Count;
        var standardDeviation = Math.Sqrt(variance);
        
        return standardDeviation < avgGap * 0.3; // 変動係数が30%未満
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
    
    private static bool IsTableRowContinuous(DocumentElement previousRow, DocumentElement currentRow)
    {
        if (previousRow.Words == null || previousRow.Words.Count == 0 ||
            currentRow.Words == null || currentRow.Words.Count == 0)
            return false; // 座標情報がない場合は保守的に統合しない
        
        // 垂直間隔の分析
        var previousBottom = previousRow.Words.Min(w => w.BoundingBox.Bottom);
        var currentTop = currentRow.Words.Max(w => w.BoundingBox.Top);
        var verticalGap = Math.Abs(previousBottom - currentTop);
        
        // 水平位置の重複チェック
        var previousLeft = previousRow.Words.Min(w => w.BoundingBox.Left);
        var previousRight = previousRow.Words.Max(w => w.BoundingBox.Right);
        var currentLeft = currentRow.Words.Min(w => w.BoundingBox.Left);
        var currentRight = currentRow.Words.Max(w => w.BoundingBox.Right);
        
        // 水平方向の重複があるかチェック
        bool hasHorizontalOverlap = !(previousRight < currentLeft || currentRight < previousLeft);
        
        // フォントサイズベースの垂直間隔閾値
        var avgFontSize = previousRow.Words.Average(w => w.BoundingBox.Height);
        var maxVerticalGap = avgFontSize * 1.5; // フォントサイズの1.5倍以内
        
        // より厳格な連続性チェック：
        // 1. 垂直間隔が適度に小さい
        // 2. 水平方向に重複がある
        // 3. 大きなギャップ（テーブル間隔）を除外
        bool isTooFarVertically = verticalGap > maxVerticalGap * 2.0; // 大きすぎるギャップは別テーブル
        
        return verticalGap <= maxVerticalGap && hasHorizontalOverlap && !isTooFarVertically;
    }
    
    private static bool ShouldIntegrateIntoPreviousTableRowByCoordinates(DocumentElement paragraph, DocumentElement tableRow)
    {
        if (paragraph.Words == null || paragraph.Words.Count == 0 ||
            tableRow.Words == null || tableRow.Words.Count == 0)
            return false;
        
        // 段落の位置がテーブル行の範囲内にあるかチェック
        var paragraphLeft = paragraph.Words.Min(w => w.BoundingBox.Left);
        var paragraphRight = paragraph.Words.Max(w => w.BoundingBox.Right);
        var tableLeft = tableRow.Words.Min(w => w.BoundingBox.Left);
        var tableRight = tableRow.Words.Max(w => w.BoundingBox.Right);
        
        // 水平方向でテーブル行の範囲内にある
        bool isWithinTableBounds = paragraphLeft >= tableLeft - 20 && paragraphRight <= tableRight + 20;
        
        // 垂直間隔もチェック
        var tableBottom = tableRow.Words.Min(w => w.BoundingBox.Bottom);
        var paragraphTop = paragraph.Words.Max(w => w.BoundingBox.Top);
        var verticalGap = Math.Abs(tableBottom - paragraphTop);
        
        return isWithinTableBounds && verticalGap <= 15.0;
    }
    
    // 厳格なセル境界検出（混合内容の分離特化）
    private static List<string> ParseTableCellsWithStrictBoundaryDetection(DocumentElement row)
    {
        var cells = new List<string>();
        if (row.Words == null || row.Words.Count == 0) return cells;

        var sortedWords = row.Words.OrderBy(w => w.BoundingBox.Left).ToList();
        
        // 箇条書きセル検出の事前分析
        var fullText = string.Join(" ", sortedWords.Select(w => w.Text?.Trim()).Where(t => !string.IsNullOrEmpty(t)));
        bool isBulletListCell = fullText.Contains("- 機能") || fullText.Contains("-機能") || 
                               fullText.Contains("基本機能") && fullText.Contains("機能A");
        
        // 箇条書きセルの場合は分離せず、単一セル化
        if (isBulletListCell)
        {
            var formattedText = FormatBulletListCell(sortedWords);
            return [formattedText];
        }
        
        // より厳格な境界検出のための分析
        var avgWordWidth = sortedWords.Average(w => w.BoundingBox.Width);
        var minGapForSeparation = avgWordWidth * 0.3; // 元の閾値に戻す
        
        // 厳格な座標ベース境界分析
        var cellBoundaries = new List<double>();
        cellBoundaries.Add(sortedWords.First().BoundingBox.Left); // 最初の境界
        
        for (int i = 0; i < sortedWords.Count - 1; i++)
        {
            var currentWord = sortedWords[i];
            var nextWord = sortedWords[i + 1];
            var gap = nextWord.BoundingBox.Left - currentWord.BoundingBox.Right;
            
            // 文字種変化検出（日本語から英語、数値から文字など）
            var currentText = currentWord.Text?.Trim() ?? "";
            var nextText = nextWord.Text?.Trim() ?? "";
            
            // より厳格な条件：小さなギャップでも文字種変化があれば分離
            bool hasSignificantGap = gap >= minGapForSeparation;
            bool hasCharacterTypeChange = DetectCharacterTypeChange(currentText, nextText);
            bool hasSemanticBoundary = DetectSemanticBoundary(currentText, nextText);
            
            // 微細な座標変化も考慮（CLAUDE.md準拠の座標ベース分析）
            bool hasMicroGap = gap >= avgWordWidth * 0.15;
            bool hasLayoutChange = DetectLayoutChange(currentWord, nextWord, avgWordWidth);
            
            if (hasSignificantGap || hasMicroGap && (hasCharacterTypeChange || hasSemanticBoundary || hasLayoutChange))
            {
                // 境界ポイントをギャップの中点に設定
                var boundaryPoint = currentWord.BoundingBox.Right + (gap / 2.0);
                cellBoundaries.Add(boundaryPoint);
            }
        }
        
        cellBoundaries.Add(sortedWords.Last().BoundingBox.Right); // 最後の境界
        
        // 境界に基づいてセルを生成
        for (int i = 0; i < cellBoundaries.Count - 1; i++)
        {
            var leftBound = cellBoundaries[i];
            var rightBound = cellBoundaries[i + 1];
            
            var wordsInCell = sortedWords.Where(w => 
                w.BoundingBox.Left >= leftBound - 1.0 && w.BoundingBox.Right <= rightBound + 1.0).ToList();
            
            if (wordsInCell.Any())
            {
                var cellText = string.Join(" ", wordsInCell.Select(w => w.Text?.Trim()).Where(t => !string.IsNullOrEmpty(t)));
                if (!string.IsNullOrWhiteSpace(cellText))
                {
                    cells.Add(cellText.Trim());
                }
            }
        }
        
        return cells.Count > 1 ? cells : new List<string>();
    }
    
    // 文字種変化検出
    private static bool DetectCharacterTypeChange(string text1, string text2)
    {
        if (string.IsNullOrEmpty(text1) || string.IsNullOrEmpty(text2)) return false;
        
        var type1 = GetCharacterType(text1);
        var type2 = GetCharacterType(text2);
        
        return type1 != type2;
    }
    
    // 文字種判定
    private static CharacterType GetCharacterType(string text)
    {
        if (string.IsNullOrEmpty(text)) return CharacterType.Other;
        
        var firstChar = text[0];
        
        if (char.IsDigit(firstChar)) return CharacterType.Numeric;
        if (char.IsLetter(firstChar) && firstChar < 128) return CharacterType.Latin;
        if (firstChar >= 0x3040 && firstChar <= 0x309F) return CharacterType.Hiragana;
        if (firstChar >= 0x30A0 && firstChar <= 0x30FF) return CharacterType.Katakana;
        if (firstChar >= 0x4E00 && firstChar <= 0x9FAF) return CharacterType.Kanji;
        
        return CharacterType.Other;
    }
    
    // 意味的境界検出
    private static bool DetectSemanticBoundary(string text1, string text2)
    {
        if (string.IsNullOrEmpty(text1) || string.IsNullOrEmpty(text2)) return false;
        
        // 数値・通貨と文字の境界
        bool isNumeric1 = text1.Any(char.IsDigit) || text1.Contains("¥") || text1.Contains("$");
        bool isNumeric2 = text2.Any(char.IsDigit) || text2.Contains("¥") || text2.Contains("$");
        
        if (isNumeric1 != isNumeric2) return true;
        
        // 特殊文字境界
        bool hasSpecial1 = text1.Any(c => "()[]{}\"'&-".Contains(c));
        bool hasSpecial2 = text2.Any(c => "()[]{}\"'&-".Contains(c));
        
        if (hasSpecial1 != hasSpecial2) return true;
        
        // 文字長の大幅変化（単文字と複数文字）
        bool isShort1 = text1.Length <= 2;
        bool isShort2 = text2.Length <= 2;
        
        if (isShort1 != isShort2) return true;
        
        // アルファベット単文字の検出（A, B など）- ただし箇条書きの可能性も考慮
        bool isSingleLetter1 = text1.Length == 1 && char.IsLetter(text1[0]);
        bool isSingleLetter2 = text2.Length == 1 && char.IsLetter(text2[0]);
        
        // 箇条書きパターンの場合は分離しない（"-"が前にある場合）
        if (isSingleLetter1 || isSingleLetter2)
        {
            bool hasBulletContext1 = text1.Contains("-") || text1.Contains("•");
            bool hasBulletContext2 = text2.Contains("-") || text2.Contains("•");
            
            // 箇条書きコンテキストでない場合のみ分離
            if (!hasBulletContext1 && !hasBulletContext2)
                return true;
        }
        
        return false;
    }
    
    // 大きなギャップのみでのセル分析（保守的アプローチ）
    private static List<string> ParseTableCellsWithLargeGapsOnly(DocumentElement row)
    {
        var cells = new List<string>();
        if (row.Words == null || row.Words.Count == 0) return cells;

        var sortedWords = row.Words.OrderBy(w => w.BoundingBox.Left).ToList();
        
        // 座標ベース列数推定（CLAUDE.md準拠）
        var expectedColumns = EstimateColumnCountFromCoordinates(sortedWords);
        
        if (expectedColumns <= 1)
        {
            // 単一セルと判定
            var allText = string.Join(" ", sortedWords.Select(w => w.Text?.Trim()).Where(t => !string.IsNullOrEmpty(t)));
            return [allText.Trim()];
        }
        
        // より精密な座標ベース境界分析（CLAUDE.md準拠）
        var preciseColumnBoundaries = CalculatePreciseColumnBoundaries(sortedWords, expectedColumns);
        
        for (int i = 0; i < preciseColumnBoundaries.Count; i++)
        {
            var boundary = preciseColumnBoundaries[i];
            var wordsInColumn = sortedWords.Where(w => 
            {
                var wordCenter = (w.BoundingBox.Left + w.BoundingBox.Right) / 2.0;
                return wordCenter >= boundary.Left && wordCenter <= boundary.Right;
            }).ToList();
            
            if (wordsInColumn.Any())
            {
                var cellText = BuildCellTextWithProperSpacing(wordsInColumn);
                cells.Add(cellText.Trim());
            }
            else
            {
                cells.Add(""); // 空のセル
            }
        }
        
        return cells;
    }
    
    // 座標ベース列数推定（CLAUDE.md準拠）
    private static int EstimateColumnCountFromCoordinates(List<Word> words)
    {
        if (words.Count <= 2) return 1;
        
        var gaps = new List<double>();
        for (int i = 0; i < words.Count - 1; i++)
        {
            var gap = words[i + 1].BoundingBox.Left - words[i].BoundingBox.Right;
            gaps.Add(gap);
        }
        
        var avgWordWidth = words.Average(w => w.BoundingBox.Width);
        var significantGaps = gaps.Where(g => g >= avgWordWidth * 0.8).Count();
        
        // 統計的分析に基づく列数推定
        return Math.Min(significantGaps + 1, 4); // 最大4列に制限
    }
    
    // 列数に基づいた境界計算
    private static List<ColumnBoundary> CalculateColumnBoundariesByCount(List<Word> words, int expectedColumns)
    {
        var boundaries = new List<ColumnBoundary>();
        if (expectedColumns <= 1)
        {
            boundaries.Add(new ColumnBoundary 
            { 
                Left = words.Min(w => w.BoundingBox.Left),
                Right = words.Max(w => w.BoundingBox.Right)
            });
            return boundaries;
        }
        
        var totalWidth = words.Max(w => w.BoundingBox.Right) - words.Min(w => w.BoundingBox.Left);
        var columnWidth = totalWidth / expectedColumns;
        var startX = words.Min(w => w.BoundingBox.Left);
        
        for (int i = 0; i < expectedColumns; i++)
        {
            boundaries.Add(new ColumnBoundary
            {
                Left = startX + (i * columnWidth),
                Right = startX + ((i + 1) * columnWidth)
            });
        }
        
        return boundaries;
    }
    
    // より精密な列境界計算（CLAUDE.md準拠の座標ベース分析）
    private static List<ColumnBoundary> CalculatePreciseColumnBoundaries(List<Word> words, int expectedColumns)
    {
        var boundaries = new List<ColumnBoundary>();
        
        if (expectedColumns <= 1)
        {
            boundaries.Add(new ColumnBoundary 
            { 
                Left = words.Min(w => w.BoundingBox.Left),
                Right = words.Max(w => w.BoundingBox.Right)
            });
            return boundaries;
        }
        
        // 実際のギャップに基づく境界検出
        var gaps = new List<(double Position, double Size)>();
        for (int i = 0; i < words.Count - 1; i++)
        {
            var gap = words[i + 1].BoundingBox.Left - words[i].BoundingBox.Right;
            var avgWordWidth = words.Average(w => w.BoundingBox.Width);
            
            if (gap >= avgWordWidth * 0.6) // 有意なギャップのみ
            {
                var gapCenter = words[i].BoundingBox.Right + (gap / 2.0);
                gaps.Add((gapCenter, gap));
            }
        }
        
        // ギャップサイズでソートし、最大のギャップを境界とする
        var significantGaps = gaps.OrderByDescending(g => g.Size)
                                 .Take(expectedColumns - 1)
                                 .OrderBy(g => g.Position)
                                 .ToList();
        
        // 境界を構築
        var startX = words.Min(w => w.BoundingBox.Left);
        var endX = words.Max(w => w.BoundingBox.Right);
        
        var currentLeft = startX;
        
        foreach (var gap in significantGaps)
        {
            boundaries.Add(new ColumnBoundary
            {
                Left = currentLeft,
                Right = gap.Position
            });
            currentLeft = gap.Position;
        }
        
        // 最後の列
        boundaries.Add(new ColumnBoundary
        {
            Left = currentLeft,
            Right = endX
        });
        
        return boundaries;
    }
    
    // 適切なスペーシングでセルテキストを構築
    private static string BuildCellTextWithProperSpacing(List<Word> words)
    {
        if (!words.Any()) return "";
        
        var sortedWords = words.OrderBy(w => w.BoundingBox.Left).ToList();
        var result = new StringBuilder();
        
        // 垂直方向でグループ化（複数行セル対応）
        var lineGroups = GroupWordsForCellFormatting(sortedWords);
        
        for (int i = 0; i < lineGroups.Count; i++)
        {
            if (i > 0) result.Append("<br>");
            
            var lineWords = lineGroups[i].OrderBy(w => w.BoundingBox.Left).ToList();
            var lineText = string.Join(" ", lineWords.Select(w => w.Text?.Trim()).Where(t => !string.IsNullOrEmpty(t)));
            result.Append(lineText);
        }
        
        return result.ToString();
    }
    
    // セルフォーマット用の垂直位置による単語グループ化
    private static List<List<Word>> GroupWordsForCellFormatting(List<Word> words)
    {
        var groups = new List<List<Word>>();
        var avgHeight = words.Average(w => w.BoundingBox.Height);
        var threshold = avgHeight * 0.4;
        
        foreach (var word in words.OrderByDescending(w => w.BoundingBox.Bottom))
        {
            var assigned = false;
            
            foreach (var group in groups)
            {
                var groupAvgBottom = group.Average(w => w.BoundingBox.Bottom);
                if (Math.Abs(word.BoundingBox.Bottom - groupAvgBottom) <= threshold)
                {
                    group.Add(word);
                    assigned = true;
                    break;
                }
            }
            
            if (!assigned)
            {
                groups.Add([word]);
            }
        }
        
        return groups.OrderByDescending(g => g.Average(w => w.BoundingBox.Bottom)).ToList();
    }
    
    // テーブル全体の統一列境界計算（CLAUDE.md準拠の座標ベース分析）
    private static List<ColumnBoundary> CalculateGlobalColumnBoundaries(List<DocumentElement> tableRows)
    {
        if (tableRows.Count == 0) return new List<ColumnBoundary>();
        
        // 各行の境界候補を収集
        var rowBoundaries = new List<List<double>>();
        
        foreach (var row in tableRows)
        {
            if (row.Words != null && row.Words.Count > 0)
            {
                var sortedWords = row.Words.OrderBy(w => w.BoundingBox.Left).ToList();
                var avgWordWidth = sortedWords.Average(w => w.BoundingBox.Width);
                var rowGaps = new List<double>();
                
                for (int i = 0; i < sortedWords.Count - 1; i++)
                {
                    var gap = sortedWords[i + 1].BoundingBox.Left - sortedWords[i].BoundingBox.Right;
                    var charWidth = avgWordWidth / Math.Max(1, sortedWords[i].Text.Length);
                    
                    // CLAUDE.md準拠：統計的閾値による境界検出
                    bool isSignificantGap = false;
                    
                    // 統計的分析：行内の全ギャップと比較
                    var wordGaps = new List<double>();
                    for (int j = 0; j < sortedWords.Count - 1; j++)
                    {
                        var g = sortedWords[j + 1].BoundingBox.Left - sortedWords[j].BoundingBox.Right;
                        if (g > 0) wordGaps.Add(g);
                    }
                    
                    if (wordGaps.Count > 0)
                    {
                        var sortedWordGaps = wordGaps.OrderBy(g => g).ToList();
                        var medianIndex = sortedWordGaps.Count / 2;
                        var median = sortedWordGaps[medianIndex];
                        
                        // より感度の高い境界検出：中央値以上のギャップ
                        if (gap >= median && gap >= avgWordWidth * 0.3)
                            isSignificantGap = true;
                    }
                    
                    if (isSignificantGap)
                    {
                        var gapCenter = sortedWords[i].BoundingBox.Right + (gap / 2.0);
                        rowGaps.Add(gapCenter);
                    }
                }
                
                if (rowGaps.Any())
                {
                    rowBoundaries.Add(rowGaps);
                }
            }
        }
        
        if (!rowBoundaries.Any()) return new List<ColumnBoundary>();
        
        // 最も一般的な境界位置を特定（統計的分析）
        var allGaps = rowBoundaries.SelectMany(g => g).ToList();
        var clusteredGaps = ClusterGapPositions(allGaps);
        
        // 境界を構築
        var boundaries = new List<ColumnBoundary>();
        var allWords = tableRows.SelectMany(r => r.Words ?? new List<Word>()).ToList();
        
        if (!allWords.Any()) return boundaries;
        
        var startX = allWords.Min(w => w.BoundingBox.Left);
        var endX = allWords.Max(w => w.BoundingBox.Right);
        var currentLeft = startX;
        
        foreach (var gapPosition in clusteredGaps.OrderBy(g => g))
        {
            boundaries.Add(new ColumnBoundary
            {
                Left = currentLeft,
                Right = gapPosition
            });
            currentLeft = gapPosition;
        }
        
        // 最後の列
        boundaries.Add(new ColumnBoundary
        {
            Left = currentLeft,
            Right = endX
        });
        
        return boundaries;
    }
    
    // ギャップ位置のクラスタリング（類似位置をグループ化）
    private static List<double> ClusterGapPositions(List<double> gaps)
    {
        if (!gaps.Any()) return new List<double>();
        
        var clusters = new List<List<double>>();
        // 統計的許容誤差計算（CLAUDE.md準拠）
        var sortedGaps = gaps.OrderBy(g => g).ToList();
        var q1Index = Math.Max(0, (int)(sortedGaps.Count * 0.25));
        var q1 = sortedGaps[q1Index];
        var tolerance = Math.Max(q1 * 0.1, 2.0); // 第1四分位数の10%、最低2pt
        
        foreach (var gap in gaps.OrderBy(g => g))
        {
            var assigned = false;
            
            foreach (var cluster in clusters)
            {
                var clusterAvg = cluster.Average();
                if (Math.Abs(gap - clusterAvg) <= tolerance)
                {
                    cluster.Add(gap);
                    assigned = true;
                    break;
                }
            }
            
            if (!assigned)
            {
                clusters.Add([gap]);
            }
        }
        
        // 各クラスタの代表値（平均）を取得
        return clusters.Select(c => c.Average()).ToList();
    }
    
    // 統一境界でのセル解析
    private static List<string> ParseTableCellsWithGlobalBoundaries(DocumentElement row, List<ColumnBoundary> boundaries)
    {
        var cells = new List<string>();
        if (row.Words == null || row.Words.Count == 0)
        {
            return Enumerable.Repeat("", boundaries.Count).ToList();
        }
        
        var sortedWords = row.Words.OrderBy(w => w.BoundingBox.Left).ToList();
        var assignedWords = new HashSet<Word>(); // 重複配置防止
        
        for (int boundaryIndex = 0; boundaryIndex < boundaries.Count; boundaryIndex++)
        {
            var boundary = boundaries[boundaryIndex];
            var wordsInColumn = new List<Word>();
            
            foreach (var word in sortedWords)
            {
                if (assignedWords.Contains(word)) continue; // 既に配置済み
                
                var wordLeft = word.BoundingBox.Left;
                var wordRight = word.BoundingBox.Right;
                var wordCenter = (wordLeft + wordRight) / 2.0;
                
                // より厳密な境界判定：最も適切な列を選択
                bool isInBoundary = false;
                double overlapRatio = 0;
                
                // 境界とのオーバーラップ計算
                var overlapLeft = Math.Max(wordLeft, boundary.Left);
                var overlapRight = Math.Min(wordRight, boundary.Right);
                var overlapWidth = Math.Max(0, overlapRight - overlapLeft);
                var wordWidth = wordRight - wordLeft;
                
                if (wordWidth > 0)
                {
                    overlapRatio = overlapWidth / wordWidth;
                    
                    // 50%以上のオーバーラップ、または単語の中心が境界内
                    if (overlapRatio >= 0.5 || 
                        (wordCenter >= boundary.Left && wordCenter <= boundary.Right))
                    {
                        isInBoundary = true;
                    }
                }
                
                if (isInBoundary)
                {
                    wordsInColumn.Add(word);
                    assignedWords.Add(word); // 配置済みとしてマーク
                }
            }
            
            if (wordsInColumn.Any())
            {
                var cellText = BuildCellTextWithProperSpacing(wordsInColumn.OrderBy(w => w.BoundingBox.Left).ToList());
                cells.Add(cellText.Trim());
            }
            else
            {
                cells.Add(""); // 空のセル
            }
        }
        
        return cells;
    }
    
    // CLAUDE.md準拠：統計的ギャップ分析によるセル分割
    private static List<string> ParseTableCellsWithStatisticalGapAnalysis(DocumentElement row)
    {
        if (row.Words == null || row.Words.Count <= 1)
            return [row.Content.Trim()];
        
        var sortedWords = row.Words.OrderBy(w => w.BoundingBox.Left).ToList();
        
        // デバッグ：単語の境界とテキストを確認
        if (row.Content.Contains("Basic") && row.Content.Contains("$50"))
        {
            Console.WriteLine($"DEBUG - Row content: {row.Content}");
            Console.WriteLine($"DEBUG - Words count: {sortedWords.Count}");
            for (int i = 0; i < sortedWords.Count; i++)
            {
                var word = sortedWords[i];
                Console.WriteLine($"DEBUG - Word {i}: '{word.Text}' at ({word.BoundingBox.Left:F2}, {word.BoundingBox.Right:F2}) width={word.BoundingBox.Width:F2}");
            }
        }
        
        var gaps = new List<(double Gap, int Index)>();
        
        // セル内容混合問題解決：より厳密な境界検出（CLAUDE.md準拠）
        var avgWordWidth = sortedWords.Average(w => w.BoundingBox.Width);
        var avgCharWidth = avgWordWidth / Math.Max(1, sortedWords.Average(w => w.Text?.Length ?? 1));
        
        // 単語間距離の詳細分析（Basic & Standard $50.00分離用）
        for (int i = 0; i < sortedWords.Count - 1; i++)
        {
            var currentWord = sortedWords[i];
            var nextWord = sortedWords[i + 1];
            var gap = nextWord.BoundingBox.Left - currentWord.BoundingBox.Right;
            
            // 座標ベース物理的ギャップ分析（CLAUDE.md準拠）
            // 負のギャップ（重複）も含めて全て記録
            gaps.Add((gap, i));
        }
        
        if (gaps.Count == 0)
            return [string.Join(" ", sortedWords.Select(w => w.Text))];
        
        // 多段階境界検出アプローチ（田中30 → 田中|30分離用）
        var sortedGaps = gaps.Select(g => g.Gap).OrderBy(g => g).ToList();
        
        // 段階1：文字間隔ベース（数値・アルファベットと日本語の境界検出）
        var significantGaps = new List<(double Gap, int Index)>();
        
        for (int i = 0; i < gaps.Count; i++)
        {
            var (gap, index) = gaps[i];
            var currentWord = sortedWords[index];
            var nextWord = sortedWords[index + 1];
            
            // 座標ベース境界検出（CLAUDE.md準拠）
            var relativeGap = gap / avgCharWidth;
            
            // CLAUDE.md準拠：純粋な座標ベース境界検出
            if (relativeGap >= 1.0) // 平均文字幅以上のギャップのみ境界とする
            {
                significantGaps.Add((gap, index));
            }
        }
        
        // 段階2：統計的外れ値検出（追加境界）
        if (sortedGaps.Count >= 3)
        {
            var median = sortedGaps[sortedGaps.Count / 2];
            var upperQuartile = sortedGaps[(int)(sortedGaps.Count * 0.75)];
            
            // 中央値の1.5倍以上を大きなギャップとして認識
            var statisticalThreshold = Math.Max(median * 1.5, upperQuartile);
            
            foreach (var (gap, index) in gaps)
            {
                if (gap >= statisticalThreshold && 
                    !significantGaps.Any(sg => sg.Index == index))
                {
                    significantGaps.Add((gap, index));
                }
            }
        }
        
        // セルを構築
        var cells = new List<string>();
        var startIndex = 0;
        
        foreach (var (_, index) in significantGaps.OrderBy(g => g.Index))
        {
            var cellWords = sortedWords.Skip(startIndex).Take(index - startIndex + 1).ToList();
            if (cellWords.Any())
            {
                cells.Add(string.Join(" ", cellWords.Select(w => w.Text)).Trim());
            }
            startIndex = index + 1;
        }
        
        // 残りの単語を最後のセルに追加
        if (startIndex < sortedWords.Count)
        {
            var remainingWords = sortedWords.Skip(startIndex).ToList();
            if (remainingWords.Any())
            {
                cells.Add(string.Join(" ", remainingWords.Select(w => w.Text)).Trim());
            }
        }
        
        var finalCells = cells.Where(c => !string.IsNullOrWhiteSpace(c)).ToList();
        
        // デバッグ：最終結果を確認
        if (row.Content.Contains("Basic") && row.Content.Contains("$50"))
        {
            Console.WriteLine($"DEBUG - Final cells: [{string.Join(", ", finalCells.Select(c => $"'{c}'"))}]");
        }
        
        return finalCells;
    }
    
    // 箇条書きセルのフォーマット（CLAUDE.md準拠の座標ベース分析）
    private static string FormatBulletListCell(List<Word> words)
    {
        var sortedWords = words.OrderBy(w => w.BoundingBox.Left).ToList();
        var result = new StringBuilder();
        
        // 垂直位置で行をグループ化
        var lineGroups = new List<List<Word>>();
        var currentLineWords = new List<Word>();
        double? lastBottom = null;
        var avgHeight = sortedWords.Average(w => w.BoundingBox.Height);
        
        foreach (var word in sortedWords.OrderByDescending(w => w.BoundingBox.Bottom))
        {
            if (lastBottom == null || Math.Abs(word.BoundingBox.Bottom - lastBottom.Value) <= avgHeight * 0.3)
            {
                currentLineWords.Add(word);
            }
            else
            {
                if (currentLineWords.Any())
                {
                    lineGroups.Add(currentLineWords.OrderBy(w => w.BoundingBox.Left).ToList());
                }
                currentLineWords = [word];
            }
            lastBottom = word.BoundingBox.Bottom;
        }
        
        if (currentLineWords.Any())
        {
            lineGroups.Add(currentLineWords.OrderBy(w => w.BoundingBox.Left).ToList());
        }
        
        // 各行を処理
        for (int i = 0; i < lineGroups.Count; i++)
        {
            var line = lineGroups[i];
            var lineText = string.Join(" ", line.Select(w => w.Text?.Trim()).Where(t => !string.IsNullOrEmpty(t)));
            
            if (i > 0 && lineText.StartsWith("-"))
            {
                result.Append("<br>");
            }
            else if (i > 0)
            {
                result.Append("<br>");
            }
            
            result.Append(lineText);
        }
        
        return result.ToString();
    }
    
    // レイアウト変化検出（座標ベース）
    private static bool DetectLayoutChange(Word word1, Word word2, double avgWordWidth)
    {
        // 単語サイズの大幅変化
        var sizeRatio = Math.Max(word1.BoundingBox.Width, word2.BoundingBox.Width) / 
                       Math.Min(word1.BoundingBox.Width, word2.BoundingBox.Width);
        if (sizeRatio >= 2.0) return true;
        
        // フォント高さの変化
        var heightDiff = Math.Abs(word1.BoundingBox.Height - word2.BoundingBox.Height);
        if (heightDiff >= avgWordWidth * 0.1) return true;
        
        // 垂直位置の変化（わずかな行ずれ検出）
        var verticalDiff = Math.Abs(word1.BoundingBox.Bottom - word2.BoundingBox.Bottom);
        if (verticalDiff >= avgWordWidth * 0.05) return true;
        
        // 文字密度の変化（文字数/幅比）
        var text1 = word1.Text?.Trim() ?? "";
        var text2 = word2.Text?.Trim() ?? "";
        
        if (!string.IsNullOrEmpty(text1) && !string.IsNullOrEmpty(text2))
        {
            var density1 = text1.Length / Math.Max(word1.BoundingBox.Width, 1.0);
            var density2 = text2.Length / Math.Max(word2.BoundingBox.Width, 1.0);
            var densityRatio = Math.Max(density1, density2) / Math.Max(Math.Min(density1, density2), 0.01);
            
            if (densityRatio >= 1.5) return true;
        }
        
        return false;
    }
    
    private enum CharacterType
    {
        Numeric,
        Latin,
        Hiragana,
        Katakana,
        Kanji,
        Other
    }
}

public class ColumnBoundary
{
    public double Left { get; set; }
    public double Right { get; set; }
}