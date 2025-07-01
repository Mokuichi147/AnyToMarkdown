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
        
        // 各行のセルを解析（グローバル境界を優先で一貫性確保）
        foreach (var row in tableRows)
        {
            var cells = new List<string>();
            
            // 優先順位1: グローバル境界による統一列構造（CLAUDE.md準拠）
            if (globalBoundaries.Any())
            {
                cells = ParseTableCellsWithGlobalBoundaries(row, globalBoundaries);
            }
            
            // 優先順位2: 統計的ギャップ分析（グローバル境界が失敗した場合のみ）
            if (cells.Count == 0 || cells.All(c => string.IsNullOrWhiteSpace(c)))
            {
                cells = ParseTableCellsWithStatisticalGapAnalysis(row);
            }
            
            // セル数の動的調整（CLAUDE.md準拠）
            // グローバル境界から期待される列数を計算
            var expectedColumns = globalBoundaries.Count;
            while (cells.Count < expectedColumns && expectedColumns > 0)
            {
                cells.Add("");
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
    
    private static List<List<Word>> GroupWordsByVerticalPosition(List<Word> words)
    {
        var groups = new List<List<Word>>();
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
                groups.Add(new List<Word> { word });
            }
        }
        
        return groups;
    }
    
    private static bool HasRegularVerticalSpacing(List<List<Word>> wordGroups)
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
                        
                        // CLAUDE.md準拠：精密境界検出（セル内容混合解決）
                        if (gap >= median * 0.8 && gap >= avgWordWidth * 0.5)
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
        
        // 段階1：座標ベース精密境界検出（CLAUDE.md準拠）
        var significantGaps = new List<(double Gap, int Index)>();
        
        for (int i = 0; i < gaps.Count; i++)
        {
            var (gap, index) = gaps[i];
            var currentWord = sortedWords[index];
            var nextWord = sortedWords[index + 1];
            
            // 単語の文字幅による動的閾値計算（CLAUDE.md準拠）
            var currentWordCharWidth = currentWord.BoundingBox.Width / Math.Max(1, currentWord.Text?.Length ?? 1);
            var nextWordCharWidth = nextWord.BoundingBox.Width / Math.Max(1, nextWord.Text?.Length ?? 1);
            var localCharWidth = (currentWordCharWidth + nextWordCharWidth) / 2.0;
            
            // 座標ベース境界検出（セル内容混合解決用）
            var relativeGap = gap / localCharWidth;
            
            // CLAUDE.md準拠：精密座標ベース境界検出（精密分離）
            if (relativeGap >= 0.3) // より精密な境界検出で年齢|職業分離
            {
                significantGaps.Add((gap, index));
            }
        }
        
        // 段階2：CLAUDE.md準拠の統計的外れ値検出（過分割防止のため保守的）
        if (sortedGaps.Count >= 3 && significantGaps.Count > 5) // 5分割以上で過分割判定
        {
            var median = sortedGaps[sortedGaps.Count / 2];
            var upperQuartile = sortedGaps[(int)(sortedGaps.Count * 0.75)];
            
            // 四分位数以上の大きなギャップのみ保持（適度な分割維持）
            var filteredGaps = significantGaps.Where(g => g.Gap >= upperQuartile * 0.8).ToList();
            
            if (filteredGaps.Count >= 3 && filteredGaps.Count <= 5) // 3-5列の適切な分割数
            {
                significantGaps = filteredGaps;
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
        
        // CLAUDE.md準拠：空列防止と最終フィルタリング
        var finalCells = cells.Where(c => !string.IsNullOrWhiteSpace(c) && c.Length > 0).ToList();
        
        // 統計的分析：過度に細分化されたセルの統合検討
        if (finalCells.Count > 6) // 6列を超える場合は過分割の可能性
        {
            // 隣接する短いセルを統合（座標ベース判定）
            var consolidatedCells = new List<string>();
            for (int i = 0; i < finalCells.Count; i++)
            {
                if (i < finalCells.Count - 1 && 
                    finalCells[i].Length <= 2 && finalCells[i + 1].Length <= 2)
                {
                    // 短いセル（記号など）は隣接セルと統合
                    consolidatedCells.Add($"{finalCells[i]} {finalCells[i + 1]}".Trim());
                    i++; // 次のセルをスキップ
                }
                else
                {
                    consolidatedCells.Add(finalCells[i]);
                }
            }
            finalCells = consolidatedCells;
        }
        
        return finalCells;
    }
}

public class ColumnBoundary
{
    public double Left { get; set; }
    public double Right { get; set; }
}