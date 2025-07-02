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
            else if (allElements[i].Type == ElementType.Paragraph && allElements[i].Content.Trim().Length <= 100) // 100文字まで拡大
            {
                // CLAUDE.md準拠：短い段落の統合判定を追加（分離セル内容対応）
                if (ShouldIntegrateIntoPreviousTableRowByCoordinates(allElements[i], consecutiveTableRows.Last()))
                {
                    var lastRow = consecutiveTableRows.Last();
                    
                    // 座標ベース統合
                    if (allElements[i].Words != null && allElements[i].Words.Count > 0 && 
                        lastRow.Words != null && lastRow.Words.Count > 0)
                    {
                        var combinedWords = new List<Word>(lastRow.Words);
                        combinedWords.AddRange(allElements[i].Words);
                        lastRow.Words = combinedWords;
                    }
                    
                    lastRow.Content = lastRow.Content + " " + allElements[i].Content.Trim();
                    continue; // 統合して続行
                }
                // 次のケースに遅して通常の段落処理を続行
            }
            else if (allElements[i].Type == ElementType.Header)
            {
                break; // ヘッダーが出現したら別のセクション
            }
            else if (allElements[i].Type == ElementType.Paragraph)
            {
                // CLAUDE.md準拠：段落の複数行テーブル統合判定
                if (ShouldIntegrateIntoPreviousTableRowByCoordinates(allElements[i], consecutiveTableRows.Last()))
                {
                    // 前のテーブル行に統合
                    var lastRow = consecutiveTableRows.Last();
                    
                    // 単語レベルでの統合：座標に基づく適切なセル配置
                    if (allElements[i].Words != null && allElements[i].Words.Count > 0 && 
                        lastRow.Words != null && lastRow.Words.Count > 0)
                    {
                        // CLAUDE.md準拠：座標ベース単語統合
                        var combinedWords = new List<Word>(lastRow.Words);
                        combinedWords.AddRange(allElements[i].Words);
                        lastRow.Words = combinedWords;
                    }
                    
                    // コンテンツも統合
                    lastRow.Content = lastRow.Content + " " + allElements[i].Content.Trim();
                }
                else
                {
                    break;
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
            // テーブル処理後の後続段落統合処理
            IntegrateSubsequentParagraphs(consecutiveTableRows, allElements, currentIndex);
            
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
        
        
        // CLAUDE.md準拠：完全に空の列を除去（強化版）
        allCells = RemoveEmptyColumns(allCells);
        
        // さらに、ほぼ空の列も削除（90%以上が空）
        allCells = RemoveMostlyEmptyColumns(allCells);
        
        maxColumns = allCells.Count > 0 ? allCells.Max(row => row.Count) : 0;
        
        if (maxColumns == 0) return "";
        
        // 統計的分析による最適列数決定
        var optimalColumnCount = DetermineOptimalColumnCount(allCells);
        if (optimalColumnCount > 0 && optimalColumnCount < maxColumns)
        {
            maxColumns = optimalColumnCount;
        }
        
        // 列数統一（最適列数まで）
        foreach (var row in allCells)
        {
            // 余分な列を削除
            while (row.Count > maxColumns)
            {
                row.RemoveAt(row.Count - 1);
            }
            
            // 不足する列を追加
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
    
    // CLAUDE.md準拠：座標ベーステーブル行連続性判定
    private static bool IsTableRowContinuous(DocumentElement previousRow, DocumentElement currentRow)
    {
        if (previousRow.Words == null || previousRow.Words.Count == 0 ||
            currentRow.Words == null || currentRow.Words.Count == 0)
            return false;
        
        // 垂直間隔の分析
        var previousBottom = previousRow.Words.Min(w => w.BoundingBox.Bottom);
        var currentTop = currentRow.Words.Max(w => w.BoundingBox.Top);
        var verticalGap = Math.Abs(previousBottom - currentTop);
        
        // 水平位置の重複チェック
        var previousLeft = previousRow.Words.Min(w => w.BoundingBox.Left);
        var previousRight = previousRow.Words.Max(w => w.BoundingBox.Right);
        var currentLeft = currentRow.Words.Min(w => w.BoundingBox.Left);
        var currentRight = currentRow.Words.Max(w => w.BoundingBox.Right);
        
        // 水平方向の重複計算
        var overlapLeft = Math.Max(previousLeft, currentLeft);
        var overlapRight = Math.Min(previousRight, currentRight);
        var overlapWidth = Math.Max(0, overlapRight - overlapLeft);
        var totalWidth = Math.Max(previousRight, currentRight) - Math.Min(previousLeft, currentLeft);
        
        bool hasSignificantOverlap = totalWidth > 0 && (overlapWidth / totalWidth) >= 0.2; // 0.3から0.2に緩和
        
        // 更なる柔軟性：水平位置の近接性も考慮
        bool isHorizontallyClose = Math.Abs(previousLeft - currentLeft) <= 50 || Math.Abs(previousRight - currentRight) <= 50;
        
        // フォントサイズベースの垂直間隔閾値（更に寛容に）
        var avgFontSize = previousRow.Words.Average(w => w.BoundingBox.Height);
        var maxVerticalGap = avgFontSize * 3.5; // 2.5倍から3.5倍にさらに拡大
        
        // CLAUDE.md準拠：より寛容な連続性判定
        return verticalGap <= maxVerticalGap && (hasSignificantOverlap || isHorizontallyClose);
    }
    
    // CLAUDE.md準拠：座標ベース段落統合判定
    private static bool ShouldIntegrateIntoPreviousTableRowByCoordinates(DocumentElement paragraph, DocumentElement tableRow)
    {
        if (paragraph.Words == null || paragraph.Words.Count == 0 ||
            tableRow.Words == null || tableRow.Words.Count == 0)
            return false;
        
        // CLAUDE.md準拠：座標範囲による統合判定
        var paragraphLeft = paragraph.Words.Min(w => w.BoundingBox.Left);
        var paragraphRight = paragraph.Words.Max(w => w.BoundingBox.Right);
        var tableLeft = tableRow.Words.Min(w => w.BoundingBox.Left);
        var tableRight = tableRow.Words.Max(w => w.BoundingBox.Right);
        
        // 水平方向の重複計算（より柔軟に）
        var overlapLeft = Math.Max(paragraphLeft, tableLeft);
        var overlapRight = Math.Min(paragraphRight, tableRight);
        var overlapWidth = Math.Max(0, overlapRight - overlapLeft);
        var paragraphWidth = paragraphRight - paragraphLeft;
        
        // 重複閾値を下げて、より多くの段落を統合対象にする
        bool hasSignificantOverlap = paragraphWidth > 0 && (overlapWidth / paragraphWidth) >= 0.2; // 0.4から0.2にさらに緩和
        
        // より寛容な範囲判定：段落がテーブル範囲内またはその近傍にあるか
        bool isWithinExtendedTableBounds = paragraphLeft >= tableLeft - 50 && paragraphRight <= tableRight + 50; // 30かぐ20に拡大
        
        // CLAUDE.md準拠：垂直間隔の統計的評価（より寛容に）
        var tableBottom = tableRow.Words.Min(w => w.BoundingBox.Bottom);
        var paragraphTop = paragraph.Words.Max(w => w.BoundingBox.Top);
        var verticalGap = Math.Abs(tableBottom - paragraphTop);
        
        // 動的垂直閾値を拡大（より多くの分離セルを統合）
        var avgRowHeight = tableRow.Words.Average(w => w.BoundingBox.Height);
        var maxVerticalGap = avgRowHeight * 5.0;  // 3.0から5.0にさらに拡大して分離セルを捕捉
        
        // 追加条件：水平位置の近接性を考慮
        bool isHorizontallyAligned = Math.Abs(paragraphLeft - tableLeft) <= 100; // より寛容な水平位置判定
        
        // より寛容な統合条件（いずれかの条件を満たす）
        bool shouldIntegrate = (hasSignificantOverlap || isWithinExtendedTableBounds || isHorizontallyAligned) && verticalGap <= maxVerticalGap;
        
        // 特別ケース：短いテキスト（30文字以下）の拡張統合条件
        if (!shouldIntegrate && paragraph.Content.Trim().Length <= 30)
        {
            // より寛容な垂直距離閾値（6倍まで拡張）
            bool isVeryClose = verticalGap <= avgRowHeight * 6.0;
            
            // 水平重複の最小条件をより緩く（Wordsプロパティを使用）
            var localParagraphLeft = paragraph.Words?.Min(w => w.BoundingBox.Left) ?? 0;
            var localTableLeft = tableRow.Words?.Min(w => w.BoundingBox.Left) ?? 0;
            var localParagraphRight = paragraph.Words?.Max(w => w.BoundingBox.Right) ?? 0;
            var localTableRight = tableRow.Words?.Max(w => w.BoundingBox.Right) ?? 0;
            var localAvgCharWidth = tableRow.Words?.Average(w => w.BoundingBox.Width) ?? 10.0;
            
            bool hasMinimalOverlap = overlapWidth > 0 || 
                                    Math.Abs(localParagraphLeft - localTableLeft) <= localAvgCharWidth * 2;
            
            // テーブル範囲内の位置判定
            bool isWithinTableHorizontalRange = 
                localParagraphLeft >= localTableLeft - localAvgCharWidth * 3 &&
                localParagraphRight <= localTableRight + localAvgCharWidth * 3;
            
            shouldIntegrate = isVeryClose && (hasMinimalOverlap || isWithinTableHorizontalRange);
        }
        
        return shouldIntegrate;
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
    
    // CLAUDE.md準拠：垂直位置による多行セル単語グループ化
    private static List<List<Word>> GroupWordsForCellFormatting(List<Word> words)
    {
        if (words.Count <= 1) return words.Select(w => new List<Word> { w }).ToList();
        
        var groups = new List<List<Word>>();
        var avgHeight = words.Average(w => w.BoundingBox.Height);
        
        // CLAUDE.md準拠：統計的垂直閾値計算
        var verticalPositions = words.Select(w => w.BoundingBox.Bottom).ToList();
        var variance = CalculateVariance(verticalPositions);
        var dynamicThreshold = Math.Max(avgHeight * 0.4, Math.Sqrt(variance) * 0.5);
        
        foreach (var word in words.OrderByDescending(w => w.BoundingBox.Bottom))
        {
            var assigned = false;
            
            foreach (var group in groups)
            {
                var groupAvgBottom = group.Average(w => w.BoundingBox.Bottom);
                if (Math.Abs(word.BoundingBox.Bottom - groupAvgBottom) <= dynamicThreshold)
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
    
    // CLAUDE.md準拠：分散計算
    private static double CalculateVariance(List<double> values)
    {
        if (values.Count <= 1) return 0;
        
        var mean = values.Average();
        return values.Sum(v => Math.Pow(v - mean, 2)) / values.Count;
    }
    
    // CLAUDE.md準拠：垂直共通空白領域による列境界検出
    private static List<ColumnBoundary> CalculateGlobalColumnBoundaries(List<DocumentElement> tableRows)
    {
        if (tableRows.Count == 0) return new List<ColumnBoundary>();
        
        // 新アプローチ：垂直共通空白領域の検出
        var verticalGaps = FindVerticalCommonGaps(tableRows);
        if (verticalGaps.Count > 0)
        {
            return BuildBoundariesFromVerticalGaps(verticalGaps, tableRows);
        }
        
        // フォールバック：行間一貫性を優先した境界検出
        var consistentBoundaries = CalculateConsistentColumnBoundaries(tableRows);
        if (consistentBoundaries.Count > 0)
        {
            return consistentBoundaries;
        }
        
        // フォールバック：列配置統一性分析
        var alignmentBoundaries = CalculateColumnBoundariesByAlignment(tableRows);
        if (alignmentBoundaries.Count > 0)
        {
            return alignmentBoundaries;
        }
        
        // 従来の方法にフォールバック：CLAUDE.md準拠の垂直座標分析
        var allWords = tableRows.SelectMany(r => r.Words ?? new List<Word>())
                               .Where(w => !string.IsNullOrEmpty(w.Text))
                               .ToList();
        
        if (allWords.Count == 0) return new List<ColumnBoundary>();
        
        // 新機能：垂直座標クラスタリングによる列境界候補の抽出
        var verticalBoundaries = ExtractVerticalColumnBoundaries(allWords);
        
        // 従来のギャップベース分析と垂直座標分析の統合
        var rowBoundaries = new List<List<double>>();
        var globalAvgWordWidth = allWords.Average(w => w.BoundingBox.Width);
        
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
                        
                        // CLAUDE.md準拠：統計的分散分析による境界検出（セル内容混合解決）
                        // 四分位数に基づく統計的閾値：より精密な境界検出
                        var q75Index = Math.Min(wordGaps.Count - 1, (int)(wordGaps.Count * 0.75));
                        var q75 = sortedWordGaps[q75Index];
                        var q25Index = (int)(wordGaps.Count * 0.25);
                        var q25 = sortedWordGaps[q25Index];
                        var iqr = q75 - q25;
                        
                        // CLAUDE.md準拠：バランスの取れた境界検出改善
                        // 目標：語句間自然区切りを除外しつつ、必要な列境界を保持
                        // 方法：統計的分析と座標ベース解析の両立
                        
                        // CLAUDE.md準拠：段階的境界検出で精度とリコールのバランスを取る
                        // フェーズ1：明らかな列境界を識別
                        // フェーズ2：語句間区切りをフィルタリング
                        
                        // CLAUDE.md準拠：バランスの取れた統計的境界検出
                        // 問題：過度に保守的で意味のある列境界を統合してしまう
                        // 解決：統計分析と座標ベース分析のバランス改善
                        
                        // 改善1：適応的統計閾値（精度とリコールのバランス）
                        var moderateThreshold = q25 + (iqr * 0.4); // Q1から40%（適度に緩和）
                        var isModerateOutlier = gap >= moderateThreshold;
                        
                        // 改善2：多段階文字幅判定（柔軟性向上）
                        var localCharWidth = avgWordWidth / Math.Max(sortedWords[i].Text.Length, 1);
                        var normalizedGap = gap / Math.Max(localCharWidth, 5.0);
                        var isRelativelySignificant = normalizedGap >= 1.2; // 文字幅の1.2倍以上（緩和）
                        var isStronglySignificant = normalizedGap >= 2.5; // 強い境界指標
                        
                        // 改善3：段階的境界判定（柔軟性確保）
                        var isLargerThanMedian = gap > median * 1.2; // 中央値の120%以上（緩和）
                        var isAbsolutelySignificant = gap >= avgWordWidth * 0.6; // 語幅の60%以上（緩和）
                        var isVerySignificant = gap >= avgWordWidth * 1.2; // 非常に大きなギャップ
                        
                        // CLAUDE.md準拠：多段階境界検出（精度とリコールのバランス）
                        if (isStronglySignificant || isVerySignificant || 
                            (isModerateOutlier && isLargerThanMedian) || 
                            (isRelativelySignificant && isAbsolutelySignificant))
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
        var tableWords = tableRows.SelectMany(r => r.Words ?? new List<Word>()).ToList();
        
        if (!tableWords.Any()) return boundaries;
        
        var startX = tableWords.Min(w => w.BoundingBox.Left);
        var endX = tableWords.Max(w => w.BoundingBox.Right);
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
        
        // 垂直座標境界との統合で精度向上
        return IntegrateVerticalAndGapBoundaries(boundaries, verticalBoundaries, allWords);
    }
    
    // CLAUDE.md準拠：垂直座標クラスタリングによる列境界抽出
    private static List<double> ExtractVerticalColumnBoundaries(List<Word> allWords)
    {
        if (allWords.Count == 0) return new List<double>();
        
        // 全単語の左端座標を収集
        var leftCoordinates = allWords.Select(w => w.BoundingBox.Left).ToList();
        
        // 座標クラスタリング：近い座標をグループ化
        var clusters = PerformCoordinateClustering(leftCoordinates);
        
        // 各クラスタの代表座標（平均）を境界候補とする
        var boundaries = clusters.Select(cluster => cluster.Average()).OrderBy(x => x).ToList();
        
        // フィルタリング：最小間隔以下の境界を除去
        var filteredBoundaries = new List<double>();
        var minBoundaryGap = allWords.Average(w => w.BoundingBox.Width) * 0.5; // 文字幅の50%を最小間隔
        
        for (int i = 0; i < boundaries.Count; i++)
        {
            if (i == 0 || boundaries[i] - filteredBoundaries.Last() >= minBoundaryGap)
            {
                filteredBoundaries.Add(boundaries[i]);
            }
        }
        
        return filteredBoundaries;
    }
    
    // CLAUDE.md準拠：座標クラスタリング（統計的距離分析）
    private static List<List<double>> PerformCoordinateClustering(List<double> coordinates)
    {
        if (coordinates.Count == 0) return new List<List<double>>();
        
        var sortedCoords = coordinates.OrderBy(x => x).ToList();
        var clusters = new List<List<double>>();
        
        // 座標間距離の統計分析
        var distances = new List<double>();
        for (int i = 0; i < sortedCoords.Count - 1; i++)
        {
            distances.Add(sortedCoords[i + 1] - sortedCoords[i]);
        }
        
        // クラスタリング閾値：距離の中央値
        var clusterThreshold = distances.Any() ? distances.OrderBy(d => d).ElementAt(distances.Count / 2) * 2.0 : 10.0;
        
        // クラスタリング実行
        var currentCluster = new List<double> { sortedCoords[0] };
        
        for (int i = 1; i < sortedCoords.Count; i++)
        {
            if (sortedCoords[i] - sortedCoords[i - 1] <= clusterThreshold)
            {
                currentCluster.Add(sortedCoords[i]);
            }
            else
            {
                clusters.Add(currentCluster);
                currentCluster = new List<double> { sortedCoords[i] };
            }
        }
        
        clusters.Add(currentCluster);
        return clusters;
    }
    
    // CLAUDE.md準拠：垂直境界とギャップ境界の統合
    private static List<ColumnBoundary> IntegrateVerticalAndGapBoundaries(
        List<ColumnBoundary> gapBoundaries, 
        List<double> verticalBoundaries, 
        List<Word> allWords)
    {
        if (allWords.Count == 0) return gapBoundaries;
        
        var startX = allWords.Min(w => w.BoundingBox.Left);
        var endX = allWords.Max(w => w.BoundingBox.Right);
        
        // 垂直境界を基準としたより精密な境界設定
        var boundaries = new List<ColumnBoundary>();
        var currentLeft = startX;
        
        foreach (var boundary in verticalBoundaries.OrderBy(x => x))
        {
            if (boundary > currentLeft)
            {
                boundaries.Add(new ColumnBoundary
                {
                    Left = currentLeft,
                    Right = boundary
                });
                currentLeft = boundary;
            }
        }
        
        // 最後の列
        if (currentLeft < endX)
        {
            boundaries.Add(new ColumnBoundary
            {
                Left = currentLeft,
                Right = endX
            });
        }
        
        // 垂直境界が不十分な場合は従来のギャップベース境界を使用
        return boundaries.Count >= 2 ? boundaries : gapBoundaries;
    }
    
    // ギャップ位置のクラスタリング（類似位置をグループ化）
    private static List<double> ClusterGapPositions(List<double> gaps)
    {
        if (!gaps.Any()) return new List<double>();
        
        var clusters = new List<List<double>>();
        // 統計的許容誤差計算（CLAUDE.md準拠）
        var sortedGaps = gaps.OrderBy(g => g).ToList();
        var mean = gaps.Average();
        var variance = gaps.Sum(g => Math.Pow(g - mean, 2)) / gaps.Count;
        var stdDev = Math.Sqrt(variance);
        
        // 標準偏差ベースの動的許容誤差（統合精度のバランス調整）
        var tolerance = Math.Max(stdDev * 0.4, Math.Max(mean * 0.06, 4.0)); // バランス調整
        
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
    
    // CLAUDE.md準拠：完全新アプローチ - 物理的レイアウト分析による境界検出
    private static List<string> ParseTableCellsWithStatisticalGapAnalysis(DocumentElement row)
    {
        if (row.Words == null || row.Words.Count <= 1)
            return [row.Content.Trim()];
        
        var sortedWords = row.Words.OrderBy(w => w.BoundingBox.Left).ToList();
        if (sortedWords.Count == 1)
            return [sortedWords[0].Text ?? ""];
        
        // CLAUDE.md準拠：新アプローチ - 物理的配置パターン分析
        var cells = AnalyzePhysicalWordLayout(sortedWords);
        
        return cells.Count > 0 ? cells : [string.Join(" ", sortedWords.Select(w => w.Text))];
    }
    
    // CLAUDE.md準拠：物理的単語配置パターンの統計分析
    private static List<string> AnalyzePhysicalWordLayout(List<Word> sortedWords)
    {
        // フェーズ1：基本座標統計の計算
        var wordPositions = sortedWords.Select(w => w.BoundingBox.Left).ToList();
        var wordWidths = sortedWords.Select(w => w.BoundingBox.Width).ToList();
        var avgWordWidth = wordWidths.Average();
        
        // フェーズ2：単語間距離の統計分析
        var interWordDistances = new List<double>();
        for (int i = 0; i < sortedWords.Count - 1; i++)
        {
            var distance = sortedWords[i + 1].BoundingBox.Left - sortedWords[i].BoundingBox.Right;
            interWordDistances.Add(Math.Max(0, distance));
        }
        
        if (interWordDistances.Count == 0) 
            return [string.Join(" ", sortedWords.Select(w => w.Text))];
        
        // フェーズ3：距離分布の四分位分析
        var sortedDistances = interWordDistances.OrderBy(d => d).ToList();
        var median = GetMedian(sortedDistances);
        var q75 = GetPercentile(sortedDistances, 0.75);
        var q90 = GetPercentile(sortedDistances, 0.90);
        
        // フェーズ4：列境界の判定閾値計算
        var baseThreshold = Math.Max(avgWordWidth * 0.3, median * 1.5);
        var strongThreshold = Math.Max(q75, avgWordWidth * 0.8);
        
        // フェーズ5：境界候補の抽出
        var boundaries = new List<int>();
        for (int i = 0; i < interWordDistances.Count; i++)
        {
            var distance = interWordDistances[i];
            
            // 強い境界指標
            if (distance >= strongThreshold)
            {
                boundaries.Add(i);
            }
            // 中程度の境界指標（追加条件付き）
            else if (distance >= baseThreshold && distance >= q75)
            {
                boundaries.Add(i);
            }
        }
        
        // フェーズ6：境界数の最適化
        if (boundaries.Count > 5)
        {
            // 最も大きい距離の上位5つのみ保持
            var boundaryDistances = boundaries.Select(i => new { Index = i, Distance = interWordDistances[i] })
                                            .OrderByDescending(x => x.Distance)
                                            .Take(5)
                                            .OrderBy(x => x.Index)
                                            .Select(x => x.Index)
                                            .ToList();
            boundaries = boundaryDistances;
        }
        
        // フェーズ7：セル構築
        var cells = new List<string>();
        var startIndex = 0;
        
        foreach (var boundaryIndex in boundaries)
        {
            var cellWords = sortedWords.Skip(startIndex).Take(boundaryIndex - startIndex + 1).ToList();
            if (cellWords.Any())
            {
                var cellText = string.Join(" ", cellWords.Select(w => w.Text?.Trim()).Where(t => !string.IsNullOrEmpty(t)));
                if (!string.IsNullOrWhiteSpace(cellText))
                {
                    cells.Add(cellText);
                }
            }
            startIndex = boundaryIndex + 1;
        }
        
        // 残りの単語
        if (startIndex < sortedWords.Count)
        {
            var remainingWords = sortedWords.Skip(startIndex).ToList();
            var cellText = string.Join(" ", remainingWords.Select(w => w.Text?.Trim()).Where(t => !string.IsNullOrEmpty(t)));
            if (!string.IsNullOrWhiteSpace(cellText))
            {
                cells.Add(cellText);
            }
        }
        
        return cells;
    }
    
    // 統計計算ヘルパー関数
    private static double GetMedian(List<double> sortedValues)
    {
        if (sortedValues.Count == 0) return 0;
        int mid = sortedValues.Count / 2;
        return sortedValues.Count % 2 == 0 
            ? (sortedValues[mid - 1] + sortedValues[mid]) / 2.0 
            : sortedValues[mid];
    }
    
    private static double GetPercentile(List<double> sortedValues, double percentile)
    {
        if (sortedValues.Count == 0) return 0;
        int index = (int)(sortedValues.Count * percentile);
        return sortedValues[Math.Min(index, sortedValues.Count - 1)];
    }
    
    // CLAUDE.md準拠：統計的文字幅計算
    private static double CalculateAverageCharacterWidth(List<Word> words)
    {
        var charWidths = new List<double>();
        
        foreach (var word in words)
        {
            if (!string.IsNullOrEmpty(word.Text) && word.BoundingBox.Width > 0)
            {
                var charWidth = word.BoundingBox.Width / word.Text.Length;
                charWidths.Add(charWidth);
            }
        }
        
        return charWidths.Any() ? charWidths.Average() : 8.0;
    }
    
    // CLAUDE.md準拠：統計的有意ギャップ検出
    private static List<(double Gap, int Index)> DetectStatisticallySignificantGaps(
        List<(double Gap, int Index)> gaps, 
        double avgCharWidth)
    {
        var significantGaps = new List<(double Gap, int Index)>();
        
        // 統計的閾値計算
        var gapValues = gaps.Select(g => g.Gap).Where(g => g > 0).ToList();
        if (!gapValues.Any()) return significantGaps;
        
        var mean = gapValues.Average();
        var variance = gapValues.Sum(g => Math.Pow(g - mean, 2)) / gapValues.Count;
        var stdDev = Math.Sqrt(variance);
        
        // CLAUDE.md準拠：より保守的な閾値（必要な境界を保持）
        var baseThreshold = avgCharWidth * 1.4;  // 1.2から1.4に調整  // 文字幅ベース閾値を下げる
        var statisticalThreshold = mean + stdDev * 0.7;  // 0.5から0.7に調整  // 統計閾値も下げる
        var adaptiveThreshold = Math.Min(baseThreshold, statisticalThreshold);
        
        foreach (var gap in gaps)
        {
            if (gap.Gap >= adaptiveThreshold)
            {
                significantGaps.Add(gap);
            }
        }
        
        return significantGaps;
    }
    
    // CLAUDE.md準拠：過分割の統計的フィルタリング
    private static List<(double Gap, int Index)> FilterExcessiveSegmentation(
        List<(double Gap, int Index)> significantGaps)
    {
        // より多くの境界を保持（上位70%）
        var sortedGaps = significantGaps.OrderByDescending(g => g.Gap).ToList();
        var keepCount = Math.Max(1, (int)(sortedGaps.Count * 0.7));
        
        return sortedGaps.Take(keepCount).OrderBy(g => g.Index).ToList();
    }
    
    
    // CLAUDE.md準拠：統計的分析による最適列数決定
    private static int DetermineOptimalColumnCount(List<List<string>> allCells)
    {
        if (allCells.Count == 0) return 0;
        
        // 各行の非空列数を計算
        var nonEmptyColumnCounts = allCells.Select(row => 
            row.Count(cell => !string.IsNullOrWhiteSpace(cell))).ToList();
        
        // 統計的分析
        var validCounts = nonEmptyColumnCounts.Where(count => count > 1).ToList();
        if (validCounts.Count == 0) return 0;
        
        // 最頻値を計算
        var mostCommonCount = validCounts
            .GroupBy(x => x)
            .OrderByDescending(g => g.Count())
            .ThenByDescending(g => g.Key) // 同頻度の場合は大きい値を選択
            .FirstOrDefault()?.Key ?? 0;
        
        // 中央値を計算
        var sortedCounts = validCounts.OrderBy(x => x).ToList();
        var median = sortedCounts.Count % 2 == 0
            ? (sortedCounts[sortedCounts.Count / 2 - 1] + sortedCounts[sortedCounts.Count / 2]) / 2.0
            : sortedCounts[sortedCounts.Count / 2];
        
        // 最頻値と中央値の調和を取る
        var harmonicMean = 2.0 * mostCommonCount * median / (mostCommonCount + median);
        
        // 最適値は最頻値を基準とし、適度な範囲内に制限
        var optimalCount = Math.Max(mostCommonCount, (int)Math.Round(median));
        return Math.Max(2, Math.Min(6, optimalCount)); // 2-6列の範囲内
    }
    
    // CLAUDE.md準拠：完全に空の列を除去
    private static List<List<string>> RemoveEmptyColumns(List<List<string>> allCells)
    {
        if (allCells.Count == 0) return allCells;
        
        var maxColumns = allCells.Max(row => row.Count);
        var columnsToKeep = new List<bool>();
        
        // 各列が空かどうかを判定
        for (int col = 0; col < maxColumns; col++)
        {
            bool hasNonEmptyContent = false;
            foreach (var row in allCells)
            {
                if (col < row.Count && !string.IsNullOrWhiteSpace(row[col]))
                {
                    hasNonEmptyContent = true;
                    break;
                }
            }
            columnsToKeep.Add(hasNonEmptyContent);
        }
        
        // 空でない列のみを保持
        var result = new List<List<string>>();
        foreach (var row in allCells)
        {
            var newRow = new List<string>();
            for (int col = 0; col < row.Count && col < columnsToKeep.Count; col++)
            {
                if (columnsToKeep[col])
                {
                    newRow.Add(row[col]);
                }
            }
            if (newRow.Count > 0) // 完全に空の行は除外
            {
                result.Add(newRow);
            }
        }
        
        return result;
    }
    
    // CLAUDE.md準拠：ほぼ空の列を除去（90%以上が空）
    private static List<List<string>> RemoveMostlyEmptyColumns(List<List<string>> allCells)
    {
        if (allCells.Count == 0) return allCells;
        
        var maxColumns = allCells.Max(row => row.Count);
        var columnsToKeep = new List<bool>();
        
        // 各列の内容密度を分析
        for (int col = 0; col < maxColumns; col++)
        {
            int nonEmptyCount = 0;
            int totalCount = 0;
            
            foreach (var row in allCells)
            {
                if (col < row.Count)
                {
                    totalCount++;
                    if (!string.IsNullOrWhiteSpace(row[col]))
                    {
                        nonEmptyCount++;
                    }
                }
            }
            
            // バランスの取れた空列除去：10%以上の内容がある列を保持
            double contentRatio = totalCount > 0 ? (double)nonEmptyCount / totalCount : 0;
            
            // 特別条件：完全に空または単一文字のみの列を除外
            bool hasOnlyMinimalContent = true;
            bool hasAnyMeaningfulContent = false;
            
            foreach (var row in allCells)
            {
                if (col < row.Count && !string.IsNullOrWhiteSpace(row[col]))
                {
                    var content = row[col].Trim();
                    if (content.Length > 0) // 何らかの内容があるか
                    {
                        hasAnyMeaningfulContent = true;
                        if (content.Length > 1) // 2文字以上の内容があるか
                        {
                            hasOnlyMinimalContent = false;
                        }
                    }
                }
            }
            
            // 改良された列保持判定：より精密な空列除去
            bool shouldKeep = false;
            
            if (contentRatio == 0)
            {
                // 完全に空の列は除外
                shouldKeep = false;
            }
            else if (hasOnlyMinimalContent && contentRatio < 0.1)
            {
                // 単一文字のみで内容率が低い列は除外
                shouldKeep = false;
            }
            else if (contentRatio >= 0.15 || (contentRatio >= 0.08 && hasAnyMeaningfulContent))
            {
                // 15%以上の内容があるか、8%以上で意味のある内容がある列を保持
                shouldKeep = true;
            }
            else
            {
                shouldKeep = false;
            }
            
            columnsToKeep.Add(shouldKeep);
        }
        
        // 内容密度の高い列のみ保持
        var result = new List<List<string>>();
        foreach (var row in allCells)
        {
            var newRow = new List<string>();
            for (int col = 0; col < row.Count && col < columnsToKeep.Count; col++)
            {
                if (columnsToKeep[col])
                {
                    newRow.Add(row[col]);
                }
            }
            if (newRow.Count > 0)
            {
                result.Add(newRow);
            }
        }
        
        return result;
    }
    
    // CLAUDE.md準拠：列配置統一性に基づく座標ベース境界検出
    private static List<ColumnBoundary> CalculateColumnBoundariesByAlignment(List<DocumentElement> tableRows)
    {
        if (tableRows.Count < 2) return new List<ColumnBoundary>();
        
        var allWords = tableRows.SelectMany(r => r.Words ?? new List<Word>())
                               .Where(w => !string.IsNullOrEmpty(w.Text?.Trim()))
                               .ToList();
        
        if (allWords.Count < 3) return new List<ColumnBoundary>();
        
        // 各行の単語を左端座標でグループ化
        var rowWordGroups = new List<List<Word>>();
        foreach (var row in tableRows)
        {
            if (row.Words != null && row.Words.Count > 0)
            {
                var sortedWords = row.Words.Where(w => !string.IsNullOrEmpty(w.Text?.Trim()))
                                          .OrderBy(w => w.BoundingBox.Left)
                                          .ToList();
                if (sortedWords.Count > 0)
                {
                    rowWordGroups.Add(sortedWords);
                }
            }
        }
        
        if (rowWordGroups.Count < 2) return new List<ColumnBoundary>();
        
        // 列配置パターン分析：各行の左端位置のクラスタリング
        var leftPositions = new List<List<double>>();
        
        foreach (var rowWords in rowWordGroups)
        {
            var positions = rowWords.Select(w => w.BoundingBox.Left).ToList();
            leftPositions.Add(positions);
        }
        
        // 共通の左端位置を統計的に特定（列配置の統一性を利用）
        var commonLeftPositions = FindCommonLeftPositions(leftPositions);
        
        if (commonLeftPositions.Count < 2) return new List<ColumnBoundary>();
        
        // 共通位置から列境界を構築
        return BuildColumnBoundariesFromPositions(commonLeftPositions, allWords);
    }
    
    // CLAUDE.md準拠：共通の左端位置を統計的に特定
    private static List<double> FindCommonLeftPositions(List<List<double>> leftPositions)
    {
        if (leftPositions.Count == 0) return new List<double>();
        
        // 全ての位置を統合してクラスタリング
        var allPositions = leftPositions.SelectMany(positions => positions)
                                       .OrderBy(pos => pos)
                                       .ToList();
        
        if (allPositions.Count == 0) return new List<double>();
        
        // クラスタリング闾値：平均文字幅の1.0倍（より細かく）
        var avgCharWidth = CalculateAverageCharWidth(allPositions);
        var clusterThreshold = avgCharWidth * 1.0;
        
        // 位置クラスタリング
        var clusters = new List<List<double>>();
        var currentCluster = new List<double> { allPositions[0] };
        
        for (int i = 1; i < allPositions.Count; i++)
        {
            if (allPositions[i] - allPositions[i - 1] <= clusterThreshold)
            {
                currentCluster.Add(allPositions[i]);
            }
            else
            {
                clusters.Add(currentCluster);
                currentCluster = new List<double> { allPositions[i] };
            }
        }
        clusters.Add(currentCluster);
        
        // 各クラスターの中央値を計算し、統一性を評価
        var commonPositions = new List<double>();
        int totalRows = leftPositions.Count;
        
        foreach (var cluster in clusters)
        {
            // この位置が何行に現れるかをカウント
            int appearanceCount = 0;
            double clusterCenter = cluster.Average();
            
            foreach (var rowPositions in leftPositions)
            {
                bool appearsInRow = rowPositions.Any(pos => Math.Abs(pos - clusterCenter) <= clusterThreshold);
                if (appearsInRow) appearanceCount++;
            }
            
            // 40%以上の行に現れる位置を共通位置として採用（より柔軟に）
            if ((double)appearanceCount / totalRows >= 0.4)
            {
                commonPositions.Add(clusterCenter);
            }
        }
        
        return commonPositions.OrderBy(pos => pos).ToList();
    }
    
    // CLAUDE.md準拠：行間一貫性に基づく列境界検出
    private static List<ColumnBoundary> CalculateConsistentColumnBoundaries(List<DocumentElement> tableRows)
    {
        if (tableRows.Count < 2) return new List<ColumnBoundary>();
        
        // 各行で有意なギャップ位置を検出
        var allRowGaps = new List<List<double>>();
        
        foreach (var row in tableRows)
        {
            if (row.Words == null || row.Words.Count < 2) continue;
            
            var sortedWords = row.Words.OrderBy(w => w.BoundingBox.Left).ToList();
            var rowGaps = new List<double>();
            var avgCharWidth = CalculateAverageCharacterWidth(sortedWords);
            
            for (int i = 0; i < sortedWords.Count - 1; i++)
            {
                var gap = sortedWords[i + 1].BoundingBox.Left - sortedWords[i].BoundingBox.Right;
                
                // CLAUDE.md準拠：行内で有意なギャップのみを記録（厳格化）
                if (gap >= avgCharWidth * 1.5) // 文字幅の1.5倍以上を列境界候補とする（精度重視）
                {
                    var gapCenter = sortedWords[i].BoundingBox.Right + (gap / 2.0);
                    rowGaps.Add(gapCenter);
                }
            }
            
            if (rowGaps.Count > 0)
            {
                allRowGaps.Add(rowGaps);
            }
        }
        
        if (allRowGaps.Count < 2) return new List<ColumnBoundary>();
        
        // CLAUDE.md準拠：行間で一貫して現れるギャップ位置を特定
        var consistentGaps = FindConsistentGapPositions(allRowGaps);
        
        if (consistentGaps.Count == 0) return new List<ColumnBoundary>();
        
        // 境界を構築
        var allWords = tableRows.SelectMany(r => r.Words ?? new List<Word>()).ToList();
        if (!allWords.Any()) return new List<ColumnBoundary>();
        
        var startX = allWords.Min(w => w.BoundingBox.Left);
        var endX = allWords.Max(w => w.BoundingBox.Right);
        var boundaries = new List<ColumnBoundary>();
        var currentLeft = startX;
        
        foreach (var gapPosition in consistentGaps.OrderBy(g => g))
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
    
    // CLAUDE.md準拠：行間一貫性のあるギャップ位置の特定
    private static List<double> FindConsistentGapPositions(List<List<double>> allRowGaps)
    {
        if (allRowGaps.Count == 0) return new List<double>();
        
        var consistentGaps = new List<double>();
        var allGaps = allRowGaps.SelectMany(gaps => gaps).OrderBy(g => g).ToList();
        
        if (allGaps.Count == 0) return consistentGaps;
        
        // 動的クラスタリング闾値
        var avgGapDistance = 0.0;
        if (allGaps.Count > 1)
        {
            var distances = new List<double>();
            for (int i = 1; i < allGaps.Count; i++)
            {
                distances.Add(allGaps[i] - allGaps[i - 1]);
            }
            avgGapDistance = distances.Average();
        }
        var clusterThreshold = Math.Max(avgGapDistance * 0.5, 10.0); // 動的または最低10ポイント
        
        // ギャップ位置のクラスタリング
        var clusters = new List<List<double>>();
        var currentCluster = new List<double> { allGaps[0] };
        
        for (int i = 1; i < allGaps.Count; i++)
        {
            if (allGaps[i] - allGaps[i - 1] <= clusterThreshold)
            {
                currentCluster.Add(allGaps[i]);
            }
            else
            {
                clusters.Add(currentCluster);
                currentCluster = new List<double> { allGaps[i] };
            }
        }
        clusters.Add(currentCluster);
        
        // 複数行で一貫して現れるクラスターのみを採用
        int totalRows = allRowGaps.Count;
        int minRequiredRows = Math.Max(2, (int)(totalRows * 0.5)); // 50%以上の行に現れる（柔軟化）
        
        foreach (var cluster in clusters)
        {
            var clusterCenter = cluster.Average();
            int appearanceCount = 0;
            
            foreach (var rowGaps in allRowGaps)
            {
                bool appearsInRow = rowGaps.Any(gap => Math.Abs(gap - clusterCenter) <= clusterThreshold);
                if (appearsInRow) appearanceCount++;
            }
            
            if (appearanceCount >= minRequiredRows)
            {
                consistentGaps.Add(clusterCenter);
            }
        }
        
        return consistentGaps.OrderBy(g => g).ToList();
    }
    
    // CLAUDE.md準拠：文字幅の統計的計算
    private static double CalculateAverageCharWidth(List<double> positions)
    {
        if (positions.Count < 2) return 10.0; // デフォルト値
        
        var gaps = new List<double>();
        for (int i = 1; i < positions.Count; i++)
        {
            gaps.Add(positions[i] - positions[i - 1]);
        }
        
        if (gaps.Count == 0) return 10.0;
        
        // 異常値を除外した平均値
        gaps.Sort();
        var median = gaps[gaps.Count / 2];
        var filteredGaps = gaps.Where(g => g <= median * 3).ToList();
        
        return filteredGaps.Count > 0 ? filteredGaps.Average() / 3.0 : 10.0;
    }
    
    // CLAUDE.md準拠：垂直共通空白領域の検出
    private static List<(double Left, double Right)> FindVerticalCommonGaps(List<DocumentElement> tableRows)
    {
        if (tableRows.Count < 2) return new List<(double, double)>();
        
        // 各行における空白領域を特定
        var rowGaps = new List<List<(double Left, double Right)>>();
        
        foreach (var row in tableRows)
        {
            if (row.Words == null || row.Words.Count == 0) continue;
            
            var sortedWords = row.Words.OrderBy(w => w.BoundingBox.Left).ToList();
            var gaps = new List<(double Left, double Right)>();
            
            // 行内の単語間空白を計算（最小閾値を下げる）
            for (int i = 0; i < sortedWords.Count - 1; i++)
            {
                var gapLeft = sortedWords[i].BoundingBox.Right;
                var gapRight = sortedWords[i + 1].BoundingBox.Left;
                var gapWidth = gapRight - gapLeft;
                
                // より小さなギャップも考慮（5ポイント以上）
                if (gapWidth >= 5.0)
                {
                    gaps.Add((gapLeft, gapRight));
                }
            }
            
            if (gaps.Count > 0)
            {
                rowGaps.Add(gaps);
            }
        }
        
        if (rowGaps.Count < 2) return new List<(double, double)>();
        
        // 共通する垂直空白領域を特定
        var commonGaps = new List<(double Left, double Right)>();
        
        // 最初の行の空白を基準として、他の行にも共通して存在するかチェック
        foreach (var referenceGap in rowGaps[0])
        {
            int commonCount = 1; // 基準行を含む
            var overlappingGaps = new List<(double Left, double Right)> { referenceGap };
            
            foreach (var otherRowGaps in rowGaps.Skip(1))
            {
                var overlapping = FindOverlappingGap(referenceGap, otherRowGaps);
                if (overlapping.HasValue)
                {
                    commonCount++;
                    overlappingGaps.Add(overlapping.Value);
                }
            }
            
            // 30%以上の行に共通する空白領域を採用（より柔軟に）
            if ((double)commonCount / rowGaps.Count >= 0.3)
            {
                // 重複する領域の共通部分を計算
                var commonLeft = overlappingGaps.Max(g => g.Left);
                var commonRight = overlappingGaps.Min(g => g.Right);
                
                if (commonRight > commonLeft)
                {
                    commonGaps.Add((commonLeft, commonRight));
                }
            }
        }
        
        // 重複する領域をマージ
        return MergeOverlappingGaps(commonGaps);
    }
    
    // CLAUDE.md準拠：空白領域の重複検出
    private static (double Left, double Right)? FindOverlappingGap(
        (double Left, double Right) referenceGap, 
        List<(double Left, double Right)> candidateGaps)
    {
        foreach (var candidate in candidateGaps)
        {
            // 重複判定：一部でも重複していれば考慮
            var overlapLeft = Math.Max(referenceGap.Left, candidate.Left);
            var overlapRight = Math.Min(referenceGap.Right, candidate.Right);
            
            if (overlapRight > overlapLeft)
            {
                // 重複面積が各領域の50%以上であれば同じ空白領域とみなす
                var overlapWidth = overlapRight - overlapLeft;
                var refWidth = referenceGap.Right - referenceGap.Left;
                var candWidth = candidate.Right - candidate.Left;
                
                if (overlapWidth >= refWidth * 0.3 || overlapWidth >= candWidth * 0.3)
                {
                    return candidate;
                }
            }
        }
        
        return null;
    }
    
    // CLAUDE.md準拠：重複する空白領域のマージ
    private static List<(double Left, double Right)> MergeOverlappingGaps(List<(double Left, double Right)> gaps)
    {
        if (gaps.Count <= 1) return gaps;
        
        var sorted = gaps.OrderBy(g => g.Left).ToList();
        var merged = new List<(double Left, double Right)>();
        var current = sorted[0];
        
        for (int i = 1; i < sorted.Count; i++)
        {
            var next = sorted[i];
            
            // 重複または隣接している場合はマージ
            if (next.Left <= current.Right + 5) // 5ポイントの許容誤差
            {
                current = (current.Left, Math.Max(current.Right, next.Right));
            }
            else
            {
                merged.Add(current);
                current = next;
            }
        }
        
        merged.Add(current);
        return merged;
    }
    
    // CLAUDE.md準拠：垂直空白領域から列境界を構築
    private static List<ColumnBoundary> BuildBoundariesFromVerticalGaps(
        List<(double Left, double Right)> verticalGaps, 
        List<DocumentElement> tableRows)
    {
        var allWords = tableRows.SelectMany(r => r.Words ?? new List<Word>()).ToList();
        if (!allWords.Any()) return new List<ColumnBoundary>();
        
        var minX = allWords.Min(w => w.BoundingBox.Left);
        var maxX = allWords.Max(w => w.BoundingBox.Right);
        var boundaries = new List<ColumnBoundary>();
        
        // 空白領域の中央を境界として使用
        var gapCenters = verticalGaps.Select(g => (g.Left + g.Right) / 2.0)
                                   .OrderBy(x => x)
                                   .ToList();
        
        var currentLeft = minX;
        
        foreach (var gapCenter in gapCenters)
        {
            boundaries.Add(new ColumnBoundary
            {
                Left = currentLeft,
                Right = gapCenter
            });
            currentLeft = gapCenter;
        }
        
        // 最後の列
        boundaries.Add(new ColumnBoundary
        {
            Left = currentLeft,
            Right = maxX
        });
        
        return boundaries;
    }
    
    // CLAUDE.md準拠：共通位置から列境界を構築
    private static List<ColumnBoundary> BuildColumnBoundariesFromPositions(List<double> positions, List<Word> allWords)
    {
        if (positions.Count < 2) return new List<ColumnBoundary>();
        
        var boundaries = new List<ColumnBoundary>();
        var minX = allWords.Min(w => w.BoundingBox.Left);
        var maxX = allWords.Max(w => w.BoundingBox.Right);
        
        for (int i = 0; i < positions.Count; i++)
        {
            double left = i == 0 ? minX : (positions[i - 1] + positions[i]) / 2;
            double right = i == positions.Count - 1 ? maxX : (positions[i] + positions[i + 1]) / 2;
            
            boundaries.Add(new ColumnBoundary
            {
                Left = left,
                Right = right
            });
        }
        
        return boundaries;
    }
    
    // CLAUDE.md準拠：テーブル処理後の後続段落統合（分離セル内容対応）
    private static void IntegrateSubsequentParagraphs(List<DocumentElement> tableRows, List<DocumentElement> allElements, int tableStartIndex)
    {
        if (tableRows.Count == 0) return;
        
        var lastTableRow = tableRows.Last();
        
        // テーブル処理範囲の終了位置を特定
        int searchStartIndex = tableStartIndex + tableRows.Count;
        
        // 後続の段落を検索して統合判定
        for (int i = searchStartIndex; i < allElements.Count && i < searchStartIndex + 5; i++) // 最大5要素先まで検索
        {
            var element = allElements[i];
            
            // 空要素はスキップ
            if (element.Type == ElementType.Empty) continue;
            
            // 段落の場合は統合判定
            if (element.Type == ElementType.Paragraph && element.Content.Trim().Length <= 50)
            {
                // CLAUDE.md準拠：座標ベース統合判定
                if (ShouldIntegrateIntoPreviousTableRowByCoordinates(element, lastTableRow))
                {
                    // 統合実行
                    if (element.Words != null && element.Words.Count > 0 && 
                        lastTableRow.Words != null && lastTableRow.Words.Count > 0)
                    {
                        var combinedWords = new List<Word>(lastTableRow.Words);
                        combinedWords.AddRange(element.Words);
                        lastTableRow.Words = combinedWords;
                    }
                    
                    lastTableRow.Content = lastTableRow.Content + " " + element.Content.Trim();
                    
                    // 統合した要素を削除
                    element.Content = ""; // 空にして後続処理で無視されるようにする
                    element.Type = ElementType.Empty;
                }
                else
                {
                    break; // 統合できない段落が見つかったら終了
                }
            }
            else if (element.Type == ElementType.Header)
            {
                break; // ヘッダーが見つかったら終了
            }
            else
            {
                break; // テーブル以外の要素が見つかったら終了
            }
        }
    }
}

public class ColumnBoundary
{
    public double Left { get; set; }
    public double Right { get; set; }
}