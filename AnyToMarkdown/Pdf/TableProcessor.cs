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
        
        
        // CLAUDE.md準拠：完全に空の列を除去
        allCells = RemoveEmptyColumns(allCells);
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
        
        // 水平方向の重複計算
        var overlapLeft = Math.Max(paragraphLeft, tableLeft);
        var overlapRight = Math.Min(paragraphRight, tableRight);
        var overlapWidth = Math.Max(0, overlapRight - overlapLeft);
        var paragraphWidth = paragraphRight - paragraphLeft;
        
        bool hasSignificantOverlap = paragraphWidth > 0 && (overlapWidth / paragraphWidth) >= 0.6;
        
        // CLAUDE.md準拠：垂直間隔の統計的評価
        var tableBottom = tableRow.Words.Min(w => w.BoundingBox.Bottom);
        var paragraphTop = paragraph.Words.Max(w => w.BoundingBox.Top);
        var verticalGap = Math.Abs(tableBottom - paragraphTop);
        
        // 動的垂直閾値（行高基準）
        var avgRowHeight = tableRow.Words.Average(w => w.BoundingBox.Height);
        var maxVerticalGap = avgRowHeight * 1.2;
        
        return hasSignificantOverlap && verticalGap <= maxVerticalGap;
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
    
    // テーブル全体の統一列境界計算（CLAUDE.md準拠：垂直座標クラスタリング強化版）
    private static List<ColumnBoundary> CalculateGlobalColumnBoundaries(List<DocumentElement> tableRows)
    {
        if (tableRows.Count == 0) return new List<ColumnBoundary>();
        
        // CLAUDE.md準拠：垂直座標による列境界検出の前処理
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
                        
                        // CLAUDE.md準拠：改善された統計的境界識別
                        // 問題：「名前 年齢」が統合、必要な境界の見落とし
                        // 解決：統計的外れ値検出の精度向上と文字幅正規化の改善
                        
                        // CLAUDE.md準拠：精密化された統計的境界検出
                        // 目標：「説明 備考」「名前 年齢」などの語句境界を正確に検出
                        // 方法：分布の変動に敏感な統計的分析
                        
                        // 改善1：分散考慮型統計閾値（異常値検出の精度向上）
                        var varianceBasedThreshold = q25 + (iqr * 0.8); // Q1から上位80%をターゲット
                        var medianDeviation = Math.Abs(gap - median);
                        var isStatisticalOutlier = gap >= varianceBasedThreshold && medianDeviation > median * 0.3;
                        
                        // 改善2：語句間隔の平均的パターン分析
                        var wordSpacingPattern = median * 0.7; // 通常の語句間隔の70%を基準
                        var isAboveNormalSpacing = gap > wordSpacingPattern;
                        
                        // 改善3：文字サイズ正規化による適応的判定
                        var normalizedGapSize = gap / Math.Max(avgWordWidth, 10.0);
                        var isSizeSignificant = normalizedGapSize >= 0.6; // 文字幅の60%以上
                        
                        // 統合判定：複数条件による境界認識
                        if (isStatisticalOutlier && isAboveNormalSpacing && isSizeSignificant)
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
    
    // CLAUDE.md準拠：純粋な座標・統計分析による境界検出
    private static List<string> ParseTableCellsWithStatisticalGapAnalysis(DocumentElement row)
    {
        if (row.Words == null || row.Words.Count <= 1)
            return [row.Content.Trim()];
        
        var sortedWords = row.Words.OrderBy(w => w.BoundingBox.Left).ToList();
        if (sortedWords.Count == 1)
            return [sortedWords[0].Text ?? ""];
        
        // CLAUDE.md準拠：座標ベースギャップ計算
        var gaps = new List<(double Gap, int Index)>();
        var avgCharWidth = CalculateAverageCharacterWidth(sortedWords);
        
        for (int i = 0; i < sortedWords.Count - 1; i++)
        {
            var currentWord = sortedWords[i];
            var nextWord = sortedWords[i + 1];
            var gap = nextWord.BoundingBox.Left - currentWord.BoundingBox.Right;
            gaps.Add((gap, i));
        }
        
        if (gaps.Count == 0)
            return [string.Join(" ", sortedWords.Select(w => w.Text))];
        
        // CLAUDE.md準拠：統計的有意ギャップ検出
        var significantGaps = DetectStatisticallySignificantGaps(gaps, avgCharWidth);
        
        // CLAUDE.md準拠：過分割の統計的フィルタリング
        if (significantGaps.Count > 3)
        {
            significantGaps = FilterExcessiveSegmentation(significantGaps);
        }
        
        // セル構築
        return BuildCellsFromSignificantGaps(sortedWords, significantGaps);
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
        var gapValues = gaps.Select(g => g.Gap).ToList();
        var mean = gapValues.Average();
        var variance = gapValues.Sum(g => Math.Pow(g - mean, 2)) / gapValues.Count;
        var stdDev = Math.Sqrt(variance);
        
        // CLAUDE.md準拠：適応的統計閾値
        var adaptiveThreshold = Math.Max(
            avgCharWidth * 1.5,      // 文字幅ベース
            mean + stdDev * 0.8      // 統計的外れ値
        );
        
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
        // 上位50%のギャップのみ保持（最も有意な境界）
        var sortedGaps = significantGaps.OrderByDescending(g => g.Gap).ToList();
        var keepCount = Math.Max(1, sortedGaps.Count / 2);
        
        return sortedGaps.Take(keepCount).OrderBy(g => g.Index).ToList();
    }
    
    // CLAUDE.md準拠：ギャップからのセル構築
    private static List<string> BuildCellsFromSignificantGaps(
        List<Word> sortedWords, 
        List<(double Gap, int Index)> significantGaps)
    {
        var cells = new List<string>();
        var startIndex = 0;
        
        foreach (var (_, index) in significantGaps.OrderBy(g => g.Index))
        {
            var cellWords = sortedWords.Skip(startIndex).Take(index - startIndex + 1).ToList();
            if (cellWords.Any())
            {
                var cellText = string.Join(" ", cellWords.Select(w => w.Text)).Trim();
                if (!string.IsNullOrWhiteSpace(cellText))
                {
                    cells.Add(cellText);
                }
            }
            startIndex = index + 1;
        }
        
        // 残りの単語
        if (startIndex < sortedWords.Count)
        {
            var remainingWords = sortedWords.Skip(startIndex).ToList();
            if (remainingWords.Any())
            {
                var cellText = string.Join(" ", remainingWords.Select(w => w.Text)).Trim();
                if (!string.IsNullOrWhiteSpace(cellText))
                {
                    cells.Add(cellText);
                }
            }
        }
        
        return cells.Count > 0 ? cells : [string.Join(" ", sortedWords.Select(w => w.Text))];
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
}

public class ColumnBoundary
{
    public double Left { get; set; }
    public double Right { get; set; }
}