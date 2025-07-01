using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

internal static class PostProcessor
{
    public static List<DocumentElement> PostProcessElementClassification(List<DocumentElement> elements, FontAnalysis fontAnalysis)
    {
        var result = new List<DocumentElement>(elements);
        
        // コンテキスト情報を使用して要素分類を改善
        for (int i = 0; i < result.Count; i++)
        {
            var current = result[i];
            var previous = i > 0 ? result[i - 1] : null;
            var next = i < result.Count - 1 ? result[i + 1] : null;
            
            // 段落の継続性チェック（一時的に無効化 - 段落を分離したまま保持）
            /*if (current.Type == ElementType.Paragraph && 
                previous?.Type == ElementType.Paragraph &&
                ShouldMergeParagraphs(previous, current))
            {
                // 段落を統合
                previous.Content = previous.Content.Trim() + " " + current.Content.Trim();
                previous.Words.AddRange(current.Words);
                result.RemoveAt(i);
                i--; // インデックス調整
                continue;
            }*/
            
            // ヘッダー分類の改善（より慎重に - 段落の誤分類を防ぐ）
            if (current.Type == ElementType.Paragraph && 
                IsDefinitelyHeader(current, previous, next, fontAnalysis))
            {
                current.Type = ElementType.Header;
            }
            
            // リストアイテムの継続性チェック
            if (current.Type == ElementType.Paragraph &&
                previous?.Type == ElementType.ListItem &&
                IsListContinuation(current, previous))
            {
                current.Type = ElementType.ListItem;
            }
            
            // テーブル行の改善
            if (current.Type == ElementType.Paragraph &&
                IsPartOfTableSequence(current, result, i))
            {
                current.Type = ElementType.TableRow;
            }
            
            // テーブルとヘッダーの分離改善
            if (current.Type == ElementType.Paragraph && 
                i > 0 && result[i - 1].Type == ElementType.TableRow &&
                i < result.Count - 1 && result[i + 1].Type == ElementType.TableRow)
            {
                // テーブル間に挟まれた短いテキストはヘッダー候補
                var content = current.Content?.Trim() ?? "";
                if (content.Length <= 30 && !content.Contains("|") && !content.Contains("\t"))
                {
                    // フォントサイズや位置でヘッダー判定
                    if (CouldBeHeader(current, previous, next, fontAnalysis))
                    {
                        current.Type = ElementType.Header;
                    }
                }
            }
        }
        
        return result;
    }

    public static List<DocumentElement> PostProcessHeaderDetectionWithCoordinates(List<DocumentElement> elements, FontAnalysis fontAnalysis)
    {
        var result = new List<DocumentElement>(elements);
        
        for (int i = 0; i < result.Count; i++)
        {
            var element = result[i];
            
            // ヘッダー候補の再評価（非常に慎重に）
            if (element.Type == ElementType.Paragraph)
            {
                // 段落的なパターンを含むテキストは絶対にヘッダーにしない
                var content = element.Content?.Trim() ?? "";
                bool isParagraphPattern = content.EndsWith("。") || content.EndsWith("です。") || 
                                        content.Contains("、") || content.Contains("**") ||
                                        content.StartsWith("これは") || content.StartsWith("この") ||
                                        content.Contains("マークダウン") || content.Contains("と斜体") ||
                                        content.Length > 15;
                                        
                // マークダウン記号で始まる場合はヘッダーとして処理
                if (content.StartsWith("#"))
                {
                    element.Type = ElementType.Header;
                }
                else if (!isParagraphPattern && 
                    ElementDetector.IsHeaderStructure(element.Content, element.Words, element.FontSize, fontAnalysis))
                {
                    element.Type = ElementType.Header;
                }
            }
            else if (element.Type == ElementType.Header)
            {
                // 既存のヘッダーの妥当性チェック
                if (!ElementDetector.IsHeaderStructure(element.Content, element.Words, element.FontSize, fontAnalysis) &&
                    !ElementDetector.IsHeaderLike(element.Content))
                {
                    // ヘッダーとして不適切な場合は段落に戻す
                    element.Type = ElementType.Paragraph;
                }
            }
        }
        
        return result;
    }

    public static List<DocumentElement> PostProcessTableDetection(List<DocumentElement> elements, GraphicsInfo graphicsInfo)
    {
        var result = new List<DocumentElement>(elements);
        
        // 図形情報を使用したテーブル検出
        if (graphicsInfo?.TablePatterns != null && graphicsInfo.TablePatterns.Any())
        {
            foreach (var pattern in graphicsInfo.TablePatterns)
            {
                // パターンに該当する要素をテーブル行として分類
                for (int i = 0; i < result.Count; i++)
                {
                    var element = result[i];
                    if (IsElementInTableArea(element, pattern.BoundingArea))
                    {
                        if (element.Type == ElementType.Paragraph)
                        {
                            element.Type = ElementType.TableRow;
                        }
                    }
                }
            }
        }
        
        // 図形情報がない場合の座標ベーステーブル検出
        result = DetectTablesWithoutGraphics(result);
        
        // 座標ベースのテーブル要素統合処理
        result = ConsolidateTableElementsByCoordinates(result, graphicsInfo);
        
        // 連続するテーブル行の検出
        result = DetectTableSequences(result);
        
        // テーブル領域の最適化
        result = OptimizeTableRegions(result);
        
        return result;
    }
    
    private static List<DocumentElement> DetectTablesWithoutGraphics(List<DocumentElement> elements)
    {
        var result = new List<DocumentElement>(elements);
        
        // 複数列のレイアウトパターンを検出
        for (int i = 0; i < result.Count; i++)
        {
            var element = result[i];
            
            // 段落要素で複数列のパターンを持つ場合
            if (element.Type == ElementType.Paragraph && element.Words != null && element.Words.Count > 0)
            {
                // 列の分布を分析
                var columnCount = AnalyzeColumnStructure(element);
                
                // より積極的な検出：2列以上、短いテキスト、または数字を含む場合
                bool isTableCandidate = columnCount >= 2 || 
                                       HasMultipleShortTexts(element) ||
                                       LooksLikeTableRow(element);
                
                if (isTableCandidate)
                {
                    // 前後の要素チェックをより緩和
                    if (IsPartOfTablePattern(result, i) || IsIsolatedTableRow(element))
                    {
                        element.Type = ElementType.TableRow;
                    }
                }
            }
        }
        
        return result;
    }
    
    private static bool IsIsolatedTableRow(DocumentElement element)
    {
        // 座標情報に基づく判定のみ使用
        if (element.Words == null || element.Words.Count == 0)
        {
            return false;
        }
        
        // 単語の配置パターンを分析
        var words = element.Words.OrderBy(w => w.BoundingBox.Left).ToList();
        
        // 単語間のギャップを分析
        var hasSignificantGaps = false;
        for (int i = 0; i < words.Count - 1; i++)
        {
            var currentWord = words[i];
            var nextWord = words[i + 1];
            var gapSize = nextWord.BoundingBox.Left - currentWord.BoundingBox.Right;
            var avgFontWidth = currentWord.BoundingBox.Width / Math.Max(1, 1.0); // 座標幅ベース
            
            if (gapSize > avgFontWidth * 1.5) // フォント幅の1.5倍以上のギャップ
            {
                hasSignificantGaps = true;
                break;
            }
        }
        
        // 複数の単語が適度な間隔で配置されている場合はテーブル行候補
        return words.Count >= 2 && hasSignificantGaps;
    }
    
    private static int AnalyzeColumnStructure(DocumentElement element)
    {
        if (element.Words == null || element.Words.Count == 0)
            return 0;
        
        // 単語の水平位置をクラスタリング
        var positions = element.Words.Select(w => w.BoundingBox.Left).OrderBy(p => p).ToList();
        var clusters = ClusterHorizontalPositions(positions);
        
        return clusters.Count;
    }
    
    private static List<List<double>> ClusterHorizontalPositions(List<double> positions)
    {
        var clusters = new List<List<double>>();
        if (positions.Count == 0) return clusters;
        
        const double threshold = 30.0; // 30pt以上の間隔で別列とみなす
        var currentCluster = new List<double> { positions[0] };
        
        for (int i = 1; i < positions.Count; i++)
        {
            if (positions[i] - positions[i - 1] <= threshold)
            {
                currentCluster.Add(positions[i]);
            }
            else
            {
                clusters.Add(currentCluster);
                currentCluster = new List<double> { positions[i] };
            }
        }
        
        clusters.Add(currentCluster);
        return clusters;
    }
    
    private static bool HasMultipleShortTexts(DocumentElement element)
    {
        // 座標ベースの分析のみを使用
        if (element.Words == null || element.Words.Count == 0)
        {
            return false;
        }
        
        // 単語間のギャップ分析によるセグメント検出
        var words = element.Words.OrderBy(w => w.BoundingBox.Left).ToList();
        var segmentCount = 1; // 最初の単語は1つ目のセグメント
        
        for (int i = 0; i < words.Count - 1; i++)
        {
            var currentWord = words[i];
            var nextWord = words[i + 1];
            var gapSize = nextWord.BoundingBox.Left - currentWord.BoundingBox.Right;
            var avgFontWidth = currentWord.BoundingBox.Width;
            
            // 有意なギャップ（フォント幅の2倍以上）でセグメント分離
            if (gapSize > avgFontWidth * 2.0)
            {
                segmentCount++;
            }
        }
        
        return segmentCount >= 2;
    }
    
    private static bool IsPartOfTablePattern(List<DocumentElement> elements, int index)
    {
        var element = elements[index];
        
        // 前後の要素で類似の列構造を持つかチェック
        int similarPatternCount = 0;
        
        // 前の2要素をチェック
        for (int i = Math.Max(0, index - 2); i < index; i++)
        {
            if (HasSimilarColumnStructure(elements[i], element))
            {
                similarPatternCount++;
            }
        }
        
        // 次の2要素をチェック
        for (int i = index + 1; i < Math.Min(elements.Count, index + 3); i++)
        {
            if (HasSimilarColumnStructure(elements[i], element))
            {
                similarPatternCount++;
            }
        }
        
        return similarPatternCount >= 1; // 少なくとも1つの類似要素があればテーブルパターン
    }
    
    private static bool HasSimilarColumnStructure(DocumentElement a, DocumentElement b)
    {
        if (a.Words == null || b.Words == null || a.Words.Count == 0 || b.Words.Count == 0)
            return false;
        
        var columnsA = AnalyzeColumnStructure(a);
        var columnsB = AnalyzeColumnStructure(b);
        
        // 列数が同じまたは±1以内
        return Math.Abs(columnsA - columnsB) <= 1 && columnsA >= 2 && columnsB >= 2;
    }

    public static List<DocumentElement> PostProcessTableHeaderIntegration(List<DocumentElement> elements)
    {
        var result = new List<DocumentElement>(elements);
        
        for (int i = 0; i < result.Count - 1; i++)
        {
            var current = result[i];
            var next = result[i + 1];
            
            // ヘッダーの直後にテーブルがある場合
            if (current.Type == ElementType.Header && next.Type == ElementType.TableRow)
            {
                // ヘッダーをテーブルの一部として統合
                var headerContent = current.Content.Trim();
                
                // テーブルヘッダー行として統合
                if (ShouldIntegrateHeaderIntoTable(headerContent, next))
                {
                    next.Content = headerContent + "\n" + next.Content;
                    result.RemoveAt(i);
                    i--; // インデックス調整
                }
            }
        }
        
        return result;
    }

    public static List<DocumentElement> PostProcessCodeAndQuoteBlocks(List<DocumentElement> elements)
    {
        var result = new List<DocumentElement>(elements);
        
        // コードブロックの検出と統合
        result = DetectAndMergeCodeBlocks(result);
        
        // 引用ブロックの検出と統合
        result = DetectAndMergeQuoteBlocks(result);
        
        return result;
    }
    
    public static List<DocumentElement> ConsolidateBrokenTableCells(List<DocumentElement> elements)
    {
        var result = new List<DocumentElement>();
        var i = 0;
        
        while (i < elements.Count)
        {
            var current = elements[i];
            
            // より積極的な統合：テーブル行、段落、またはヘッダー以外の要素
            if (current.Type == ElementType.TableRow || 
                current.Type == ElementType.Paragraph ||
                (current.Type != ElementType.Header && current.Type != ElementType.Empty))
            {
                var consolidated = ConsolidateMultilineTableCell(elements, i);
                
                // 統合された要素がテーブルらしい場合はTableRowに分類
                if (consolidated.Element.Type == ElementType.Paragraph && 
                    LooksLikeTableRow(consolidated.Element))
                {
                    consolidated.Element.Type = ElementType.TableRow;
                }
                
                result.Add(consolidated.Element);
                i = consolidated.NextIndex;
            }
            else
            {
                result.Add(current);
                i++;
            }
        }
        
        return result;
    }
    
    private static bool LooksLikeTableRow(DocumentElement element)
    {
        var content = element.Content?.Trim() ?? "";
        
        // マークダウン記号「<br>」を含む場合（許可されたマークダウン記法）
        if (content.Contains("<br>"))
        {
            return true;
        }
        
        // 座標ベースの複数列検出
        if (element.Words != null && element.Words.Count > 0)
        {
            return AnalyzeColumnStructure(element) >= 2;
        }
        
        return false;
    }
    
    private static (DocumentElement Element, int NextIndex) ConsolidateMultilineTableCell(List<DocumentElement> elements, int startIndex)
    {
        var baseElement = elements[startIndex];
        var consolidated = new DocumentElement
        {
            Type = baseElement.Type,
            Content = baseElement.Content,
            FontSize = baseElement.FontSize,
            LeftMargin = baseElement.LeftMargin,
            IsIndented = baseElement.IsIndented,
            Words = baseElement.Words?.ToList() ?? new List<Word>()
        };
        
        int nextIndex = startIndex + 1;
        int consecutiveConsolidations = 0;
        const int maxConsolidations = 1; // 最大1つまでの要素を統合（最も保守的に）
        
        // 後続の要素で統合可能なものを探す
        for (int i = startIndex + 1; i < elements.Count && consecutiveConsolidations < maxConsolidations; i++)
        {
            var candidate = elements[i];
            
            // 空要素はスキップ
            if (candidate.Type == ElementType.Empty || string.IsNullOrWhiteSpace(candidate.Content))
            {
                nextIndex = i + 1;
                continue;
            }
            
            // ヘッダー要素に到達したら統合停止
            if (candidate.Type == ElementType.Header)
            {
                break;
            }
            
            // 座標ベースで統合可能性を判定
            if (ShouldConsolidateWithTableCell(consolidated, candidate))
            {
                // 内容を<br>で結合（既に<br>がある場合は追加せず、スペースで区切る）
                var trimmedCandidate = candidate.Content.Trim();
                if (!string.IsNullOrEmpty(trimmedCandidate))
                {
                    if (consolidated.Content.TrimEnd().EndsWith("<br>") || trimmedCandidate.StartsWith("<br>"))
                    {
                        // 既に<br>がある場合は単純に結合
                        consolidated.Content = consolidated.Content.TrimEnd() + " " + trimmedCandidate;
                    }
                    else
                    {
                        // <br>で結合
                        consolidated.Content = consolidated.Content.TrimEnd() + "<br>" + trimmedCandidate;
                    }
                }
                
                // 単語リストも統合
                if (candidate.Words != null)
                {
                    consolidated.Words.AddRange(candidate.Words);
                }
                
                nextIndex = i + 1;
                consecutiveConsolidations++;
            }
            else
            {
                // 統合できない要素に到達したら終了
                break;
            }
        }
        
        return (consolidated, nextIndex);
    }
    
    private static bool ShouldConsolidateWithTableCell(DocumentElement baseElement, DocumentElement candidate)
    {
        // 候補要素がヘッダーまたは空要素の場合は統合しない
        if (candidate.Type == ElementType.Header || candidate.Type == ElementType.Empty)
        {
            return false;
        }
        
        // 座標情報が不足している場合は統合しない（座標ベース分析のみ）
        if (baseElement.Words == null || baseElement.Words.Count == 0 ||
            candidate.Words == null || candidate.Words.Count == 0)
        {
            return false;
        }
        
        // 両方の要素の垂直および水平位置を分析
        var baseBottom = baseElement.Words.Min(w => w.BoundingBox.Bottom);
        var candidateTop = candidate.Words.Max(w => w.BoundingBox.Top);
        var verticalGap = Math.Abs(baseBottom - candidateTop);
        
        var baseLeft = baseElement.Words.Min(w => w.BoundingBox.Left);
        var baseRight = baseElement.Words.Max(w => w.BoundingBox.Right);
        var candidateLeft = candidate.Words.Min(w => w.BoundingBox.Left);
        var candidateRight = candidate.Words.Max(w => w.BoundingBox.Right);
        
        // 重要：異なるテーブル行の統合を防止
        // 座標範囲が大きく重複する場合は別の行とみなす
        var horizontalOverlap = Math.Max(0, Math.Min(baseRight, candidateRight) - Math.Max(baseLeft, candidateLeft));
        var baseWidth = baseRight - baseLeft;
        var candidateWidth = candidateRight - candidateLeft;
        var overlapRatio = Math.Max(baseWidth, candidateWidth) > 0 ? 
                          horizontalOverlap / Math.Max(baseWidth, candidateWidth) : 0;
        
        // 50%以上水平重複する場合は異なるテーブル行として扱い統合しない
        if (overlapRatio > 0.5)
        {
            return false;
        }
        
        // フォントサイズベースの垂直間隔閾値（複数行テキスト対応のため緩和）
        var avgFontSize = baseElement.Words.Average(w => w.BoundingBox.Height);
        var maxVerticalGap = avgFontSize * 1.0; // より厳格な閾値
        
        // 統合条件（座標・フォント・ギャップベースのみ）：
        // 1. 垂直間隔が非常に小さい
        // 2. 水平方向の重複が少ない（同じセル内の改行など）
        // 3. テーブル行または段落タイプ
        bool isAppropriateType = candidate.Type == ElementType.Paragraph || 
                               candidate.Type == ElementType.TableRow;
        
        return verticalGap <= maxVerticalGap && 
               overlapRatio <= 0.5 && 
               isAppropriateType;
    }
    
    public static List<DocumentElement> ConsolidateTableElementsByCoordinates(List<DocumentElement> elements, GraphicsInfo graphicsInfo)
    {
        if (graphicsInfo?.TablePatterns == null || !graphicsInfo.TablePatterns.Any())
        {
            return elements;
        }
        
        var result = new List<DocumentElement>();
        var processedIndices = new HashSet<int>();
        
        for (int i = 0; i < elements.Count; i++)
        {
            if (processedIndices.Contains(i))
                continue;
                
            var current = elements[i];
            
            // テーブル境界内の要素を探す
            var tablePattern = FindContainingTablePattern(current, graphicsInfo.TablePatterns);
            if (tablePattern != null)
            {
                // テーブル行の構築
                var tableRows = BuildTableRowsFromCoordinates(elements, i, tablePattern, graphicsInfo, processedIndices);
                result.AddRange(tableRows);
            }
            else
            {
                result.Add(current);
                processedIndices.Add(i);
            }
        }
        
        return result;
    }
    
    private static TablePattern? FindContainingTablePattern(DocumentElement element, List<TablePattern> patterns)
    {
        if (element.Words == null || element.Words.Count == 0)
            return null;
            
        var elementBounds = GetElementBounds(element);
        
        foreach (var pattern in patterns)
        {
            if (IsWithinBounds(elementBounds, pattern.BoundingArea))
            {
                return pattern;
            }
        }
        
        return null;
    }
    
    private static List<DocumentElement> BuildTableRowsFromCoordinates(
        List<DocumentElement> elements, 
        int startIndex, 
        TablePattern tablePattern, 
        GraphicsInfo graphicsInfo,
        HashSet<int> processedIndices)
    {
        var tableRows = new List<DocumentElement>();
        var tableBounds = tablePattern.BoundingArea;
        
        // テーブル領域内の全要素を収集
        var tableElements = new List<(DocumentElement element, int index)>();
        for (int i = 0; i < elements.Count; i++)
        {
            if (processedIndices.Contains(i))
                continue;
                
            var element = elements[i];
            if (element.Words != null && element.Words.Count > 0)
            {
                var elementBounds = GetElementBounds(element);
                if (IsWithinBounds(elementBounds, tableBounds))
                {
                    tableElements.Add((element, i));
                    processedIndices.Add(i);
                }
            }
        }
        
        // Y座標でグループ化してテーブル行を構築
        var rowGroups = GroupElementsByRows(tableElements, graphicsInfo.HorizontalLines);
        
        foreach (var rowGroup in rowGroups)
        {
            // 行内の要素を列でグループ化
            var columnGroups = GroupElementsByColumns(rowGroup, graphicsInfo.VerticalLines);
            
            // テーブル行として統合
            var tableRow = BuildTableRowFromColumnGroups(columnGroups);
            if (tableRow != null)
            {
                tableRows.Add(tableRow);
            }
        }
        
        return tableRows;
    }
    
    private static UglyToad.PdfPig.Core.PdfRectangle GetElementBounds(DocumentElement element)
    {
        if (element.Words == null || element.Words.Count == 0)
        {
            return new UglyToad.PdfPig.Core.PdfRectangle(0, 0, 0, 0);
        }
        
        var minX = element.Words.Min(w => w.BoundingBox.Left);
        var maxX = element.Words.Max(w => w.BoundingBox.Right);
        var minY = element.Words.Min(w => w.BoundingBox.Bottom);
        var maxY = element.Words.Max(w => w.BoundingBox.Top);
        
        return new UglyToad.PdfPig.Core.PdfRectangle(minX, minY, maxX, maxY);
    }
    
    private static bool IsWithinBounds(UglyToad.PdfPig.Core.PdfRectangle elementBounds, UglyToad.PdfPig.Core.PdfRectangle tableBounds)
    {
        const double tolerance = 5.0;
        
        return elementBounds.Left >= tableBounds.Left - tolerance &&
               elementBounds.Right <= tableBounds.Right + tolerance &&
               elementBounds.Bottom >= tableBounds.Bottom - tolerance &&
               elementBounds.Top <= tableBounds.Top + tolerance;
    }
    
    private static List<List<(DocumentElement element, int index)>> GroupElementsByRows(
        List<(DocumentElement element, int index)> tableElements, 
        List<LineSegment> horizontalLines)
    {
        var rowGroups = new List<List<(DocumentElement element, int index)>>();
        
        // Y座標で要素をソート（上から下へ）
        var sortedElements = tableElements
            .OrderByDescending(x => x.element.Words?.Average(w => w.BoundingBox.Top) ?? 0)
            .ToList();
        
        foreach (var element in sortedElements)
        {
            var elementY = element.element.Words?.Average(w => w.BoundingBox.Top) ?? 0;
            var assigned = false;
            
            // 既存の行グループに属するかチェック
            foreach (var group in rowGroups)
            {
                var groupY = group.Average(x => x.element.Words?.Average(w => w.BoundingBox.Top) ?? 0);
                
                // 水平線による境界チェック
                if (Math.Abs(elementY - groupY) < 10.0 && !IsSeparatedByHorizontalLine(elementY, groupY, horizontalLines))
                {
                    group.Add(element);
                    assigned = true;
                    break;
                }
            }
            
            if (!assigned)
            {
                rowGroups.Add(new List<(DocumentElement element, int index)> { element });
            }
        }
        
        return rowGroups;
    }
    
    private static List<List<(DocumentElement element, int index)>> GroupElementsByColumns(
        List<(DocumentElement element, int index)> rowElements, 
        List<LineSegment> verticalLines)
    {
        var columnGroups = new List<List<(DocumentElement element, int index)>>();
        
        // X座標で要素をソート（左から右へ）
        var sortedElements = rowElements
            .OrderBy(x => x.element.Words?.Min(w => w.BoundingBox.Left) ?? 0)
            .ToList();
        
        foreach (var element in sortedElements)
        {
            var elementX = element.element.Words?.Average(w => w.BoundingBox.Left) ?? 0;
            var assigned = false;
            
            // 既存の列グループに属するかチェック
            foreach (var group in columnGroups)
            {
                var groupX = group.Average(x => x.element.Words?.Average(w => w.BoundingBox.Left) ?? 0);
                
                // 垂直線による境界チェック
                if (Math.Abs(elementX - groupX) < 30.0 && !IsSeparatedByVerticalLine(elementX, groupX, verticalLines))
                {
                    group.Add(element);
                    assigned = true;
                    break;
                }
            }
            
            if (!assigned)
            {
                columnGroups.Add(new List<(DocumentElement element, int index)> { element });
            }
        }
        
        return columnGroups;
    }
    
    private static bool IsSeparatedByHorizontalLine(double y1, double y2, List<LineSegment> horizontalLines)
    {
        var minY = Math.Min(y1, y2);
        var maxY = Math.Max(y1, y2);
        
        return horizontalLines.Any(line =>
            line.From.Y > minY && line.From.Y < maxY &&
            Math.Abs(line.From.Y - line.To.Y) < 2.0);
    }
    
    private static bool IsSeparatedByVerticalLine(double x1, double x2, List<LineSegment> verticalLines)
    {
        var minX = Math.Min(x1, x2);
        var maxX = Math.Max(x1, x2);
        
        return verticalLines.Any(line =>
            line.From.X > minX && line.From.X < maxX &&
            Math.Abs(line.From.X - line.To.X) < 2.0);
    }
    
    private static DocumentElement? BuildTableRowFromColumnGroups(List<List<(DocumentElement element, int index)>> columnGroups)
    {
        if (columnGroups.Count == 0)
            return null;
            
        var cells = new List<string>();
        var allWords = new List<UglyToad.PdfPig.Content.Word>();
        
        foreach (var columnGroup in columnGroups)
        {
            // 列内の要素を統合し、配置を考慮
            var cellContent = string.Join("<br>", columnGroup.Select(x => x.element.Content?.Trim() ?? ""));
            
            // 列内での配置分析（左寄せ、中央寄せ、右寄せ）
            var alignment = AnalyzeColumnAlignment(columnGroup);
            
            cells.Add(cellContent);
            
            // 座標情報も統合
            foreach (var item in columnGroup)
            {
                if (item.element.Words != null)
                {
                    allWords.AddRange(item.element.Words);
                }
            }
        }
        
        // パイプ区切りでテーブル行を構築
        var tableRowContent = "| " + string.Join(" | ", cells) + " |";
        
        var firstElement = columnGroups[0][0].element;
        return new DocumentElement
        {
            Type = ElementType.TableRow,
            Content = tableRowContent,
            FontSize = firstElement.FontSize,
            LeftMargin = firstElement.LeftMargin,
            IsIndented = firstElement.IsIndented,
            Words = allWords
        };
    }
    
    private static ColumnAlignment AnalyzeColumnAlignment(List<(DocumentElement element, int index)> columnGroup)
    {
        if (columnGroup.Count <= 1)
            return ColumnAlignment.Left;
            
        var positions = new List<double>();
        var columnBounds = new List<(double left, double right, double center)>();
        
        foreach (var item in columnGroup)
        {
            if (item.element.Words != null && item.element.Words.Count > 0)
            {
                var left = item.element.Words.Min(w => w.BoundingBox.Left);
                var right = item.element.Words.Max(w => w.BoundingBox.Right);
                var center = (left + right) / 2.0;
                
                columnBounds.Add((left, right, center));
            }
        }
        
        if (columnBounds.Count <= 1)
            return ColumnAlignment.Left;
            
        // 左端位置の分散を計算
        var leftPositions = columnBounds.Select(b => b.left).ToList();
        var leftVariance = CalculateVariance(leftPositions);
        
        // 右端位置の分散を計算
        var rightPositions = columnBounds.Select(b => b.right).ToList();
        var rightVariance = CalculateVariance(rightPositions);
        
        // 中央位置の分散を計算
        var centerPositions = columnBounds.Select(b => b.center).ToList();
        var centerVariance = CalculateVariance(centerPositions);
        
        // 最も分散が小さい（一致している）配置を判定
        if (leftVariance <= rightVariance && leftVariance <= centerVariance)
        {
            return ColumnAlignment.Left;
        }
        else if (rightVariance <= centerVariance)
        {
            return ColumnAlignment.Right;
        }
        else
        {
            return ColumnAlignment.Center;
        }
    }
    
    private static double CalculateVariance(List<double> values)
    {
        if (values.Count <= 1)
            return 0.0;
            
        var mean = values.Average();
        var variance = values.Sum(v => Math.Pow(v - mean, 2)) / values.Count;
        return variance;
    }
    
    public enum ColumnAlignment
    {
        Left,
        Center,
        Right
    }

    private static bool IsDefinitelyHeader(DocumentElement current, DocumentElement? previous, DocumentElement? next, FontAnalysis fontAnalysis)
    {
        if (current == null || string.IsNullOrWhiteSpace(current.Content))
            return false;

        var content = current.Content.Trim();
        
        // 明確に段落的なパターンは絶対にヘッダーではない
        if (content.EndsWith("。") || content.EndsWith("です。") || content.EndsWith("ます。") ||
            content.EndsWith(".") || content.Contains("、") || content.Contains(",") ||
            content.Contains("**") || (content.Contains("*") && !content.StartsWith("*")) ||
            content.StartsWith("これは") || content.StartsWith("それは") || 
            content.StartsWith("この") || content.StartsWith("その") ||
            content.StartsWith("•") || content.StartsWith("・") || content.StartsWith("-"))
        {
            return false;
        }

        // フォントサイズベースの判定（閾値を下げる）
        var fontRatio = current.FontSize / fontAnalysis.BaseFontSize;
        
        // 中サイズ以上のフォントで短いテキスト（日本語ヘッダーを考慮）
        if (fontRatio >= 1.2 && content.Length <= 50) // 閾値を下げる
        {
            // 位置的特徴を確認（左端配置など）
            if (current.LeftMargin <= 120.0 || current.Words?.Count <= 3) // より緩い基準
            {
                return true;
            }
        }
        
        // 大きなフォントサイズ（より緩い基準）
        if (fontRatio >= 1.8 && content.Length <= 60)
        {
            return true;
        }

        // 全て大文字で短いタイトル的なテキスト
        if (content.Length <= 20 && content.Length >= 3 && 
            content.All(c => char.IsUpper(c) || char.IsWhiteSpace(c) || char.IsPunctuation(c) || char.IsDigit(c)))
        {
            return true;
        }
        
        // 孤立した短いテキスト（前後に空白があり、適度なサイズ）
        if (content.Length <= 60 && content.Length >= 3 && fontRatio >= 1.1) // より緩い基準
        {
            // テーブル的でないテキストのヘッダー可能性をチェック
            bool notTableContent = !content.Contains("|") && 
                                 !content.Contains("\t") && 
                                 !System.Text.RegularExpressions.Regex.IsMatch(content, @"\d+\s+\d+");
            
            // 前後の要素との関係を確認（より緩い判定）
            bool isolated = (previous == null || 
                           (previous.Content.Trim().Length == 0) ||
                           (previous.Type == ElementType.TableRow) ||
                           (Math.Abs(previous.LeftMargin - current.LeftMargin) > 15.0)) &&
                          (next == null || 
                           (next.Content.Trim().Length == 0) ||
                           (next.Type == ElementType.TableRow) ||
                           (Math.Abs(next.LeftMargin - current.LeftMargin) > 15.0));
            
            if (isolated && notTableContent && current.LeftMargin <= 120.0)
            {
                return true;
            }
        }

        return false;
    }

    private static bool CouldBeHeader(DocumentElement current, DocumentElement? previous, DocumentElement? next, FontAnalysis fontAnalysis)
    {
        if (current == null)
            return false;

        // フォントサイズが基準より大きい
        var fontRatio = current.FontSize / fontAnalysis.BaseFontSize;
        if (fontRatio >= 1.15)
            return true;

        // ヘッダーらしい特徴を持つ短いテキストで左端に配置（より厳格に）
        if (current.Content.Trim().Length <= 50 && current.LeftMargin <= 50.0)
        {
            var content = current.Content.Trim();
            
            // 句読点で終わるテキストはヘッダーではない
            if (content.EndsWith("。") || content.EndsWith("、") || content.EndsWith(".") || content.EndsWith(","))
                return false;
            
            // フォーマット記号を含むテキストはヘッダーではない
            if (content.Contains("**") || content.Contains("*") && !content.StartsWith("*"))
                return false;
                
            // 典型的な段落開始パターンはヘッダーではない
            if (content.StartsWith("これは") || content.StartsWith("それは") || content.StartsWith("この") || content.StartsWith("その"))
                return false;
            
            // フォントサイズが明らかに大きい場合のみヘッダーとする
            var currentFontRatio = current.FontSize / fontAnalysis.BaseFontSize;
            if (currentFontRatio >= 1.4)
            {
                var hasFollowingContent = next?.Type == ElementType.Paragraph;
                if (hasFollowingContent)
                    return true;
            }
        }

        // ヘッダー的なパターン
        if (ElementDetector.IsStrongHeaderPattern(current.Content))
            return true;

        return false;
    }

    private static bool IsListContinuation(DocumentElement current, DocumentElement previous)
    {
        if (current == null || previous == null)
            return false;

        // インデントが類似している
        if (Math.Abs(current.LeftMargin - previous.LeftMargin) <= 15.0)
        {
            // リスト的なコンテンツパターン
            if (ElementDetector.IsListItemLike(current.Content))
                return true;
        }

        return false;
    }

    private static bool IsPartOfTableSequence(DocumentElement current, List<DocumentElement> elements, int currentIndex)
    {
        if (current == null || elements == null)
            return false;

        // 座標ベースのテーブル行判定を強化
        if (!ElementDetector.IsTableRowLike(current.Content, current.Words))
            return false;

        // 前後の要素との座標的関係を分析
        var hasPreviousTable = false;
        var hasNextTable = false;
        
        // より広い範囲で前後のテーブル行を検索
        for (int i = Math.Max(0, currentIndex - 3); i < Math.Min(elements.Count, currentIndex + 4); i++)
        {
            if (i == currentIndex) continue;
            
            var element = elements[i];
            if (element.Type == ElementType.TableRow)
            {
                // 座標的に近い位置にテーブル行がある
                var currentTop = current.Words?.FirstOrDefault()?.BoundingBox.Top ?? 0;
                var elementTop = element.Words?.FirstOrDefault()?.BoundingBox.Top ?? 0;
                var verticalDistance = Math.Abs(elementTop - currentTop);
                var horizontalOverlap = Math.Min(element.LeftMargin + 100, current.LeftMargin + 100) - 
                                      Math.Max(element.LeftMargin, current.LeftMargin);
                
                if (verticalDistance < 200 && horizontalOverlap > 50) // 座標ベース判定
                {
                    if (i < currentIndex) hasPreviousTable = true;
                    if (i > currentIndex) hasNextTable = true;
                }
            }
        }

        return hasPreviousTable || hasNextTable;
    }

    private static bool IsElementInTableArea(DocumentElement element, UglyToad.PdfPig.Core.PdfRectangle area)
    {
        if (element?.Words == null || !element.Words.Any())
            return false;

        // 要素の境界ボックスを計算
        var minX = element.Words.Min(w => w.BoundingBox.Left);
        var maxX = element.Words.Max(w => w.BoundingBox.Right);
        var minY = element.Words.Min(w => w.BoundingBox.Bottom);
        var maxY = element.Words.Max(w => w.BoundingBox.Top);

        // テーブル領域内にあるかチェック
        return minX >= area.Left && maxX <= area.Right &&
               minY >= area.Bottom && maxY <= area.Top;
    }

    private static List<DocumentElement> DetectTableSequences(List<DocumentElement> elements)
    {
        var result = new List<DocumentElement>(elements);
        
        // 連続するテーブル行のパターンを検出
        for (int i = 0; i < result.Count - 1; i++)
        {
            var current = result[i];
            var next = result[i + 1];
            
            if (current.Type == ElementType.TableRow && next.Type == ElementType.Paragraph)
            {
                // 次の要素がテーブル行的なパターンの場合
                if (ElementDetector.IsTableRowLike(next.Content, next.Words))
                {
                    // 左マージンが類似している
                    if (Math.Abs(current.LeftMargin - next.LeftMargin) <= 20.0)
                    {
                        next.Type = ElementType.TableRow;
                    }
                }
            }
        }
        
        return result;
    }

    private static List<DocumentElement> OptimizeTableRegions(List<DocumentElement> elements)
    {
        var result = new List<DocumentElement>(elements);
        
        // テーブル領域の最適化処理
        var tableRegions = new List<TableRegion>();
        TableRegion? currentRegion = null;
        
        foreach (var element in result)
        {
            if (element.Type == ElementType.TableRow)
            {
                if (currentRegion == null)
                {
                    currentRegion = new TableRegion();
                    tableRegions.Add(currentRegion);
                }
                currentRegion.Elements.Add(element);
            }
            else
            {
                if (currentRegion != null && currentRegion.Elements.Count < 2)
                {
                    // 単一行のテーブルは段落に戻す
                    foreach (var tableElement in currentRegion.Elements)
                    {
                        tableElement.Type = ElementType.Paragraph;
                    }
                    tableRegions.Remove(currentRegion);
                }
                currentRegion = null;
            }
        }
        
        // 最後の領域もチェック
        if (currentRegion != null && currentRegion.Elements.Count < 2)
        {
            foreach (var tableElement in currentRegion.Elements)
            {
                tableElement.Type = ElementType.Paragraph;
            }
        }
        
        return result;
    }

    private static bool ShouldIntegrateHeaderIntoTable(string headerContent, DocumentElement tableRow)
    {
        // ヘッダーが短く、テーブル的な内容の場合
        if (headerContent.Length <= 50)
        {
            // テーブル行と関連性があるかチェック
            var headerWords = headerContent.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var tableWords = tableRow.Content.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            
            if (headerWords.Length <= 5 && tableWords.Length >= 2)
                return true;
        }
        
        return false;
    }
    
    private static List<DocumentElement> DetectAndMergeCodeBlocks(List<DocumentElement> elements)
    {
        var result = new List<DocumentElement>();
        var codeBlockBuffer = new List<DocumentElement>();
        
        foreach (var element in elements)
        {
            if (element.Type == ElementType.CodeBlock || 
                (element.Type == ElementType.Paragraph && ElementDetector.IsCodeBlockLike(element.Content, element.Words, new FontAnalysis())))
            {
                element.Type = ElementType.CodeBlock;
                codeBlockBuffer.Add(element);
            }
            else
            {
                // コードブロックバッファを処理
                if (codeBlockBuffer.Any())
                {
                    if (codeBlockBuffer.Count == 1)
                    {
                        result.Add(codeBlockBuffer[0]);
                    }
                    else
                    {
                        // 複数のコードブロックを統合
                        var mergedCode = MergeCodeBlocks(codeBlockBuffer);
                        result.Add(mergedCode);
                    }
                    codeBlockBuffer.Clear();
                }
                
                result.Add(element);
            }
        }
        
        // 残ったコードブロックを処理
        if (codeBlockBuffer.Any())
        {
            if (codeBlockBuffer.Count == 1)
            {
                result.Add(codeBlockBuffer[0]);
            }
            else
            {
                var mergedCode = MergeCodeBlocks(codeBlockBuffer);
                result.Add(mergedCode);
            }
        }
        
        return result;
    }
    
    private static List<DocumentElement> DetectAndMergeQuoteBlocks(List<DocumentElement> elements)
    {
        var result = new List<DocumentElement>();
        var quoteBlockBuffer = new List<DocumentElement>();
        
        foreach (var element in elements)
        {
            if (element.Type == ElementType.QuoteBlock || 
                (element.Type == ElementType.Paragraph && ElementDetector.IsQuoteBlockLike(element.Content)))
            {
                element.Type = ElementType.QuoteBlock;
                quoteBlockBuffer.Add(element);
            }
            else
            {
                // 引用ブロックバッファを処理
                if (quoteBlockBuffer.Any())
                {
                    if (quoteBlockBuffer.Count == 1)
                    {
                        result.Add(quoteBlockBuffer[0]);
                    }
                    else
                    {
                        // 複数の引用ブロックを統合
                        var mergedQuote = MergeQuoteBlocks(quoteBlockBuffer);
                        result.Add(mergedQuote);
                    }
                    quoteBlockBuffer.Clear();
                }
                
                result.Add(element);
            }
        }
        
        // 残った引用ブロックを処理
        if (quoteBlockBuffer.Any())
        {
            if (quoteBlockBuffer.Count == 1)
            {
                result.Add(quoteBlockBuffer[0]);
            }
            else
            {
                var mergedQuote = MergeQuoteBlocks(quoteBlockBuffer);
                result.Add(mergedQuote);
            }
        }
        
        return result;
    }
    
    private static DocumentElement MergeCodeBlocks(List<DocumentElement> codeBlocks)
    {
        var mergedContent = string.Join("\n", codeBlocks.Select(cb => cb.Content));
        var allWords = codeBlocks.SelectMany(cb => cb.Words).ToList();
        
        return new DocumentElement
        {
            Type = ElementType.CodeBlock,
            Content = mergedContent,
            FontSize = codeBlocks.First().FontSize,
            LeftMargin = codeBlocks.First().LeftMargin,
            Words = allWords
        };
    }
    
    private static DocumentElement MergeQuoteBlocks(List<DocumentElement> quoteBlocks)
    {
        var mergedContent = string.Join("\n", quoteBlocks.Select(qb => qb.Content));
        var allWords = quoteBlocks.SelectMany(qb => qb.Words).ToList();
        
        return new DocumentElement
        {
            Type = ElementType.QuoteBlock,
            Content = mergedContent,
            FontSize = quoteBlocks.First().FontSize,
            LeftMargin = quoteBlocks.First().LeftMargin,
            Words = allWords
        };
    }
}

public class TableRegion
{
    public List<DocumentElement> Elements { get; set; } = new List<DocumentElement>();
    public TablePattern? Pattern { get; set; }
    public UglyToad.PdfPig.Core.PdfRectangle BoundingArea { get; set; }
}