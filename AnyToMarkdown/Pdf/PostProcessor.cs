using System.Linq;
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
        
        // 座標ベースのテーブル要素統合処理
        result = ConsolidateTableElementsByCoordinates(result, graphicsInfo);
        
        // 連続するテーブル行の検出
        result = DetectTableSequences(result);
        
        // テーブル領域の最適化
        result = OptimizeTableRegions(result);
        
        return result;
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
        // この処理は無効化し、代わりに座標ベースのテーブル検出を使用
        // 座標情報なしでは正確な判定が困難なため
        return elements;
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
    
    private static bool ShouldMergeIntoTableCell(DocumentElement tableRow, DocumentElement paragraph)
    {
        var paragraphText = paragraph.Content.Trim();
        
        // 空の段落は統合しない
        if (string.IsNullOrWhiteSpace(paragraphText))
        {
            return false;
        }
        
        // 段落が短く、テーブル行の続きと思われる場合
        if (paragraphText.Length < 100 && !paragraphText.Contains("|"))
        {
            // 段落がテーブル行に近い位置にある
            if (Math.Abs(tableRow.LeftMargin - paragraph.LeftMargin) < 30.0)
            {
                return true;
            }
        }
        
        // 段落が明らかに独立した内容でない場合（句読点で終わらない）
        if (!paragraphText.EndsWith("。") && !paragraphText.EndsWith(".") && 
            !paragraphText.EndsWith("！") && !paragraphText.EndsWith("!") &&
            paragraphText.Length < 50)
        {
            return true;
        }
        
        return false;
    }

    private static bool ShouldMergeParagraphs(DocumentElement previous, DocumentElement current)
    {
        if (previous == null || current == null)
            return false;

        // フォントサイズが似ている
        if (Math.Abs(previous.FontSize - current.FontSize) > 2.0)
            return false;

        // 左マージンが似ている
        if (Math.Abs(previous.LeftMargin - current.LeftMargin) > 10.0)
            return false;

        // 両方とも短いテキスト（継続の可能性）
        if (previous.Content.Trim().Length < 50 && current.Content.Trim().Length < 50)
            return true;

        // 前の段落が文で終わっていない（継続の可能性）
        var prevContent = previous.Content.Trim();
        if (!prevContent.EndsWith(".") && !prevContent.EndsWith("。") && 
            !prevContent.EndsWith("!") && !prevContent.EndsWith("！") &&
            !prevContent.EndsWith("?") && !prevContent.EndsWith("？"))
            return true;

        return false;
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

        // 非常に大きなフォントサイズのみ（2.0倍以上）
        var fontRatio = current.FontSize / fontAnalysis.BaseFontSize;
        if (fontRatio >= 2.0 && content.Length <= 30)
        {
            // 強いヘッダーパターンも確認
            if (ElementDetector.IsStrongHeaderPattern(content))
                return true;
        }

        // 全て大文字で短いタイトル的なテキスト
        if (content.Length <= 20 && content.Length >= 3 && 
            content.All(c => char.IsUpper(c) || char.IsWhiteSpace(c) || char.IsPunctuation(c) || char.IsDigit(c)))
        {
            return true;
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

        // 前後にテーブル行がある
        var hasPreviousTable = currentIndex > 0 && elements[currentIndex - 1].Type == ElementType.TableRow;
        var hasNextTable = currentIndex < elements.Count - 1 && elements[currentIndex + 1].Type == ElementType.TableRow;

        if (hasPreviousTable || hasNextTable)
        {
            // テーブル行的なコンテンツ
            if (ElementDetector.IsTableRowLike(current.Content, current.Words))
                return true;
        }

        return false;
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
    
    public bool Contains(DocumentElement element)
    {
        if (element?.Words == null || !element.Words.Any())
            return false;

        var elementMinX = element.Words.Min(w => w.BoundingBox.Left);
        var elementMaxX = element.Words.Max(w => w.BoundingBox.Right);
        var elementMinY = element.Words.Min(w => w.BoundingBox.Bottom);
        var elementMaxY = element.Words.Max(w => w.BoundingBox.Top);

        return elementMinX >= BoundingArea.Left && elementMaxX <= BoundingArea.Right &&
               elementMinY >= BoundingArea.Bottom && elementMaxY <= BoundingArea.Top;
    }
}

public class HeaderCoordinateAnalysis
{
    public double AverageLeftMargin { get; set; }
    public double VerticalSpacing { get; set; }
    public List<double> FontSizes { get; set; } = new List<double>();
    public List<HeaderLevelInfo> CoordinateLevels { get; set; } = new List<HeaderLevelInfo>();
}

public class HeaderLevelInfo
{
    public int Level { get; set; }
    public double FontSize { get; set; }
    public double LeftMargin { get; set; }
    public string Pattern { get; set; } = "";
    public double Coordinate { get; set; }
    public int Count { get; set; }
    public double AvgFontSize { get; set; }
    public double Consistency { get; set; }
}

public class HeaderCandidate
{
    public DocumentElement Element { get; set; } = null!;
    public int SuggestedLevel { get; set; }
    public double Confidence { get; set; }
    public string Reason { get; set; } = "";
    public double LeftPosition { get; set; }
    public double FontSize { get; set; }
    public double FontSizeRatio { get; set; }
    public bool IsCurrentlyHeader { get; set; }
}
