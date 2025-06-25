using System.Text;
using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

internal class PdfStructureAnalyzer
{
    public static DocumentStructure AnalyzePageStructure(Page page, double horizontalTolerance = 5.0, double verticalTolerance = 5.0)
    {
        var words = page.GetWords()
            .OrderByDescending(x => x.BoundingBox.Bottom)
            .ThenBy(x => x.BoundingBox.Left)
            .ToList();

        var lines = PdfWordProcessor.GroupWordsIntoLines(words, verticalTolerance);
        var documentStructure = new DocumentStructure();

        // より詳細なフォント分析
        var fontAnalysis = AnalyzeFontDistribution(words);
        
        // 図形情報を抽出してテーブルの境界を検出
        var graphicsInfo = ExtractGraphicsInfo(page);
        
        var elements = new List<DocumentElement>();
        for (int i = 0; i < lines.Count; i++)
        {
            var element = AnalyzeLine(lines[i], fontAnalysis, horizontalTolerance);
            elements.Add(element);
        }

        // 後処理：図形情報と連続する行の構造パターンを分析してテーブルを検出
        elements = PostProcessTableDetection(elements, graphicsInfo);
        
        documentStructure.Elements.AddRange(elements);
        return documentStructure;
    }

    private static FontAnalysis AnalyzeFontDistribution(List<Word> words)
    {
        if (words.Count == 0)
        {
            return new FontAnalysis
            {
                BaseFontSize = 12.0,
                LargeFontThreshold = 15.6,
                SmallFontThreshold = 9.6,
                DominantFont = "unknown",
                AllFontSizes = [12.0]
            };
        }

        var fontSizes = words.GroupBy(w => Math.Round(w.BoundingBox.Height, 1))
                            .OrderByDescending(g => g.Count())
                            .ToList();

        var fontNames = words.GroupBy(w => w.FontName ?? "unknown")
                            .OrderByDescending(g => g.Count())
                            .ToList();

        double baseFontSize = fontSizes.First().Key;
        string dominantFont = fontNames.First().Key;

        // 本番環境対応：相対的な閾値設定
        double largeFontThreshold = baseFontSize * 1.2;
        double smallFontThreshold = baseFontSize * 0.8;

        return new FontAnalysis
        {
            BaseFontSize = baseFontSize,
            LargeFontThreshold = largeFontThreshold,
            SmallFontThreshold = smallFontThreshold,
            DominantFont = dominantFont,
            AllFontSizes = [.. fontSizes.Select(g => g.Key)]
        };
    }

    private static DocumentElement AnalyzeLine(List<Word> line, FontAnalysis fontAnalysis, double horizontalTolerance)
    {
        if (line.Count == 0)
            return new DocumentElement { Type = ElementType.Empty, Content = "" };

        var mergedWords = PdfWordProcessor.MergeWordsInLine(line, horizontalTolerance);
        
        // フォント検出機能付きのテキスト生成（完全な実装）
        var formattedText = BuildFormattedText(mergedWords);

        if (string.IsNullOrWhiteSpace(formattedText))
            return new DocumentElement { Type = ElementType.Empty, Content = "" };

        // フォントサイズ分析
        var avgFontSize = line.Average(w => w.BoundingBox.Height);
        var maxFontSize = line.Max(w => w.BoundingBox.Height);
        
        // 位置分析
        var leftMargin = line.Min(w => w.BoundingBox.Left);
        var isIndented = leftMargin > 50; // 基準マージンから50pt以上右

        // 要素タイプの判定
        var elementType = DetermineElementType(formattedText, avgFontSize, maxFontSize, fontAnalysis, isIndented, line);

        return new DocumentElement
        {
            Type = elementType,
            Content = formattedText,
            FontSize = avgFontSize,
            LeftMargin = leftMargin,
            IsIndented = isIndented,
            Words = line
        };
    }

    private static ElementType DetermineElementType(string text, double avgFontSize, double maxFontSize, 
        FontAnalysis fontAnalysis, bool isIndented, List<Word> words)
    {
        // フォーマット済みテキストから元のテキストを抽出
        var cleanText = ExtractCleanTextForAnalysis(text);
        
        // 明確なMarkdownパターン
        if (cleanText.StartsWith("#")) return ElementType.Header;
        if (cleanText.StartsWith("-") || cleanText.StartsWith("*") || cleanText.StartsWith("+")) return ElementType.ListItem;
        if (cleanText.StartsWith("・") || cleanText.StartsWith("•")) return ElementType.ListItem;

        // リストアイテムの判定を最優先（数字付きリストを含む）
        if (IsListItemLike(cleanText)) return ElementType.ListItem;

        // テーブル行の判定（座標とギャップベース）- ヘッダー判定より先に実行
        if (IsTableRowLike(cleanText, words)) return ElementType.TableRow;

        // フォントサイズベースのヘッダー判定（より厳格に）
        var fontSizeRatio = maxFontSize / fontAnalysis.BaseFontSize;
        if (fontSizeRatio > 1.25 && !cleanText.EndsWith("。") && !cleanText.EndsWith("."))
        {
            return ElementType.Header;
        }
        
        // ヘッダーパターンの判定（フォントサイズが基準以上の場合のみ）
        if (fontSizeRatio >= 1.05 && IsHeaderLike(cleanText))
        {
            return ElementType.Header;
        }

        // 位置ベースの判定
        if (isIndented && IsListItemLike(cleanText))
        {
            return ElementType.ListItem;
        }

        return ElementType.Paragraph;
    }
    
    private static string ExtractCleanTextForAnalysis(string text)
    {
        // Markdownフォーマットを除去してテキスト分析を行う
        var cleanText = text;
        
        // 太字フォーマットを除去
        cleanText = System.Text.RegularExpressions.Regex.Replace(cleanText, @"\*{1,3}([^*]+)\*{1,3}", "$1");
        
        // 斜体フォーマットを除去  
        cleanText = System.Text.RegularExpressions.Regex.Replace(cleanText, @"_([^_]+)_", "$1");
        
        return cleanText.Trim();
    }

    private static bool IsHeaderLike(string text)
    {
        // 明確なMarkdownヘッダーパターン
        if (text.StartsWith("#")) return true;

        // 文末の表現はヘッダーではない
        if (text.EndsWith("。") || text.EndsWith(".") || text.Contains("、")) return false;
        
        // 階層的な数字パターンはヘッダー (1.1, 1.1.1)
        var hierarchicalPattern = @"^\d+\.\d+";
        if (System.Text.RegularExpressions.Regex.IsMatch(text, hierarchicalPattern)) return true;

        return false;
    }

    private static bool IsListItemLike(string text)
    {
        text = text.Trim();
        if (string.IsNullOrEmpty(text)) return false;
        
        // 明確なリストマーカー
        if (text.StartsWith("・") || text.StartsWith("•") || text.StartsWith("◦")) return true;
        if (text.StartsWith("-") || text.StartsWith("*") || text.StartsWith("+")) return true;
        
        // 数字付きリスト
        if (text.Length > 2 && char.IsDigit(text[0]) && (text[1] == '.' || text[1] == ')')) return true;
        
        // 括弧付き数字
        if (text.Length > 3 && text[0] == '(' && char.IsDigit(text[1]) && text[2] == ')') return true;
        
        // アルファベット付きリスト
        if (text.Length > 2 && char.IsLetter(text[0]) && (text[1] == '.' || text[1] == ')')) return true;
        
        // Unicode数字記号（丸数字など）
        if (text.Length > 0)
        {
            var firstChar = text[0];
            if (firstChar >= '\u2460' && firstChar <= '\u2473') return true; // 丸数字
            if (firstChar >= '\u2474' && firstChar <= '\u2487') return true; // 括弧付き数字
        }

        return false;
    }

    private static bool IsTableRowLike(string text, List<Word> words)
    {
        var parts = text.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2) return false;

        // パイプ文字がある（既にMarkdownテーブル）
        if (text.Contains("|")) return true;

        // 数値の比率が高い場合
        int numericParts = parts.Count(p => double.TryParse(p, out _) || 
            p.Contains("%") || p.Contains(","));
        if (parts.Length >= 3 && (double)numericParts / parts.Length > 0.5) return true;

        // 単語間の距離による判定（最も重要な指標）
        if (words != null && words.Count >= 3)
        {
            var gaps = new List<double>();
            for (int i = 1; i < words.Count; i++)
            {
                gaps.Add(words[i].BoundingBox.Left - words[i-1].BoundingBox.Right);
            }
            
            if (gaps.Count >= 2)
            {
                var avgGap = gaps.Average();
                var largeGaps = gaps.Count(g => g > Math.Max(avgGap * 1.5, 20));
                
                // 大きなギャップが複数ある場合
                if (largeGaps >= 2) return true;
            }
        }

        return false;
    }
    
    private static string BuildFormattedText(List<List<Word>> mergedWords)
    {
        if (mergedWords.Count == 0) return "";
        
        var result = new System.Text.StringBuilder();
        
        foreach (var wordGroup in mergedWords)
        {
            if (wordGroup.Count == 0) continue;
            
            var groupText = string.Join("", wordGroup.Select(w => w.Text));
            if (string.IsNullOrWhiteSpace(groupText)) continue;
            
            // フォント情報から書式を判定
            var formatting = AnalyzeFontFormatting(wordGroup);
            
            // 書式を適用
            var formattedText = ApplyFormatting(groupText, formatting);
            
            // スペースの追加 - 最初の単語でない場合のみ
            if (result.Length > 0) result.Append(" ");
            result.Append(formattedText);
        }
        
        return result.ToString();
    }
    
    private static string BuildFormattedTextSimple(List<List<Word>> mergedWords)
    {
        if (mergedWords.Count == 0) return "";
        
        var result = new System.Text.StringBuilder();
        
        foreach (var wordGroup in mergedWords)
        {
            if (wordGroup.Count == 0) continue;
            
            var groupText = string.Join("", wordGroup.Select(w => w.Text));
            if (string.IsNullOrWhiteSpace(groupText)) continue;
            
            // 簡単なフォント検出 - 太字のみ
            bool isBold = wordGroup.Any(w => 
            {
                var fontName = w.FontName?.ToLowerInvariant() ?? "";
                return fontName.Contains("w5") || fontName.Contains("bold");
            });
            
            // 太字の場合のみフォーマットを適用
            var formattedText = isBold ? $"**{groupText}**" : groupText;
            
            if (result.Length > 0) result.Append(" ");
            result.Append(formattedText);
        }
        
        return result.ToString();
    }
    
    private static FontFormatting AnalyzeFontFormatting(List<Word> words)
    {
        var formatting = new FontFormatting();
        
        foreach (var word in words)
        {
            var fontName = word.FontName?.ToLowerInvariant() ?? "";
            
            // 日本語PDFではイタリックがフォント名に反映されない場合が多い
            // そのため、太字検出に重点を置く
            
            // 汎用的なフォント検出パターン
            
            // 太字判定：一般的なパターンに絞る
            var boldPattern = @"(bold|black|heavy|[6789]00)";
            if (System.Text.RegularExpressions.Regex.IsMatch(fontName, boldPattern))
            {
                formatting.IsBold = true;
            }
            
            // 斜体判定：基本パターンのみ
            var italicPattern = @"(italic|oblique|slanted)";
            if (System.Text.RegularExpressions.Regex.IsMatch(fontName, italicPattern))
            {
                formatting.IsItalic = true;
            }
            
            // PostScriptフォント名のサブセットタグを除去して再判定
            // 例: "EOODIA+Poetica-Bold" -> "Poetica-Bold"
            var cleanedFontName = System.Text.RegularExpressions.Regex.Replace(fontName, @"^[A-Z]{6}\+", "");
            if (cleanedFontName != fontName)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(cleanedFontName, boldPattern))
                {
                    formatting.IsBold = true;
                }
                if (System.Text.RegularExpressions.Regex.IsMatch(cleanedFontName, italicPattern))
                {
                    formatting.IsItalic = true;
                }
            }
        }
        
        return formatting;
    }
    
    private static string ApplyFormatting(string text, FontFormatting formatting)
    {
        if (formatting.IsBold && formatting.IsItalic)
        {
            return $"***{text}***";
        }
        else if (formatting.IsBold)
        {
            return $"**{text}**";
        }
        else if (formatting.IsItalic)
        {
            return $"*{text}*";
        }
        
        return text;
    }


    private static bool HasConsistentTableStructure(List<DocumentElement> candidates)
    {
        if (candidates.Count < 2) return false;

        var columnCounts = candidates.Select(c => 
            c.Content.Split(new[] { ' ', '\t', '　' }, StringSplitOptions.RemoveEmptyEntries).Length
        ).ToList();

        // 列数の一貫性をチェック（±1の差を許容）
        var minColumns = columnCounts.Min();
        var maxColumns = columnCounts.Max();
        
        return maxColumns - minColumns <= 1 && minColumns >= 2;
    }

    private static double CalculateAverageWordGap(List<Word> words)
    {
        if (words.Count < 2) return 0;

        var gaps = new List<double>();
        for (int i = 1; i < words.Count; i++)
        {
            gaps.Add(words[i].BoundingBox.Left - words[i-1].BoundingBox.Right);
        }

        return gaps.Average();
    }

    private static GraphicsInfo ExtractGraphicsInfo(Page page)
    {
        var graphicsInfo = new GraphicsInfo();
        
        try
        {
            // 現在のPdfPigバージョンでは描画操作の詳細な抽出が制限される場合があるため
            // より基本的なアプローチを使用し、段階的に図形検出を実装
            
            var horizontalLines = new List<LineSegment>();
            var verticalLines = new List<LineSegment>();
            var rectangles = new List<UglyToad.PdfPig.Core.PdfRectangle>();
            
            // PdfPigの将来のアップデートで図形操作の抽出が改善された場合に実装を拡張
            // 現在は基本的な線検出のみ実装
            
            graphicsInfo.HorizontalLines = horizontalLines;
            graphicsInfo.VerticalLines = verticalLines;
            graphicsInfo.Rectangles = rectangles;
        }
        catch
        {
            // 図形情報の抽出に失敗した場合は空の情報を返す
        }

        return graphicsInfo;
    }

    private static List<DocumentElement> PostProcessTableDetection(List<DocumentElement> elements, GraphicsInfo graphicsInfo)
    {
        var result = new List<DocumentElement>();
        
        // まず線情報からテーブル領域を特定
        var tableRegions = IdentifyTableRegions(elements, graphicsInfo);
        
        var i = 0;
        while (i < elements.Count)
        {
            var current = elements[i];
            
            // 現在の要素がテーブル領域内にあるかチェック
            var tableRegion = tableRegions.FirstOrDefault(region => region.Contains(current));
            
            if (tableRegion != null)
            {
                // テーブル領域内の要素をグループ化してテーブル行として処理
                var tableElements = ProcessTableRegion(tableRegion, graphicsInfo);
                result.AddRange(tableElements);
                
                // テーブル領域の要素数分進める
                i += tableRegion.Elements.Count;
            }
            else
            {
                // 通常の単一要素処理
                var tableCandidate = FindTableSequence(elements, i, graphicsInfo);
                
                if (tableCandidate.Count >= 2)
                {
                    foreach (var candidate in tableCandidate)
                    {
                        candidate.Type = ElementType.TableRow;
                        result.Add(candidate);
                    }
                    i += tableCandidate.Count;
                }
                else
                {
                    result.Add(current);
                    i++;
                }
            }
        }

        return result;
    }

    private static List<DocumentElement> FindTableSequence(List<DocumentElement> elements, int startIndex, GraphicsInfo graphicsInfo)
    {
        var tableCandidate = new List<DocumentElement>();
        var i = startIndex;

        while (i < elements.Count)
        {
            var current = elements[i];
            
            // 空の要素はスキップ
            if (current.Type == ElementType.Empty)
            {
                i++;
                continue;
            }
            
            // テーブル行の可能性をチェック（図形情報も考慮）
            if (IsLikelyTableRow(current, graphicsInfo))
            {
                tableCandidate.Add(current);
            }
            else
            {
                // 表形式でない行に遭遇したら終了
                break;
            }
            
            i++;
        }

        // 構造の一貫性をチェック
        if (tableCandidate.Count >= 2 && HasConsistentTableStructure(tableCandidate))
        {
            return tableCandidate;
        }

        return new List<DocumentElement>();
    }

    private static bool IsLikelyTableRow(DocumentElement element, GraphicsInfo graphicsInfo)
    {
        var text = element.Content.Trim();
        
        // 明らかにヘッダーではない短い行
        if (text.Length < 100 && !text.EndsWith("。") && !text.EndsWith("."))
        {
            // 複数の単語/要素に分かれている
            var parts = text.Split(new[] { ' ', '\t', '　' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 2 && parts.Length <= 10) // 2-10列が一般的
            {
                // 図形情報による補強判定
                if (HasTableBoundaries(element, graphicsInfo))
                {
                    return true;
                }

                // 単語間に適切な間隔がある（Words情報を使用）
                if (element.Words != null && element.Words.Count >= 2)
                {
                    var avgGap = CalculateAverageWordGap(element.Words);
                    return avgGap > 10; // 10pt以上の間隔
                }
                return true;
            }
        }

        return false;
    }

    private static bool HasTableBoundaries(DocumentElement element, GraphicsInfo graphicsInfo)
    {
        if (element.Words == null || element.Words.Count == 0) return false;

        var elementBounds = new UglyToad.PdfPig.Core.PdfRectangle(
            element.Words.Min(w => w.BoundingBox.Left),
            element.Words.Min(w => w.BoundingBox.Bottom),
            element.Words.Max(w => w.BoundingBox.Right),
            element.Words.Max(w => w.BoundingBox.Top)
        );

        // 要素の周囲に水平線があるかチェック
        var hasHorizontalBoundary = graphicsInfo.HorizontalLines.Any(line =>
            Math.Abs(line.From.Y - elementBounds.Bottom) < 5 ||
            Math.Abs(line.From.Y - elementBounds.Top) < 5);

        // 要素の周囲に垂直線があるかチェック
        var hasVerticalBoundary = graphicsInfo.VerticalLines.Any(line =>
            line.From.X >= elementBounds.Left - 5 && line.From.X <= elementBounds.Right + 5 &&
            line.From.Y <= elementBounds.Top + 5 && line.To.Y >= elementBounds.Bottom - 5);

        return hasHorizontalBoundary || hasVerticalBoundary;
    }

    private static List<TableRegion> IdentifyTableRegions(List<DocumentElement> elements, GraphicsInfo graphicsInfo)
    {
        var tableRegions = new List<TableRegion>();
        
        if (graphicsInfo.HorizontalLines.Count < 2 || graphicsInfo.VerticalLines.Count < 2)
        {
            return tableRegions; // 十分な線がない場合は空を返す
        }
        
        // 水平線と垂直線から矩形グリッドを検出
        var gridCells = DetectGridCells(graphicsInfo);
        
        foreach (var cell in gridCells)
        {
            // セル内に含まれる要素を検索
            var cellElements = elements.Where(element => IsElementInCell(element, cell)).ToList();
            
            if (cellElements.Count > 0)
            {
                var region = new TableRegion
                {
                    Bounds = cell,
                    Elements = cellElements
                };
                tableRegions.Add(region);
            }
        }
        
        return MergeAdjacentTableRegions(tableRegions);
    }

    private static List<UglyToad.PdfPig.Core.PdfRectangle> DetectGridCells(GraphicsInfo graphicsInfo)
    {
        var cells = new List<UglyToad.PdfPig.Core.PdfRectangle>();
        
        // 水平線と垂直線をソート
        var hLines = graphicsInfo.HorizontalLines.OrderBy(l => l.From.Y).ToList();
        var vLines = graphicsInfo.VerticalLines.OrderBy(l => l.From.X).ToList();
        
        // 隣接する線の組み合わせでセルを作成
        for (int i = 0; i < hLines.Count - 1; i++)
        {
            for (int j = 0; j < vLines.Count - 1; j++)
            {
                var topLine = hLines[i + 1];
                var bottomLine = hLines[i];
                var leftLine = vLines[j];
                var rightLine = vLines[j + 1];
                
                // 線が交差してセルを形成するかチェック
                if (LinesFormCell(topLine, bottomLine, leftLine, rightLine))
                {
                    var cell = new UglyToad.PdfPig.Core.PdfRectangle(
                        leftLine.From.X,
                        bottomLine.From.Y,
                        rightLine.From.X,
                        topLine.From.Y
                    );
                    cells.Add(cell);
                }
            }
        }
        
        return cells;
    }

    private static bool LinesFormCell(LineSegment top, LineSegment bottom, LineSegment left, LineSegment right)
    {
        var tolerance = 5.0;
        
        // 線が適切な位置にあるかチェック
        return Math.Abs(top.From.Y - top.To.Y) < tolerance && // 水平線
               Math.Abs(bottom.From.Y - bottom.To.Y) < tolerance && // 水平線
               Math.Abs(left.From.X - left.To.X) < tolerance && // 垂直線
               Math.Abs(right.From.X - right.To.X) < tolerance && // 垂直線
               top.From.Y > bottom.From.Y && // 上下関係
               right.From.X > left.From.X; // 左右関係
    }

    private static bool IsElementInCell(DocumentElement element, UglyToad.PdfPig.Core.PdfRectangle cell)
    {
        if (element.Words == null || element.Words.Count == 0) return false;
        
        var elementBounds = new UglyToad.PdfPig.Core.PdfRectangle(
            element.Words.Min(w => w.BoundingBox.Left),
            element.Words.Min(w => w.BoundingBox.Bottom),
            element.Words.Max(w => w.BoundingBox.Right),
            element.Words.Max(w => w.BoundingBox.Top)
        );
        
        // 要素がセル内に含まれるかチェック（マージンを考慮）
        var margin = 2.0;
        return elementBounds.Left >= cell.Left - margin &&
               elementBounds.Right <= cell.Right + margin &&
               elementBounds.Bottom >= cell.Bottom - margin &&
               elementBounds.Top <= cell.Top + margin;
    }

    private static List<TableRegion> MergeAdjacentTableRegions(List<TableRegion> regions)
    {
        // 隣接するテーブル領域をマージして、より大きなテーブルを形成
        // 簡略化のため、現在は元のリストをそのまま返す
        return regions;
    }

    private static List<DocumentElement> ProcessTableRegion(TableRegion region, GraphicsInfo graphicsInfo)
    {
        var result = new List<DocumentElement>();
        
        // テーブル領域内の要素をセル単位でグループ化
        var cellGroups = GroupElementsByCells(region, graphicsInfo);
        
        // 各行を処理
        foreach (var rowGroup in cellGroups.GroupBy(g => Math.Round(g.Key.Bottom, 0)).OrderByDescending(g => g.Key))
        {
            var cellContents = rowGroup.OrderBy(g => g.Key.Left)
                                      .Select(g => string.Join(" ", g.Value.Select(e => e.Content)))
                                      .ToList();
            
            if (cellContents.Any(c => !string.IsNullOrWhiteSpace(c)))
            {
                var tableRow = new DocumentElement
                {
                    Type = ElementType.TableRow,
                    Content = string.Join(" | ", cellContents),
                    Words = rowGroup.SelectMany(g => g.Value.SelectMany(e => e.Words ?? [])).ToList()
                };
                result.Add(tableRow);
            }
        }
        
        return result;
    }

    private static Dictionary<UglyToad.PdfPig.Core.PdfRectangle, List<DocumentElement>> GroupElementsByCells(TableRegion region, GraphicsInfo graphicsInfo)
    {
        var cellGroups = new Dictionary<UglyToad.PdfPig.Core.PdfRectangle, List<DocumentElement>>();
        
        // 各要素を適切なセルに割り当て
        foreach (var element in region.Elements)
        {
            var cell = FindContainingCell(element, graphicsInfo);
            if (cell.HasValue)
            {
                if (!cellGroups.ContainsKey(cell.Value))
                {
                    cellGroups[cell.Value] = new List<DocumentElement>();
                }
                cellGroups[cell.Value].Add(element);
            }
        }
        
        return cellGroups;
    }

    private static UglyToad.PdfPig.Core.PdfRectangle? FindContainingCell(DocumentElement element, GraphicsInfo graphicsInfo)
    {
        if (element.Words == null || element.Words.Count == 0) return null;
        
        var elementCenter = new UglyToad.PdfPig.Core.PdfPoint(
            element.Words.Average(w => w.BoundingBox.Left + w.BoundingBox.Width / 2),
            element.Words.Average(w => w.BoundingBox.Bottom + w.BoundingBox.Height / 2)
        );
        
        // 最も近いセルを検索
        var gridCells = DetectGridCells(graphicsInfo);
        return gridCells.FirstOrDefault(cell => 
            elementCenter.X >= cell.Left && elementCenter.X <= cell.Right &&
            elementCenter.Y >= cell.Bottom && elementCenter.Y <= cell.Top);
    }
}

public class TableRegion
{
    public UglyToad.PdfPig.Core.PdfRectangle Bounds { get; set; }
    public List<DocumentElement> Elements { get; set; } = new List<DocumentElement>();
    
    public bool Contains(DocumentElement element)
    {
        return Elements.Contains(element);
    }
}

public class GraphicsInfo
{
    public List<LineSegment> HorizontalLines { get; set; } = new List<LineSegment>();
    public List<LineSegment> VerticalLines { get; set; } = new List<LineSegment>();
    public List<UglyToad.PdfPig.Core.PdfRectangle> Rectangles { get; set; } = new List<UglyToad.PdfPig.Core.PdfRectangle>();
}

public class LineSegment
{
    public UglyToad.PdfPig.Core.PdfPoint From { get; set; }
    public UglyToad.PdfPig.Core.PdfPoint To { get; set; }
}

public class FontFormatting
{
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
}

public class FontAnalysis
{
    public double BaseFontSize { get; set; }
    public double LargeFontThreshold { get; set; }
    public double SmallFontThreshold { get; set; }
    public string DominantFont { get; set; } = "";
    public List<double> AllFontSizes { get; set; } = [];
}

public class DocumentStructure
{
    public List<DocumentElement> Elements { get; set; } = [];
}

public class DocumentElement
{
    public ElementType Type { get; set; }
    public string Content { get; set; } = "";
    public double FontSize { get; set; }
    public double LeftMargin { get; set; }
    public bool IsIndented { get; set; }
    public List<Word> Words { get; set; } = [];
}

public enum ElementType
{
    Empty,
    Header,
    Paragraph,
    ListItem,
    TableRow
}