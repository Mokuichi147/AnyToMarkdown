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
        var fontAnalysis = FontAnalyzer.AnalyzeFontDistribution(words);
        
        // 図形情報を抽出してテーブルの境界を検出
        var graphicsInfo = GraphicsProcessor.ExtractGraphicsInfo(page);
        
        var elements = new List<DocumentElement>();
        for (int i = 0; i < lines.Count; i++)
        {
            var element = LineAnalyzer.AnalyzeLine(lines[i], fontAnalysis, horizontalTolerance);
            elements.Add(element);
        }

        // 事前処理：<br>タグで分割されたテーブル内容の統合
        elements = PostProcessor.ConsolidateBrokenTableCells(elements);
        
        // 後処理：コンテキスト情報を活用した要素分類の改善（テーブル検出前）
        elements = PostProcessor.PostProcessElementClassification(elements, fontAnalysis);
        
        // ヘッダーの座標ベース検出とレベル修正
        elements = PostProcessor.PostProcessHeaderDetectionWithCoordinates(elements, fontAnalysis);
        
        // 後処理：コードブロックと引用ブロックの検出
        elements = PostProcessor.PostProcessCodeAndQuoteBlocks(elements);
        
        // 後処理：テーブルヘッダーの統合処理
        elements = PostProcessor.PostProcessTableHeaderIntegration(elements);
        
        // 後処理：図形情報と連続する行の構造パターンを分析してテーブルを検出
        elements = PostProcessor.PostProcessTableDetection(elements, graphicsInfo);
        
        // 座標ベーステーブル列アライメント（CLAUDE.md準拠）- 一時的に無効化
        // elements = PostProcessor.PostProcessTableAlignment(elements);
        
        documentStructure.Elements.AddRange(elements);
        documentStructure.FontAnalysis = fontAnalysis;
        return documentStructure;
    }
    
    
    private static string ExtractCleanTextForAnalysis(string text)
    {
        // Markdownフォーマットを除去してテキスト分析を行う
        var cleanText = text;

        // 太字フォーマットを除去（複数回実行して入れ子を処理）
        while (cleanText.Contains("**") || cleanText.Contains("*"))
        {
            var before = cleanText;
            cleanText = System.Text.RegularExpressions.Regex.Replace(cleanText, @"\*{1,3}([^*]*)\*{1,3}", "$1");
            if (before == cleanText) break; // 変化がなければ終了
        }

        // 斜体フォーマットを除去  
        cleanText = System.Text.RegularExpressions.Regex.Replace(cleanText, @"_([^_]+)_", "$1");

        // 余分なスペースを統合
        cleanText = System.Text.RegularExpressions.Regex.Replace(cleanText, @"\s+", " ");

        return cleanText.Trim();
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
            var gap = words[i].BoundingBox.Left - words[i-1].BoundingBox.Right;
            if (gap > 0) // 負のギャップは無視（重複文字の場合）
                gaps.Add(gap);
        }

        return gaps.Count > 0 ? gaps.Average() : 0;
    }

    private static (List<LineSegment> horizontalLines, List<LineSegment> verticalLines, List<UglyToad.PdfPig.Core.PdfRectangle> rectangles) ExtractLinesFromPaths(IEnumerable<object> paths)
    {
        var horizontalLines = new List<LineSegment>();
        var verticalLines = new List<LineSegment>();
        var rectangles = new List<UglyToad.PdfPig.Core.PdfRectangle>();
        
        try
        {
            foreach (var path in paths)
            {
                // パス情報から線とエクスタンスを抽出
                var pathInfo = AnalyzePath(path);
                
                // 水平線の検出（Y座標が同じで、X座標が異なる）
                var horizontalPathLines = pathInfo.lines.Where(line => 
                    Math.Abs(line.From.Y - line.To.Y) < 2.0 && 
                    Math.Abs(line.From.X - line.To.X) > 10.0);
                
                // 垂直線の検出（X座標が同じで、Y座標が異なる）
                var verticalPathLines = pathInfo.lines.Where(line => 
                    Math.Abs(line.From.X - line.To.X) < 2.0 && 
                    Math.Abs(line.From.Y - line.To.Y) > 10.0);
                
                horizontalLines.AddRange(horizontalPathLines);
                verticalLines.AddRange(verticalPathLines);
                rectangles.AddRange(pathInfo.rectangles);
            }
        }
        catch
        {
            // パス解析に失敗した場合は空のリストを返す
        }
        
        return (horizontalLines, verticalLines, rectangles);
    }
    
    private static (List<LineSegment> lines, List<UglyToad.PdfPig.Core.PdfRectangle> rectangles) AnalyzePath(object path)
    {
        var lines = new List<LineSegment>();
        var rectangles = new List<UglyToad.PdfPig.Core.PdfRectangle>();
        
        try
        {
            // PdfPigのパス情報からライン要素を抽出
            // 実際の実装はPdfPigのAPIに依存するため、
            // ここでは汎用的な座標ベースの線検出を行う
            
            // パスオブジェクトの型とプロパティを動的に解析
            var pathType = path.GetType();
            
            // パスの座標情報を取得（リフレクションを使用）
            var commands = GetPathCommands(path);
            if (commands != null)
            {
                var currentPoint = new UglyToad.PdfPig.Core.PdfPoint(0, 0);
                
                for (int i = 0; i < commands.Count - 1; i++)
                {
                    var cmd1 = commands[i];
                    var cmd2 = commands[i + 1];
                    
                    // 線分の検出
                    if (IsLineCommand(cmd1, cmd2))
                    {
                        var lineSegment = CreateLineSegment(cmd1, cmd2);
                        if (lineSegment != null)
                        {
                            lines.Add(lineSegment);
                        }
                    }
                    
                    // 矩形の検出
                    if (IsRectanglePattern(commands, i))
                    {
                        var rectangle = CreateRectangle(commands, i);
                        if (rectangle.HasValue)
                        {
                            rectangles.Add(rectangle.Value);
                        }
                    }
                }
            }
        }
        catch
        {
            // パス解析失敗時は空のリストを返す
        }
        
        return (lines, rectangles);
    }
    
    private static List<object>? GetPathCommands(object path)
    {
        try
        {
            // パスオブジェクトからコマンドリストを取得
            // 実際の実装はPdfPigのAPIに応じて調整
            var commandsProperty = path.GetType().GetProperty("Commands");
            return commandsProperty?.GetValue(path) as List<object>;
        }
        catch
        {
            return null;
        }
    }
    
    private static bool IsLineCommand(object cmd1, object cmd2)
    {
        // 線描画コマンドかどうかを判定
        try
        {
            var type1 = cmd1.GetType().Name;
            var type2 = cmd2.GetType().Name;
            
            // MoveTo → LineTo パターン
            return (type1.Contains("Move") && type2.Contains("Line")) ||
                   (type1.Contains("Line") && type2.Contains("Line"));
        }
        catch
        {
            return false;
        }
    }
    
    private static LineSegment? CreateLineSegment(object cmd1, object cmd2)
    {
        try
        {
            var point1 = GetCommandPoint(cmd1);
            var point2 = GetCommandPoint(cmd2);
            
            if (point1.HasValue && point2.HasValue)
            {
                return new LineSegment
                {
                    From = point1.Value,
                    To = point2.Value,
                    Thickness = 1.0
                };
            }
        }
        catch
        {
            // 線セグメント作成失敗
        }
        
        return null;
    }
    
    private static UglyToad.PdfPig.Core.PdfPoint? GetCommandPoint(object command)
    {
        try
        {
            var xProperty = command.GetType().GetProperty("X");
            var yProperty = command.GetType().GetProperty("Y");
            
            if (xProperty != null && yProperty != null)
            {
                var x = Convert.ToDouble(xProperty.GetValue(command));
                var y = Convert.ToDouble(yProperty.GetValue(command));
                return new UglyToad.PdfPig.Core.PdfPoint(x, y);
            }
        }
        catch
        {
            // ポイント取得失敗
        }
        
        return null;
    }
    
    private static bool IsRectanglePattern(List<object> commands, int startIndex)
    {
        // 矩形パターンの検出（4つの線で構成される閉じた図形）
        if (startIndex + 4 >= commands.Count) return false;
        
        try
        {
            var points = new List<UglyToad.PdfPig.Core.PdfPoint>();
            for (int i = 0; i < 5 && startIndex + i < commands.Count; i++)
            {
                var point = GetCommandPoint(commands[startIndex + i]);
                if (point.HasValue)
                {
                    points.Add(point.Value);
                }
            }
            
            // 4つの角が矩形を形成するかチェック
            return points.Count >= 4 && IsRectangleShape(points);
        }
        catch
        {
            return false;
        }
    }
    
    private static UglyToad.PdfPig.Core.PdfRectangle? CreateRectangle(List<object> commands, int startIndex)
    {
        try
        {
            var points = new List<UglyToad.PdfPig.Core.PdfPoint>();
            for (int i = 0; i < 4 && startIndex + i < commands.Count; i++)
            {
                var point = GetCommandPoint(commands[startIndex + i]);
                if (point.HasValue)
                {
                    points.Add(point.Value);
                }
            }
            
            if (points.Count >= 4)
            {
                var minX = points.Min(p => p.X);
                var maxX = points.Max(p => p.X);
                var minY = points.Min(p => p.Y);
                var maxY = points.Max(p => p.Y);
                
                return new UglyToad.PdfPig.Core.PdfRectangle(minX, minY, maxX, maxY);
            }
        }
        catch
        {
            // 矩形作成失敗
        }
        
        return null;
    }
    
    private static bool IsRectangleShape(List<UglyToad.PdfPig.Core.PdfPoint> points)
    {
        if (points.Count < 4) return false;
        
        try
        {
            // 矩形の特徴：対角の点が等しい距離にある
            var distinctX = points.Select(p => Math.Round(p.X, 1)).Distinct().Count();
            var distinctY = points.Select(p => Math.Round(p.Y, 1)).Distinct().Count();
            
            // 2つの異なるX座標と2つの異なるY座標を持つ
            return distinctX == 2 && distinctY == 2;
        }
        catch
        {
            return false;
        }
    }
    
    private static List<TablePattern> AnalyzeTablePatterns(List<LineSegment> horizontalLines, List<LineSegment> verticalLines, List<UglyToad.PdfPig.Core.PdfRectangle> rectangles)
    {
        var patterns = new List<TablePattern>();
        
        try
        {
            // 線の密度が高い領域を特定
            var tableAreas = IdentifyTableAreas(horizontalLines, verticalLines);
            
            foreach (var area in tableAreas)
            {
                var pattern = AnalyzeSingleTablePattern(area, horizontalLines, verticalLines, rectangles);
                if (pattern != null)
                {
                    patterns.Add(pattern);
                }
            }
            
            // 矩形からも表パターンを検出
            foreach (var rectangle in rectangles)
            {
                var rectPattern = AnalyzeRectangleTablePattern(rectangle, horizontalLines, verticalLines);
                if (rectPattern != null)
                {
                    patterns.Add(rectPattern);
                }
            }
        }
        catch
        {
            // パターン分析失敗時は空のリストを返す
        }
        
        return patterns;
    }
    
    private static List<UglyToad.PdfPig.Core.PdfRectangle> IdentifyTableAreas(List<LineSegment> horizontalLines, List<LineSegment> verticalLines)
    {
        var areas = new List<UglyToad.PdfPig.Core.PdfRectangle>();
        
        try
        {
            // 線の交点と密度を分析して表領域を特定
            var intersections = FindLineIntersections(horizontalLines, verticalLines);
            
            if (intersections.Count >= 4) // 最低4つの交点で矩形を形成
            {
                var clusteredAreas = ClusterIntersectionsIntoAreas(intersections);
                areas.AddRange(clusteredAreas);
            }
        }
        catch
        {
            // 領域特定失敗
        }
        
        return areas;
    }
    
    private static List<UglyToad.PdfPig.Core.PdfPoint> FindLineIntersections(List<LineSegment> horizontalLines, List<LineSegment> verticalLines)
    {
        var intersections = new List<UglyToad.PdfPig.Core.PdfPoint>();
        
        foreach (var hLine in horizontalLines)
        {
            foreach (var vLine in verticalLines)
            {
                var intersection = CalculateIntersection(hLine, vLine);
                if (intersection.HasValue)
                {
                    intersections.Add(intersection.Value);
                }
            }
        }
        
        return intersections;
    }
    
    private static UglyToad.PdfPig.Core.PdfPoint? CalculateIntersection(LineSegment horizontal, LineSegment vertical)
    {
        try
        {
            // 水平線と垂直線の交点を計算
            var hY = (horizontal.From.Y + horizontal.To.Y) / 2;
            var vX = (vertical.From.X + vertical.To.X) / 2;
            
            // 交点が両方の線分の範囲内にあるかチェック
            var hMinX = Math.Min(horizontal.From.X, horizontal.To.X);
            var hMaxX = Math.Max(horizontal.From.X, horizontal.To.X);
            var vMinY = Math.Min(vertical.From.Y, vertical.To.Y);
            var vMaxY = Math.Max(vertical.From.Y, vertical.To.Y);
            
            if (vX >= hMinX && vX <= hMaxX && hY >= vMinY && hY <= vMaxY)
            {
                return new UglyToad.PdfPig.Core.PdfPoint(vX, hY);
            }
        }
        catch
        {
            // 交点計算失敗
        }
        
        return null;
    }
    
    private static List<UglyToad.PdfPig.Core.PdfRectangle> ClusterIntersectionsIntoAreas(List<UglyToad.PdfPig.Core.PdfPoint> intersections)
    {
        var areas = new List<UglyToad.PdfPig.Core.PdfRectangle>();
        
        try
        {
            if (intersections.Count < 4) return areas;
            
            // 交点を領域にクラスター化
            var minX = intersections.Min(p => p.X);
            var maxX = intersections.Max(p => p.X);
            var minY = intersections.Min(p => p.Y);
            var maxY = intersections.Max(p => p.Y);
            
            // 表領域として妥当なサイズかチェック
            if (maxX - minX > 50 && maxY - minY > 20)
            {
                areas.Add(new UglyToad.PdfPig.Core.PdfRectangle(minX, minY, maxX, maxY));
            }
        }
        catch
        {
            // クラスタリング失敗
        }
        
        return areas;
    }
    
    private static TablePattern? AnalyzeSingleTablePattern(UglyToad.PdfPig.Core.PdfRectangle area, List<LineSegment> horizontalLines, List<LineSegment> verticalLines, List<UglyToad.PdfPig.Core.PdfRectangle> rectangles)
    {
        try
        {
            var areaHLines = horizontalLines.Where(line => IsLineInArea(line, area)).ToList();
            var areaVLines = verticalLines.Where(line => IsLineInArea(line, area)).ToList();
            
            var borderType = DetermineBorderType(areaHLines, areaVLines, area);
            var confidence = CalculateConfidence(areaHLines, areaVLines, area);
            
            if (confidence > 0.3) // 信頼度閾値
            {
                return new TablePattern
                {
                    BorderType = borderType,
                    BoundingArea = area,
                    BorderLines = GetBorderLines(areaHLines, areaVLines, area),
                    InternalLines = GetInternalLines(areaHLines, areaVLines, area),
                    Confidence = confidence,
                    EstimatedColumns = EstimateColumns(areaVLines, area),
                    EstimatedRows = EstimateRows(areaHLines, area)
                };
            }
        }
        catch
        {
            // パターン分析失敗
        }
        
        return null;
    }
    
    private static bool IsLineInArea(LineSegment line, UglyToad.PdfPig.Core.PdfRectangle area)
    {
        var tolerance = 5.0;
        return line.From.X >= area.Left - tolerance && line.From.X <= area.Right + tolerance &&
               line.From.Y >= area.Bottom - tolerance && line.From.Y <= area.Top + tolerance &&
               line.To.X >= area.Left - tolerance && line.To.X <= area.Right + tolerance &&
               line.To.Y >= area.Bottom - tolerance && line.To.Y <= area.Top + tolerance;
    }
    
    private static TableBorderType DetermineBorderType(List<LineSegment> horizontalLines, List<LineSegment> verticalLines, UglyToad.PdfPig.Core.PdfRectangle area)
    {
        try
        {
            var hasTopBorder = HasBorderLine(horizontalLines, area.Top, area.Left, area.Right, true);
            var hasBottomBorder = HasBorderLine(horizontalLines, area.Bottom, area.Left, area.Right, true);
            var hasLeftBorder = HasBorderLine(verticalLines, area.Left, area.Bottom, area.Top, false);
            var hasRightBorder = HasBorderLine(verticalLines, area.Right, area.Bottom, area.Top, false);
            
            var borderCount = (hasTopBorder ? 1 : 0) + (hasBottomBorder ? 1 : 0) + 
                             (hasLeftBorder ? 1 : 0) + (hasRightBorder ? 1 : 0);
            
            if (borderCount >= 4) return TableBorderType.FullBorder;
            if (hasTopBorder && hasBottomBorder && !hasLeftBorder && !hasRightBorder) return TableBorderType.TopBottomOnly;
            if (hasTopBorder && horizontalLines.Count == 1) return TableBorderType.HeaderSeparator;
            if (horizontalLines.Count > 2 && verticalLines.Count > 2) return TableBorderType.GridLines;
            if (borderCount > 0) return TableBorderType.PartialBorder;
        }
        catch
        {
            // 境界タイプ判定失敗
        }
        
        return TableBorderType.None;
    }
    
    private static bool HasBorderLine(List<LineSegment> lines, double position, double start, double end, bool isHorizontal)
    {
        var tolerance = 3.0;
        
        return lines.Any(line =>
        {
            if (isHorizontal)
            {
                return Math.Abs((line.From.Y + line.To.Y) / 2 - position) < tolerance &&
                       Math.Min(line.From.X, line.To.X) <= start + tolerance &&
                       Math.Max(line.From.X, line.To.X) >= end - tolerance;
            }
            else
            {
                return Math.Abs((line.From.X + line.To.X) / 2 - position) < tolerance &&
                       Math.Min(line.From.Y, line.To.Y) <= start + tolerance &&
                       Math.Max(line.From.Y, line.To.Y) >= end - tolerance;
            }
        });
    }
    
    private static double CalculateConfidence(List<LineSegment> horizontalLines, List<LineSegment> verticalLines, UglyToad.PdfPig.Core.PdfRectangle area)
    {
        try
        {
            var lineCount = horizontalLines.Count + verticalLines.Count;
            var areaSize = (area.Width * area.Height) / 10000; // 正規化
            
            // 線の数と領域サイズに基づく信頼度
            var lineDensity = lineCount / Math.Max(areaSize, 1);
            var confidence = Math.Min(lineDensity / 2.0, 1.0);
            
            // 最低限の線数要件
            if (lineCount < 2) confidence *= 0.5;
            
            return confidence;
        }
        catch
        {
            return 0.0;
        }
    }
    
    private static List<LineSegment> GetBorderLines(List<LineSegment> horizontalLines, List<LineSegment> verticalLines, UglyToad.PdfPig.Core.PdfRectangle area)
    {
        var borderLines = new List<LineSegment>();
        var tolerance = 5.0;
        
        // 境界に近い線を境界線として分類
        borderLines.AddRange(horizontalLines.Where(line =>
            Math.Abs((line.From.Y + line.To.Y) / 2 - area.Top) < tolerance ||
            Math.Abs((line.From.Y + line.To.Y) / 2 - area.Bottom) < tolerance));
            
        borderLines.AddRange(verticalLines.Where(line =>
            Math.Abs((line.From.X + line.To.X) / 2 - area.Left) < tolerance ||
            Math.Abs((line.From.X + line.To.X) / 2 - area.Right) < tolerance));
        
        return borderLines;
    }
    
    private static List<LineSegment> GetInternalLines(List<LineSegment> horizontalLines, List<LineSegment> verticalLines, UglyToad.PdfPig.Core.PdfRectangle area)
    {
        var internalLines = new List<LineSegment>();
        var tolerance = 5.0;
        
        // 境界から離れた線を内部線として分類
        internalLines.AddRange(horizontalLines.Where(line =>
            Math.Abs((line.From.Y + line.To.Y) / 2 - area.Top) >= tolerance &&
            Math.Abs((line.From.Y + line.To.Y) / 2 - area.Bottom) >= tolerance));
            
        internalLines.AddRange(verticalLines.Where(line =>
            Math.Abs((line.From.X + line.To.X) / 2 - area.Left) >= tolerance &&
            Math.Abs((line.From.X + line.To.X) / 2 - area.Right) >= tolerance));
        
        return internalLines;
    }
    
    private static int EstimateColumns(List<LineSegment> verticalLines, UglyToad.PdfPig.Core.PdfRectangle area)
    {
        try
        {
            // 垂直線の位置から列数を推定
            var columnPositions = verticalLines
                .Select(line => (line.From.X + line.To.X) / 2)
                .Distinct()
                .OrderBy(x => x)
                .ToList();
            
            return Math.Max(columnPositions.Count + 1, 1);
        }
        catch
        {
            return 1;
        }
    }
    
    private static int EstimateRows(List<LineSegment> horizontalLines, UglyToad.PdfPig.Core.PdfRectangle area)
    {
        try
        {
            // 水平線の位置から行数を推定
            var rowPositions = horizontalLines
                .Select(line => (line.From.Y + line.To.Y) / 2)
                .Distinct()
                .OrderBy(y => y)
                .ToList();
            
            return Math.Max(rowPositions.Count + 1, 1);
        }
        catch
        {
            return 1;
        }
    }
    
    private static TablePattern? AnalyzeRectangleTablePattern(UglyToad.PdfPig.Core.PdfRectangle rectangle, List<LineSegment> horizontalLines, List<LineSegment> verticalLines)
    {
        try
        {
            // 矩形がテーブルの境界を表している可能性を分析
            var confidence = rectangle.Width > 50 && rectangle.Height > 20 ? 0.6 : 0.3;
            
            return new TablePattern
            {
                BorderType = TableBorderType.FullBorder,
                BoundingArea = rectangle,
                BorderLines = CreateRectangleBorderLines(rectangle),
                InternalLines = new List<LineSegment>(),
                Confidence = confidence,
                EstimatedColumns = 2, // デフォルト値
                EstimatedRows = 2     // デフォルト値
            };
        }
        catch
        {
            return null;
        }
    }
    
    private static List<LineSegment> CreateRectangleBorderLines(UglyToad.PdfPig.Core.PdfRectangle rectangle)
    {
        var lines = new List<LineSegment>();
        
        try
        {
            // 矩形の4辺を線セグメントとして作成
            var topLeft = new UglyToad.PdfPig.Core.PdfPoint(rectangle.Left, rectangle.Top);
            var topRight = new UglyToad.PdfPig.Core.PdfPoint(rectangle.Right, rectangle.Top);
            var bottomLeft = new UglyToad.PdfPig.Core.PdfPoint(rectangle.Left, rectangle.Bottom);
            var bottomRight = new UglyToad.PdfPig.Core.PdfPoint(rectangle.Right, rectangle.Bottom);
            
            lines.Add(new LineSegment { From = topLeft, To = topRight, Type = LineType.TableBorder });
            lines.Add(new LineSegment { From = topRight, To = bottomRight, Type = LineType.TableBorder });
            lines.Add(new LineSegment { From = bottomRight, To = bottomLeft, Type = LineType.TableBorder });
            lines.Add(new LineSegment { From = bottomLeft, To = topLeft, Type = LineType.TableBorder });
        }
        catch
        {
            // 境界線作成失敗
        }
        
        return lines;
    }
    
    private static bool IsHorizontalLine(LineSegment line)
    {
        var tolerance = 2.0; // Y座標の許容差
        return Math.Abs(line.From.Y - line.To.Y) <= tolerance && Math.Abs(line.From.X - line.To.X) > tolerance;
    }
    
    private static bool IsVerticalLine(LineSegment line)
    {
        var tolerance = 2.0; // X座標の許容差
        return Math.Abs(line.From.X - line.To.X) <= tolerance && Math.Abs(line.From.Y - line.To.Y) > tolerance;
    }
    

    private static bool IsElementInTableArea(DocumentElement element, UglyToad.PdfPig.Core.PdfRectangle area)
    {
        try
        {
            if (element.Words?.Count > 0)
            {
                // 要素の単語がテーブル領域内にあるかチェック
                var elementBounds = GetElementBounds(element);
                var tolerance = 10.0;

                return elementBounds.Left >= area.Left - tolerance &&
                       elementBounds.Right <= area.Right + tolerance &&
                       elementBounds.Bottom >= area.Bottom - tolerance &&
                       elementBounds.Top <= area.Top + tolerance;
            }
        }
        catch
        {
            // 境界チェック失敗
        }

        return false;
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
    
    
    private static bool LineIntersects(LineSegment line1, LineSegment line2, double tolerance)
    {
        // 線分の交点計算（簡易版）
        
        // 水平線と垂直線の場合
        if (IsHorizontalLine(line1) && IsVerticalLine(line2))
        {
            var hLine = line1;
            var vLine = line2;
            
            var hLeft = Math.Min(hLine.From.X, hLine.To.X);
            var hRight = Math.Max(hLine.From.X, hLine.To.X);
            var hY = hLine.From.Y;
            
            var vBottom = Math.Min(vLine.From.Y, vLine.To.Y);
            var vTop = Math.Max(vLine.From.Y, vLine.To.Y);
            var vX = vLine.From.X;
            
            return vX >= hLeft - tolerance && vX <= hRight + tolerance &&
                   hY >= vBottom - tolerance && hY <= vTop + tolerance;
        }
        
        if (IsVerticalLine(line1) && IsHorizontalLine(line2))
        {
            return LineIntersects(line2, line1, tolerance);
        }
        
        return false;
    }
    
    
    private static UglyToad.PdfPig.Core.PdfRectangle GetElementBounds(DocumentElement element)
    {
        if (element.Words == null || !element.Words.Any())
        {
            return new UglyToad.PdfPig.Core.PdfRectangle(0, 0, 0, 0);
        }

        var left = element.Words.Min(w => w.BoundingBox.Left);
        var right = element.Words.Max(w => w.BoundingBox.Right);
        var bottom = element.Words.Min(w => w.BoundingBox.Bottom);
        var top = element.Words.Max(w => w.BoundingBox.Top);

        return new UglyToad.PdfPig.Core.PdfRectangle(left, bottom, right, top);
    }
}


