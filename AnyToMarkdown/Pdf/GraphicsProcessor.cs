using System.Linq;
using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

public class GraphicsInfo
{
    public List<LineSegment> HorizontalLines { get; set; } = new List<LineSegment>();
    public List<LineSegment> VerticalLines { get; set; } = new List<LineSegment>();
    public List<UglyToad.PdfPig.Core.PdfRectangle> Rectangles { get; set; } = new List<UglyToad.PdfPig.Core.PdfRectangle>();
    public List<TablePattern> TablePatterns { get; set; } = new List<TablePattern>();
}

public class LineSegment
{
    public UglyToad.PdfPig.Core.PdfPoint From { get; set; }
    public UglyToad.PdfPig.Core.PdfPoint To { get; set; }
    public double Thickness { get; set; } = 1.0;
    public LineType Type { get; set; } = LineType.Unknown;
}

public class TablePattern
{
    public TableBorderType BorderType { get; set; }
    public UglyToad.PdfPig.Core.PdfRectangle BoundingArea { get; set; }
    public List<LineSegment> BorderLines { get; set; } = new List<LineSegment>();
    public List<LineSegment> InternalLines { get; set; } = new List<LineSegment>();
    public double Confidence { get; set; }
    public int EstimatedColumns { get; set; }
    public int EstimatedRows { get; set; }
}

public enum LineType
{
    Unknown,
    Horizontal,
    Vertical,
    Diagonal,
    TableBorder,
    TableInternal,
    HeaderSeparator,
    RowSeparator,
    ColumnSeparator
}

public enum TableBorderType
{
    None,
    Full,
    Partial,
    Rectangle,
    FullBorder,      // 全体を囲う
    TopBottomOnly,   // 上下のみ
    HeaderSeparator, // ヘッダー下のみ
    GridLines,       // グリッド線
    PartialBorder    // 部分的な境界
}

internal static class GraphicsProcessor
{
    public static GraphicsInfo ExtractGraphicsInfo(Page page)
    {
        var graphicsInfo = new GraphicsInfo();
        
        try
        {
            var horizontalLines = new List<LineSegment>();
            var verticalLines = new List<LineSegment>();
            var rectangles = new List<UglyToad.PdfPig.Core.PdfRectangle>();
            
            // PdfPigのパス情報から実際の線要素を抽出
            try
            {
                var paths = page.Paths;
                var actualLines = ExtractLinesFromPaths(paths);
                horizontalLines.AddRange(actualLines.horizontalLines);
                verticalLines.AddRange(actualLines.verticalLines);
                rectangles.AddRange(actualLines.rectangles);
            }
            catch
            {
                // パス情報の取得に失敗した場合は単語位置から推測
                var words = page.GetWords();
                var tableStructure = InferTableStructureFromWordPositions(words);
                horizontalLines.AddRange(tableStructure.horizontalLines);
                verticalLines.AddRange(tableStructure.verticalLines);
                rectangles.AddRange(tableStructure.rectangles);
            }
            
            // 線のパターンから表構造を分析
            var tablePatterns = AnalyzeTablePatterns(horizontalLines, verticalLines, rectangles);
            
            graphicsInfo.HorizontalLines = horizontalLines;
            graphicsInfo.VerticalLines = verticalLines;
            graphicsInfo.Rectangles = rectangles;
            graphicsInfo.TablePatterns = tablePatterns;
        }
        catch
        {
            // 図形情報の抽出に失敗した場合は空の情報を返す
        }

        return graphicsInfo;
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
            var cmd1Type = cmd1.GetType().Name;
            var cmd2Type = cmd2.GetType().Name;
            
            // PdfPigの線描画コマンドパターンを検出
            return (cmd1Type.Contains("MoveTo") && cmd2Type.Contains("LineTo")) ||
                   (cmd1Type.Contains("Line") && cmd2Type.Contains("Line"));
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
            // コマンドオブジェクトから座標を抽出
            var point1 = ExtractPointFromCommand(cmd1);
            var point2 = ExtractPointFromCommand(cmd2);
            
            if (point1.HasValue && point2.HasValue)
            {
                return new LineSegment
                {
                    From = point1.Value,
                    To = point2.Value,
                    Type = DetermineLineType(point1.Value, point2.Value)
                };
            }
        }
        catch
        {
            // 座標抽出失敗
        }
        
        return null;
    }
    
    private static UglyToad.PdfPig.Core.PdfPoint? ExtractPointFromCommand(object command)
    {
        try
        {
            // コマンドオブジェクトから座標情報を抽出（リフレクション使用）
            var xProperty = command.GetType().GetProperty("X");
            var yProperty = command.GetType().GetProperty("Y");
            
            if (xProperty != null && yProperty != null)
            {
                var x = Convert.ToDouble(xProperty.GetValue(command));
                var y = Convert.ToDouble(yProperty.GetValue(command));
                return new UglyToad.PdfPig.Core.PdfPoint(x, y);
            }
            
            // 代替アプローチ：Point プロパティを探す
            var pointProperty = command.GetType().GetProperty("Point");
            if (pointProperty != null)
            {
                return pointProperty.GetValue(command) as UglyToad.PdfPig.Core.PdfPoint?;
            }
        }
        catch
        {
            // 座標抽出失敗
        }
        
        return null;
    }
    
    private static LineType DetermineLineType(UglyToad.PdfPig.Core.PdfPoint from, UglyToad.PdfPig.Core.PdfPoint to)
    {
        var deltaX = Math.Abs(to.X - from.X);
        var deltaY = Math.Abs(to.Y - from.Y);
        
        if (deltaX < 2.0 && deltaY > 10.0)
            return LineType.Vertical;
        else if (deltaY < 2.0 && deltaX > 10.0)
            return LineType.Horizontal;
        else
            return LineType.Diagonal;
    }
    
    private static bool IsRectanglePattern(List<object> commands, int startIndex)
    {
        // 矩形パターンの検出ロジック
        // 4つの連続する線コマンドが矩形を形成するかチェック
        if (startIndex + 4 >= commands.Count) return false;
        
        try
        {
            // 簡単な矩形パターン検出
            var cmd1 = commands[startIndex];
            var cmd2 = commands[startIndex + 1];
            var cmd3 = commands[startIndex + 2];
            var cmd4 = commands[startIndex + 3];
            
            return IsLineCommand(cmd1, cmd2) && 
                   IsLineCommand(cmd2, cmd3) && 
                   IsLineCommand(cmd3, cmd4);
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
            // 連続する線コマンドから矩形を構築
            var points = new List<UglyToad.PdfPig.Core.PdfPoint>();
            
            for (int i = 0; i < 4 && startIndex + i < commands.Count; i++)
            {
                var point = ExtractPointFromCommand(commands[startIndex + i]);
                if (point.HasValue)
                {
                    points.Add(point.Value);
                }
            }
            
            if (points.Count >= 2)
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
    
    private static (List<LineSegment> horizontalLines, List<LineSegment> verticalLines, List<UglyToad.PdfPig.Core.PdfRectangle> rectangles) InferTableStructureFromWordPositions(IEnumerable<UglyToad.PdfPig.Content.Word> words)
    {
        var horizontalLines = new List<LineSegment>();
        var verticalLines = new List<LineSegment>();
        var rectangles = new List<UglyToad.PdfPig.Core.PdfRectangle>();
        
        try
        {
            var wordList = words.ToList();
            if (!wordList.Any()) return (horizontalLines, verticalLines, rectangles);
            
            // 単語の位置から表構造を推測
            var groups = GroupWordsByRows(wordList);
            
            foreach (var group in groups)
            {
                if (group.Count < 2) continue;
                
                // 行の境界線を推測
                var minY = group.Min(w => w.BoundingBox.Bottom);
                var maxY = group.Max(w => w.BoundingBox.Top);
                var leftX = group.Min(w => w.BoundingBox.Left);
                var rightX = group.Max(w => w.BoundingBox.Right);
                
                // 水平線を追加（行の上下）
                horizontalLines.Add(new LineSegment
                {
                    From = new UglyToad.PdfPig.Core.PdfPoint(leftX, minY),
                    To = new UglyToad.PdfPig.Core.PdfPoint(rightX, minY),
                    Type = LineType.Horizontal
                });
                
                horizontalLines.Add(new LineSegment
                {
                    From = new UglyToad.PdfPig.Core.PdfPoint(leftX, maxY),
                    To = new UglyToad.PdfPig.Core.PdfPoint(rightX, maxY),
                    Type = LineType.Horizontal
                });
                
                // 列の境界線を推測
                var sortedWords = group.OrderBy(w => w.BoundingBox.Left).ToList();
                for (int i = 0; i < sortedWords.Count - 1; i++)
                {
                    var gap = sortedWords[i + 1].BoundingBox.Left - sortedWords[i].BoundingBox.Right;
                    if (gap > 20) // 大きなギャップがある場合
                    {
                        var x = sortedWords[i].BoundingBox.Right + gap / 2;
                        verticalLines.Add(new LineSegment
                        {
                            From = new UglyToad.PdfPig.Core.PdfPoint(x, minY),
                            To = new UglyToad.PdfPig.Core.PdfPoint(x, maxY),
                            Type = LineType.Vertical
                        });
                    }
                }
            }
        }
        catch
        {
            // エラー時は空のリストを返す
        }
        
        return (horizontalLines, verticalLines, rectangles);
    }
    
    private static List<List<UglyToad.PdfPig.Content.Word>> GroupWordsByRows(List<UglyToad.PdfPig.Content.Word> words)
    {
        var groups = new List<List<UglyToad.PdfPig.Content.Word>>();
        var tolerance = 5.0;
        
        foreach (var word in words.OrderByDescending(w => w.BoundingBox.Bottom))
        {
            var placed = false;
            
            foreach (var group in groups)
            {
                var avgY = group.Average(w => w.BoundingBox.Bottom);
                if (Math.Abs(word.BoundingBox.Bottom - avgY) <= tolerance)
                {
                    group.Add(word);
                    placed = true;
                    break;
                }
            }
            
            if (!placed)
            {
                groups.Add(new List<UglyToad.PdfPig.Content.Word> { word });
            }
        }
        
        return groups;
    }
    
    private static List<TablePattern> AnalyzeTablePatterns(List<LineSegment> horizontalLines, List<LineSegment> verticalLines, List<UglyToad.PdfPig.Core.PdfRectangle> rectangles)
    {
        var patterns = new List<TablePattern>();
        
        try
        {
            // 線の交差点から表のセル構造を分析
            var intersections = FindLineIntersections(horizontalLines, verticalLines);
            
            foreach (var area in GetTableAreas(intersections))
            {
                var pattern = AnalyzeSingleTablePattern(area, horizontalLines, verticalLines, rectangles);
                if (pattern != null)
                {
                    patterns.Add(pattern);
                }
            }
            
            // 矩形から表パターンを検出
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
            // 表パターン解析失敗
        }
        
        return patterns;
    }
    
    private static List<UglyToad.PdfPig.Core.PdfPoint> FindLineIntersections(List<LineSegment> horizontalLines, List<LineSegment> verticalLines)
    {
        var intersections = new List<UglyToad.PdfPig.Core.PdfPoint>();
        
        foreach (var hLine in horizontalLines)
        {
            foreach (var vLine in verticalLines)
            {
                var intersection = FindIntersection(hLine, vLine);
                if (intersection.HasValue)
                {
                    intersections.Add(intersection.Value);
                }
            }
        }
        
        return intersections;
    }
    
    private static UglyToad.PdfPig.Core.PdfPoint? FindIntersection(LineSegment horizontal, LineSegment vertical)
    {
        // 水平線と垂直線の交差点を計算
        if (horizontal.Type == LineType.Horizontal && vertical.Type == LineType.Vertical)
        {
            var hY = horizontal.From.Y;
            var vX = vertical.From.X;
            
            // 交差範囲内かチェック
            var hMinX = Math.Min(horizontal.From.X, horizontal.To.X);
            var hMaxX = Math.Max(horizontal.From.X, horizontal.To.X);
            var vMinY = Math.Min(vertical.From.Y, vertical.To.Y);
            var vMaxY = Math.Max(vertical.From.Y, vertical.To.Y);
            
            if (vX >= hMinX && vX <= hMaxX && hY >= vMinY && hY <= vMaxY)
            {
                return new UglyToad.PdfPig.Core.PdfPoint(vX, hY);
            }
        }
        
        return null;
    }
    
    private static List<UglyToad.PdfPig.Core.PdfRectangle> GetTableAreas(List<UglyToad.PdfPig.Core.PdfPoint> intersections)
    {
        var areas = new List<UglyToad.PdfPig.Core.PdfRectangle>();
        
        if (intersections.Count < 4) return areas;
        
        // 交差点から矩形領域を構築
        var sortedByX = intersections.OrderBy(p => p.X).ToList();
        var sortedByY = intersections.OrderBy(p => p.Y).ToList();
        
        for (int i = 0; i < sortedByX.Count - 1; i++)
        {
            for (int j = 0; j < sortedByY.Count - 1; j++)
            {
                var point1 = sortedByX[i];
                var point2 = sortedByY[j];
                
                if (Math.Abs(point1.X - point2.X) > 50 && Math.Abs(point1.Y - point2.Y) > 20)
                {
                    var minX = Math.Min(point1.X, point2.X);
                    var maxX = Math.Max(point1.X, point2.X);
                    var minY = Math.Min(point1.Y, point2.Y);
                    var maxY = Math.Max(point1.Y, point2.Y);
                    
                    areas.Add(new UglyToad.PdfPig.Core.PdfRectangle(minX, minY, maxX, maxY));
                }
            }
        }
        
        return areas;
    }
    
    private static TablePattern? AnalyzeSingleTablePattern(UglyToad.PdfPig.Core.PdfRectangle area, List<LineSegment> horizontalLines, List<LineSegment> verticalLines, List<UglyToad.PdfPig.Core.PdfRectangle> rectangles)
    {
        try
        {
            // 指定領域内の線を抽出
            var areaHLines = horizontalLines.Where(line => IsLineInArea(line, area)).ToList();
            var areaVLines = verticalLines.Where(line => IsLineInArea(line, area)).ToList();
            
            if (areaHLines.Count < 2 || areaVLines.Count < 2) return null;
            
            return new TablePattern
            {
                BoundingArea = area,
                BorderLines = GetBorderLines(areaHLines, areaVLines, area),
                InternalLines = GetInternalLines(areaHLines, areaVLines, area),
                BorderType = DetermineBorderType(areaHLines, areaVLines),
                EstimatedColumns = areaVLines.Count + 1,
                EstimatedRows = areaHLines.Count + 1,
                Confidence = CalculateTableConfidence(areaHLines, areaVLines, area)
            };
        }
        catch
        {
            return null;
        }
    }
    
    private static TablePattern? AnalyzeRectangleTablePattern(UglyToad.PdfPig.Core.PdfRectangle rectangle, List<LineSegment> horizontalLines, List<LineSegment> verticalLines)
    {
        try
        {
            // 矩形領域から表パターンを作成
            return new TablePattern
            {
                BoundingArea = rectangle,
                BorderType = TableBorderType.Rectangle,
                EstimatedColumns = 2, // デフォルト
                EstimatedRows = 2,   // デフォルト
                Confidence = 0.7
            };
        }
        catch
        {
            return null;
        }
    }
    
    private static bool IsLineInArea(LineSegment line, UglyToad.PdfPig.Core.PdfRectangle area)
    {
        // 線が指定領域内にあるかチェック
        return line.From.X >= area.Left && line.From.X <= area.Right &&
               line.From.Y >= area.Bottom && line.From.Y <= area.Top &&
               line.To.X >= area.Left && line.To.X <= area.Right &&
               line.To.Y >= area.Bottom && line.To.Y <= area.Top;
    }
    
    private static List<LineSegment> GetBorderLines(List<LineSegment> horizontalLines, List<LineSegment> verticalLines, UglyToad.PdfPig.Core.PdfRectangle area)
    {
        var borderLines = new List<LineSegment>();
        
        // 境界線を抽出
        var tolerance = 2.0;
        
        borderLines.AddRange(horizontalLines.Where(line => 
            Math.Abs(line.From.Y - area.Top) < tolerance || Math.Abs(line.From.Y - area.Bottom) < tolerance));
        
        borderLines.AddRange(verticalLines.Where(line => 
            Math.Abs(line.From.X - area.Left) < tolerance || Math.Abs(line.From.X - area.Right) < tolerance));
        
        return borderLines;
    }
    
    private static List<LineSegment> GetInternalLines(List<LineSegment> horizontalLines, List<LineSegment> verticalLines, UglyToad.PdfPig.Core.PdfRectangle area)
    {
        var internalLines = new List<LineSegment>();
        
        // 内部線を抽出
        var tolerance = 2.0;
        
        internalLines.AddRange(horizontalLines.Where(line => 
            line.From.Y > area.Bottom + tolerance && line.From.Y < area.Top - tolerance));
        
        internalLines.AddRange(verticalLines.Where(line => 
            line.From.X > area.Left + tolerance && line.From.X < area.Right - tolerance));
        
        return internalLines;
    }
    
    private static TableBorderType DetermineBorderType(List<LineSegment> horizontalLines, List<LineSegment> verticalLines)
    {
        if (horizontalLines.Count >= 3 && verticalLines.Count >= 3)
            return TableBorderType.Full;
        else if (horizontalLines.Count >= 2 || verticalLines.Count >= 2)
            return TableBorderType.Partial;
        else
            return TableBorderType.None;
    }
    
    private static double CalculateTableConfidence(List<LineSegment> horizontalLines, List<LineSegment> verticalLines, UglyToad.PdfPig.Core.PdfRectangle area)
    {
        var score = 0.0;
        
        // 線の数によるスコア
        score += Math.Min(horizontalLines.Count * 0.2, 0.4);
        score += Math.Min(verticalLines.Count * 0.2, 0.4);
        
        // 線の規則性によるスコア
        if (AreHorizontalLinesRegular(horizontalLines))
            score += 0.1;
        
        if (AreVerticalLinesRegular(verticalLines))
            score += 0.1;
        
        return Math.Min(score, 1.0);
    }
    
    private static bool AreHorizontalLinesRegular(List<LineSegment> lines)
    {
        if (lines.Count < 2) return false;
        
        var sortedLines = lines.OrderBy(l => l.From.Y).ToList();
        var gaps = new List<double>();
        
        for (int i = 0; i < sortedLines.Count - 1; i++)
        {
            gaps.Add(sortedLines[i + 1].From.Y - sortedLines[i].From.Y);
        }
        
        if (!gaps.Any()) return false;
        
        var avgGap = gaps.Average();
        var variance = gaps.Sum(g => Math.Pow(g - avgGap, 2)) / gaps.Count;
        
        return variance < avgGap * 0.2; // 規則性の閾値
    }
    
    private static bool AreVerticalLinesRegular(List<LineSegment> lines)
    {
        if (lines.Count < 2) return false;
        
        var sortedLines = lines.OrderBy(l => l.From.X).ToList();
        var gaps = new List<double>();
        
        for (int i = 0; i < sortedLines.Count - 1; i++)
        {
            gaps.Add(sortedLines[i + 1].From.X - sortedLines[i].From.X);
        }
        
        if (!gaps.Any()) return false;
        
        var avgGap = gaps.Average();
        var variance = gaps.Sum(g => Math.Pow(g - avgGap, 2)) / gaps.Count;
        
        return variance < avgGap * 0.2; // 規則性の閾値
    }
}